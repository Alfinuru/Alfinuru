Option Explicit
'==============================================================================
' modRegistres - lignes des registres
'
'   NouvelleLigne       place le curseur sur la première ligne libre du registre
'                       actif (ou choisi) et attribue un identifiant (DLM-, GAR-...)
'   SupprimerExemples   efface les lignes d'exemple (ligne grisée sous l'en-tête et
'                       lignes dont l'identifiant commence par « EXEMPLE »)
'   ArchiverClotures    copie vers « Archive » (valeurs et formats) les garanties en
'                       mainlevée et les factures soldées, puis libère leurs lignes
'
' Les registres sont des blocs de formules pré-étendues : on n'y supprime jamais de
' ligne (cela réduirait la capacité et décalerait les plages) ; on efface seulement
' les cellules de saisie, les formules restent en place.
'==============================================================================

Private Const RG_FEUILLE As Long = 0
Private Const RG_EXEMPLE As Long = 1
Private Const RG_PREMIERE As Long = 2
Private Const RG_DERNIERE As Long = 3
Private Const RG_DERCOL As Long = 4
Private Const RG_ID As Long = 5
Private Const RG_PREFIXE As Long = 6
Private Const RG_CHIFFRES As Long = 7

Public Sub NouvelleLigne()
    Const NOM As String = "NouvelleLigne"
    Dim regle As String, ws As Worksheet, r As Long, rLibre As Long, nbCol As Long, idCol As Long
    Dim prefixe As String, nouvelID As String

    On Error GoTo Echec
    regle = RegistreCible()
    If Len(regle) = 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(Champ(regle, RG_FEUILLE))
    nbCol = NumCol(Champ(regle, RG_DERCOL))
    idCol = NumCol(Champ(regle, RG_ID))
    prefixe = Champ(regle, RG_PREFIXE)
    For r = CLng(Champ(regle, RG_PREMIERE)) To CLng(Champ(regle, RG_DERNIERE))
        If LigneVide(ws, r, nbCol) Then
            rLibre = r
            Exit For
        End If
    Next r
    If rLibre = 0 Then
        Dlg_Msg "Registre « " & ws.Name & " » complet (" & Champ(regle, RG_DERNIERE) & " lignes) : " & _
                "archiver les lignes clôturées ou étendre les formules.", vbExclamation
        Exit Sub
    End If
    If Len(prefixe) > 0 Then
        nouvelID = ProchainID(ws, regle, prefixe)
        Deproteger ws
        ws.Cells(rLibre, idCol).Value = nouvelID
    End If
    Terminer
    On Error Resume Next
    ws.Activate
    ws.Cells(rLibre, IIf(Len(prefixe) > 0, idCol + 1, idCol)).Select
    On Error GoTo Echec
    Journaliser NOM, "Ligne " & rLibre & " de « " & ws.Name & " »" & IIf(Len(nouvelID) > 0, " : " & nouvelID, "") & "."
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

' Registre de la feuille active, sinon choix de l'utilisateur ("" si annulé).
Private Function RegistreCible() As String
    Dim regle As Variant, liste As String, i As Long, choix As String, t() As String, n As Long
    For Each regle In Registres()
        If ActiveSheet.Name = Champ(regle, RG_FEUILLE) Then
            RegistreCible = regle
            Exit Function
        End If
    Next regle
    ReDim t(1 To Registres().Count)
    For Each regle In Registres()
        n = n + 1
        t(n) = regle
        liste = liste & vbCrLf & n & " = " & Champ(regle, RG_FEUILLE)
    Next regle
    choix = Trim$(Dlg_Saisie("Dans quel registre ajouter une ligne ?" & liste, "1"))
    If Len(choix) = 0 Then Exit Function
    If Not IsNumeric(choix) Then
        Dlg_Msg "Choix non valide : " & choix, vbExclamation
        Exit Function
    End If
    i = CLng(Val(choix))
    If i < 1 Or i > n Then
        Dlg_Msg "Choix non valide : " & choix, vbExclamation
        Exit Function
    End If
    RegistreCible = t(i)
End Function

' Ligne vide : aucune cellule de saisie (sans formule) renseignée.
Private Function LigneVide(ByVal ws As Worksheet, ByVal r As Long, ByVal nbCol As Long) As Boolean
    Dim c As Long
    For c = 1 To nbCol
        If Not ws.Cells(r, c).HasFormula Then
            If Len(Txt(ws.Cells(r, c).Value)) > 0 Then Exit Function
        End If
    Next c
    LigneVide = True
End Function

' Prochain identifiant libre : PREFIXE + (plus grand numéro existant + 1).
Private Function ProchainID(ByVal ws As Worksheet, ByVal regle As String, ByVal prefixe As String) As String
    Dim r As Long, v As String, n As Long, maxi As Long, idCol As Long, nbChiffres As Long
    idCol = NumCol(Champ(regle, RG_ID))
    nbChiffres = CLng(Champ(regle, RG_CHIFFRES))
    For r = CLng(Champ(regle, RG_PREMIERE)) To CLng(Champ(regle, RG_DERNIERE))
        v = UCase$(Txt(ws.Cells(r, idCol).Value))
        If Left$(v, Len(prefixe)) = UCase$(prefixe) Then
            n = CLng(Val(Mid$(v, Len(prefixe) + 1)))
            If n > maxi Then maxi = n
        End If
    Next r
    ProchainID = prefixe & Format$(maxi + 1, String$(nbChiffres, "0"))
End Function

Public Sub SupprimerExemples()
    Const NOM As String = "SupprimerExemples"
    Dim regle As Variant, ws As Worksheet, r As Long, nbCol As Long, idCol As Long
    Dim cibles As New Collection, it As Variant, detail As String, nb As Long

    On Error GoTo Echec
    ' 1. Repérage : ligne d'exemple sous l'en-tête + lignes « EXEMPLE... »
    For Each regle In Registres()
        Set ws = ThisWorkbook.Worksheets(Champ(regle, RG_FEUILLE))
        nbCol = NumCol(Champ(regle, RG_DERCOL))
        idCol = NumCol(Champ(regle, RG_ID))
        r = CLng(Champ(regle, RG_EXEMPLE))
        If Not LigneVide(ws, r, nbCol) Then cibles.Add Array(ws.Name, r, nbCol)
        For r = CLng(Champ(regle, RG_PREMIERE)) To CLng(Champ(regle, RG_DERNIERE))
            If Left$(UCase$(Txt(ws.Cells(r, idCol).Value)), 7) = "EXEMPLE" Then cibles.Add Array(ws.Name, r, nbCol)
        Next r
    Next regle
    If cibles.Count = 0 Then
        Journaliser NOM, "Aucune ligne d'exemple."
        Dlg_Msg "Aucune ligne d'exemple à effacer.", vbInformation
        Exit Sub
    End If
    For Each it In cibles
        detail = detail & vbCrLf & "   · " & it(0) & " ligne " & it(1)
        nb = nb + 1
        If nb >= 15 Then
            detail = detail & vbCrLf & "   · ..."
            Exit For
        End If
    Next it
    If Dlg_Msg(cibles.Count & " ligne(s) d'exemple vont être effacées (les formules sont conservées) :" & detail & _
               vbCrLf & vbCrLf & "Continuer ?", vbYesNo + vbQuestion) <> vbYes Then
        Journaliser NOM, "Annulé par l'utilisateur."
        Exit Sub
    End If

    ' 2. Effacement des seules cellules de saisie
    OptimiserDebut
    For Each it In cibles
        Set ws = ThisWorkbook.Worksheets(CStr(it(0)))
        Deproteger ws
        ViderSaisiesLigne ws, CLng(it(1)), CLng(it(2))
    Next it
    Application.Calculate
    Terminer
    Journaliser NOM, cibles.Count & " ligne(s) d'exemple effacée(s)."
    Dlg_Msg cibles.Count & " ligne(s) d'exemple effacée(s).", vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

Public Sub ArchiverClotures()
    Const NOM As String = "ArchiverClotures"
    Dim wsG As Worksheet, wsF As Worksheet, wsA As Worksheet, r As Long
    Dim selG As New Collection, selF As New Collection, it As Variant
    Dim x As Double, y As Double, rA As Long

    On Error GoTo Echec
    Application.Calculate
    Set wsG = ThisWorkbook.Worksheets(GAR_FEUILLE)
    Set wsF = ThisWorkbook.Worksheets(FAC_FEUILLE)
    Set wsA = ThisWorkbook.Worksheets(FEUILLE_ARCHIVE)

    ' 1. Repérage des lignes clôturées
    For r = 6 To 205
        If Txt(wsG.Cells(r, GAR_COL_STATUT).Value) = GAR_STATUT_CLOS And Not LigneVide(wsG, r, GAR_NB_COL) Then selG.Add r
        If EnNombre(wsF.Cells(r, FAC_COL_RESTE).Value, x) And EnNombre(wsF.Cells(r, FAC_COL_TTC).Value, y) Then
            If x = 0 And y > 0 Then selF.Add r
        End If
    Next r
    If selG.Count + selF.Count = 0 Then
        Journaliser NOM, "Aucune ligne clôturée."
        Dlg_Msg "Aucune ligne clôturée à archiver.", vbInformation
        Exit Sub
    End If
    If Dlg_Msg(selG.Count & " garantie(s) en « " & GAR_STATUT_CLOS & " » et " & selF.Count & _
               " facture(s) soldée(s) vont être déplacées vers la feuille « " & FEUILLE_ARCHIVE & " »." & _
               vbCrLf & vbCrLf & "Continuer ?", vbYesNo + vbQuestion) <> vbYes Then
        Journaliser NOM, "Annulé par l'utilisateur."
        Exit Sub
    End If

    ' 2. Copie (valeurs + formats) puis libération des lignes (formules conservées)
    OptimiserDebut
    Deproteger wsA
    Deproteger wsG
    Deproteger wsF
    rA = ProchaineLigneArchive(wsA, ARCH_COL_GAR)
    For Each it In selG
        CopierVersArchive wsG, CLng(it), GAR_NB_COL, wsA, rA, ARCH_COL_GAR
        rA = rA + 1
    Next it
    rA = ProchaineLigneArchive(wsA, ARCH_COL_FAC)
    For Each it In selF
        CopierVersArchive wsF, CLng(it), FAC_NB_COL, wsA, rA, ARCH_COL_FAC
        rA = rA + 1
    Next it
    For Each it In selG
        ViderSaisiesLigne wsG, CLng(it), GAR_NB_COL
    Next it
    For Each it In selF
        ViderSaisiesLigne wsF, CLng(it), FAC_NB_COL
    Next it
    Application.Calculate
    Terminer
    Journaliser NOM, selG.Count & " garantie(s) et " & selF.Count & " facture(s) archivée(s)."
    Dlg_Msg selG.Count & " garantie(s) et " & selF.Count & " facture(s) archivée(s)." & vbCrLf & _
            "Leurs lignes sont libérées dans les registres (« Nouvelle ligne » les réutilise).", vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

Private Function ProchaineLigneArchive(ByVal wsA As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = DerniereLigne(wsA, col) + 1
    If r < ARCH_PREMIERE Then r = ARCH_PREMIERE
    ProchaineLigneArchive = r
End Function

' Copie une ligne de registre (valeurs figées et formats de nombre) + date d'archivage.
Private Sub CopierVersArchive(ByVal ws As Worksheet, ByVal r As Long, ByVal nbCol As Long, _
                              ByVal wsA As Worksheet, ByVal rA As Long, ByVal colA As Long)
    Dim c As Long, v As Variant
    For c = 1 To nbCol
        v = ws.Cells(r, c).Value
        If VarType(v) = vbString Then
            wsA.Cells(rA, colA + c - 1).Value = Texte(v)
        ElseIf IsError(v) Then
            wsA.Cells(rA, colA + c - 1).Value = Texte(v)
        Else
            wsA.Cells(rA, colA + c - 1).Value = v
        End If
        wsA.Cells(rA, colA + c - 1).NumberFormat = ws.Cells(r, c).NumberFormat
    Next c
    wsA.Cells(rA, colA + nbCol).Value = Now
    wsA.Cells(rA, colA + nbCol).NumberFormat = "dd/mm/yyyy hh:mm"
    With wsA.Range(wsA.Cells(rA, colA), wsA.Cells(rA, colA + nbCol))
        .Font.Name = "Arial"
        .Font.Size = 9
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
    End With
End Sub
