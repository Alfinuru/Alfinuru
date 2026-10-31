Attribute VB_Name = "Mod_Pilotage"
'===============================================================================
' MODULE       : Mod_Pilotage
' CLASSEUR     : Pilotage_Financements_ASER.xlsm
' RÔLE         : Macros opérationnelles de gestion du portefeuille :
'                1. Actualiser
'                2. GenererPlanPaiement (avec persistance des saisies AC/DFP)
'                3. ArchiverClotures
'                4. NouvelleLigne
'                5. SupprimerExemples
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' MACRO 2 : Actualiser
' Force le recalcul complet, met à jour Plan_Paiement et le Tableau_de_bord,
' horodate la cellule Dernière actualisation et journalise l'opération.
'-------------------------------------------------------------------------------
Public Sub Actualiser()
    On Error GoTo ErrHandler

    Dim wsTDB As Worksheet
    Dim chObj As ChartObject

    Call ActiverModeMacro(True)

    ' 1. Recalcul intégral des formules de tous les onglets
    Application.CalculateFull

    ' 2. Reconstruction et tri du Plan de Paiement (en conservant les saisies)
    Call GenererPlanPaiement_Interne(False)

    ' 3. Mise à jour de l'horodatage et rafraîchissement des graphiques du Tableau de bord
    Set wsTDB = ThisWorkbook.Worksheets("Tableau_de_bord")
    Call DeprotegerFeuille(wsTDB)
    wsTDB.Range("DerniereActualisation").Value = Now
    wsTDB.Range("DerniereActualisation").NumberFormat = "dd/mm/yyyy hh:mm"

    For Each chObj In wsTDB.ChartObjects
        chObj.Chart.Refresh
    Next chObj
    Call ProtegerFeuille(wsTDB)

    Application.Calculate
    Call EcrireJournal("Actualiser", "Recalcul complet, mise à jour de Plan_Paiement et du Tableau_de_bord effectués.")

    Call ActiverModeMacro(False)
    MsgBox "Actualisation complète effectuée avec succès (" & Format$(Now, "dd/mm/yyyy hh:mm") & ").", _
           vbInformation, "Pilotage Financements ASER"
    Exit Sub

ErrHandler:
    Dim msgErr As String
    msgErr = Err.Description & " (Code " & Err.Number & ")"
    Call InitialiserProtectionClasseur
    Call ActiverModeMacro(False)
    Call EcrireJournal("Actualiser", "ERREUR : " & msgErr)
    MsgBox "Erreur lors de l'actualisation : " & vbCrLf & msgErr, vbCritical, "Erreur VBA - ASER"
End Sub

'-------------------------------------------------------------------------------
' MACRO 3 : GenererPlanPaiement
' Construit et trie Plan_Paiement (Jours avant DLM croissant, Retard contractuel
' décroissant, Reste à payer décroissant) en conservant les saisies existantes
' par rapprochement sur le N° facture.
'-------------------------------------------------------------------------------
Public Sub GenererPlanPaiement()
    On Error GoTo ErrHandler

    Call ActiverModeMacro(True)
    Call GenererPlanPaiement_Interne(True)
    Call ActiverModeMacro(False)

    MsgBox "Le Plan de Paiement priorisé a été régénéré et trié avec conservation des arbitrages AC/DFP.", _
           vbInformation, "Plan de Paiement - ASER"
    Exit Sub

ErrHandler:
    Dim msgErr As String
    msgErr = Err.Description & " (Code " & Err.Number & ")"
    Call InitialiserProtectionClasseur
    Call ActiverModeMacro(False)
    Call EcrireJournal("GenererPlanPaiement", "ERREUR : " & msgErr)
    MsgBox "Erreur lors de la génération du Plan de Paiement : " & vbCrLf & msgErr, vbCritical, "Erreur VBA - ASER"
End Sub

'-------------------------------------------------------------------------------
' Procédure interne appelée par GenererPlanPaiement, Actualiser, RapportMensuel
' et ArchiverClotures. En cas d'erreur, la feuille est reprotégée et l'erreur
' est remontée à la macro appelante.
'-------------------------------------------------------------------------------
Public Sub GenererPlanPaiement_Interne(Optional ByVal journaliserSeul As Boolean = True)
    Dim wsFac As Worksheet
    Dim wsPP As Worksheet
    Dim loFac As ListObject
    Dim loPP As ListObject
    Dim dictSaisies As Object
    Dim i As Long, j As Long, k As Long, nFactures As Long
    Dim cleFacture As String
    Dim arrSaisie(1 To 4) As Variant
    Dim vRecup As Variant
    Dim nErr As Long, dErr As String, sErr As String

    Set wsFac = ThisWorkbook.Worksheets("Factures_Decomptes")
    Set wsPP = ThisWorkbook.Worksheets("Plan_Paiement")
    Set loFac = wsFac.ListObjects("T_Factures")
    Set loPP = wsPP.ListObjects("T_PlanPaiement")

    On Error GoTo ErrInterne
    Application.Calculate
    Call DeprotegerFeuille(wsPP)

    ' 1. Mémoriser les saisies existantes de T_PlanPaiement dans un Dictionary (liaison tardive)
    Set dictSaisies = CreateObject("Scripting.Dictionary")
    dictSaisies.CompareMode = 1 ' vbTextCompare

    If Not loPP.DataBodyRange Is Nothing Then
        Dim idxPPNum As Long, idxPPDate As Long, idxPPAC As Long, idxPPDFP As Long, idxPPComm As Long
        idxPPNum = loPP.ListColumns("N° facture").Index
        idxPPDate = loPP.ListColumns("Date de paiement proposée").Index
        idxPPAC = loPP.ListColumns("Validation AC (O/N)").Index
        idxPPDFP = loPP.ListColumns("Validation DFP (O/N)").Index
        idxPPComm = loPP.ListColumns("Commentaire").Index

        For i = 1 To loPP.ListRows.Count
            cleFacture = TexteCellule(loPP.DataBodyRange.Cells(i, idxPPNum).Value)
            If Len(cleFacture) > 0 Then
                arrSaisie(1) = loPP.DataBodyRange.Cells(i, idxPPDate).Value
                arrSaisie(2) = loPP.DataBodyRange.Cells(i, idxPPAC).Value
                arrSaisie(3) = loPP.DataBodyRange.Cells(i, idxPPDFP).Value
                arrSaisie(4) = loPP.DataBodyRange.Cells(i, idxPPComm).Value
                dictSaisies(cleFacture) = arrSaisie
            End If
        Next i

        ' Vider les lignes existantes de T_PlanPaiement
        For i = loPP.ListRows.Count To 1 Step -1
            loPP.ListRows(i).Delete
        Next i
    End If

    ' 2. Collecter toutes les factures non soldées (Étape en cours <> "Payée" et Reste à payer > 0)
    nFactures = 0
    If Not loFac.DataBodyRange Is Nothing Then
        Dim idxFNum As Long, idxFDLM As Long
        Dim idxFJoursDLM As Long, idxFReste As Long, idxFRetard As Long, idxFEtape As Long

        idxFNum = loFac.ListColumns("N° facture/décompte").Index
        idxFDLM = loFac.ListColumns("DLM liée").Index
        idxFJoursDLM = loFac.ListColumns("Jours avant DLM").Index
        idxFReste = loFac.ListColumns("Reste à payer").Index
        idxFRetard = loFac.ListColumns("Retard contractuel (j)").Index
        idxFEtape = loFac.ListColumns("Étape en cours").Index

        Dim candidats() As Variant
        Dim dictVus As Object
        ReDim candidats(1 To 3, 1 To loFac.ListRows.Count)
        Dim numCand() As String
        Dim valCand() As Variant
        ReDim numCand(1 To loFac.ListRows.Count)
        ReDim valCand(1 To loFac.ListRows.Count)
        Set dictVus = CreateObject("Scripting.Dictionary")
        dictVus.CompareMode = 1

        For i = 1 To loFac.ListRows.Count
            Dim etapeStr As String
            Dim resteVal As Double
            Dim numStr As String

            numStr = TexteCellule(loFac.DataBodyRange.Cells(i, idxFNum).Value)
            etapeStr = TexteCellule(loFac.DataBodyRange.Cells(i, idxFEtape).Value)
            resteVal = NombreCellule(loFac.DataBodyRange.Cells(i, idxFReste).Value, 0)

            ' Un N° facture en double ne figure qu'une fois dans le plan (rapprochement INDEX/EQUIV)
            If Len(numStr) > 0 And StrComp(etapeStr, "Payée", vbTextCompare) <> 0 And resteVal > 0 _
               And Not dictVus.Exists(numStr) Then
                dictVus(numStr) = True
                nFactures = nFactures + 1
                numCand(nFactures) = numStr
                valCand(nFactures) = loFac.DataBodyRange.Cells(i, idxFNum).Value
                candidats(1, nFactures) = NombreCellule(loFac.DataBodyRange.Cells(i, idxFJoursDLM).Value, 9999)
                candidats(2, nFactures) = NombreCellule(loFac.DataBodyRange.Cells(i, idxFRetard).Value, 0)
                candidats(3, nFactures) = resteVal
            End If
        Next i

        ' 3. Tri multi-critères en mémoire (compatible toutes versions Excel < 2021) :
        '    Clé 1 : Jours avant DLM (croissant)
        '    Clé 2 : Retard contractuel (décroissant)
        '    Clé 3 : Reste à payer (décroissant)
        If nFactures > 1 Then
            Dim doitEchanger As Boolean
            Dim tmpVal As Variant
            Dim tmpNum As String
            Dim tmpV As Variant

            For i = 1 To nFactures - 1
                For j = i + 1 To nFactures
                    doitEchanger = False
                    If candidats(1, j) < candidats(1, i) Then
                        doitEchanger = True
                    ElseIf candidats(1, j) = candidats(1, i) Then
                        If candidats(2, j) > candidats(2, i) Then
                            doitEchanger = True
                        ElseIf candidats(2, j) = candidats(2, i) Then
                            If candidats(3, j) > candidats(3, i) Then doitEchanger = True
                        End If
                    End If

                    If doitEchanger Then
                        For k = 1 To 3
                            tmpVal = candidats(k, i)
                            candidats(k, i) = candidats(k, j)
                            candidats(k, j) = tmpVal
                        Next k
                        tmpNum = numCand(i)
                        numCand(i) = numCand(j)
                        numCand(j) = tmpNum
                        tmpV = valCand(i)
                        valCand(i) = valCand(j)
                        valCand(j) = tmpV
                    End If
                Next j
            Next i
        End If

        ' 4. Insérer les lignes triées dans T_PlanPaiement et restaurer les saisies
        For i = 1 To nFactures
            Dim lrNew As ListRow
            Set lrNew = AjouterLigneTableau(loPP)
            cleFacture = numCand(i)

            With lrNew.Range
                .Cells(1, loPP.ListColumns("Rang").Index).Value = i
                ' N° recopié avec son type d'origine (nécessaire au rapprochement EQUIV) ;
                ' un texte commençant par "=" est neutralisé (injection de formule)
                If VarType(valCand(i)) = vbString And Left$(CStr(valCand(i)), 1) = "=" Then
                    .Cells(1, loPP.ListColumns("N° facture").Index).Value = "'" & CStr(valCand(i))
                Else
                    .Cells(1, loPP.ListColumns("N° facture").Index).Value = valCand(i)
                End If
                .Cells(1, loPP.ListColumns("Entreprise").Index).Formula = _
                    "=INDEX(T_Factures[Entreprise],MATCH(T_PlanPaiement[[#This Row],[N° facture]],T_Factures[N° facture/décompte],0))"
                .Cells(1, loPP.ListColumns("Bailleur").Index).Formula = _
                    "=INDEX(T_Factures[Bailleur],MATCH(T_PlanPaiement[[#This Row],[N° facture]],T_Factures[N° facture/décompte],0))"
                .Cells(1, loPP.ListColumns("DLM liée").Index).Formula = _
                    "=INDEX(T_Factures[DLM liée],MATCH(T_PlanPaiement[[#This Row],[N° facture]],T_Factures[N° facture/décompte],0))"
                .Cells(1, loPP.ListColumns("Jours avant DLM").Index).Formula = _
                    "=INDEX(T_Factures[Jours avant DLM],MATCH(T_PlanPaiement[[#This Row],[N° facture]],T_Factures[N° facture/décompte],0))"
                .Cells(1, loPP.ListColumns("Reste à payer").Index).Formula = _
                    "=INDEX(T_Factures[Reste à payer],MATCH(T_PlanPaiement[[#This Row],[N° facture]],T_Factures[N° facture/décompte],0))"
                ' Cumul indépendant de la position du tableau dans la feuille
                .Cells(1, loPP.ListColumns("Cumul").Index).Formula = _
                    "=SUM(INDEX(T_PlanPaiement[Reste à payer],1):T_PlanPaiement[[#This Row],[Reste à payer]])"

                ' Restaurer les saisies d'arbitrage si le N° facture existait déjà
                If dictSaisies.Exists(cleFacture) Then
                    vRecup = dictSaisies(cleFacture)
                    .Cells(1, loPP.ListColumns("Date de paiement proposée").Index).Value = vRecup(1)
                    .Cells(1, loPP.ListColumns("Validation AC (O/N)").Index).Value = vRecup(2)
                    .Cells(1, loPP.ListColumns("Validation DFP (O/N)").Index).Value = vRecup(3)
                    .Cells(1, loPP.ListColumns("Commentaire").Index).Value = vRecup(4)
                End If
                .Cells(1, loPP.ListColumns("Date de paiement proposée").Index).NumberFormat = "dd/mm/yyyy"

                ' Verrouiller les colonnes calculées, déverrouiller les 4 colonnes de saisie
                .Locked = True
                .Cells(1, loPP.ListColumns("Date de paiement proposée").Index).Locked = False
                .Cells(1, loPP.ListColumns("Validation AC (O/N)").Index).Locked = False
                .Cells(1, loPP.ListColumns("Validation DFP (O/N)").Index).Locked = False
                .Cells(1, loPP.ListColumns("Commentaire").Index).Locked = False
            End With
        Next i
    End If

    Call ProtegerFeuille(wsPP)

    If journaliserSeul Then
        Call EcrireJournal("GenererPlanPaiement", _
            "Plan de paiement construit et trié (" & nFactures & " facture(s) non soldée(s), saisies AC/DFP conservées).")
    End If
    Exit Sub

ErrInterne:
    nErr = Err.Number: dErr = Err.Description: sErr = Err.Source
    Call ProtegerFeuille(wsPP)
    Err.Raise nErr, sErr, dErr
End Sub

'-------------------------------------------------------------------------------
' MACRO 7 : ArchiverClotures
' Déplace les garanties au statut "Mainlevée" et les factures au statut "Payée"
' vers l'onglet Archive après confirmation de l'utilisateur.
' Les colonnes sont rapprochées par leur intitulé (et non par leur position).
'-------------------------------------------------------------------------------
Public Sub ArchiverClotures()
    On Error GoTo ErrHandler

    Dim wsGar As Worksheet, wsFac As Worksheet, wsArc As Worksheet
    Dim loGar As ListObject, loFac As ListObject
    Dim loArcGar As ListObject, loArcFac As ListObject
    Dim dtRef As Date
    Dim i As Long
    Dim nbGarClot As Long, nbFacClot As Long
    Dim idxStatGar As Long, idxEtapeFac As Long
    Dim rep As VbMsgBoxResult

    Set wsGar = ThisWorkbook.Worksheets("Garanties")
    Set wsFac = ThisWorkbook.Worksheets("Factures_Decomptes")
    Set wsArc = ThisWorkbook.Worksheets("Archive")

    Set loGar = wsGar.ListObjects("T_Garanties")
    Set loFac = wsFac.ListObjects("T_Factures")
    Set loArcGar = wsArc.ListObjects("T_Archive_Garanties")
    Set loArcFac = wsArc.ListObjects("T_Archive_Factures")

    idxStatGar = loGar.ListColumns("Statut").Index
    idxEtapeFac = loFac.ListColumns("Étape en cours").Index

    Application.Calculate

    ' 1. Compter les lignes éligibles à l'archivage
    If Not loGar.DataBodyRange Is Nothing Then
        For i = 1 To loGar.ListRows.Count
            If StrComp(TexteCellule(loGar.DataBodyRange.Cells(i, idxStatGar).Value), "Mainlevée", vbTextCompare) = 0 Then
                nbGarClot = nbGarClot + 1
            End If
        Next i
    End If

    If Not loFac.DataBodyRange Is Nothing Then
        For i = 1 To loFac.ListRows.Count
            If StrComp(TexteCellule(loFac.DataBodyRange.Cells(i, idxEtapeFac).Value), "Payée", vbTextCompare) = 0 Then
                nbFacClot = nbFacClot + 1
            End If
        Next i
    End If

    If (nbGarClot + nbFacClot) = 0 Then
        MsgBox "Aucune ligne clôturée à archiver (0 garantie en « Mainlevée », 0 facture « Payée »).", _
               vbInformation, "Archivage - ASER"
        Exit Sub
    End If

    ' 2. Demander confirmation à l'utilisateur
    rep = MsgBox("Confirmez-vous l'archivage des éléments clôturés suivants ?" & vbCrLf & vbCrLf & _
                 "• Garanties au statut « Mainlevée » : " & nbGarClot & vbCrLf & _
                 "• Factures/décomptes au statut « Payée » : " & nbFacClot & vbCrLf & vbCrLf & _
                 "Ces lignes seront déplacées vers l'onglet Archive avec la date d'archivage.", _
                 vbQuestion + vbYesNo, "Confirmation d'archivage - ASER")
    If rep <> vbYes Then Exit Sub

    Call ActiverModeMacro(True)
    dtRef = ObtenirDateRef()

    Call DeprotegerFeuille(wsGar)
    Call DeprotegerFeuille(wsFac)
    Call DeprotegerFeuille(wsArc)

    ' 3. Déplacer les garanties en Mainlevée (du bas vers le haut)
    If Not loGar.DataBodyRange Is Nothing Then
        For i = loGar.ListRows.Count To 1 Step -1
            If StrComp(TexteCellule(loGar.DataBodyRange.Cells(i, idxStatGar).Value), "Mainlevée", vbTextCompare) = 0 Then
                Call CopierVersArchive(loGar, i, loArcGar, dtRef)
                loGar.ListRows(i).Delete
            End If
        Next i
    End If

    ' 4. Déplacer les factures Payées (du bas vers le haut)
    If Not loFac.DataBodyRange Is Nothing Then
        For i = loFac.ListRows.Count To 1 Step -1
            If StrComp(TexteCellule(loFac.DataBodyRange.Cells(i, idxEtapeFac).Value), "Payée", vbTextCompare) = 0 Then
                Call CopierVersArchive(loFac, i, loArcFac, dtRef)
                loFac.ListRows(i).Delete
            End If
        Next i
    End If

    Call ProtegerFeuille(wsGar)
    Call ProtegerFeuille(wsFac)
    Call ProtegerFeuille(wsArc)

    ' 5. Mettre à jour le Plan de paiement
    Call GenererPlanPaiement_Interne(False)

    Call EcrireJournal("ArchiverClotures", _
        "Archivage effectué : " & nbGarClot & " garantie(s) en mainlevée et " & nbFacClot & " facture(s) soldée(s).")

    Call ActiverModeMacro(False)
    MsgBox "Archivage terminé : " & nbGarClot & " garantie(s) et " & nbFacClot & " facture(s) déplacées dans l'onglet Archive.", _
           vbInformation, "Archivage - ASER"
    Exit Sub

ErrHandler:
    Dim msgErr As String
    msgErr = Err.Description & " (Code " & Err.Number & ")"
    Call InitialiserProtectionClasseur
    Call ActiverModeMacro(False)
    Call EcrireJournal("ArchiverClotures", "ERREUR : " & msgErr)
    MsgBox "Erreur lors de l'archivage : " & vbCrLf & msgErr, vbCritical, "Erreur VBA - ASER"
End Sub

'-------------------------------------------------------------------------------
' Copie une ligne d'un registre vers sa table d'archive en rapprochant les
' colonnes par intitulé, puis renseigne la Date d'archivage.
'-------------------------------------------------------------------------------
Private Sub CopierVersArchive(ByVal loSrc As ListObject, ByVal numLigne As Long, _
                              ByVal loArc As ListObject, ByVal dtRef As Date)
    Dim lrA As ListRow
    Dim c As Long
    Dim idxSrc As Long
    Dim nomCol As String
    Dim celA As Range

    Set lrA = AjouterLigneTableau(loArc)
    For c = 1 To loArc.ListColumns.Count
        nomCol = loArc.ListColumns(c).Name
        Set celA = lrA.Range.Cells(1, c)
        If StrComp(nomCol, "Date d'archivage", vbTextCompare) = 0 Then
            celA.Value = dtRef
            celA.NumberFormat = "dd/mm/yyyy"
        Else
            idxSrc = IndexColonne(loSrc, nomCol)
            If idxSrc > 0 Then
                celA.Value = loSrc.DataBodyRange.Cells(numLigne, idxSrc).Value
                celA.NumberFormat = loSrc.DataBodyRange.Cells(numLigne, idxSrc).NumberFormat
            End If
        End If
    Next c
End Sub

'-------------------------------------------------------------------------------
' Retourne l'index d'une colonne de tableau par son intitulé (0 si absente)
'-------------------------------------------------------------------------------
Private Function IndexColonne(ByVal lo As ListObject, ByVal nomCol As String) As Long
    Dim c As Long
    IndexColonne = 0
    For c = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(c).Name, nomCol, vbTextCompare) = 0 Then
            IndexColonne = c
            Exit Function
        End If
    Next c
End Function

'-------------------------------------------------------------------------------
' MACRO 8 : NouvelleLigne
' Ajoute une ligne au registre actif (DLM, Garanties, Factures, Plan) avec un
' ID auto-incrémenté et les formules calculées. Depuis tout autre onglet, la
' macro demande le registre cible (les tableaux techniques ne sont jamais visés).
'-------------------------------------------------------------------------------
Public Sub NouvelleLigne()
    On Error GoTo ErrHandler

    Dim wsCible As Worksheet
    Dim loCible As ListObject
    Dim prefixeID As String
    Dim choix As String
    Dim lrNew As ListRow
    Dim nouvelID As String
    Dim numMax As Long
    Dim i As Long, c As Long
    Dim valID As String

    Set loCible = Nothing
    If ActiveWorkbook Is ThisWorkbook Then
        On Error Resume Next
        Set loCible = ActiveCell.ListObject
        On Error GoTo ErrHandler
        If Not loCible Is Nothing Then
            If Len(PrefixeRegistre(loCible.Name)) = 0 Then Set loCible = Nothing
        End If
        If loCible Is Nothing Then
            Select Case ActiveSheet.Name
                Case "DLM_Prorogations": Set loCible = ThisWorkbook.Worksheets("DLM_Prorogations").ListObjects("T_DLM")
                Case "Garanties": Set loCible = ThisWorkbook.Worksheets("Garanties").ListObjects("T_Garanties")
                Case "Factures_Decomptes": Set loCible = ThisWorkbook.Worksheets("Factures_Decomptes").ListObjects("T_Factures")
                Case "Plan_Amelioration": Set loCible = ThisWorkbook.Worksheets("Plan_Amelioration").ListObjects("T_Plan")
            End Select
        End If
    End If

    ' Hors d'un registre de saisie : demander le registre cible
    If loCible Is Nothing Then
        choix = InputBox("Dans quel registre souhaitez-vous ajouter une nouvelle ligne ?" & vbCrLf & vbCrLf & _
                         "1 = DLM_Prorogations (T_DLM)" & vbCrLf & _
                         "2 = Garanties (T_Garanties)" & vbCrLf & _
                         "3 = Factures_Decomptes (T_Factures)" & vbCrLf & _
                         "4 = Plan_Amelioration (T_Plan)", _
                         "Ajout d'une nouvelle ligne - ASER", "1")
        Select Case Trim$(choix)
            Case "1": Set loCible = ThisWorkbook.Worksheets("DLM_Prorogations").ListObjects("T_DLM")
            Case "2": Set loCible = ThisWorkbook.Worksheets("Garanties").ListObjects("T_Garanties")
            Case "3": Set loCible = ThisWorkbook.Worksheets("Factures_Decomptes").ListObjects("T_Factures")
            Case "4": Set loCible = ThisWorkbook.Worksheets("Plan_Amelioration").ListObjects("T_Plan")
            Case Else: Exit Sub
        End Select
    End If

    Set wsCible = loCible.Parent
    prefixeID = PrefixeRegistre(loCible.Name)

    Call ActiverModeMacro(True)
    Call DeprotegerFeuille(wsCible)

    ' Calculer le prochain numéro d'ID auto-incrémenté
    numMax = 0
    If Not loCible.DataBodyRange Is Nothing Then
        For i = 1 To loCible.ListRows.Count
            valID = TexteCellule(loCible.DataBodyRange.Cells(i, 1).Value)
            Dim posTiret As Long
            posTiret = InStrRev(valID, "-")
            If posTiret > 0 Then
                Dim suffixe As String
                suffixe = Trim$(Mid$(valID, posTiret + 1))
                If Len(suffixe) > 0 And Len(suffixe) < 9 And IsNumeric(suffixe) Then
                    If CLng(Val(suffixe)) > numMax Then numMax = CLng(Val(suffixe))
                End If
            End If
        Next i
    End If
    nouvelID = prefixeID & Format$(numMax + 1, "000")

    Set lrNew = AjouterLigneTableau(loCible)
    lrNew.Range.Cells(1, 1).Value = nouvelID

    ' Colonnes de saisie déverrouillées (les colonnes calculées restent verrouillées)
    For c = 2 To loCible.ListColumns.Count
        If Not lrNew.Range.Cells(1, c).HasFormula Then lrNew.Range.Cells(1, c).Locked = False
    Next c
    lrNew.Range.Cells(1, 1).Locked = True

    ' Initialisations par défaut selon le registre
    If loCible.Name = "T_Factures" Then
        lrNew.Range.Cells(1, loCible.ListColumns("Délai contractuel de paiement").Index).Formula = "=DelaiPaiementDefaut"
        lrNew.Range.Cells(1, loCible.ListColumns("Délai contractuel de paiement").Index).Locked = False
        lrNew.Range.Cells(1, loCible.ListColumns("Montant payé").Index).Value = 0
    ElseIf loCible.Name = "T_Plan" Then
        lrNew.Range.Cells(1, loCible.ListColumns("Statut").Index).Value = "Non démarré"
        lrNew.Range.Cells(1, loCible.ListColumns("% avancement").Index).Value = 0
    ElseIf loCible.Name = "T_DLM" Then
        lrNew.Range.Cells(1, loCible.ListColumns("Nb prorogations obtenues").Index).Value = 0
        lrNew.Range.Cells(1, loCible.ListColumns("Nb relances").Index).Value = 0
    End If

    Call ProtegerFeuille(wsCible)
    Call EcrireJournal("NouvelleLigne", "Ajout de la ligne " & nouvelID & " dans " & loCible.Name & " (" & wsCible.Name & ").")
    Call ActiverModeMacro(False)

    wsCible.Activate
    lrNew.Range.Cells(1, 2).Select

    MsgBox "Nouvelle ligne ajoutée dans « " & wsCible.Name & " » avec l'identifiant : " & nouvelID, _
           vbInformation, "Nouvelle Ligne - ASER"
    Exit Sub

ErrHandler:
    Dim msgErr As String
    msgErr = Err.Description & " (Code " & Err.Number & ")"
    Call InitialiserProtectionClasseur
    Call ActiverModeMacro(False)
    Call EcrireJournal("NouvelleLigne", "ERREUR : " & msgErr)
    MsgBox "Erreur lors de l'ajout d'une nouvelle ligne : " & vbCrLf & msgErr, vbCritical, "Erreur VBA - ASER"
End Sub

'-------------------------------------------------------------------------------
' Préfixe d'ID des registres de saisie ("" pour les tableaux techniques)
'-------------------------------------------------------------------------------
Private Function PrefixeRegistre(ByVal nomTable As String) As String
    Select Case nomTable
        Case "T_DLM": PrefixeRegistre = "DLM-"
        Case "T_Garanties": PrefixeRegistre = "GAR-"
        Case "T_Factures": PrefixeRegistre = "FAC-"
        Case "T_Plan": PrefixeRegistre = "PA-"
        Case Else: PrefixeRegistre = ""
    End Select
End Function

'-------------------------------------------------------------------------------
' MACRO 9 : SupprimerExemples
' Supprime toutes les lignes préfixées par « EXEMPLE - » dans tous les registres
' du classeur après confirmation de l'utilisateur.
'-------------------------------------------------------------------------------
Public Sub SupprimerExemples()
    On Error GoTo ErrHandler

    Dim rep As VbMsgBoxResult
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long, c As Long
    Dim nbSuppr As Long
    Dim estExemple As Boolean
    Dim valCell As String

    rep = MsgBox("Voulez-vous supprimer définitivement toutes les lignes préfixées « EXEMPLE - » dans l'ensemble des onglets ?" & vbCrLf & vbCrLf & _
                 "Cette action prépare le classeur pour la mise en production.", _
                 vbQuestion + vbYesNo + vbDefaultButton2, "Suppression des exemples - ASER")
    If rep <> vbYes Then Exit Sub

    Call ActiverModeMacro(True)
    nbSuppr = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.ListObjects.Count > 0 Then
            Call DeprotegerFeuille(ws)
            For Each lo In ws.ListObjects
                If Not lo.DataBodyRange Is Nothing Then
                    For i = lo.ListRows.Count To 1 Step -1
                        estExemple = False
                        For c = 1 To lo.ListColumns.Count
                            valCell = TexteCellule(lo.DataBodyRange.Cells(i, c).Value)
                            If Left$(UCase$(valCell), 9) = "EXEMPLE -" Then
                                estExemple = True
                                Exit For
                            End If
                        Next c

                        If estExemple Then
                            lo.ListRows(i).Delete
                            nbSuppr = nbSuppr + 1
                        End If
                    Next i
                End If
            Next lo
            Call ProtegerFeuille(ws)
        End If
    Next ws

    Application.CalculateFull
    Call EcrireJournal("SupprimerExemples", nbSuppr & " ligne(s) d'exemple (« EXEMPLE - ») supprimée(s).")
    Call ActiverModeMacro(False)

    MsgBox nbSuppr & " ligne(s) « EXEMPLE - » ont été supprimées avec succès.", _
           vbInformation, "Suppression des exemples - ASER"
    Exit Sub

ErrHandler:
    Dim msgErr As String
    msgErr = Err.Description & " (Code " & Err.Number & ")"
    Call InitialiserProtectionClasseur
    Call ActiverModeMacro(False)
    Call EcrireJournal("SupprimerExemples", "ERREUR : " & msgErr)
    MsgBox "Erreur lors de la suppression des exemples : " & vbCrLf & msgErr, vbCritical, "Erreur VBA - ASER"
End Sub
