Option Explicit
'==============================================================================
' modMoteur - moteur générique de pilotage (alertes, contrôle qualité, filtres,
' date de référence, instantané de revue, export PDF, notification d'ouverture)
'
'   ActualiserTout          recalcul + tableau de bord + alertes + contrôle qualité
'   SyntheseAlertes         feuille « Alertes » : toutes les lignes en alerte
'   ControleQualite         feuille « Contrôle_Qualité » : doublons, formules
'                           écrasées, dates, champs obligatoires, incohérences
'   FiltrerAlertes          n'affiche dans chaque registre que les lignes en alerte
'                           (toutes les règles de la feuille, pas seulement la 1re)
'   AfficherTout            réaffiche toutes les lignes et retire les filtres
'   FigerDateReference      fige la date de référence pour une revue
'   LibererDateReference    rétablit la date automatique (aujourd'hui)
'   ArchiverRevue           instantané figé (valeurs) du plan d'action
'   ExporterRevuePDF        PDF : tableau de bord + plan d'action + alertes
'
' La structure du classeur (feuilles, colonnes, jetons) est décrite dans modConfig.
'==============================================================================

Private Const NB_COL_ALERTES As Long = 9
Private Const NB_COL_QC As Long = 7

' Champs d'une règle d'alerte (séparateur « | »)
Private Const RA_FEUILLE As Long = 0
Private Const RA_ENTETE As Long = 1
Private Const RA_PREMIERE As Long = 2
Private Const RA_DERNIERE As Long = 3
Private Const RA_ID As Long = 4
Private Const RA_LIBELLE As Long = 5
Private Const RA_ALERTE As Long = 6
Private Const RA_MOTIF As Long = 7
Private Const RA_CRITIQUE As Long = 8
Private Const RA_ATTENTION As Long = 9
Private Const RA_RESP As Long = 10
Private Const RA_JOURS As Long = 11

' Champs d'une règle de contrôle qualité
Private Const RC_FEUILLE As Long = 0
Private Const RC_ENTETE As Long = 1
Private Const RC_PREMIERE As Long = 2
Private Const RC_DERNIERE As Long = 3
Private Const RC_ID As Long = 4
Private Const RC_CALCUL As Long = 5
Private Const RC_DATES As Long = 6
Private Const RC_OBLIG As Long = 7
Private Const RC_TEMOINS As Long = 8

Private mQC As Collection        ' anomalies accumulées par ControleQualite

'==============================================================================
' 1. COMMANDES (boutons du Sommaire, Alt+F8)
'==============================================================================

Public Sub ActualiserTout()
    Const NOM As String = "ActualiserTout"
    Dim t As Variant, nQC As Long
    On Error GoTo Echec
    OptimiserDebut
    Recalculer
    MettreAJourTableauDeBord
    nQC = ConstruireControle(True)
    t = ConstruireAlertes(True)
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(FEUILLE_ALERTES).Activate
    Terminer
    Journaliser NOM, "Date de référence " & DateReferenceTexte() & " ; " & t(0) & " alerte(s) critique(s), " & _
                     t(1) & " à surveiller ; " & nQC & " anomalie(s) de saisie."
    Dlg_Msg "Classeur recalculé - date de référence : " & DateReferenceTexte() & vbCrLf & vbCrLf & _
            t(0) & " alerte(s) critique(s), " & t(1) & " à surveiller (feuille « " & FEUILLE_ALERTES & " »)." & vbCrLf & _
            nQC & " anomalie(s) de saisie (feuille « " & FEUILLE_QC & " »).", vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

Public Sub SyntheseAlertes()
    Const NOM As String = "SyntheseAlertes"
    Dim t As Variant
    On Error GoTo Echec
    OptimiserDebut
    Application.Calculate
    t = ConstruireAlertes(False)
    Terminer
    Journaliser NOM, t(0) & " alerte(s) critique(s), " & t(1) & " à surveiller."
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

Public Sub ControleQualite()
    Const NOM As String = "ControleQualite"
    Dim n As Long
    On Error GoTo Echec
    OptimiserDebut
    Application.Calculate
    n = ConstruireControle(False)
    Terminer
    Journaliser NOM, n & " anomalie(s) détectée(s)."
    Dlg_Msg n & " anomalie(s) détectée(s)." & vbCrLf & "Détail et liens : feuille « " & FEUILLE_QC & " ».", _
            IIf(n = 0, vbInformation, vbExclamation)
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

' N'affiche, dans chaque registre, que les lignes en alerte selon TOUTES ses règles.
Public Sub FiltrerAlertes()
    Const NOM As String = "FiltrerAlertes"
    Dim feuilles As New Collection, f As Variant, regle As Variant, ws As Worksheet
    Dim r As Long, r1 As Long, r2 As Long, a As Long, b As Long, cAl As Long, n As Long
    Dim garder() As Boolean, msg As String, dRef As Date, total As Long

    On Error GoTo Echec
    dRef = DateReference()
    For Each regle In ReglesAlerte()
        AjouterCle feuilles, Champ(regle, RA_FEUILLE)
    Next regle
    OptimiserDebut
    Application.Calculate
    For Each f In feuilles
        If FeuilleExiste(CStr(f)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(f))
            BornesFeuille CStr(f), r1, r2
            ReDim garder(r1 To r2)
            n = 0
            For Each regle In ReglesAlerte()
                If Champ(regle, RA_FEUILLE) = CStr(f) Then
                    cAl = NumCol(Champ(regle, RA_ALERTE))
                    For r = CLng(Champ(regle, RA_PREMIERE)) To CLng(Champ(regle, RA_DERNIERE))
                        If Not garder(r) Then
                            If Correspond(ws.Cells(r, cAl).Value, Champ(regle, RA_CRITIQUE), dRef) Or _
                               Correspond(ws.Cells(r, cAl).Value, Champ(regle, RA_ATTENTION), dRef) Then
                                garder(r) = True
                                n = n + 1
                            End If
                        End If
                    Next r
                End If
            Next regle
            If n > 0 Then
                Deproteger ws
                If EstFiltree(ws) Then ws.ShowAllData
                ws.Rows(r1 & ":" & r2).Hidden = False
                r = r1
                Do While r <= r2                         ' masquage par blocs contigus
                    If Not garder(r) Then
                        a = r
                        Do While r <= r2
                            If garder(r) Then Exit Do
                            r = r + 1
                        Loop
                        b = r - 1
                        ws.Rows(a & ":" & b).Hidden = True
                    Else
                        r = r + 1
                    End If
                Loop
                msg = msg & vbCrLf & "   · " & ws.Name & " : " & n & " ligne(s)"
                total = total + n
            End If
        End If
    Next f
    Terminer
    Journaliser NOM, total & " ligne(s) en alerte affichée(s)." & Replace(msg, vbCrLf, " ")
    If Len(msg) = 0 Then
        Dlg_Msg "Aucune ligne en alerte : aucun registre filtré.", vbInformation
    Else
        Dlg_Msg "Seules les lignes en alerte restent affichées :" & msg & vbCrLf & vbCrLf & _
                "Bouton « Retirer les filtres » pour tout réafficher.", vbInformation
    End If
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

' Réaffiche les lignes masquées des registres et retire les filtres automatiques.
Public Sub AfficherTout()
    Const NOM As String = "AfficherTout"
    Dim ws As Worksheet, n As Long, r1 As Long, r2 As Long, modifie As Boolean
    On Error GoTo Echec
    OptimiserDebut
    For Each ws In ThisWorkbook.Worksheets
        modifie = False
        If EstFiltree(ws) Then
            Deproteger ws
            ws.ShowAllData
            modifie = True
        End If
        BornesFeuille ws.Name, r1, r2
        If r1 > 0 Then
            If MasqueesDans(ws, r1, r2) Then
                Deproteger ws
                ws.Rows(r1 & ":" & r2).Hidden = False
                modifie = True
            End If
        End If
        If modifie Then n = n + 1
    Next ws
    Terminer
    Journaliser NOM, "Lignes réaffichées sur " & n & " feuille(s)."
    Dlg_Msg "Toutes les lignes sont réaffichées (" & n & " feuille(s) concernée(s)).", vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

Public Sub FigerDateReference()
    Const NOM As String = "FigerDateReference"
    Dim s As String, d As Date
    On Error GoTo Echec
    s = Dlg_Saisie("Date de référence à figer pour la revue (jj/mm/aaaa) :" & vbCrLf & _
                   "Jours restants, alertes et priorités seront calculés à cette date.", _
                   Format$(DateReference(), "dd/mm/yyyy"))
    If Len(Trim$(s)) = 0 Then Exit Sub
    If Not LireDate(s, d) Then
        Dlg_Msg "Date non reconnue : " & s & vbCrLf & "Format attendu : jj/mm/aaaa.", vbExclamation
        Exit Sub
    End If
    OptimiserDebut
    EcrireDateReference d
    Recalculer
    Terminer
    Journaliser NOM, "Date de référence figée au " & Format$(d, "dd/mm/yyyy") & "."
    Dlg_Msg "Date de référence figée au " & Format$(d, "dd/mm/yyyy") & "." & vbCrLf & MSG_LIBERER, vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

Public Sub LibererDateReference()
    Const NOM As String = "LibererDateReference"
    On Error GoTo Echec
    OptimiserDebut
    RetablirDateAuto
    Recalculer
    Terminer
    Journaliser NOM, "Date de référence automatique rétablie (" & DateReferenceTexte() & ")."
    Dlg_Msg "Date de référence rétablie : " & DateReferenceTexte() & ".", vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

Public Sub ArchiverRevue()
    Const NOM As String = "ArchiverRevue"
    Dim cp As Worksheet, racine As String, nomF As String, k As Long

    On Error GoTo Echec
    If Not FeuilleExiste(FEUILLE_REVUE) Then
        Dlg_Msg "Feuille « " & FEUILLE_REVUE & " » introuvable.", vbExclamation
        Exit Sub
    End If
    racine = "Hist_" & Format$(DateReference(), "yyyy-mm-dd")
    nomF = racine
    k = 1
    Do While FeuilleExiste(nomF)
        k = k + 1
        nomF = racine & "_" & k
    Loop
    If Dlg_Msg("Créer l'instantané figé « " & nomF & " » de la feuille « " & FEUILLE_REVUE & " » ?", _
               vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    OptimiserDebut
    Set cp = CreerInstantane(nomF)
    cp.Activate
    Terminer
    Journaliser NOM, "Instantané « " & nomF & " » créé."
    Dlg_Msg "Instantané « " & nomF & " » créé (feuille protégée sans mot de passe).", vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

' Copie la feuille de revue, remplace les formules par leurs valeurs et protège la copie.
' En cas d'échec, la copie partielle est supprimée.
Private Function CreerInstantane(ByVal nomF As String) As Worksheet
    Dim cp As Worksheet, cel As Range, nb As Long, numErr As Long, descErr As String
    Application.DisplayAlerts = False
    Recalculer
    nb = ThisWorkbook.Worksheets.Count
    ThisWorkbook.Worksheets(FEUILLE_REVUE).Copy After:=ThisWorkbook.Worksheets(nb)
    Set cp = ThisWorkbook.Worksheets(nb + 1)
    On Error GoTo Echec
    cp.Unprotect Password:=""
    cp.Name = nomF
    For Each cel In cp.UsedRange.Cells
        If cel.HasFormula Then FigerCellule cel
    Next cel
    cp.Rows(1).Insert
    With cp.Range("A1")
        .Value = "INSTANTANÉ FIGÉ - date de référence " & DateReferenceTexte() & _
                 " - archivé le " & DateHeure(Now) & " (valeurs, sans formules)"
        .Font.Bold = True
        .Font.Color = RGB(192, 0, 0)
    End With
    On Error Resume Next
    cp.Tab.Color = RGB(128, 128, 128)
    On Error GoTo Echec
    cp.Protect Password:="", Contents:=True, DrawingObjects:=False
    Application.DisplayAlerts = True
    Set CreerInstantane = cp
    Exit Function
Echec:
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    cp.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Err.Raise numErr, "CreerInstantane", descErr
End Function

' Remplace la formule par sa valeur (le texte reste du texte, l'erreur devient un libellé).
Private Sub FigerCellule(ByVal cel As Range)
    Dim v As Variant
    v = cel.Value
    cel.Formula = ""
    If IsError(v) Then
        cel.Value = "'" & TexteErreur(v)
    ElseIf VarType(v) = vbString Then
        If Len(v) > 0 Then cel.Value = "'" & v
    ElseIf Not IsEmpty(v) Then
        cel.Value = v
    End If
End Sub

Public Sub ExporterRevuePDF()
    Const NOM As String = "ExporterRevuePDF"
    Dim noms() As Variant, p As Variant, n As Long, chemin As String, t As Variant

    On Error GoTo Echec
    OptimiserDebut
    Application.Calculate
    MettreAJourTableauDeBord
    t = ConstruireAlertes(True)
    Terminer
    ReDim noms(0 To 0)
    For Each p In Split(FEUILLES_PDF, ";")
        If Len(Trim$(CStr(p))) > 0 Then
            If FeuilleExiste(Trim$(CStr(p))) Then
                If ThisWorkbook.Worksheets(Trim$(CStr(p))).Visible = xlSheetVisible Then
                    ReDim Preserve noms(0 To n)
                    noms(n) = Trim$(CStr(p))
                    n = n + 1
                End If
            End If
        End If
    Next p
    If n = 0 Then Exit Sub

    chemin = Dlg_FichierPDF("Revue_Pilotage_" & Format$(DateReference(), "yyyy-mm-dd") & ".pdf", _
                            "Exporter la revue en PDF")
    If Len(chemin) = 0 Then Exit Sub
    ExporterFeuillesPDF noms, chemin, Not gModeTest
    ThisWorkbook.Worksheets(CStr(noms(0))).Select
    Journaliser NOM, "PDF : " & chemin & " (" & Join(noms, ", ") & ")."
    Exit Sub
Echec:
    On Error Resume Next
    ThisWorkbook.Worksheets(FEUILLE_ACCUEIL).Select
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

' Exporte un groupe de feuilles visibles en un seul PDF.
Public Sub ExporterFeuillesPDF(ByVal noms As Variant, ByVal chemin As String, ByVal ouvrir As Boolean)
    If gModeTest Then Trace "PDF", chemin & " <- " & Join(noms, ";")
    ThisWorkbook.Activate                     ' la sélection de feuilles exige le classeur actif
    ThisWorkbook.Worksheets(noms).Select
    ActiveSheet.ExportAsFixedFormat Type:=0, Filename:=chemin, Quality:=0, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=ouvrir
    ThisWorkbook.Worksheets(CStr(noms(LBound(noms)))).Select
End Sub

' Message à l'ouverture (ThisWorkbook.Workbook_Open) : uniquement s'il existe des
' alertes critiques ou une date de référence invalide ; le classeur n'est pas modifié.
Public Sub NotificationOuverture()
    Dim al As Collection, nC As Long, nA As Long, it As Variant, cle As String, msg As String
    Dim cles() As String, cpt() As Long, n As Long, i As Long, trouve As Boolean

    Application.Calculate
    If Not DateReferenceValide() Then
        Dlg_Msg "La date de référence (Paramètres!C5) n'est pas une date : les délais et alertes sont faux." & _
                vbCrLf & MSG_LIBERER, vbExclamation
        Exit Sub
    End If
    Set al = CollecterAlertes(nC, nA)
    If nC = 0 Then Exit Sub
    ReDim cles(0 To 0)
    ReDim cpt(0 To 0)
    For Each it In al
        If it(0) = "CRITIQUE" Then
            cle = it(1) & " - " & it(2)
            trouve = False
            For i = 0 To n - 1
                If cles(i) = cle Then
                    cpt(i) = cpt(i) + 1
                    trouve = True
                    Exit For
                End If
            Next i
            If Not trouve Then
                ReDim Preserve cles(0 To n)
                ReDim Preserve cpt(0 To n)
                cles(n) = cle
                cpt(n) = 1
                n = n + 1
            End If
        End If
    Next it
    For i = 0 To n - 1
        If i < 12 Then msg = msg & vbCrLf & "   · " & cles(i) & " : " & cpt(i)
    Next i
    If n > 12 Then msg = msg & vbCrLf & "   · ... et " & (n - 12) & " autre(s) motif(s)"
    Dlg_Msg nC & " alerte(s) critique(s) à la date de référence du " & DateReferenceTexte() & " :" & msg & _
            vbCrLf & vbCrLf & "Détail : bouton « Synthèse des alertes ».", vbExclamation
End Sub

' Comptage sans interface : renvoie « critiques;à surveiller ».
Public Function CompterAlertes() As String
    Dim al As Collection, nC As Long, nA As Long
    Set al = CollecterAlertes(nC, nA)
    CompterAlertes = nC & ";" & nA
End Function

'==============================================================================
' 2. TABLEAU DE BORD
'==============================================================================

' Horodatage et ajustement du graphique « bailleurs » au nombre de bailleurs.
Public Sub MettreAJourTableauDeBord()
    Dim ws As Worksheet, co As Object, nb As Long, r As Long
    If Not FeuilleExiste(FEUILLE_TDB) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(FEUILLE_TDB)
    Deproteger ws
    ws.Range("G2").Value = Now
    ws.Range("G2").NumberFormat = "dd/mm/yyyy hh:mm"
    For r = 25 To 34
        If Len(Txt(ws.Cells(r, 2).Value)) > 0 Then nb = r - 24
    Next r
    If nb < 1 Then Exit Sub
    On Error Resume Next                         ' graphiques : sans incidence sur les données
    For Each co In ws.ChartObjects
        If co.Chart.HasTitle Then
            If InStr(1, co.Chart.ChartTitle.Text, "bailleur", vbTextCompare) > 0 Then
                co.Chart.SeriesCollection(1).Values = ws.Range("C25").Resize(nb, 1)
                co.Chart.SeriesCollection(1).XValues = ws.Range("B25").Resize(nb, 1)
            End If
        End If
    Next co
    On Error GoTo 0
End Sub

'==============================================================================
' 3. SYNTHÈSE DES ALERTES
'==============================================================================

Private Function CollecterAlertes(ByRef nCrit As Long, ByRef nAtt As Long) As Collection
    Dim res As New Collection, regle As Variant, ws As Worksheet
    Dim r As Long, r1 As Long, r2 As Long, cAl As Long, cJ As Long, cR As Long, k As Long
    Dim idSpec() As String, idNum() As Long, idTxt As String, grav As String
    Dim crit As String, att As String, v As Variant, x As Double, jours As Variant, resp As String
    Dim dRef As Date

    dRef = DateReference()
    nCrit = 0
    nAtt = 0
    For Each regle In ReglesAlerte()
        If FeuilleExiste(Champ(regle, RA_FEUILLE)) Then
            Set ws = ThisWorkbook.Worksheets(Champ(regle, RA_FEUILLE))
            r1 = CLng(Champ(regle, RA_PREMIERE))
            r2 = CLng(Champ(regle, RA_DERNIERE))
            cAl = NumCol(Champ(regle, RA_ALERTE))
            cJ = NumCol(Champ(regle, RA_JOURS))
            cR = NumCol(Champ(regle, RA_RESP))
            crit = Champ(regle, RA_CRITIQUE)
            att = Champ(regle, RA_ATTENTION)
            idSpec = Split(Champ(regle, RA_ID), ",")
            ReDim idNum(0 To UBound(idSpec))
            For k = 0 To UBound(idSpec)
                idNum(k) = NumCol(idSpec(k))
            Next k
            For r = r1 To r2
                v = ws.Cells(r, cAl).Value
                grav = ""
                If Correspond(v, crit, dRef) Then
                    grav = "CRITIQUE"
                    nCrit = nCrit + 1
                ElseIf Correspond(v, att, dRef) Then
                    grav = "ATTENTION"
                    nAtt = nAtt + 1
                End If
                If Len(grav) > 0 Then
                    ' identifiant : 1re colonne ID non vide, sinon numéro de ligne (ligne jamais ignorée)
                    idTxt = ""
                    For k = 0 To UBound(idNum)
                        idTxt = Txt(ws.Cells(r, idNum(k)).Value)
                        If Len(idTxt) > 0 Then Exit For
                    Next k
                    If Len(idTxt) = 0 Then idTxt = "(ligne " & r & ")"
                    jours = Empty
                    If cJ > 0 Then
                        If EnNombre(ws.Cells(r, cJ).Value, x) Then jours = Round(x, 0)
                    End If
                    resp = ""
                    If cR > 0 Then resp = Txt(ws.Cells(r, cR).Value)
                    res.Add Array(grav, ws.Name, Champ(regle, RA_MOTIF), idTxt, _
                                  Libelle(ws, r, Champ(regle, RA_LIBELLE)), Txt(v), jours, resp, _
                                  ws.Cells(r, cAl).Address(False, False))
                End If
            Next r
        End If
    Next regle
    Set CollecterAlertes = res
End Function

' Construit la feuille « Alertes ». Renvoie Array(nb critiques, nb à surveiller).
Private Function ConstruireAlertes(ByVal silencieux As Boolean) As Variant
    Dim al As Collection, ws As Worksheet, prec As Object, it As Variant
    Dim nCrit As Long, nAtt As Long, n As Long, i As Long, j As Long
    Dim data() As Variant

    Set al = CollecterAlertes(nCrit, nAtt)
    n = al.Count
    Set prec = FeuilleActive()
    Set ws = PreparerFeuille(FEUILLE_ALERTES, RGB(192, 0, 0))
    With ws
        .Range("A1").Value = "Synthèse des alertes - " & APP_TITRE
        .Range("A2").Value = "Date de référence : " & DateReferenceTexte() & _
            "   ·   générée le " & DateHeure(Now) & _
            "   ·   " & nCrit & " critique(s), " & nAtt & " à surveiller"
        .Range("A4").Resize(1, NB_COL_ALERTES).Value = Array("Gravité", "Feuille", "Motif", "ID", "Libellé", _
            "Statut / valeur", "Jours", "Responsable / détenteur", "Lien")
        If n > 0 Then
            ReDim data(1 To n, 1 To NB_COL_ALERTES)
            ' critiques d'abord, puis par nombre de jours croissant (tri par insertion stable)
            i = 0
            For Each it In al
                i = i + 1
                data(i, 1) = it(0)
                For j = 2 To 6
                    data(i, j) = Texte(it(j - 1))
                Next j
                data(i, 7) = it(6)
                data(i, 8) = Texte(it(7))
                data(i, 9) = "=HYPERLINK(""#'" & Replace(it(1), "'", "''") & "'!" & it(8) & """,""Ouvrir"")"
            Next it
            TrierAlertes data, n
            .Range("A5").Resize(n, NB_COL_ALERTES).Value = data
            For i = 5 To n + 4
                If .Cells(i, 1).Value = "CRITIQUE" Then
                    .Cells(i, 1).Interior.Color = RGB(255, 199, 206)
                    .Cells(i, 1).Font.Color = RGB(156, 0, 6)
                Else
                    .Cells(i, 1).Interior.Color = RGB(255, 235, 156)
                    .Cells(i, 1).Font.Color = RGB(156, 87, 0)
                End If
            Next i
        Else
            .Range("A5").Value = "Aucune alerte selon les règles en vigueur."
        End If
    End With
    MettreEnForme ws, NB_COL_ALERTES, IIf(n > 0, n, 1), Array(12, 22, 30, 16, 48, 30, 8, 28, 9)
    If silencieux Then
        If Not prec Is Nothing Then prec.Activate
    Else
        ws.Activate
        ws.Range("A5").Select
    End If
    ConstruireAlertes = Array(nCrit, nAtt)
End Function

' Tri en mémoire : CRITIQUE avant ATTENTION, puis jours croissants (vides en dernier).
Private Sub TrierAlertes(ByRef data() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, k As Long, tmp() As Variant
    ReDim tmp(1 To NB_COL_ALERTES)
    For i = 2 To n
        For k = 1 To NB_COL_ALERTES
            tmp(k) = data(i, k)
        Next k
        j = i - 1
        Do While j >= 1
            If Not AvantAlerte(tmp, data, j) Then Exit Do
            For k = 1 To NB_COL_ALERTES
                data(j + 1, k) = data(j, k)
            Next k
            j = j - 1
        Loop
        For k = 1 To NB_COL_ALERTES
            data(j + 1, k) = tmp(k)
        Next k
    Next i
End Sub

Private Function AvantAlerte(ByRef a() As Variant, ByRef data() As Variant, ByVal j As Long) As Boolean
    Dim ga As Long, gb As Long
    ga = IIf(a(1) = "CRITIQUE", 0, 1)
    gb = IIf(data(j, 1) = "CRITIQUE", 0, 1)
    If ga <> gb Then
        AvantAlerte = (ga < gb)
    ElseIf IsEmpty(a(7)) Then
        AvantAlerte = False
    ElseIf IsEmpty(data(j, 7)) Then
        AvantAlerte = True
    Else
        AvantAlerte = (a(7) < data(j, 7))
    End If
End Function

'==============================================================================
' 4. CONTRÔLE QUALITÉ DES SAISIES
'==============================================================================

' Ajoute une anomalie (utilisé aussi par modConfig.ControlesSpecifiques).
' Gravité : BLOQUANT, À CORRIGER, À COMPLÉTER ou INFO.
Public Sub QC_Ajouter(ByVal feuille As String, ByVal adresse As String, ByVal id As String, _
                      ByVal controle As String, ByVal detail As String, ByVal gravite As String)
    If mQC Is Nothing Then Set mQC = New Collection
    mQC.Add Array(gravite, feuille, adresse, id, controle, detail)
End Sub

Private Function ConstruireControle(ByVal silencieux As Boolean) As Long
    Dim regle As Variant, ws As Worksheet, prec As Object, it As Variant
    Dim ordre As Variant, g As Variant, n As Long, i As Long, j As Long
    Dim data() As Variant

    Set mQC = New Collection
    For Each regle In ReglesControle()
        If FeuilleExiste(Champ(regle, RC_FEUILLE)) Then ControlerRegle CStr(regle)
    Next regle
    ControlesSpecifiques
    If DateReferenceFigee() And DateReferenceValide() Then
        QC_Ajouter "", "", "", "Date de référence figée", _
            "Calculs établis au " & DateReferenceTexte() & " (et non à la date du jour). " & MSG_LIBERER, "INFO"
    End If

    n = mQC.Count
    Set prec = FeuilleActive()
    Set ws = PreparerFeuille(FEUILLE_QC, RGB(237, 125, 49))
    With ws
        .Range("A1").Value = "Contrôle qualité des saisies - " & APP_TITRE
        .Range("A2").Value = "Date de référence : " & DateReferenceTexte() & _
            "   ·   généré le " & DateHeure(Now) & "   ·   " & n & " anomalie(s)"
        .Range("A4").Resize(1, NB_COL_QC).Value = Array("Gravité", "Feuille", "Cellule", "ID", "Contrôle", "Détail", "Lien")
        If n > 0 Then
            ReDim data(1 To n, 1 To NB_COL_QC)
            ordre = Array("BLOQUANT", "À CORRIGER", "À COMPLÉTER", "INFO")
            i = 0
            For Each g In ordre
                For Each it In mQC
                    If it(0) = g Then
                        i = i + 1
                        data(i, 1) = it(0)
                        For j = 2 To 6
                            data(i, j) = Texte(it(j - 1))
                        Next j
                        If Len(it(1)) > 0 And Len(it(2)) > 0 Then
                            data(i, 7) = "=HYPERLINK(""#'" & Replace(it(1), "'", "''") & "'!" & it(2) & """,""Ouvrir"")"
                        End If
                    End If
                Next it
            Next g
            .Range("A5").Resize(n, NB_COL_QC).Value = data
            For i = 5 To n + 4
                Select Case .Cells(i, 1).Value
                    Case "BLOQUANT"
                        .Cells(i, 1).Interior.Color = RGB(255, 199, 206)
                    Case "À CORRIGER"
                        .Cells(i, 1).Interior.Color = RGB(255, 235, 156)
                    Case "À COMPLÉTER"
                        .Cells(i, 1).Interior.Color = RGB(221, 235, 247)
                    Case Else
                        .Cells(i, 1).Interior.Color = RGB(242, 242, 242)
                End Select
            Next i
        Else
            .Range("A5").Value = "Aucune anomalie détectée."
        End If
    End With
    MettreEnForme ws, NB_COL_QC, IIf(n > 0, n, 1), Array(13, 24, 10, 16, 34, 70, 9)
    If silencieux Then
        If Not prec Is Nothing Then prec.Activate
    Else
        ws.Activate
        ws.Range("A5").Select
    End If
    ConstruireControle = n
End Function

Private Sub ControlerRegle(ByVal regle As String)
    Dim ws As Worksheet, hdr As Long, r1 As Long, r2 As Long, lastC As Long, i As Long, c As Long, k As Long
    Dim idCol As Long, calc As Variant, dts As Variant, obl As Variant, tem As Variant
    Dim vals As Variant, frm As Variant, entetes As Variant, vus As New Collection
    Dim idTxt As String, cle As String, s As String, x As Double, v As Variant
    Dim premiere As Long, utilisee As Boolean, exemple As Boolean

    Set ws = ThisWorkbook.Worksheets(Champ(regle, RC_FEUILLE))
    hdr = CLng(Champ(regle, RC_ENTETE))
    r1 = CLng(Champ(regle, RC_PREMIERE))
    r2 = CLng(Champ(regle, RC_DERNIERE))
    idCol = NumCol(Champ(regle, RC_ID))
    calc = ListeColonnes(Champ(regle, RC_CALCUL))
    dts = ListeColonnes(Champ(regle, RC_DATES))
    obl = ListeColonnes(Champ(regle, RC_OBLIG))
    If Len(Champ(regle, RC_TEMOINS)) > 0 Then
        tem = ListeColonnes(Champ(regle, RC_TEMOINS))
    Else
        tem = obl
    End If

    lastC = DerniereColonne(ws, hdr)
    If idCol > lastC Then lastC = idCol
    lastC = MaxCol(lastC, calc)
    lastC = MaxCol(lastC, dts)
    lastC = MaxCol(lastC, obl)
    lastC = MaxCol(lastC, tem)
    vals = ws.Range(ws.Cells(r1, 1), ws.Cells(r2, lastC)).Value
    frm = ws.Range(ws.Cells(r1, 1), ws.Cells(r2, lastC)).Formula
    entetes = ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, lastC)).Value

    For i = 1 To r2 - r1 + 1
        ' Ligne utilisée : identifiant ou au moins une colonne témoin renseignée
        idTxt = Txt(vals(i, idCol))
        utilisee = (Len(idTxt) > 0)
        If Not utilisee Then
            For k = 0 To UBound(tem)
                If tem(k) > 0 Then
                    If Len(Txt(vals(i, tem(k)))) > 0 And Left$(CStr(frm(i, tem(k))), 1) <> "=" Then
                        utilisee = True
                        Exit For
                    End If
                End If
            Next k
        End If
        If utilisee Then
            If Len(idTxt) = 0 Then
                idTxt = "ligne " & (r1 + i - 1)
                QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, idCol), idTxt, "Identifiant manquant", _
                    "Ligne renseignée sans « " & Entete(entetes, idCol) & " » : à compléter (référence des alertes et des recherches).", "À COMPLÉTER"
            Else
                ' Doublon d'identifiant
                cle = UCase$(idTxt)
                On Error Resume Next
                vus.Add r1 + i - 1, cle
                If Err.Number <> 0 Then
                    Err.Clear
                    premiere = vus.Item(cle)
                    On Error GoTo 0
                    QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, idCol), idTxt, "Identifiant en doublon", _
                        "Même identifiant qu'en ligne " & premiere & " : les recherches (INDEX/EQUIV) ne lisent que la première.", "BLOQUANT"
                End If
                On Error GoTo 0
            End If
            ' Ligne d'exemple : identifiant « EX » ou mention « exemple »
            exemple = (UCase$(idTxt) = "EX")
            ' Formules supprimées ou écrasées
            For k = 0 To UBound(calc)
                If calc(k) > 0 Then
                    If Left$(CStr(frm(i, calc(k))), 1) <> "=" Then
                        If Len(Txt(vals(i, calc(k)))) = 0 Then
                            s = "Formule supprimée"
                        Else
                            s = "Formule remplacée par une valeur saisie"
                        End If
                        QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, calc(k)), idTxt, s, _
                            "Colonne calculée « " & Entete(entetes, calc(k)) & " » : recopier la formule de la ligne précédente.", "BLOQUANT"
                    End If
                End If
            Next k
            ' Erreurs de calcul et références inconnues
            For c = 1 To lastC
                v = vals(i, c)
                If IsError(v) Then
                    QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, c), idTxt, "Erreur de calcul", _
                        TexteErreur(v) & " dans « " & Entete(entetes, c) & " ».", "BLOQUANT"
                ElseIf VarType(v) = vbString Then
                    s = UCase$(v)
                    If InStr(s, "EXEMPLE") > 0 Then exemple = True
                    If Left$(s, 14) = "MARCHÉ INCONNU" Or Left$(s, 12) = "CODE INCONNU" Or Left$(s, 19) = "CODE MARCHÉ INCONNU" Then
                        QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, c), idTxt, "Référence inconnue", _
                            "« " & Entete(entetes, c) & " » : code absent du référentiel (" & CStr(v) & ").", "À CORRIGER"
                    End If
                End If
            Next c
            If exemple Then
                QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, idCol), idTxt, "Ligne d'exemple", _
                    "Ligne d'exemple encore présente dans le registre : elle est comptée dans les indicateurs.", "INFO"
            End If
            ' Dates
            For k = 0 To UBound(dts)
                If dts(k) > 0 Then
                    v = vals(i, dts(k))
                    If Not IsError(v) And Len(Txt(v)) > 0 Then
                        If VarType(v) = vbString Then
                            QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, dts(k)), idTxt, "Date non reconnue", _
                                "« " & Entete(entetes, dts(k)) & " » saisie en texte (" & CStr(v) & ") : ignorée par les calculs.", "À CORRIGER"
                        ElseIf EnNombre(v, x) Then
                            If x < CDbl(DateSerial(2000, 1, 1)) Or x > CDbl(DateSerial(2060, 12, 31)) Then
                                QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, dts(k)), idTxt, "Date hors plage", _
                                    "« " & Entete(entetes, dts(k)) & " » : " & Txt(v) & " (plage plausible 2000-2060).", "À CORRIGER"
                            End If
                        End If
                    End If
                End If
            Next k
            ' Champs obligatoires
            For k = 0 To UBound(obl)
                If obl(k) > 0 Then
                    If Len(Txt(vals(i, obl(k)))) = 0 Then
                        QC_Ajouter ws.Name, Adr(ws, r1 + i - 1, obl(k)), idTxt, "Champ obligatoire vide", _
                            "« " & Entete(entetes, obl(k)) & " » non renseigné.", "À COMPLÉTER"
                    End If
                End If
            Next k
        End If
    Next i
End Sub

'==============================================================================
' 5. UTILITAIRES PRIVÉS
'==============================================================================

' Lignes min / max couvertes par les règles d'alerte d'une feuille (0 si aucune).
Private Sub BornesFeuille(ByVal nom As String, ByRef r1 As Long, ByRef r2 As Long)
    Dim regle As Variant, a As Long, b As Long
    r1 = 0
    r2 = 0
    For Each regle In ReglesAlerte()
        If Champ(regle, RA_FEUILLE) = nom Then
            a = CLng(Champ(regle, RA_PREMIERE))
            b = CLng(Champ(regle, RA_DERNIERE))
            If r1 = 0 Or a < r1 Then r1 = a
            If b > r2 Then r2 = b
        End If
    Next regle
End Sub

Private Function MasqueesDans(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long) As Boolean
    Dim r As Long
    For r = r1 To r2
        If ws.Rows(r).Hidden Then
            MasqueesDans = True
            Exit Function
        End If
    Next r
End Function

' Dernière colonne renseignée de la ligne d'en-tête (plafonnée à 80 colonnes).
Private Function DerniereColonne(ByVal ws As Worksheet, ByVal ligne As Long) As Long
    Dim c As Long, fin As Long
    fin = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If fin > 80 Then fin = 80
    For c = fin To 1 Step -1
        If Len(Txt(ws.Cells(ligne, c).Value)) > 0 Then
            DerniereColonne = c
            Exit Function
        End If
    Next c
    DerniereColonne = 1
End Function

Private Function ListeColonnes(ByVal spec As String) As Variant
    Dim t() As String, res() As Long, k As Long
    spec = Trim$(spec)
    If Len(spec) = 0 Then
        ReDim res(0 To 0)
        res(0) = 0
    Else
        t = Split(spec, ",")
        ReDim res(0 To UBound(t))
        For k = 0 To UBound(t)
            res(k) = NumCol(t(k))
        Next k
    End If
    ListeColonnes = res
End Function

Private Function MaxCol(ByVal actuel As Long, ByVal cols As Variant) As Long
    Dim k As Long
    MaxCol = actuel
    For k = LBound(cols) To UBound(cols)
        If cols(k) > MaxCol Then MaxCol = cols(k)
    Next k
End Function

' Jetons : « X » préfixe (sans tenir compte de la casse) ; « =X » égalité stricte ;
' « #>0 » nombre strictement positif ; « #<=REF » date antérieure ou égale à la date de référence.
Private Function Correspond(ByVal v As Variant, ByVal jetons As String, ByVal dRef As Date) As Boolean
    Dim t As Variant, s As String, j As String, x As Double
    If Len(jetons) = 0 Then Exit Function
    If IsError(v) Then Exit Function
    s = UCase$(Txt(v))
    If Len(s) = 0 Then Exit Function
    For Each t In Split(jetons, ";")
        j = UCase$(Trim$(CStr(t)))
        If Len(j) > 0 Then
            If j = "#>0" Then
                If EnNombre(v, x) Then
                    If x > 0 Then Correspond = True
                End If
            ElseIf j = "#<=REF" Then
                If EnNombre(v, x) Then
                    If x <= CDbl(dRef) Then Correspond = True
                End If
            ElseIf Left$(j, 1) = "=" Then
                If s = Mid$(j, 2) Then Correspond = True
            ElseIf Left$(s, Len(j)) = j Then
                Correspond = True
            End If
            If Correspond Then Exit Function
        End If
    Next t
End Function

' Libellé : « C,B » = première valeur non vide ; « B+E » = valeurs jointes par « - ».
Private Function Libelle(ByVal ws As Worksheet, ByVal r As Long, ByVal spec As String) As String
    Dim p As Variant, v As String, res As String
    If Len(spec) = 0 Then Exit Function
    If InStr(spec, "+") > 0 Then
        For Each p In Split(spec, "+")
            v = Txt(ws.Cells(r, NumCol(CStr(p))).Value)
            If Len(v) > 0 Then
                If Len(res) > 0 Then res = res & " - "
                res = res & v
            End If
        Next p
    Else
        For Each p In Split(spec, ",")
            v = Txt(ws.Cells(r, NumCol(CStr(p))).Value)
            If Len(v) > 0 Then
                res = v
                Exit For
            End If
        Next p
    End If
    If Len(res) > 250 Then res = Left$(res, 247) & "..."
    Libelle = res
End Function

Private Function Entete(ByVal entetes As Variant, ByVal c As Long) As String
    Entete = Replace(Replace(Txt(entetes(1, c)), vbLf, " "), vbCr, "")
    If Len(Entete) = 0 Then Entete = "colonne " & LettreCol(c)
End Function

Private Function Adr(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Adr = LettreCol(c) & r
End Function

Private Function PreparerFeuille(ByVal nom As String, ByVal couleur As Long) As Worksheet
    Dim ws As Worksheet
    If FeuilleExiste(nom) Then
        Set ws = ThisWorkbook.Worksheets(nom)
        Deproteger ws
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
    End If
    On Error Resume Next
    ws.Tab.Color = couleur
    On Error GoTo 0
    Set PreparerFeuille = ws
End Function

Private Sub MettreEnForme(ByVal ws As Worksheet, ByVal nbCol As Long, ByVal nbLig As Long, ByVal largeurs As Variant)
    Dim k As Long, prec As Object
    With ws
        .Cells.Font.Name = "Arial"
        .Cells.Font.Size = 9
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(31, 56, 100)
        .Range("A2").Font.Italic = True
        .Range("A2").Font.Color = RGB(89, 89, 89)
        With .Range("A4").Resize(1, nbCol)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 56, 100)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        With .Range("A4").Resize(nbLig + 1, nbCol)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(217, 217, 217)
        End With
        With .Range("A5").Resize(nbLig, nbCol)
            .VerticalAlignment = xlTop
            .WrapText = True
        End With
        For k = 0 To UBound(largeurs)
            .Columns(k + 1).ColumnWidth = largeurs(k)
        Next k
        .Range("A4").Resize(nbLig + 1, nbCol).AutoFilter
    End With
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$4:$4"
    End With
    If Not gModeTest Then
        Set prec = FeuilleActive()
        ws.Activate
        ActiveWindow.FreezePanes = False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = 4
        ActiveWindow.FreezePanes = True
        If Not prec Is Nothing Then prec.Activate
    End If
    On Error GoTo 0
End Sub
