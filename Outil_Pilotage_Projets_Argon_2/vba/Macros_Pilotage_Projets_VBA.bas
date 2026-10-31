Attribute VB_Name = "Mod_Pilotage_Projets_Turbo"
Option Explicit

' ==========================================================================================
' MODULE VBA HAUTE PERFORMANCE — PILOTAGE DES PROJETS, DLM, GARANTIES (J-60) & FACTURES
' Conçu pour accélérer le traitement en mémoire (Variant Arrays) et automatiser :
'   1. Plan d'amélioration (0-30 jours, 1-3 mois, 3-12 mois)
'   2. Suivi des DLM et prorogations (BID, BOAD, FVC, ADFD) + Dépôt immédiat ADFD & Escalade DG
'   3. Registre des garanties (bonne exécution, cautions d'avance) avec alerte à 60 jours
'   4. Registre des factures/décomptes et plan de paiement priorisé par échéance DLM (AC & DFP)
' ==========================================================================================

' --- 1. MOTEUR D'OPTIMISATION DE LA PERFORMANCE EXCEL (MODE TURBO) ---
Public Sub Activer_Mode_Turbo()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayStatusBar = True
        .StatusBar = "Traitement haute performance en cours..."
    End With
End Sub

Public Sub Desactiver_Mode_Turbo()
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
        .StatusBar = False
    End With
End Sub

' --- 2. ACTUALISATION GLOBALE DU COCKPIT EXÉCUTIF DG (Raccourci : Ctrl+Shift+A) ---
Public Sub Actualiser_Cockpit_Global()
    Dim tStart As Double
    tStart = Timer
    On Error GoTo ErrHandler
    
    Call Activer_Mode_Turbo
    
    ' Recalcul intégral des formules multi-feuilles et des tableaux de bord
    Application.CalculateFullRebuild
    
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("00_DASHBOARD")
    wsDash.Activate
    
    Call Desactiver_Mode_Turbo
    MsgBox "Cockpit Exécutif DG actualisé avec succès en " & Format((Timer - tStart) * 1000, "0.0") & " ms." & vbCrLf & vbCrLf & _
           "• Avancement global (0-12 mois) : " & Format(wsDash.Range("A5").Value, "0.0%") & vbCrLf & _
           "• Dossiers DLM en Escalade DG : " & wsDash.Range("D5").Value & vbCrLf & _
           "• Garanties en alerte (<= 60j) : " & wsDash.Range("G5").Value & vbCrLf & _
           "• Volume Factures / Décomptes : " & Format(wsDash.Range("J5").Value, "#,##0") & " FCFA", _
           vbInformation, "Cockpit DG — Haute Performance"
    Exit Sub

ErrHandler:
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur lors de l'actualisation : " & Err.Description, vbCritical, "Erreur VBA"
End Sub

' --- 3. CONTRÔLE DES DLM, DÉPÔT IMMÉDIAT ADFD & ESCALADE DG (Raccourci : Ctrl+Shift+D) ---
Public Sub Verifier_DLM_Et_Escalade_DG()
    Dim wsDLM As Worksheet
    Dim arrData As Variant
    Dim i As Long, nbEscalade As Long, nbADFD As Long
    Dim msgEscalade As String
    
    On Error GoTo ErrHandler
    Call Activer_Mode_Turbo
    
    Set wsDLM = ThisWorkbook.Sheets("02_SUIVI_DLM")
    wsDLM.Calculate
    
    ' Lecture en mémoire (Variant Array) pour une exécution instantanée
    arrData = wsDLM.Range("A9:L24").Value
    nbEscalade = 0
    nbADFD = 0
    msgEscalade = ""
    
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        If Trim(CStr(arrData(i, 2))) <> "" Then
            ' Mise à jour de la date de contrôle au jour courant si souhaité
            wsDLM.Cells(8 + i, 9).Value = Date
            
            If UCase(Trim(CStr(arrData(i, 2)))) = "ADFD" Then
                nbADFD = nbADFD + 1
            End If
            
            If InStr(1, UCase(CStr(arrData(i, 11))), "ESCALADE") > 0 Or Val(arrData(i, 6)) <= 45 Or UCase(Trim(CStr(arrData(i, 2)))) = "ADFD" Then
                nbEscalade = nbEscalade + 1
                msgEscalade = msgEscalade & "  - [" & arrData(i, 1) & " | " & arrData(i, 2) & "] " & _
                              Left(CStr(arrData(i, 3)), 38) & " (J-" & arrData(i, 6) & " | Pièce : " & arrData(i, 7) & ")" & vbCrLf
            End If
        End If
    Next i
    
    wsDLM.Calculate
    Call Desactiver_Mode_Turbo
    
    If nbEscalade > 0 Then
        MsgBox "ALERTE DLM & ESCALADE VERS LE DG (" & nbEscalade & " dossier(s) critique(s)) :" & vbCrLf & vbCrLf & _
               "Action immédiate requise (dont Dépôt immédiat ADFD) :" & vbCrLf & msgEscalade, _
               vbExclamation, "Escalade Direction Générale — Suivi DLM (BID, BOAD, FVC, ADFD)"
    Else
        MsgBox "Aucun dossier DLM en dépassement critique. Toutes les échéances sont à plus de 45 jours.", vbInformation, "Suivi DLM"
    End If
    Exit Sub

ErrHandler:
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur dans Verifier_DLM_Et_Escalade_DG : " & Err.Description, vbCritical
End Sub

' --- 4. CONTRÔLE DES GARANTIES À 60 JOURS AVANT ÉCHÉANCE (Raccourci : Ctrl+Shift+G) ---
Public Sub Controler_Garanties_60Jours()
    Dim wsGar As Worksheet
    Dim arrGar As Variant
    Dim i As Long, nbAlertes As Long
    Dim msgAlertes As String
    Dim montantAlerte As Double
    
    On Error GoTo ErrHandler
    Call Activer_Mode_Turbo
    
    Set wsGar = ThisWorkbook.Sheets("03_REGISTRE_GARANTIES")
    
    ' Actualise la date de référence (Colonne G) à la date du jour sur toutes les lignes actives
    For i = 9 To 24
        If Trim(CStr(wsGar.Cells(i, 2).Value)) <> "" Then
            wsGar.Cells(i, 7).Value = Date
        End If
    Next i
    wsGar.Calculate
    
    arrGar = wsGar.Range("A9:L24").Value
    nbAlertes = 0
    montantAlerte = 0
    msgAlertes = ""
    
    For i = LBound(arrGar, 1) To UBound(arrGar, 1)
        If Trim(CStr(arrGar(i, 2))) <> "" Then
            If Val(arrGar(i, 9)) <= 60 Then
                nbAlertes = nbAlertes + 1
                montantAlerte = montantAlerte + Val(arrGar(i, 6))
                msgAlertes = msgAlertes & "  - [" & arrGar(i, 1) & " | " & arrGar(i, 4) & "] " & _
                             arrGar(i, 2) & " -> J-" & arrGar(i, 9) & " (" & Format(arrGar(i, 6), "#,##0") & " FCFA)" & vbCrLf
            End If
        End If
    Next i
    
    Call Desactiver_Mode_Turbo
    
    If nbAlertes > 0 Then
        MsgBox "ALERTE GARANTIES À 60 JOURS : " & nbAlertes & " garantie(s)/caution(s) à renouveler immédiatement !" & vbCrLf & _
               "Montant total sous alerte : " & Format(montantAlerte, "#,##0") & " FCFA" & vbCrLf & vbCrLf & _
               msgAlertes, vbExclamation, "Registre des Garanties — Seuil 60 Jours"
    Else
        MsgBox "Toutes les garanties de bonne exécution et cautions d'avance ont une validité supérieure à 60 jours.", vbInformation, "Registre des Garanties"
    End If
    Exit Sub

ErrHandler:
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur dans Controler_Garanties_60Jours : " & Err.Description, vbCritical
End Sub

' --- 5. PLAN DE PAIEMENT PRIORISÉ PAR ÉCHÉANCE DE DLM AVEC AC & DFP (Raccourci : Ctrl+Shift+P) ---
Public Sub Generer_Plan_Paiement_Priorise_AC_DFP()
    Dim wsFac As Worksheet
    Dim lastRow As Long
    
    On Error GoTo ErrHandler
    Call Activer_Mode_Turbo
    
    Set wsFac = ThisWorkbook.Sheets("04_FACTURES_DECOMPTES")
    wsFac.Activate
    wsFac.Calculate
    
    ' Tri automatique multi-critères :
    ' 1. Par Priorité DLM croissante (P1 -> P2 -> P3 -> P4)
    ' 2. Par Date d'Échéance DLM croissante (Colonne H)
    ' 3. Par Montant décroissant (Colonne D)
    With wsFac.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsFac.Range("K9:K24"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=wsFac.Range("H9:H24"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=wsFac.Range("D9:D24"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange wsFac.Range("A8:M24")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    Call Desactiver_Mode_Turbo
    MsgBox "Le Plan de Paiement a été priorisé avec succès par échéance de DLM pour arbitrage conjoint AC & DFP." & vbCrLf & vbCrLf & _
           "• Urgences P1 (DLM <= 45j) : " & Format(wsFac.Range("D5").Value, "#,##0") & " FCFA" & vbCrLf & _
           "• Dossiers en instance AC/DFP : " & wsFac.Range("G5").Value, _
           vbInformation, "Plan de Paiement Priorisé (AC & DFP)"
    Exit Sub

ErrHandler:
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur lors du tri du plan de paiement : " & Err.Description, vbCritical
End Sub

' --- 6. FILTRAGE RAPIDE DU PLAN D'AMÉLIORATION PAR HORIZON (Raccourci : Ctrl+Shift+H) ---
Public Sub Filtrer_Plan_Amelioration_Par_Horizon()
    Dim wsPlan As Worksheet
    Dim choix As String
    
    Set wsPlan = ThisWorkbook.Sheets("01_PLAN_AMELIORATION")
    wsPlan.Activate
    
    choix = InputBox("Entrez l'horizon à filtrer :" & vbCrLf & _
                     "  1 = 0-30 jours (Sécuriser les financements)" & vbCrLf & _
                     "  2 = 1-3 mois (Gouvernance, source unique, RACI, entreprises, fiscalité)" & vbCrLf & _
                     "  3 = 3-12 mois (Maturité, SIG, DAO standardisés, transfert exploitants)" & vbCrLf & _
                     "  0 = Afficher tous les horizons", "Filtre Plan d'Amélioration", "0")
    If choix = "" Then Exit Sub
    
    Call Activer_Mode_Turbo
    If wsPlan.AutoFilterMode Then wsPlan.AutoFilterMode = False
    wsPlan.Range("A8:L25").AutoFilter
    
    Select Case Trim(choix)
        Case "1": wsPlan.Range("A8:L25").AutoFilter Field:=2, Criteria1:="0-30 jours"
        Case "2": wsPlan.Range("A8:L25").AutoFilter Field:=2, Criteria1:="1-3 mois"
        Case "3": wsPlan.Range("A8:L25").AutoFilter Field:=2, Criteria1:="3-12 mois"
        Case Else
            If wsPlan.FilterMode Then wsPlan.ShowAllData
    End Select
    Call Desactiver_Mode_Turbo
End Sub

' --- 7. AUDIT QUALITÉ DU PORTEFEUILLE EN MÉMOIRE (Raccourci : Ctrl+Shift+Q) ---
Public Sub Audit_Qualite_Portefeuille_Complet()
    Dim wsDLM As Worksheet, wsGar As Worksheet, wsFac As Worksheet
    Dim anomalies As Long, rapport As String
    Dim i As Long
    
    Call Activer_Mode_Turbo
    Set wsDLM = ThisWorkbook.Sheets("02_SUIVI_DLM")
    Set wsGar = ThisWorkbook.Sheets("03_REGISTRE_GARANTIES")
    Set wsFac = ThisWorkbook.Sheets("04_FACTURES_DECOMPTES")
    
    anomalies = 0
    rapport = ""
    
    ' Vérification DLM sans détenteur ou sans échéance
    For i = 9 To 24
        If Trim(CStr(wsDLM.Cells(i, 2).Value)) <> "" Then
            If Not IsDate(wsDLM.Cells(i, 5).Value) Or Trim(CStr(wsDLM.Cells(i, 8).Value)) = "" Then
                anomalies = anomalies + 1
                rapport = rapport & "• [02_SUIVI_DLM Ligne " & i & "] Échéance ou Détenteur manquant." & vbCrLf
            End If
        End If
    Next i
    
    ' Vérification Garanties sans montant ou date d'échéance
    For i = 9 To 24
        If Trim(CStr(wsGar.Cells(i, 2).Value)) <> "" Then
            If Val(wsGar.Cells(i, 6).Value) <= 0 Or Not IsDate(wsGar.Cells(i, 8).Value) Then
                anomalies = anomalies + 1
                rapport = rapport & "• [03_REGISTRE_GARANTIES Ligne " & i & "] Montant ou Échéance invalide." & vbCrLf
            End If
        End If
    Next i
    
    ' Vérification Factures sans échéance DLM associée
    For i = 9 To 24
        If Trim(CStr(wsFac.Cells(i, 2).Value)) <> "" Then
            If Not IsDate(wsFac.Cells(i, 8).Value) Then
                anomalies = anomalies + 1
                rapport = rapport & "• [04_FACTURES_DECOMPTES Ligne " & i & "] Date d'échéance DLM manquante pour priorisation." & vbCrLf
            End If
        End If
    Next i
    
    Call Desactiver_Mode_Turbo
    If anomalies = 0 Then
        MsgBox "Audit de cohérence terminé : 0 anomalie détectée sur les 4 outils.", vbInformation, "Audit Qualité Portefeuille"
    Else
        MsgBox anomalies & " anomalie(s) détectée(s) :" & vbCrLf & vbCrLf & rapport, vbExclamation, "Audit Qualité Portefeuille"
    End If
End Sub

' --- 8. EXPORT AUTOMATIQUE DU RAPPORT EXÉCUTIF EN PDF (Raccourci : Ctrl+Shift+E) ---
Public Sub Exporter_Rapport_Executif_PDF()
    Dim pdfPath As String
    On Error GoTo ErrHandler
    
    Call Activer_Mode_Turbo
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & "Rapport_Pilotage_DG_" & Format(Date, "YYYYMMDD") & ".pdf"
    
    ThisWorkbook.Sheets(Array("00_DASHBOARD", "01_PLAN_AMELIORATION", "02_SUIVI_DLM", "03_REGISTRE_GARANTIES", "04_FACTURES_DECOMPTES")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    ThisWorkbook.Sheets("00_DASHBOARD").Select
    
    Call Desactiver_Mode_Turbo
    MsgBox "Rapport Exécutif exporté avec succès :" & vbCrLf & pdfPath, vbInformation, "Export PDF"
    Exit Sub

ErrHandler:
    ThisWorkbook.Sheets("00_DASHBOARD").Select
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur lors de l'export PDF : " & Err.Description, vbCritical
End Sub
