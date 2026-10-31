Attribute VB_Name = "Mod_Pilotage_Projets_Turbo"
Option Explicit

' ==========================================================================================
' MODULE VBA HAUTE PERFORMANCE — PILOTAGE DES PROJETS, DLM, GARANTIES (J-60) & FACTURES
' Conçu pour accélérer le traitement en mémoire (Variant Arrays) et automatiser :
'   1. Plan d'amélioration (0-30 jours, 1-3 mois, 3-12 mois)
'   2. Suivi des DLM et prorogations (BID, BOAD, FVC, ADFD) + Dépôt immédiat ADFD & Escalade DG
'   3. Registre des garanties (bonne exécution, cautions d'avance) avec alerte à 60 jours
'   4. Registre des factures/décomptes et plan de paiement priorisé par échéance DLM (AC & DFP)
' Raccourcis (activés à l'ouverture du classeur, désactivés quand il n'est plus actif) :
'   Ctrl+Shift+A / D / G / P / H / Q / E
' ==========================================================================================

Private Const LIGNE_ENTETE As Long = 8
Private Const LIGNE_DEBUT As Long = 9
Private Const LONGUEUR_MAX_MSG As Long = 850

Private mCalcState As XlCalculation
Private mTurboActif As Boolean

' --- 1. MOTEUR D'OPTIMISATION DE LA PERFORMANCE EXCEL (MODE TURBO) ---
' Le mode de calcul d'origine est mémorisé puis restauré (il n'est plus forcé en automatique).
Public Sub Activer_Mode_Turbo()
    On Error Resume Next
    If Not mTurboActif Then
        mCalcState = Application.Calculation
        mTurboActif = True
    End If
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayStatusBar = True
        .StatusBar = "Traitement haute performance en cours..."
    End With
End Sub

Public Sub Desactiver_Mode_Turbo()
    On Error Resume Next
    With Application
        If mTurboActif Then
            .Calculation = mCalcState
        Else
            .Calculation = xlCalculationAutomatic
        End If
        .EnableEvents = True
        .ScreenUpdating = True
        .StatusBar = False
    End With
    mTurboActif = False
End Sub

' --- RACCOURCIS CLAVIER (appelés par ThisWorkbook) ---
Public Sub Activer_Raccourcis()
    Dim pref As String
    pref = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    On Error Resume Next
    Application.OnKey "^+a", pref & "Actualiser_Cockpit_Global"
    Application.OnKey "^+d", pref & "Verifier_DLM_Et_Escalade_DG"
    Application.OnKey "^+g", pref & "Controler_Garanties_60Jours"
    Application.OnKey "^+p", pref & "Generer_Plan_Paiement_Priorise_AC_DFP"
    Application.OnKey "^+h", pref & "Filtrer_Plan_Amelioration_Par_Horizon"
    Application.OnKey "^+q", pref & "Audit_Qualite_Portefeuille_Complet"
    Application.OnKey "^+e", pref & "Exporter_Rapport_Executif_PDF"
End Sub

Public Sub Desactiver_Raccourcis()
    On Error Resume Next
    Application.OnKey "^+a"
    Application.OnKey "^+d"
    Application.OnKey "^+g"
    Application.OnKey "^+p"
    Application.OnKey "^+h"
    Application.OnKey "^+q"
    Application.OnKey "^+e"
End Sub

' --- FONCTIONS DE LECTURE SÉCURISÉE ---
' Texte d'une valeur de cellule (#N/A, #REF!, vide -> "")
Public Function TexteSur(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or IsArray(v) Then
        TexteSur = ""
    Else
        TexteSur = Trim$(CStr(v))
    End If
End Function

' Vrai si la valeur est un nombre exploitable (ni vide, ni texte, ni erreur)
Public Function EstNombre(ByVal v As Variant) As Boolean
    EstNombre = False
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or IsArray(v) Then Exit Function
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbDate
            EstNombre = True
    End Select
End Function

' Valeur numérique d'une cellule (Val() ignorait les décimales en paramètres régionaux français)
Public Function NombreSur(ByVal v As Variant, Optional ByVal valeurDefaut As Double = 0) As Double
    If EstNombre(v) Then
        NombreSur = CDbl(v)
    ElseIf VarType(v) = vbString And IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
        NombreSur = CDbl(v)
    Else
        NombreSur = valeurDefaut
    End If
End Function

' Vrai uniquement pour une vraie date Excel (une date saisie en texte casse les formules)
Public Function EstDate(ByVal v As Variant) As Boolean
    If IsError(v) Then
        EstDate = False
    Else
        EstDate = (VarType(v) = vbDate)
    End If
End Function

' Formatage sécurisé pour les messages
Public Function FormatSur(ByVal v As Variant, ByVal fmt As String) As String
    If IsError(v) Or IsEmpty(v) Then
        FormatSur = "n/d"
    ElseIf EstNombre(v) Then
        FormatSur = Format$(v, fmt)
    Else
        FormatSur = TexteSur(v)
    End If
End Function

' Libellé des jours restants (J-15, ou dépassement)
Public Function TexteJours(ByVal v As Variant) As String
    If Not EstNombre(v) Then
        TexteJours = "date non renseignée"
    ElseIf CDbl(v) < 0 Then
        TexteJours = "dépassée de " & Format$(-CDbl(v), "0") & " j"
    Else
        TexteJours = "J-" & Format$(CDbl(v), "0")
    End If
End Function

' Dernière ligne de saisie d'un registre : ligne précédant la ligne TOTAL / SYNTHÈSE
' (les registres peuvent être agrandis par insertion de lignes sans modifier le code)
Public Function DerniereLigneRegistre(ByVal ws As Worksheet) As Long
    Dim r As Long, fin As Long
    Dim v As String
    With ws.UsedRange
        fin = .Row + .Rows.Count - 1
    End With
    For r = LIGNE_DEBUT To fin
        v = UCase$(TexteSur(ws.Cells(r, 1).Value))
        If v = "TOTAL" Or v = "SYNTHÈSE" Then
            DerniereLigneRegistre = r - 1
            Exit Function
        End If
    Next r
    DerniereLigneRegistre = fin
End Function

' Coupe un message trop long pour MsgBox (limite d'environ 1 024 caractères)
Public Function MessageTronque(ByVal s As String) As String
    If Len(s) > LONGUEUR_MAX_MSG Then
        MessageTronque = Left$(s, LONGUEUR_MAX_MSG) & vbCrLf & "  (... liste tronquée, voir l'onglet concerné)"
    Else
        MessageTronque = s
    End If
End Function

' Dossier d'export : dossier du classeur s'il est enregistré localement, sinon dossier par défaut
Public Function DossierExport() As String
    Dim d As String
    d = ThisWorkbook.Path
    If Len(d) = 0 Or LCase$(Left$(d, 4)) = "http" Then d = Application.DefaultFilePath
    If Len(d) = 0 Or LCase$(Left$(d, 4)) = "http" Then d = CurDir$
    If Right$(d, 1) = Application.PathSeparator Then d = Left$(d, Len(d) - 1)
    DossierExport = d
End Function

' --- 2. ACTUALISATION GLOBALE DU COCKPIT EXÉCUTIF DG (Raccourci : Ctrl+Shift+A) ---
Public Sub Actualiser_Cockpit_Global()
    Dim tStart As Double
    Dim wsDash As Worksheet
    Dim msgErr As String
    tStart = Timer
    On Error GoTo ErrHandler

    Call Activer_Mode_Turbo

    ' Recalcul intégral des formules multi-feuilles et des tableaux de bord
    Application.CalculateFullRebuild

    Set wsDash = ThisWorkbook.Worksheets("00_DASHBOARD")
    ThisWorkbook.Activate
    wsDash.Activate

    Call Desactiver_Mode_Turbo
    MsgBox "Cockpit Exécutif DG actualisé avec succès en " & Format$((Timer - tStart) * 1000, "0.0") & " ms." & vbCrLf & vbCrLf & _
           "• Avancement global (0-12 mois) : " & FormatSur(wsDash.Range("A5").Value, "0.0%") & vbCrLf & _
           "• Dossiers DLM en Escalade DG : " & FormatSur(wsDash.Range("D5").Value, "0") & vbCrLf & _
           "• Garanties en alerte (<= 60j) : " & FormatSur(wsDash.Range("G5").Value, "0") & vbCrLf & _
           "• Volume Factures / Décomptes : " & FormatSur(wsDash.Range("J5").Value, "#,##0") & " FCFA", _
           vbInformation, "Cockpit DG — Haute Performance"
    Exit Sub

ErrHandler:
    msgErr = Err.Description & " (code " & Err.Number & ")"
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur lors de l'actualisation : " & msgErr, vbCritical, "Erreur VBA"
End Sub

' --- 3. CONTRÔLE DES DLM, DÉPÔT IMMÉDIAT ADFD & ESCALADE DG (Raccourci : Ctrl+Shift+D) ---
' L'escalade reprend la colonne « Escalade DG [Auto] » (même règle que le tableau de bord :
' bailleur ADFD ou DLM <= 45 jours). La « Date Dernière Action » saisie n'est plus écrasée.
Public Sub Verifier_DLM_Et_Escalade_DG()
    Dim wsDLM As Worksheet
    Dim arrData As Variant
    Dim i As Long, derniere As Long, nbEscalade As Long, nbADFD As Long
    Dim msgEscalade As String, bailleur As String, msgErr As String

    On Error GoTo ErrHandler
    Set wsDLM = ThisWorkbook.Worksheets("02_SUIVI_DLM")
    Call Activer_Mode_Turbo
    wsDLM.Calculate

    derniere = DerniereLigneRegistre(wsDLM)
    If derniere >= LIGNE_DEBUT Then
        ' Lecture en mémoire (Variant Array) pour une exécution instantanée
        arrData = wsDLM.Range(wsDLM.Cells(LIGNE_DEBUT, 1), wsDLM.Cells(derniere, 12)).Value
        For i = 1 To UBound(arrData, 1)
            bailleur = UCase$(TexteSur(arrData(i, 2)))
            If Len(bailleur) > 0 Then
                If InStr(1, UCase$(TexteSur(arrData(i, 11))), "ESCALADE") > 0 Then
                    nbEscalade = nbEscalade + 1
                    If bailleur = "ADFD" Then nbADFD = nbADFD + 1
                    msgEscalade = msgEscalade & "  - [" & TexteSur(arrData(i, 1)) & " | " & bailleur & "] " & _
                                  Left$(TexteSur(arrData(i, 3)), 38) & " (" & TexteJours(arrData(i, 6)) & _
                                  " | Pièce : " & TexteSur(arrData(i, 7)) & ")" & vbCrLf
                End If
            End If
        Next i
    End If

    Call Desactiver_Mode_Turbo

    If nbEscalade > 0 Then
        MsgBox MessageTronque("ALERTE DLM & ESCALADE VERS LE DG (" & nbEscalade & " dossier(s) critique(s), dont " & _
               nbADFD & " ADFD) :" & vbCrLf & vbCrLf & _
               "Action immédiate requise (dont Dépôt immédiat ADFD) :" & vbCrLf & msgEscalade), _
               vbExclamation, "Escalade Direction Générale — Suivi DLM (BID, BOAD, FVC, ADFD)"
    Else
        MsgBox "Aucun dossier DLM en escalade : aucune échéance à 45 jours ou moins et aucun dossier ADFD.", _
               vbInformation, "Suivi DLM"
    End If
    Exit Sub

ErrHandler:
    msgErr = Err.Description & " (code " & Err.Number & ")"
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur dans Verifier_DLM_Et_Escalade_DG : " & msgErr, vbCritical, "Erreur VBA"
End Sub

' --- 4. CONTRÔLE DES GARANTIES À 60 JOURS AVANT ÉCHÉANCE (Raccourci : Ctrl+Shift+G) ---
Public Sub Controler_Garanties_60Jours()
    Dim wsGar As Worksheet
    Dim arrGar As Variant
    Dim i As Long, r As Long, derniere As Long, nbAlertes As Long, nbSansDate As Long
    Dim msgAlertes As String, msgSansDate As String, statut As String, msgErr As String
    Dim montantAlerte As Double

    On Error GoTo ErrHandler
    Set wsGar = ThisWorkbook.Worksheets("03_REGISTRE_GARANTIES")
    Call Activer_Mode_Turbo

    derniere = DerniereLigneRegistre(wsGar)
    If derniere >= LIGNE_DEBUT Then
        ' Horodate le contrôle (colonne G) sur toutes les lignes actives
        For r = LIGNE_DEBUT To derniere
            If Len(TexteSur(wsGar.Cells(r, 2).Value)) > 0 Then
                wsGar.Cells(r, 7).Value = Date
                wsGar.Cells(r, 7).NumberFormat = "dd/mm/yyyy"
            End If
        Next r
        wsGar.Calculate

        arrGar = wsGar.Range(wsGar.Cells(LIGNE_DEBUT, 1), wsGar.Cells(derniere, 12)).Value
        For i = 1 To UBound(arrGar, 1)
            If Len(TexteSur(arrGar(i, 2))) > 0 Then
                statut = UCase$(TexteSur(arrGar(i, 10)))
                If Not EstNombre(arrGar(i, 9)) Then
                    nbSansDate = nbSansDate + 1
                    msgSansDate = msgSansDate & "  - [" & TexteSur(arrGar(i, 1)) & "] " & TexteSur(arrGar(i, 2)) & vbCrLf
                ElseIf InStr(statut, "ALERTE") > 0 Or InStr(statut, "EXPIR") > 0 Then
                    nbAlertes = nbAlertes + 1
                    montantAlerte = montantAlerte + NombreSur(arrGar(i, 6), 0)
                    msgAlertes = msgAlertes & "  - [" & TexteSur(arrGar(i, 1)) & " | " & TexteSur(arrGar(i, 4)) & "] " & _
                                 TexteSur(arrGar(i, 2)) & " -> " & TexteJours(arrGar(i, 9)) & _
                                 " (" & FormatSur(arrGar(i, 6), "#,##0") & " FCFA)" & vbCrLf
                End If
            End If
        Next i
    End If

    Call Desactiver_Mode_Turbo

    If nbSansDate > 0 Then
        msgSansDate = vbCrLf & nbSansDate & " garantie(s) sans date d'échéance exploitable (à compléter) :" & vbCrLf & msgSansDate
    End If
    If nbAlertes > 0 Then
        MsgBox MessageTronque("ALERTE GARANTIES À 60 JOURS : " & nbAlertes & " garantie(s)/caution(s) à renouveler immédiatement !" & vbCrLf & _
               "Montant total sous alerte : " & Format$(montantAlerte, "#,##0") & " FCFA" & vbCrLf & vbCrLf & _
               msgAlertes & msgSansDate), vbExclamation, "Registre des Garanties — Seuil 60 Jours"
    Else
        MsgBox MessageTronque("Aucune garantie de bonne exécution ou caution d'avance n'arrive à échéance dans les 60 jours." & _
               vbCrLf & msgSansDate), IIf(nbSansDate > 0, vbExclamation, vbInformation), "Registre des Garanties"
    End If
    Exit Sub

ErrHandler:
    msgErr = Err.Description & " (code " & Err.Number & ")"
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur dans Controler_Garanties_60Jours : " & msgErr, vbCritical, "Erreur VBA"
End Sub

' --- 5. PLAN DE PAIEMENT PRIORISÉ PAR ÉCHÉANCE DE DLM AVEC AC & DFP (Raccourci : Ctrl+Shift+P) ---
' Tri en mémoire des seules lignes renseignées (les lignes vides restent en fin de registre) :
'   1. Priorité DLM croissante (P1 -> P2 -> P3 -> P4)
'   2. Date d'échéance DLM croissante (colonne H)
'   3. Montant décroissant (colonne D)
' Seules les colonnes de saisie sont déplacées ; les formules [Auto] (I, J, K) restent en place.
Public Sub Generer_Plan_Paiement_Priorise_AC_DFP()
    Dim wsFac As Worksheet
    Dim arr As Variant, sortieAH() As Variant, sortieLM() As Variant
    Dim ordre() As Long
    Dim i As Long, j As Long, c As Long, n As Long, derniere As Long, tmp As Long
    Dim msgErr As String

    On Error GoTo ErrHandler
    Set wsFac = ThisWorkbook.Worksheets("04_FACTURES_DECOMPTES")
    Call Activer_Mode_Turbo
    wsFac.Calculate

    derniere = DerniereLigneRegistre(wsFac)
    n = derniere - LIGNE_DEBUT + 1
    If n >= 2 Then
        arr = wsFac.Range(wsFac.Cells(LIGNE_DEBUT, 1), wsFac.Cells(derniere, 13)).Value
        ReDim ordre(1 To n)
        For i = 1 To n
            ordre(i) = i
        Next i
        ' Tri par insertion (stable : l'ordre de saisie est conservé à égalité)
        For i = 2 To n
            tmp = ordre(i)
            j = i - 1
            Do While j >= 1
                If PrecedeFacture(arr, tmp, ordre(j)) Then
                    ordre(j + 1) = ordre(j)
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            ordre(j + 1) = tmp
        Next i

        ReDim sortieAH(1 To n, 1 To 8)
        ReDim sortieLM(1 To n, 1 To 2)
        For i = 1 To n
            For c = 1 To 8
                sortieAH(i, c) = arr(ordre(i), c)
            Next c
            sortieLM(i, 1) = arr(ordre(i), 12)
            sortieLM(i, 2) = arr(ordre(i), 13)
        Next i
        ' Effacement préalable des seules valeurs (formats conservés) : les cellules vides
        ' du tableau trié sont ainsi réellement vides après réécriture
        wsFac.Range(wsFac.Cells(LIGNE_DEBUT, 1), wsFac.Cells(derniere, 8)).ClearContents
        wsFac.Range(wsFac.Cells(LIGNE_DEBUT, 12), wsFac.Cells(derniere, 13)).ClearContents
        wsFac.Range(wsFac.Cells(LIGNE_DEBUT, 1), wsFac.Cells(derniere, 8)).Value = sortieAH
        wsFac.Range(wsFac.Cells(LIGNE_DEBUT, 12), wsFac.Cells(derniere, 13)).Value = sortieLM
    End If

    wsFac.Calculate
    Call Desactiver_Mode_Turbo
    ThisWorkbook.Activate
    wsFac.Activate
    MsgBox "Le Plan de Paiement a été priorisé avec succès par échéance de DLM pour arbitrage conjoint AC & DFP." & vbCrLf & vbCrLf & _
           "• Urgences P1 (DLM <= 45j) : " & FormatSur(wsFac.Range("D5").Value, "#,##0") & " FCFA" & vbCrLf & _
           "• Dossiers en instance AC/DFP : " & FormatSur(wsFac.Range("G5").Value, "0"), _
           vbInformation, "Plan de Paiement Priorisé (AC & DFP)"
    Exit Sub

ErrHandler:
    msgErr = Err.Description & " (code " & Err.Number & ")"
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur lors du tri du plan de paiement : " & msgErr, vbCritical, "Erreur VBA"
End Sub

' Rang de priorité d'une facture (1 à 4 ; 9 si inconnu ; 99 si ligne vide)
Public Function RangPriorite(ByRef arr As Variant, ByVal r As Long) As Long
    Dim k As String
    If Len(TexteSur(arr(r, 2))) = 0 Then
        RangPriorite = 99
        Exit Function
    End If
    k = UCase$(TexteSur(arr(r, 11)))
    If Left$(k, 2) Like "P#" Then
        RangPriorite = CLng(Mid$(k, 2, 1))
    Else
        RangPriorite = 9
    End If
End Function

' Vrai si la facture a doit précéder la facture b dans le plan de paiement
Public Function PrecedeFacture(ByRef arr As Variant, ByVal a As Long, ByVal b As Long) As Boolean
    Dim pa As Long, pb As Long
    Dim da As Double, db As Double
    pa = RangPriorite(arr, a)
    pb = RangPriorite(arr, b)
    If pa <> pb Then
        PrecedeFacture = (pa < pb)
        Exit Function
    End If
    da = NombreSur(arr(a, 8), 1E+15)
    db = NombreSur(arr(b, 8), 1E+15)
    If da <> db Then
        PrecedeFacture = (da < db)
        Exit Function
    End If
    PrecedeFacture = (NombreSur(arr(a, 4), 0) > NombreSur(arr(b, 4), 0))
End Function

' --- 6. FILTRAGE RAPIDE DU PLAN D'AMÉLIORATION PAR HORIZON (Raccourci : Ctrl+Shift+H) ---
Public Sub Filtrer_Plan_Amelioration_Par_Horizon()
    Dim wsPlan As Worksheet
    Dim plage As Range
    Dim choix As String, msgErr As String
    Dim derniere As Long

    On Error GoTo ErrHandler
    Set wsPlan = ThisWorkbook.Worksheets("01_PLAN_AMELIORATION")
    ThisWorkbook.Activate
    wsPlan.Activate

    choix = InputBox("Entrez l'horizon à filtrer :" & vbCrLf & _
                     "  1 = 0-30 jours (Sécuriser les financements)" & vbCrLf & _
                     "  2 = 1-3 mois (Gouvernance, source unique, RACI, entreprises, fiscalité)" & vbCrLf & _
                     "  3 = 3-12 mois (Maturité, SIG, DAO standardisés, transfert exploitants)" & vbCrLf & _
                     "  0 = Afficher tous les horizons", "Filtre Plan d'Amélioration", "0")
    choix = Trim$(choix)
    If choix = "" Then Exit Sub
    If choix <> "0" And choix <> "1" And choix <> "2" And choix <> "3" Then
        MsgBox "Choix « " & choix & " » non reconnu : saisissez 0, 1, 2 ou 3.", vbExclamation, "Filtre Plan d'Amélioration"
        Exit Sub
    End If

    derniere = DerniereLigneRegistre(wsPlan)
    If derniere < LIGNE_DEBUT Then derniere = LIGNE_DEBUT
    Set plage = wsPlan.Range(wsPlan.Cells(LIGNE_ENTETE, 1), wsPlan.Cells(derniere, 12))

    Call Activer_Mode_Turbo
    If wsPlan.AutoFilterMode Then wsPlan.AutoFilterMode = False
    Select Case choix
        Case "1": plage.AutoFilter Field:=2, Criteria1:="0-30 jours"
        Case "2": plage.AutoFilter Field:=2, Criteria1:="1-3 mois"
        Case "3": plage.AutoFilter Field:=2, Criteria1:="3-12 mois"
        Case Else: plage.AutoFilter
    End Select
    Call Desactiver_Mode_Turbo
    Exit Sub

ErrHandler:
    msgErr = Err.Description & " (code " & Err.Number & ")"
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur dans Filtrer_Plan_Amelioration_Par_Horizon : " & msgErr, vbCritical, "Erreur VBA"
End Sub

' --- 7. AUDIT QUALITÉ DU PORTEFEUILLE EN MÉMOIRE (Raccourci : Ctrl+Shift+Q) ---
' Une date saisie en texte (ex. « 15/11/2026 » mal reconnu) est signalée : elle rend les
' formules [Auto] inopérantes.
Public Sub Audit_Qualite_Portefeuille_Complet()
    Dim wsDLM As Worksheet, wsGar As Worksheet, wsFac As Worksheet
    Dim anomalies As Long, rapport As String, msgErr As String
    Dim i As Long, derniere As Long

    On Error GoTo ErrHandler
    Set wsDLM = ThisWorkbook.Worksheets("02_SUIVI_DLM")
    Set wsGar = ThisWorkbook.Worksheets("03_REGISTRE_GARANTIES")
    Set wsFac = ThisWorkbook.Worksheets("04_FACTURES_DECOMPTES")
    Call Activer_Mode_Turbo

    ' Vérification DLM sans OWNER ou sans échéance (date réelle)
    derniere = DerniereLigneRegistre(wsDLM)
    For i = LIGNE_DEBUT To derniere
        If Len(TexteSur(wsDLM.Cells(i, 2).Value)) > 0 Then
            If Not EstDate(wsDLM.Cells(i, 5).Value) Or Len(TexteSur(wsDLM.Cells(i, 8).Value)) = 0 Then
                anomalies = anomalies + 1
                rapport = rapport & "• [02_SUIVI_DLM Ligne " & i & "] Échéance (date) ou OWNER manquant." & vbCrLf
            End If
        End If
    Next i

    ' Vérification Garanties sans montant ou date d'échéance
    derniere = DerniereLigneRegistre(wsGar)
    For i = LIGNE_DEBUT To derniere
        If Len(TexteSur(wsGar.Cells(i, 2).Value)) > 0 Then
            If NombreSur(wsGar.Cells(i, 6).Value, 0) <= 0 Or Not EstDate(wsGar.Cells(i, 8).Value) Then
                anomalies = anomalies + 1
                rapport = rapport & "• [03_REGISTRE_GARANTIES Ligne " & i & "] Montant ou Échéance (date) invalide." & vbCrLf
            End If
        End If
    Next i

    ' Vérification Factures sans échéance DLM associée
    derniere = DerniereLigneRegistre(wsFac)
    For i = LIGNE_DEBUT To derniere
        If Len(TexteSur(wsFac.Cells(i, 2).Value)) > 0 Then
            If Not EstDate(wsFac.Cells(i, 8).Value) Then
                anomalies = anomalies + 1
                rapport = rapport & "• [04_FACTURES_DECOMPTES Ligne " & i & "] Date d'échéance DLM manquante pour priorisation." & vbCrLf
            End If
        End If
    Next i

    Call Desactiver_Mode_Turbo
    If anomalies = 0 Then
        MsgBox "Audit de cohérence terminé : 0 anomalie détectée sur les 4 outils.", vbInformation, "Audit Qualité Portefeuille"
    Else
        MsgBox MessageTronque(anomalies & " anomalie(s) détectée(s) :" & vbCrLf & vbCrLf & rapport), vbExclamation, "Audit Qualité Portefeuille"
    End If
    Exit Sub

ErrHandler:
    msgErr = Err.Description & " (code " & Err.Number & ")"
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur dans Audit_Qualite_Portefeuille_Complet : " & msgErr, vbCritical, "Erreur VBA"
End Sub

' --- 8. EXPORT AUTOMATIQUE DU RAPPORT EXÉCUTIF EN PDF (Raccourci : Ctrl+Shift+E) ---
Public Sub Exporter_Rapport_Executif_PDF()
    Dim pdfPath As String, msgErr As String
    On Error GoTo ErrHandler

    Call Activer_Mode_Turbo
    pdfPath = DossierExport() & Application.PathSeparator & "Rapport_Pilotage_DG_" & Format$(Date, "yyyymmdd") & ".pdf"

    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(Array("00_DASHBOARD", "01_PLAN_AMELIORATION", "02_SUIVI_DLM", "03_REGISTRE_GARANTIES", "04_FACTURES_DECOMPTES")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    ThisWorkbook.Worksheets("00_DASHBOARD").Select

    Call Desactiver_Mode_Turbo
    MsgBox "Rapport Exécutif exporté avec succès :" & vbCrLf & pdfPath, vbInformation, "Export PDF"
    Exit Sub

ErrHandler:
    msgErr = Err.Description & " (code " & Err.Number & ")"
    On Error Resume Next
    ThisWorkbook.Worksheets("00_DASHBOARD").Select
    On Error GoTo 0
    Call Desactiver_Mode_Turbo
    MsgBox "Erreur lors de l'export PDF : " & msgErr & vbCrLf & vbCrLf & _
           "(Vérifiez qu'un PDF du même nom n'est pas déjà ouvert.)", vbCritical, "Erreur VBA"
End Sub
