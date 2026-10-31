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
    Call ActiverModeMacro(False)
    Call EcrireJournal("Actualiser", "ERREUR : " & Err.Description & " (Code " & Err.Number & ")")
    MsgBox "Erreur lors de l'actualisation : " & vbCrLf & Err.Description, vbCritical, "Erreur VBA - ASER"
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
    Call ActiverModeMacro(False)
    Call EcrireJournal("GenererPlanPaiement", "ERREUR : " & Err.Description & " (Code " & Err.Number & ")")
    MsgBox "Erreur lors de la génération du Plan de Paiement : " & vbCrLf & Err.Description, vbCritical, "Erreur VBA - ASER"
End Sub

'-------------------------------------------------------------------------------
' Procédure interne appelée par GenererPlanPaiement, Actualiser et ArchiverClotures
'-------------------------------------------------------------------------------
Public Sub GenererPlanPaiement_Interne(Optional ByVal journaliserSeul As Boolean = True)
    Dim wsFac As Worksheet
    Dim wsPP As Worksheet
    Dim loFac As ListObject
    Dim loPP As ListObject
    Dim dictSaisies As Object
    Dim i As Long, j As Long, nFactures As Long
    Dim cleFacture As String
    Dim arrSaisie(1 To 4) As Variant
    Dim vRecup As Variant
    
    Set wsFac = ThisWorkbook.Worksheets("Factures_Decomptes")
    Set wsPP = ThisWorkbook.Worksheets("Plan_Paiement")
    Set loFac = wsFac.ListObjects("T_Factures")
    Set loPP = wsPP.ListObjects("T_PlanPaiement")
    
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
            cleFacture = Trim$(CStr(loPP.DataBodyRange.Cells(i, idxPPNum).Value & ""))
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
    If Not loFac.DataBodyRange Is Nothing Then
        Dim idxFNum As Long, idxFEnt As Long, idxFBail As Long, idxFDLM As Long
        Dim idxFJoursDLM As Long, idxFReste As Long, idxFRetard As Long, idxFEtape As Long
        
        idxFNum = loFac.ListColumns("N° facture/décompte").Index
        idxFEnt = loFac.ListColumns("Entreprise").Index
        idxFBail = loFac.ListColumns("Bailleur").Index
        idxFDLM = loFac.ListColumns("DLM liée").Index
        idxFJoursDLM = loFac.ListColumns("Jours avant DLM").Index
        idxFReste = loFac.ListColumns("Reste à payer").Index
        idxFRetard = loFac.ListColumns("Retard contractuel (j)").Index
        idxFEtape = loFac.ListColumns("Étape en cours").Index
        
        Dim candidats() As Variant
        ReDim candidats(1 To 8, 1 To loFac.ListRows.Count)
        nFactures = 0
        
        For i = 1 To loFac.ListRows.Count
            Dim etapeStr As String
            Dim resteVal As Double
            Dim numStr As String
            
            numStr = Trim$(CStr(loFac.DataBodyRange.Cells(i, idxFNum).Value & ""))
            etapeStr = Trim$(CStr(loFac.DataBodyRange.Cells(i, idxFEtape).Value & ""))
            resteVal = Val(Replace(CStr(loFac.DataBodyRange.Cells(i, idxFReste).Value & "0"), ",", "."))
            If IsNumeric(loFac.DataBodyRange.Cells(i, idxFReste).Value) Then
                resteVal = CDbl(loFac.DataBodyRange.Cells(i, idxFReste).Value)
            End If
            
            If Len(numStr) > 0 And StrComp(etapeStr, "Payée", vbTextCompare) <> 0 And resteVal > 0 Then
                nFactures = nFactures + 1
                candidats(1, nFactures) = numStr
                candidats(2, nFactures) = loFac.DataBodyRange.Cells(i, idxFEnt).Value
                candidats(3, nFactures) = loFac.DataBodyRange.Cells(i, idxFBail).Value
                candidats(4, nFactures) = loFac.DataBodyRange.Cells(i, idxFDLM).Value
                candidats(5, nFactures) = IIf(IsNumeric(loFac.DataBodyRange.Cells(i, idxFJoursDLM).Value), _
                                              CDbl(loFac.DataBodyRange.Cells(i, idxFJoursDLM).Value), 9999)
                candidats(6, nFactures) = resteVal
                candidats(7, nFactures) = IIf(IsNumeric(loFac.DataBodyRange.Cells(i, idxFRetard).Value), _
                                              CDbl(loFac.DataBodyRange.Cells(i, idxFRetard).Value), 0)
            End If
        Next i
        
        ' 3. Tri multi-critères en mémoire (compatible toutes versions Excel < 2021) :
        '    Clé 1 : Jours avant DLM (croissant)
        '    Clé 2 : Retard contractuel (décroissant)
        '    Clé 3 : Reste à payer (décroissant)
        If nFactures > 1 Then
            Dim k As Long
            Dim doitEchanger As Boolean
            Dim tmpVal As Variant
            
            For i = 1 To nFactures - 1
                For j = i + 1 To nFactures
                    doitEchanger = False
                    If candidats(5, j) < candidats(5, i) Then
                        doitEchanger = True
                    ElseIf candidats(5, j) = candidats(5, i) Then
                        If candidats(7, j) > candidats(7, i) Then
                            doitEchanger = True
                        ElseIf candidats(7, j) = candidats(7, i) Then
                            If candidats(6, j) > candidats(6, i) Then
                                doitEchanger = True
                            End If
                        End If
                    End If
                    
                    If doitEchanger Then
                        For k = 1 To 7
                            tmpVal = candidats(k, i)
                            candidats(k, i) = candidats(k, j)
                            candidats(k, j) = tmpVal
                        Next k
                    End If
                Next j
            Next i
        End If
        
        ' 4. Insérer les lignes triées dans T_PlanPaiement et restaurer les saisies
        Dim cumulGlissant As Double
        cumulGlissant = 0
        
        For i = 1 To nFactures
            Dim lrNew As ListRow
            Set lrNew = loPP.ListRows.Add
            cleFacture = CStr(candidats(1, i))
            cumulGlissant = cumulGlissant + CDbl(candidats(6, i))
            
            With lrNew.Range
                .Cells(1, loPP.ListColumns("Rang").Index).Value = i
                .Cells(1, loPP.ListColumns("N° facture").Index).Value = cleFacture
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
                .Cells(1, loPP.ListColumns("Cumul").Index).Formula = _
                    "=SUM($G$6:G" & .Row & ")"
                
                ' Restaurer les saisies d'arbitrage si le N° facture existait déjà
                If dictSaisies.Exists(cleFacture) Then
                    vRecup = dictSaisies(cleFacture)
                    .Cells(1, loPP.ListColumns("Date de paiement proposée").Index).Value = vRecup(1)
                    .Cells(1, loPP.ListColumns("Validation AC (O/N)").Index).Value = vRecup(2)
                    .Cells(1, loPP.ListColumns("Validation DFP (O/N)").Index).Value = vRecup(3)
                    .Cells(1, loPP.ListColumns("Commentaire").Index).Value = vRecup(4)
                End If
                
                ' Déverrouiller les 4 colonnes de saisie
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
End Sub

'-------------------------------------------------------------------------------
' MACRO 7 : ArchiverClotures
' Déplace les garanties au statut "Mainlevée" et les factures au statut "Payée"
' vers l'onglet Archive après confirmation de l'utilisateur.
'-------------------------------------------------------------------------------
Public Sub ArchiverClotures()
    On Error GoTo ErrHandler
    
    Dim wsGar As Worksheet, wsFac As Worksheet, wsArc As Worksheet
    Dim loGar As ListObject, loFac As ListObject
    Dim loArcGar As ListObject, loArcFac As ListObject
    Dim dtRef As Date
    Dim i As Long, c As Long
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
    
    ' 1. Compter les lignes éligibles à l'archivage
    If Not loGar.DataBodyRange Is Nothing Then
        For i = 1 To loGar.ListRows.Count
            If StrComp(Trim$(CStr(loGar.DataBodyRange.Cells(i, idxStatGar).Value & "")), "Mainlevée", vbTextCompare) = 0 Then
                nbGarClot = nbGarClot + 1
            End If
        Next i
    End If
    
    If Not loFac.DataBodyRange Is Nothing Then
        For i = 1 To loFac.ListRows.Count
            If StrComp(Trim$(CStr(loFac.DataBodyRange.Cells(i, idxEtapeFac).Value & "")), "Payée", vbTextCompare) = 0 Then
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
            If StrComp(Trim$(CStr(loGar.DataBodyRange.Cells(i, idxStatGar).Value & "")), "Mainlevée", vbTextCompare) = 0 Then
                Dim lrAG As ListRow
                Set lrAG = loArcGar.ListRows.Add
                For c = 1 To 17
                    lrAG.Range.Cells(1, c).Value = loGar.DataBodyRange.Cells(i, c).Value
                Next c
                lrAG.Range.Cells(1, 18).Value = dtRef
                lrAG.Range.Cells(1, 18).NumberFormat = "dd/mm/yyyy"
                loGar.ListRows(i).Delete
            End If
        Next i
    End If
    
    ' 4. Déplacer les factures Payées (du bas vers le haut)
    If Not loFac.DataBodyRange Is Nothing Then
        For i = loFac.ListRows.Count To 1 Step -1
            If StrComp(Trim$(CStr(loFac.DataBodyRange.Cells(i, idxEtapeFac).Value & "")), "Payée", vbTextCompare) = 0 Then
                Dim lrAF As ListRow
                Set lrAF = loArcFac.ListRows.Add
                For c = 1 To 15
                    lrAF.Range.Cells(1, c).Value = loFac.DataBodyRange.Cells(i, c).Value
                Next c
                lrAF.Range.Cells(1, 16).Value = dtRef
                lrAF.Range.Cells(1, 16).NumberFormat = "dd/mm/yyyy"
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
    Call ActiverModeMacro(False)
    Call EcrireJournal("ArchiverClotures", "ERREUR : " & Err.Description & " (Code " & Err.Number & ")")
    MsgBox "Erreur lors de l'archivage : " & vbCrLf & Err.Description, vbCritical, "Erreur VBA - ASER"
End Sub

'-------------------------------------------------------------------------------
' MACRO 8 : NouvelleLigne
' Ajoute une ligne au tableau actif (ou au registre choisi par l'utilisateur si
' appelé depuis le Tableau de bord) avec un ID auto-incrémenté.
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
    Dim i As Long
    Dim valID As String
    
    Set wsCible = ActiveSheet
    
    ' Si l'utilisateur lance la macro depuis le Tableau de bord (sans ListObject), lui demander le registre cible
    If wsCible.ListObjects.Count = 0 Then
        choix = InputBox("Dans quel registre souhaitez-vous ajouter une nouvelle ligne ?" & vbCrLf & vbCrLf & _
                         "1 = DLM_Prorogations (T_DLM)" & vbCrLf & _
                         "2 = Garanties (T_Garanties)" & vbCrLf & _
                         "3 = Factures_Decomptes (T_Factures)" & vbCrLf & _
                         "4 = Plan_Amelioration (T_Plan)", _
                         "Ajout d'une nouvelle ligne - ASER", "1")
        Select Case Trim$(choix)
            Case "1": Set wsCible = ThisWorkbook.Worksheets("DLM_Prorogations")
            Case "2": Set wsCible = ThisWorkbook.Worksheets("Garanties")
            Case "3": Set wsCible = ThisWorkbook.Worksheets("Factures_Decomptes")
            Case "4": Set wsCible = ThisWorkbook.Worksheets("Plan_Amelioration")
            Case Else: Exit Sub
        End Select
    End If
    
    Set loCible = wsCible.ListObjects(1)
    
    Select Case loCible.Name
        Case "T_DLM": prefixeID = "DLM-"
        Case "T_Garanties": prefixeID = "GAR-"
        Case "T_Factures": prefixeID = "FAC-"
        Case "T_Plan": prefixeID = "PA-"
        Case Else: prefixeID = "LIG-"
    End Select
    
    Call ActiverModeMacro(True)
    Call DeprotegerFeuille(wsCible)
    
    ' Calculer le prochain numéro d'ID auto-incrémenté
    numMax = 0
    If Not loCible.DataBodyRange Is Nothing Then
        For i = 1 To loCible.ListRows.Count
            valID = Trim$(CStr(loCible.DataBodyRange.Cells(i, 1).Value & ""))
            Dim posTiret As Long
            posTiret = InStrRev(valID, "-")
            If posTiret > 0 Then
                Dim numCandidat As Long
                numCandidat = Val(Mid$(valID, posTiret + 1))
                If numCandidat > numMax Then numMax = numCandidat
            End If
        Next i
    End If
    nouvelID = prefixeID & Format$(numMax + 1, "000")
    
    Set lrNew = loCible.ListRows.Add
    lrNew.Range.Cells(1, 1).Value = nouvelID
    
    ' Initialisations par défaut selon le registre
    If loCible.Name = "T_Factures" Then
        lrNew.Range.Cells(1, loCible.ListColumns("Délai contractuel de paiement").Index).Formula = "=DelaiPaiementDefaut"
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
    
    MsgBox "Nouvelle ligne ajoutée dans « " & wsCible.Name & " » avec l'identifiant : " & nouvelID, _
           vbInformation, "Nouvelle Ligne - ASER"
    Exit Sub

ErrHandler:
    Call ActiverModeMacro(False)
    Call EcrireJournal("NouvelleLigne", "ERREUR : " & Err.Description & " (Code " & Err.Number & ")")
    MsgBox "Erreur lors de l'ajout d'une nouvelle ligne : " & vbCrLf & Err.Description, vbCritical, "Erreur VBA - ASER"
End Sub

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
                 vbQuestion + vbYesNo, "Suppression des exemples - ASER")
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
                            valCell = Trim$(CStr(lo.DataBodyRange.Cells(i, c).Value & ""))
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
    Call ActiverModeMacro(False)
    Call EcrireJournal("SupprimerExemples", "ERREUR : " & Err.Description & " (Code " & Err.Number & ")")
    MsgBox "Erreur lors de la suppression des exemples : " & vbCrLf & Err.Description, vbCritical, "Erreur VBA - ASER"
End Sub
