Attribute VB_Name = "Mod_Communication_Reporting"
'===============================================================================
' MODULE       : Mod_Communication_Reporting
' CLASSEUR     : Pilotage_Financements_ASER.xlsm
' RÔLE         : Macros de communication Outlook et d'export PDF :
'                4. PreparerRelances (brouillons Outlook groupés par détenteur,
'                   mode .Display uniquement, mise à jour après confirmation)
'                5. NoteEscaladeDG (peuplement T_EscaladeDG + export PDF)
'                6. RapportMensuel (export PDF consolidé du portefeuille)
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' MACRO 4 : PreparerRelances
' Génère un brouillon Outlook par détenteur regroupant ses pièces manquantes
' et relances dues (Prochaine relance <= DateRef).
' Aucun envoi automatique (.Display uniquement).
' Date dernière relance et Nb relances ne sont mis à jour qu'après confirmation.
'-------------------------------------------------------------------------------
Public Sub PreparerRelances()
    On Error GoTo ErrHandler
    
    Dim wsDLM As Worksheet
    Dim loDLM As ListObject
    Dim dtRef As Date
    Dim dictCorps As Object
    Dim dictLignes As Object
    Dim i As Long
    
    Dim idxAccord As Long, idxProjet As Long, idxBailleur As Long
    Dim idxDLMVig As Long, idxJoursRest As Long, idxPiece As Long
    Dim idxDetenteur As Long, idxDernRel As Long, idxNbRel As Long, idxProchRel As Long
    
    Set wsDLM = ThisWorkbook.Worksheets("DLM_Prorogations")
    Set loDLM = wsDLM.ListObjects("T_DLM")
    
    If loDLM.DataBodyRange Is Nothing Then
        MsgBox "Le registre DLM_Prorogations est vide.", vbInformation, "Relances Outlook - ASER"
        Exit Sub
    End If
    
    dtRef = ObtenirDateRef()
    Set dictCorps = CreateObject("Scripting.Dictionary")
    dictCorps.CompareMode = 1 ' vbTextCompare
    Set dictLignes = CreateObject("Scripting.Dictionary")
    dictLignes.CompareMode = 1
    
    idxBailleur = loDLM.ListColumns("Bailleur").Index
    idxAccord = loDLM.ListColumns("N° accord").Index
    idxProjet = loDLM.ListColumns("Projet").Index
    idxDLMVig = loDLM.ListColumns("DLM en vigueur").Index
    idxJoursRest = loDLM.ListColumns("Jours restants").Index
    idxPiece = loDLM.ListColumns("Pièce manquante").Index
    idxDetenteur = loDLM.ListColumns("Détenteur de la pièce").Index
    idxDernRel = loDLM.ListColumns("Date dernière relance").Index
    idxNbRel = loDLM.ListColumns("Nb relances").Index
    idxProchRel = loDLM.ListColumns("Prochaine relance").Index
    
    ' 1. Identifier les dossiers ayant une pièce manquante et une relance due (<= DateRef)
    Dim nbDossiersRelance As Long
    nbDossiersRelance = 0
    
    For i = 1 To loDLM.ListRows.Count
        Dim pieceStr As String
        Dim detStr As String
        Dim vProch As Variant
        Dim relanceDue As Boolean
        
        pieceStr = Trim$(CStr(loDLM.DataBodyRange.Cells(i, idxPiece).Value & ""))
        detStr = Trim$(CStr(loDLM.DataBodyRange.Cells(i, idxDetenteur).Value & ""))
        vProch = loDLM.DataBodyRange.Cells(i, idxProchRel).Value
        
        relanceDue = False
        If Len(pieceStr) > 0 And Len(detStr) > 0 Then
            If IsDate(vProch) Then
                If CDate(vProch) <= dtRef Then relanceDue = True
            ElseIf IsNumeric(vProch) Then
                If CDbl(vProch) > 0 And CDbl(vProch) <= CDbl(dtRef) Then relanceDue = True
            ElseIf Len(Trim$(CStr(loDLM.DataBodyRange.Cells(i, idxDernRel).Value & ""))) = 0 Then
                relanceDue = True
            End If
        End If
        
        If relanceDue Then
            nbDossiersRelance = nbDossiersRelance + 1
            Dim ligneHTML As String
            ligneHTML = "<tr>" & _
                        "<td style='padding:6px;border:1px solid #cbd5e1;'>" & CStr(loDLM.DataBodyRange.Cells(i, idxBailleur).Value & "") & "</td>" & _
                        "<td style='padding:6px;border:1px solid #cbd5e1;'><b>" & CStr(loDLM.DataBodyRange.Cells(i, idxAccord).Value & "") & "</b></td>" & _
                        "<td style='padding:6px;border:1px solid #cbd5e1;'>" & CStr(loDLM.DataBodyRange.Cells(i, idxProjet).Value & "") & "</td>" & _
                        "<td style='padding:6px;border:1px solid #cbd5e1;text-align:center;'>" & Format$(loDLM.DataBodyRange.Cells(i, idxDLMVig).Value, "dd/mm/yyyy") & "</td>" & _
                        "<td style='padding:6px;border:1px solid #cbd5e1;text-align:right;color:#c00000;'><b>" & CStr(loDLM.DataBodyRange.Cells(i, idxJoursRest).Value & "") & " j</b></td>" & _
                        "<td style='padding:6px;border:1px solid #cbd5e1;color:#b45309;'><b>" & pieceStr & "</b></td>" & _
                        "</tr>"
            
            If Not dictCorps.Exists(detStr) Then
                dictCorps(detStr) = ligneHTML
                dictLignes(detStr) = CStr(i)
            Else
                dictCorps(detStr) = dictCorps(detStr) & ligneHTML
                dictLignes(detStr) = dictLignes(detStr) & ";" & CStr(i)
            End If
        End If
    Next i
    
    If nbDossiersRelance = 0 Then
        MsgBox "Aucune relance de pièce n'est arrivée à échéance à la date de référence (" & Format$(dtRef, "dd/mm/yyyy") & ").", _
               vbInformation, "Relances Outlook - ASER"
        Exit Sub
    End If
    
    ' 2. Créer un brouillon Outlook par détenteur en liaison tardive (.Display uniquement)
    Dim olApp As Object
    Dim olMail As Object
    Dim outlookDispo As Boolean
    Dim cleDet As Variant
    Dim emailDest As String
    Dim corpsComplet As String
    
    On Error Resume Next
    Set olApp = CreateObject("Outlook.Application")
    outlookDispo = Not (olApp Is Nothing)
    On Error GoTo ErrHandler
    
    For Each cleDet In dictCorps.Keys
        emailDest = ObtenirEmailDetenteur(CStr(cleDet))
        corpsComplet = "<html><body style='font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#0f172a;'>" & _
                       "<p>Bonjour <b>" & CStr(cleDet) & "</b>,</p>" & _
                       "<p>Dans le cadre du suivi des Dates Limites de Décaissement (DLM) et des demandes de prorogation des conventions de financement de l'ASER, " & _
                       "nous vous remercions de bien vouloir nous transmettre en urgence la ou les pièce(s) listée(s) ci-dessous :</p>" & _
                       "<table style='border-collapse:collapse;width:100%;font-size:10pt;'>" & _
                       "<tr style='background-color:#1b365d;color:#ffffff;'>" & _
                       "<th style='padding:6px;border:1px solid #cbd5e1;'>Bailleur</th>" & _
                       "<th style='padding:6px;border:1px solid #cbd5e1;'>N° Accord</th>" & _
                       "<th style='padding:6px;border:1px solid #cbd5e1;'>Projet</th>" & _
                       "<th style='padding:6px;border:1px solid #cbd5e1;'>DLM en vigueur</th>" & _
                       "<th style='padding:6px;border:1px solid #cbd5e1;'>Jours restants</th>" & _
                       "<th style='padding:6px;border:1px solid #cbd5e1;'>Pièce manquante attendue</th></tr>" & _
                       dictCorps(cleDet) & "</table>" & _
                       "<p>Restant à votre disposition pour tout arbitrage, nous vous prions d'agréer l'expression de nos salutations distinguées.</p>" & _
                       "<p><b>Contrôle de Gestion & Pilotage des Financements — ASER</b></p></body></html>"
        
        If outlookDispo Then
            Set olMail = olApp.CreateItem(0) ' olMailItem = 0
            With olMail
                .To = emailDest
                .Subject = "[URGENT - ASER] Relance pièces justificatives DLM / Prorogation (" & Format$(dtRef, "dd/mm/yyyy") & ")"
                .HTMLBody = corpsComplet
                .Display ' STRICTEMENT .Display, jamais .Send
            End With
        End If
    Next cleDet
    
    If Not outlookDispo Then
        MsgBox "Microsoft Outlook n'a pas pu être ouvert sur ce poste (" & dictCorps.Count & " détenteur(s) identifié(s))." & vbCrLf & _
               "Vous pouvez néanmoins confirmer ci-après la mise à jour des compteurs de relance.", _
               vbExclamation, "Outlook indisponible - ASER"
    End If
    
    ' 3. Demander confirmation à l'utilisateur avant de mettre à jour Date dernière relance et Nb relances
    Dim rep As VbMsgBoxResult
    rep = MsgBox(dictCorps.Count & " brouillon(s) de relance couvrant " & nbDossiersRelance & " dossier(s) ont été préparés." & vbCrLf & vbCrLf & _
                 "Confirmez-vous la mise à jour de « Date dernière relance » (" & Format$(dtRef, "dd/mm/yyyy") & ") " & _
                 "et l'incrémentation de « Nb relances » (+1) dans DLM_Prorogations ?", _
                 vbQuestion + vbYesNo, "Confirmation de mise à jour des relances - ASER")
                 
    If rep = vbYes Then
        Call ActiverModeMacro(True)
        Call DeprotegerFeuille(wsDLM)
        
        For Each cleDet In dictLignes.Keys
            Dim indices() As String
            Dim idxStr As Variant
            Dim numLig As Long
            indices = Split(CStr(dictLignes(cleDet)), ";")
            For Each idxStr In indices
                numLig = CLng(idxStr)
                loDLM.DataBodyRange.Cells(numLig, idxDernRel).Value = dtRef
                loDLM.DataBodyRange.Cells(numLig, idxDernRel).NumberFormat = "dd/mm/yyyy"
                loDLM.DataBodyRange.Cells(numLig, idxNbRel).Value = _
                    Val(CStr(loDLM.DataBodyRange.Cells(numLig, idxNbRel).Value & "0")) + 1
            Next idxStr
        Next cleDet
        
        Call ProtegerFeuille(wsDLM)
        Application.Calculate
        Call EcrireJournal("PreparerRelances", dictCorps.Count & " brouillon(s) préparé(s) et " & nbDossiersRelance & " compteur(s) de relance mis à jour.")
        Call ActiverModeMacro(False)
    Else
        Call EcrireJournal("PreparerRelances", dictCorps.Count & " brouillon(s) affiché(s) sans mise à jour des compteurs (annulé par l'utilisateur).")
    End If
    Exit Sub

ErrHandler:
    Call ActiverModeMacro(False)
    Call EcrireJournal("PreparerRelances", "ERREUR : " & Err.Description & " (Code " & Err.Number & ")")
    MsgBox "Erreur lors de la préparation des relances : " & vbCrLf & Err.Description, vbCritical, "Erreur VBA - ASER"
End Sub

'-------------------------------------------------------------------------------
' MACRO 5 : NoteEscaladeDG
' Alimente l'onglet dédié Note_Escalade_DG (T_EscaladeDG) avec les accords où
' Escalade DG requise = "Oui" et exporte la note en PDF dans le dossier du classeur.
'-------------------------------------------------------------------------------
Public Sub NoteEscaladeDG()
    On Error GoTo ErrHandler
    
    Dim wsDLM As Worksheet
    Dim wsEsc As Worksheet
    Dim loDLM As ListObject
    Dim loEsc As ListObject
    Dim dtRef As Date
    Dim i As Long, nbEsc As Long
    Dim cheminPDF As String
    Dim dossierBase As String
    
    Call ActiverModeMacro(True)
    
    Set wsDLM = ThisWorkbook.Worksheets("DLM_Prorogations")
    Set wsEsc = ThisWorkbook.Worksheets("Note_Escalade_DG")
    Set loDLM = wsDLM.ListObjects("T_DLM")
    Set loEsc = wsEsc.ListObjects("T_EscaladeDG")
    dtRef = ObtenirDateRef()
    
    Call DeprotegerFeuille(wsEsc)
    Call DeprotegerFeuille(wsDLM)
    
    ' 1. Vider T_EscaladeDG
    If Not loEsc.DataBodyRange Is Nothing Then
        For i = loEsc.ListRows.Count To 1 Step -1
            loEsc.ListRows(i).Delete
        Next i
    End If
    
    ' 2. Parcourir T_DLM et extraire les lignes où Escalade DG requise = "Oui"
    nbEsc = 0
    If Not loDLM.DataBodyRange Is Nothing Then
        Dim idxEscReq As Long, idxDateEsc As Long
        idxEscReq = loDLM.ListColumns("Escalade DG requise").Index
        idxDateEsc = loDLM.ListColumns("Date escalade DG").Index
        
        For i = 1 To loDLM.ListRows.Count
            If StrComp(Trim$(CStr(loDLM.DataBodyRange.Cells(i, idxEscReq).Value & "")), "Oui", vbTextCompare) = 0 Then
                nbEsc = nbEsc + 1
                Dim lrE As ListRow
                Set lrE = loEsc.ListRows.Add
                With lrE.Range
                    .Cells(1, loEsc.ListColumns("Bailleur").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("Bailleur").Index).Value
                    .Cells(1, loEsc.ListColumns("N° accord").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("N° accord").Index).Value
                    .Cells(1, loEsc.ListColumns("Projet").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("Projet").Index).Value
                    .Cells(1, loEsc.ListColumns("DLM en vigueur").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("DLM en vigueur").Index).Value
                    .Cells(1, loEsc.ListColumns("DLM en vigueur").Index).NumberFormat = "dd/mm/yyyy"
                    .Cells(1, loEsc.ListColumns("Jours restants").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("Jours restants").Index).Value
                    .Cells(1, loEsc.ListColumns("Reste à décaisser FCFA").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("Reste à décaisser FCFA").Index).Value
                    .Cells(1, loEsc.ListColumns("Reste à décaisser FCFA").Index).NumberFormat = "#,##0"
                    .Cells(1, loEsc.ListColumns("Pièce manquante").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("Pièce manquante").Index).Value
                    .Cells(1, loEsc.ListColumns("Détenteur de la pièce").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("Détenteur de la pièce").Index).Value
                    .Cells(1, loEsc.ListColumns("Nb relances").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("Nb relances").Index).Value
                    .Cells(1, loEsc.ListColumns("Date dernière relance").Index).Value = loDLM.DataBodyRange.Cells(i, loDLM.ListColumns("Date dernière relance").Index).Value
                    .Cells(1, loEsc.ListColumns("Date dernière relance").Index).NumberFormat = "dd/mm/yyyy"
                End With
                
                ' Renseigner Date escalade DG si vide
                If Len(Trim$(CStr(loDLM.DataBodyRange.Cells(i, idxDateEsc).Value & ""))) = 0 Then
                    loDLM.DataBodyRange.Cells(i, idxDateEsc).Value = dtRef
                    loDLM.DataBodyRange.Cells(i, idxDateEsc).NumberFormat = "dd/mm/yyyy"
                End If
            End If
        Next i
    End If
    
    Call ProtegerFeuille(wsDLM)
    Call ProtegerFeuille(wsEsc)
    
    ' 3. Exporter l'onglet Note_Escalade_DG en PDF
    dossierBase = ThisWorkbook.Path
    If Len(dossierBase) = 0 Then dossierBase = CurDir$
    cheminPDF = dossierBase & Application.PathSeparator & "Note_Escalade_DG_ASER_" & Format$(dtRef, "yyyymmdd") & ".pdf"
    
    wsEsc.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminPDF, Quality:=xlQualityStandard, _
                              IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
                              
    Call EcrireJournal("NoteEscaladeDG", "Note d'escalade DG générée (" & nbEsc & " dossier(s)) et exportée en PDF : " & cheminPDF)
    Call ActiverModeMacro(False)
    
    MsgBox "Note d'escalade DG mise à jour (" & nbEsc & " dossier(s) en escalade) et exportée en PDF :" & vbCrLf & vbCrLf & _
           cheminPDF, vbInformation, "Note d'Escalade DG - ASER"
    Exit Sub

ErrHandler:
    Call ActiverModeMacro(False)
    Call EcrireJournal("NoteEscaladeDG", "ERREUR : " & Err.Description & " (Code " & Err.Number & ")")
    MsgBox "Erreur lors de la génération de la Note d'Escalade DG : " & vbCrLf & Err.Description, vbCritical, "Erreur VBA - ASER"
End Sub

'-------------------------------------------------------------------------------
' MACRO 6 : RapportMensuel
' Exporte en PDF le Tableau de bord, les DLM et les Garanties pour la revue
' mensuelle du portefeuille.
'-------------------------------------------------------------------------------
Public Sub RapportMensuel()
    On Error GoTo ErrHandler
    
    Dim dtRef As Date
    Dim dossierBase As String
    Dim cheminPDF As String
    Dim feuillesRapport As Variant
    
    Call ActiverModeMacro(True)
    
    ' 1. S'assurer que les calculs et le plan de paiement sont à jour
    Application.CalculateFull
    Call GenererPlanPaiement_Interne(False)
    
    dtRef = ObtenirDateRef()
    dossierBase = ThisWorkbook.Path
    If Len(dossierBase) = 0 Then dossierBase = CurDir$
    cheminPDF = dossierBase & Application.PathSeparator & "Rapport_Mensuel_Portefeuille_ASER_" & Format$(dtRef, "yyyymmdd") & ".pdf"
    
    ' 2. Exporter groupé en PDF (Tableau_de_bord, DLM_Prorogations, Garanties, Note_Escalade_DG)
    feuillesRapport = Array("Tableau_de_bord", "DLM_Prorogations", "Garanties", "Note_Escalade_DG")
    ThisWorkbook.Worksheets(feuillesRapport).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminPDF, Quality:=xlQualityStandard, _
                                    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ThisWorkbook.Worksheets("Tableau_de_bord").Select
    
    Call EcrireJournal("RapportMensuel", "Rapport mensuel PDF exporté : " & cheminPDF)
    Call ActiverModeMacro(False)
    
    MsgBox "Le Rapport Mensuel du Portefeuille a été généré en PDF :" & vbCrLf & vbCrLf & _
           cheminPDF, vbInformation, "Rapport Mensuel - ASER"
    Exit Sub

ErrHandler:
    On Error Resume Next
    ThisWorkbook.Worksheets("Tableau_de_bord").Select
    Call ActiverModeMacro(False)
    Call EcrireJournal("RapportMensuel", "ERREUR : " & Err.Description & " (Code " & Err.Number & ")")
    MsgBox "Erreur lors de l'export du Rapport Mensuel : " & vbCrLf & Err.Description, vbCritical, "Erreur VBA - ASER"
End Sub
