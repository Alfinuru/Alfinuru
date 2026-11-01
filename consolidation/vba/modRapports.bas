Option Explicit
'==============================================================================
' modRapports - note d'escalade au Directeur général (feuille + PDF)
'
'   NoteEscaladeDG   liste les accords dont « Escalade DG requise » = OUI (DLM échue,
'                    DLM sous le seuil d'escalade, ou relances épuisées), avec
'                    l'historique des relances tracées dans le Journal, puis exporte
'                    la note en PDF dans le dossier du classeur.
' (La revue mensuelle en PDF est produite par modMoteur.ExporterRevuePDF.)
'==============================================================================

Private Const LIGNE_ENTETE_NOTE As Long = 8
Private Const NB_COL_NOTE As Long = 11

Public Sub NoteEscaladeDG()
    Const NOM As String = "NoteEscaladeDG"
    Dim wsD As Worksheet, ws As Worksheet, r As Long, rr As Long, n As Long
    Dim fichier As String, total As Double, x As Double, id As String

    On Error GoTo Echec
    OptimiserDebut
    Application.Calculate
    Set wsD = ThisWorkbook.Worksheets(DLM_FEUILLE)
    Set ws = ThisWorkbook.Worksheets(FEUILLE_NOTE)
    Deproteger ws
    ws.Range(ws.Cells(LIGNE_ENTETE_NOTE + 1, 1), ws.Cells(ws.Rows.Count, NB_COL_NOTE)).Clear
    ws.Range("B4:K4").ClearContents
    ws.Range("B6:K6").ClearContents
    ws.Range("B4").Value = "Date de référence : " & DateReferenceTexte() & "   ·   note établie le " & Format$(Date, "dd/mm/yyyy")

    ' 1. Accords en escalade
    rr = LIGNE_ENTETE_NOTE
    For r = DLM_PREMIERE To DLM_DERNIERE
        If Txt(wsD.Cells(r, DLM_COL_ESCALADE).Value) = "OUI" Then
            rr = rr + 1
            n = n + 1
            id = Txt(wsD.Cells(r, DLM_COL_ID).Value)
            ws.Cells(rr, 1).Value = Texte(wsD.Cells(r, DLM_COL_BAILLEUR).Value)
            ws.Cells(rr, 2).Value = Texte(id & " - " & Txt(wsD.Cells(r, DLM_COL_ACCORD).Value))
            ws.Cells(rr, 3).Value = Texte(wsD.Cells(r, DLM_COL_PROJET).Value)
            ws.Cells(rr, 4).Value = wsD.Cells(r, DLM_COL_DLM).Value
            ws.Cells(rr, 5).Value = wsD.Cells(r, DLM_COL_JOURS).Value
            ws.Cells(rr, 6).Value = wsD.Cells(r, DLM_COL_SOLDE_FCFA).Value
            ws.Cells(rr, 7).Value = Texte(wsD.Cells(r, DLM_COL_PIECE).Value)
            ws.Cells(rr, 8).Value = Texte(wsD.Cells(r, DLM_COL_DETENTEUR).Value)
            ws.Cells(rr, 9).Value = wsD.Cells(r, DLM_COL_NB).Value
            ws.Cells(rr, 10).Value = wsD.Cells(r, DLM_COL_DERNIERE).Value
            ws.Cells(rr, 11).Value = Texte(Txt(wsD.Cells(r, DLM_COL_NIVEAU).Value) & " - relances : " & HistoriqueRelances(id))
            If EnNombre(wsD.Cells(r, DLM_COL_SOLDE_FCFA).Value, x) Then total = total + x
        End If
    Next r

    If n = 0 Then
        ws.Cells(LIGNE_ENTETE_NOTE + 1, 1).Value = "Aucun accord ne requiert d'escalade à la date de référence."
        Terminer
        Journaliser NOM, "Aucune escalade requise."
        Dlg_Msg "Aucun accord ne requiert d'escalade DG : aucun PDF produit.", vbInformation
        Exit Sub
    End If

    ' 2. Mise en forme
    ws.Range("B6").Value = n & " accord(s) en escalade ; solde non décaissé cumulé : " & _
                           Format$(total, "#,##0") & " FCFA."
    With ws.Range(ws.Cells(LIGNE_ENTETE_NOTE + 1, 1), ws.Cells(rr + 1, NB_COL_NOTE))
        .Font.Name = "Arial"
        .Font.Size = 9
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    With ws.Range(ws.Cells(LIGNE_ENTETE_NOTE + 1, 1), ws.Cells(rr, NB_COL_NOTE)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(LIGNE_ENTETE_NOTE + 1, 4), ws.Cells(rr, 4)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(LIGNE_ENTETE_NOTE + 1, 10), ws.Cells(rr, 10)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(LIGNE_ENTETE_NOTE + 1, 6), ws.Cells(rr + 1, 6)).NumberFormat = "#,##0"
    ws.Cells(rr + 1, 5).Value = "Total"
    ws.Cells(rr + 1, 5).Font.Bold = True
    ws.Cells(rr + 1, 6).Value = total
    ws.Cells(rr + 1, 6).Font.Bold = True

    ' 3. Mise en page et export PDF
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(rr + 1, NB_COL_NOTE)).Address
        .PrintTitleRows = "$" & LIGNE_ENTETE_NOTE & ":$" & LIGNE_ENTETE_NOTE
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "Page &P / &N"
    End With
    On Error GoTo Echec
    If gModeTest Then
        fichier = gDossierTest & "/Note_Escalade_DG_" & Horodatage() & ".pdf"
    Else
        fichier = DossierExport() & Application.PathSeparator & "Note_Escalade_DG_" & Horodatage() & ".pdf"
    End If
    If gModeTest Then Trace "PDF", fichier & " <- " & ws.Name
    ws.ExportAsFixedFormat Type:=0, Filename:=fichier, Quality:=0, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Terminer
    Journaliser NOM, n & " accord(s) en escalade ; PDF : " & fichier
    Dlg_Msg "Note d'escalade générée (" & n & " accord(s))." & vbCrLf & vbCrLf & "PDF : " & fichier, vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

' Dates des relances confirmées pour un accord (lues dans le Journal).
Private Function HistoriqueRelances(ByVal accord As String) As String
    Dim ws As Worksheet, r As Long, der As Long, res As String, m As String
    Set ws = ThisWorkbook.Worksheets(FEUILLE_JOURNAL)
    der = DerniereLigne(ws, 1)
    For r = 5 To der
        m = Txt(ws.Cells(r, 3).Value)
        If m = "PreparerRelances" Or m = "EnregistrerRelance" Then
            If InStr(1, Txt(ws.Cells(r, 4).Value) & " |", "| Accord: " & accord & " |", vbTextCompare) > 0 Then
                If Len(res) > 0 Then res = res & " ; "
                res = res & Left$(Txt(ws.Cells(r, 1).Value), 10)
            End If
        End If
    Next r
    If Len(res) = 0 Then res = "aucune relance tracée par les macros"
    HistoriqueRelances = res
End Function
