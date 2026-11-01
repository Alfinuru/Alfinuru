Option Explicit
'==============================================================================
' modRelances - relances des détenteurs de pièces manquantes (DLM)
'
'   PreparerRelances     un brouillon Outlook par détenteur, regroupant ses pièces
'                        dont la relance est due (aucun envoi automatique) ; la date
'                        et le nombre de relances ne sont mis à jour qu'après
'                        confirmation, ligne d'historique dans le Journal.
'   EnregistrerRelance   enregistrement manuel (lignes sélectionnées dans « DLM »
'                        ou identifiants saisis).
'
' Outlook est appelé en liaison tardive (aucune référence à cocher) ; sans Outlook
' (Mac, poste sans messagerie), la macro le signale et ne modifie rien.
'==============================================================================

Private Const DETENTEUR_INCONNU As String = "(détenteur non renseigné)"
Private Const OL_MAIL_ITEM As Long = 0

Public Sub PreparerRelances()
    Const NOM As String = "PreparerRelances"
    Dim ws As Worksheet, r As Long, dRef As Date, x As Double, i As Long, k As Long
    Dim noms() As String, lignes() As String, n As Long, det As String, trouve As Boolean
    Dim olApp As Object, email As String, avert As String, rep As Long
    Dim nbBrouillons As Long, nbMaj As Long, lst As Variant, sel As Collection

    On Error GoTo Echec
    Set ws = ThisWorkbook.Worksheets(DLM_FEUILLE)
    Application.Calculate
    dRef = DateReference()

    ' 1. Relances dues regroupées par détenteur (sans Scripting.Dictionary : compatible Mac)
    ReDim noms(0 To 0)
    ReDim lignes(0 To 0)
    For r = DLM_PREMIERE To DLM_DERNIERE
        If Len(Txt(ws.Cells(r, DLM_COL_PIECE).Value)) > 0 Then
            If EnNombre(ws.Cells(r, DLM_COL_PROCHAINE).Value, x) Then
                If x <= CDbl(dRef) Then
                    det = Txt(ws.Cells(r, DLM_COL_DETENTEUR).Value)
                    If Len(det) = 0 Then det = DETENTEUR_INCONNU
                    trouve = False
                    For i = 0 To n - 1
                        If UCase$(noms(i)) = UCase$(det) Then
                            lignes(i) = lignes(i) & ";" & r
                            trouve = True
                            Exit For
                        End If
                    Next i
                    If Not trouve Then
                        ReDim Preserve noms(0 To n)
                        ReDim Preserve lignes(0 To n)
                        noms(n) = det
                        lignes(n) = CStr(r)
                        n = n + 1
                    End If
                End If
            End If
        End If
    Next r
    If n = 0 Then
        Journaliser NOM, "Aucune relance due au " & Format$(dRef, "dd/mm/yyyy") & "."
        Dlg_Msg "Aucune relance due au " & Format$(dRef, "dd/mm/yyyy") & ".", vbInformation
        Exit Sub
    End If

    ' 2. Outlook (liaison tardive)
    If Not gModeTest Then
        On Error Resume Next
        Set olApp = CreateObject("Outlook.Application")
        On Error GoTo Echec
        If olApp Is Nothing Then
            Journaliser NOM, "Échec : Outlook indisponible (" & n & " détenteur(s) à relancer)."
            Dlg_Msg "Outlook n'est pas disponible sur ce poste : aucun brouillon créé." & vbCrLf & _
                    "Utiliser « Enregistrer une relance » après un envoi par un autre moyen.", vbExclamation
            Exit Sub
        End If
    End If

    ' 3. Un brouillon par détenteur, puis confirmation de mise à jour
    For i = 0 To n - 1
        lst = Split(lignes(i), ";")
        email = EmailDetenteur(noms(i))
        CreerBrouillon olApp, email, "Relance : pièces attendues pour la prorogation de DLM", CorpsRelance(ws, lst)
        nbBrouillons = nbBrouillons + 1
        avert = ""
        If Len(email) = 0 Then avert = vbCrLf & "ATTENTION : e-mail du détenteur non renseigné (Paramètres, colonne AD)."
        rep = Dlg_Msg("Brouillon affiché pour : " & noms(i) & " (" & (UBound(lst) + 1) & " pièce(s))." & avert & _
                      vbCrLf & vbCrLf & "Mettre à jour « Dernière relance » (aujourd'hui) et « Nb relances » (+1) ?" & _
                      vbCrLf & vbCrLf & "Oui = vous envoyez cette relance" & vbCrLf & "Non = ne rien mettre à jour" & _
                      vbCrLf & "Annuler = arrêter la préparation", vbYesNoCancel + vbQuestion + vbSystemModal)
        If rep = vbCancel Then Exit For
        If rep = vbYes Then
            Set sel = New Collection
            For k = 0 To UBound(lst)
                sel.Add CLng(lst(k))
            Next k
            AppliquerRelance ws, sel, NOM, noms(i)
            nbMaj = nbMaj + sel.Count
        End If
    Next i
    Terminer
    Journaliser NOM, nbBrouillons & " brouillon(s) préparé(s) ; " & nbMaj & " ligne(s) mise(s) à jour après confirmation."
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

' Brouillon affiché (jamais envoyé) ; en recette, simple trace.
Private Sub CreerBrouillon(ByVal olApp As Object, ByVal dest As String, ByVal objet As String, ByVal html As String)
    Dim mail As Object
    If gModeTest Then
        Trace "BROUILLON", dest & " | " & objet & " | " & (Len(html) - Len(Replace(html, "<tr>", ""))) \ 4 & " ligne(s)"
        Exit Sub
    End If
    Set mail = olApp.CreateItem(OL_MAIL_ITEM)
    mail.Display                                  ' affiché d'abord pour conserver la signature
    mail.To = dest
    mail.Subject = objet
    mail.HTMLBody = InsererDansCorps(mail.HTMLBody, html)
End Sub

' Corps HTML de la relance (tableau des pièces attendues).
Private Function CorpsRelance(ByVal ws As Worksheet, ByVal lst As Variant) As String
    Dim h As String, k As Long, r As Long
    h = "<div style='font-family:Arial;font-size:10pt'><p>Madame, Monsieur,</p>"
    h = h & "<p>Dans le cadre du suivi des dates limites de mobilisation (DLM) des accords de financement, " & _
            "les pièces suivantes restent attendues :</p>"
    h = h & "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-family:Arial;font-size:9pt'>"
    h = h & "<tr style='background:#1F3864;color:#FFFFFF'><th>Bailleur</th><th>Accord</th><th>Projet</th>" & _
            "<th>Pièce attendue</th><th>DLM actuelle</th><th>Jours restants</th><th>Relances antérieures</th></tr>"
    For k = 0 To UBound(lst)
        r = CLng(lst(k))
        h = h & "<tr><td>" & HtmlEsc(Txt(ws.Cells(r, DLM_COL_BAILLEUR).Value)) & "</td><td>" & _
                HtmlEsc(Txt(ws.Cells(r, DLM_COL_ACCORD).Value)) & " (" & HtmlEsc(Txt(ws.Cells(r, DLM_COL_ID).Value)) & _
                ")</td><td>" & HtmlEsc(Txt(ws.Cells(r, DLM_COL_PROJET).Value)) & "</td><td>" & _
                HtmlEsc(Txt(ws.Cells(r, DLM_COL_PIECE).Value)) & "</td><td>" & HtmlEsc(Txt(ws.Cells(r, DLM_COL_DLM).Value)) & _
                "</td><td align='right'>" & HtmlEsc(Txt(ws.Cells(r, DLM_COL_JOURS).Value)) & "</td><td align='right'>" & _
                HtmlEsc(Txt(ws.Cells(r, DLM_COL_NB).Value)) & "</td></tr>"
    Next k
    h = h & "</table><p>Nous vous remercions de bien vouloir nous les transmettre dans les meilleurs délais.</p>"
    h = h & "<p>Cordialement,</p></div>"
    CorpsRelance = h
End Function

' Insère le contenu juste après la balise <body> (la signature Outlook reste en dessous).
Private Function InsererDansCorps(ByVal html As String, ByVal contenu As String) As String
    Dim p As Long, q As Long
    p = InStr(1, html, "<body", vbTextCompare)
    If p > 0 Then q = InStr(p, html, ">")
    If q > 0 Then
        InsererDansCorps = Left$(html, q) & contenu & Mid$(html, q + 1)
    Else
        InsererDansCorps = contenu & html
    End If
End Function

Public Sub EnregistrerRelance()
    Const NOM As String = "EnregistrerRelance"
    Dim ws As Worksheet, cel As Range, zone As Range, lignes As New Collection
    Dim cId As Long, r As Long, saisie As String, idv As Variant, liste As String, it As Variant, trouve As Boolean

    On Error GoTo Echec
    If Not FeuilleExiste(DLM_FEUILLE) Then
        Dlg_Msg "Feuille « " & DLM_FEUILLE & " » introuvable.", vbExclamation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(DLM_FEUILLE)
    cId = NumCol(DLM_COL_ID)

    ' 1) lignes sélectionnées sur la feuille DLM
    If ActiveSheet.Name = ws.Name And TypeName(Selection) = "Range" Then
        Set zone = Application.Intersect(Selection.EntireRow, ws.Columns(cId))
        If Not zone Is Nothing Then
            For Each cel In zone.Cells
                If cel.Row >= DLM_PREMIERE And cel.Row <= DLM_DERNIERE Then
                    If Len(Txt(cel.Value)) > 0 Then AjouterLigne lignes, cel.Row
                End If
            Next cel
        End If
    End If

    ' 2) sinon, saisie des identifiants
    If lignes.Count = 0 Then
        saisie = Dlg_Saisie("Identifiant(s) du dossier DLM relancé (colonne " & DLM_COL_ID & _
                            "), séparés par « ; » :" & vbCrLf & "Exemple : DLM-01;DLM-04")
        If Len(Trim$(saisie)) = 0 Then Exit Sub
        For Each idv In Split(saisie, ";")
            If Len(Trim$(CStr(idv))) > 0 Then
                trouve = False
                For r = DLM_PREMIERE To DLM_DERNIERE
                    If UCase$(Txt(ws.Cells(r, cId).Value)) = UCase$(Trim$(CStr(idv))) Then
                        AjouterLigne lignes, r
                        trouve = True
                        Exit For
                    End If
                Next r
                If Not trouve Then liste = liste & " " & Trim$(CStr(idv))
            End If
        Next idv
        If Len(liste) > 0 Then
            Dlg_Msg "Identifiant(s) introuvable(s) :" & liste, vbExclamation
            If lignes.Count = 0 Then Exit Sub
        End If
    End If

    liste = ""
    For Each it In lignes
        liste = liste & vbCrLf & "   · " & Txt(ws.Cells(CLng(it), cId).Value)
        If Len(Txt(ws.Cells(CLng(it), DLM_COL_PIECE).Value)) = 0 Then liste = liste & " (aucune pièce manquante saisie)"
    Next it
    If Dlg_Msg("Enregistrer une relance datée du " & Format$(Date, "dd/mm/yyyy") & " pour :" & liste, _
               vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    AppliquerRelance ws, lignes, NOM, ""
    Terminer
    Journaliser NOM, lignes.Count & " relance(s) enregistrée(s)."
    Dlg_Msg lignes.Count & " relance(s) enregistrée(s).", vbInformation
    Exit Sub
Echec:
    SignalerErreur NOM, Err.Number, Err.Description
End Sub

' Dernière relance = aujourd'hui, Nb relances + 1, ligne d'historique dans le Journal
' (format exploité par la note d'escalade : « | Accord: <ID> | »).
Private Sub AppliquerRelance(ByVal ws As Worksheet, ByVal lignes As Collection, ByVal macro As String, ByVal det As String)
    Dim it As Variant, r As Long, x As Double
    OptimiserDebut
    Deproteger ws
    For Each it In lignes
        r = CLng(it)
        ws.Cells(r, DLM_COL_DERNIERE).Value = Date
        If EnNombre(ws.Cells(r, DLM_COL_NB).Value, x) Then
            ws.Cells(r, DLM_COL_NB).Value = x + 1
        Else
            ws.Cells(r, DLM_COL_NB).Value = 1
        End If
        Journaliser macro, "Relance confirmée | Accord: " & Txt(ws.Cells(r, DLM_COL_ID).Value) & " | Détenteur: " & _
                           IIf(Len(det) > 0, det, Txt(ws.Cells(r, DLM_COL_DETENTEUR).Value))
    Next it
    Application.Calculate
    Terminer
End Sub

Private Sub AjouterLigne(ByVal col As Collection, ByVal r As Long)
    On Error Resume Next
    col.Add r, CStr(r)
    Err.Clear
End Sub
