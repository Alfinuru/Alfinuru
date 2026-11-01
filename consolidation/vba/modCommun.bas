Option Explicit
Option Private Module
'==============================================================================
' modCommun - services partagés du classeur consolidé
'   - dialogues (MsgBox / InputBox / choix de fichier) centralisés : ils
'     peuvent être scriptés par la recette automatisée (mode test) ;
'   - état d'Excel (affichage, événements) avec gestion des appels imbriqués ;
'   - protection des feuilles : seules les feuilles déprotégées par la macro
'     sont reprotégées (une feuille ôtée volontairement de la protection reste
'     telle quelle) ;
'   - journal des macros (feuille « Journal ») ;
'   - conversions sûres (#N/A, texte, dates indépendantes des paramètres régionaux).
' Aucun appel Windows spécifique (pas de Scripting.Dictionary ni d'API Declare) :
' compatible Excel 32/64 bits, Windows et Mac.
'==============================================================================

' --- Mode test : renseigné uniquement par le module de recette ----------------
Public gModeTest As Boolean
Public gReponses As Collection      ' réponses scriptées, dans l'ordre des questions
Public gTrace As String             ' messages affichés, saisies demandées, brouillons, PDF
Public gDossierTest As String       ' dossier des PDF produits en recette

Private mNiveau As Long
Private mScreen As Boolean
Private mEvents As Boolean
Private mDeprotegees As Collection

'==============================================================================
' 1. DIALOGUES
'==============================================================================

Public Function Dlg_Msg(ByVal message As String, Optional ByVal boutons As Long = 0) As Long
    If gModeTest Then
        Trace "MSG", message
        Dlg_Msg = ReponseTest(boutons)
    Else
        Dlg_Msg = MsgBox(Left$(message, 1020), boutons, APP_TITRE)
    End If
End Function

' Renvoie "" si l'utilisateur annule.
Public Function Dlg_Saisie(ByVal message As String, Optional ByVal defaut As String = "") As String
    If gModeTest Then
        Trace "SAISIE", message
        Dlg_Saisie = Depiler(defaut)
    Else
        Dlg_Saisie = InputBox(message, APP_TITRE, defaut)
    End If
End Function

' Chemin du PDF choisi par l'utilisateur ("" si annulé).
Public Function Dlg_FichierPDF(ByVal nomPropose As String, ByVal titre As String) As String
    Dim v As Variant
    If gModeTest Then
        Trace "FICHIER", nomPropose
        Dlg_FichierPDF = gDossierTest & "/" & nomPropose
        Exit Function
    End If
    v = Application.GetSaveAsFilename(InitialFileName:=DossierExport() & Application.PathSeparator & nomPropose, _
                                      FileFilter:="PDF (*.pdf), *.pdf", Title:=titre)
    If VarType(v) = vbBoolean Then
        Dlg_FichierPDF = ""
    Else
        Dlg_FichierPDF = CStr(v)
    End If
End Function

Public Sub Trace(ByVal genre As String, ByVal message As String)
    gTrace = gTrace & "[" & genre & "] " & Replace(Replace(message, vbCrLf, " / "), vbLf, " / ") & vbLf
End Sub

Private Function Depiler(ByVal defaut As String) As String
    Depiler = defaut
    If gReponses Is Nothing Then Exit Function
    If gReponses.Count = 0 Then Exit Function
    Depiler = CStr(gReponses.Item(1))
    gReponses.Remove 1
End Function

Private Function ReponseTest(ByVal boutons As Long) As Long
    Dim r As String, typ As Long
    typ = boutons Mod 8
    If typ = vbYesNo Or typ = vbYesNoCancel Or typ = vbOKCancel Or typ = vbRetryCancel Then
        r = UCase$(Depiler("OUI"))
        If r = "NON" Then
            ReponseTest = vbNo
        ElseIf r = "ANNULER" Then
            ReponseTest = vbCancel
        ElseIf typ = vbOKCancel Then
            ReponseTest = vbOK
        Else
            ReponseTest = vbYes
        End If
    Else
        ReponseTest = vbOK
    End If
End Function

'==============================================================================
' 2. ÉTAT D'EXCEL ET PROTECTION
'==============================================================================

Public Sub OptimiserDebut()
    On Error Resume Next
    If mNiveau = 0 Then
        mScreen = Application.ScreenUpdating
        mEvents = Application.EnableEvents
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If
    mNiveau = mNiveau + 1
End Sub

Public Sub OptimiserFin()
    On Error Resume Next
    If mNiveau = 0 Then Exit Sub
    mNiveau = mNiveau - 1
    If mNiveau = 0 Then
        Application.EnableEvents = mEvents
        Application.ScreenUpdating = mScreen
    End If
End Sub

' Après une erreur : rétablit l'état initial quel que soit le niveau d'imbrication.
Public Sub OptimiserReinitialiser()
    If mNiveau > 0 Then
        mNiveau = 1
        OptimiserFin
    End If
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Ôte la protection (sans mot de passe) et mémorise la feuille pour la reprotéger.
Public Sub Deproteger(ByVal ws As Worksheet)
    If mDeprotegees Is Nothing Then Set mDeprotegees = New Collection
    If ws.ProtectContents Then
        ws.Unprotect Password:=""
        On Error Resume Next
        mDeprotegees.Add ws.Name, ws.Name
        On Error GoTo 0
    End If
End Sub

' Reprotège les seules feuilles déprotégées par la macro en cours.
Public Sub Reproteger()
    Dim n As Variant
    If mDeprotegees Is Nothing Then Exit Sub
    On Error Resume Next
    For Each n In mDeprotegees
        ProtegerFeuille ThisWorkbook.Worksheets(CStr(n))
    Next n
    Set mDeprotegees = New Collection
End Sub

' Protection sans mot de passe : saisie dans les cellules déverrouillées, filtres,
' largeur des colonnes et hauteur des lignes autorisés.
Public Sub ProtegerFeuille(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
    If Err.Number <> 0 Then
        Err.Clear
        ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
End Sub

' Fin normale d'une macro.
Public Sub Terminer()
    Reproteger
    OptimiserFin
End Sub

' Fin sur erreur : état rétabli, feuilles reprotégées, message, journal.
Public Sub SignalerErreur(ByVal macro As String, ByVal numero As Long, ByVal description As String)
    Reproteger
    OptimiserReinitialiser
    Dlg_Msg "La macro « " & macro & " » s'est arrêtée." & vbCrLf & vbCrLf & _
            "Erreur " & numero & " : " & description, vbExclamation
    Journaliser macro, "ERREUR " & numero & " : " & description
End Sub

'==============================================================================
' 3. JOURNAL DES MACROS
'==============================================================================

' Ajoute une ligne au Journal ; ne bloque jamais la macro appelante.
Public Sub Journaliser(ByVal macro As String, ByVal resultat As String)
    Dim ws As Worksheet, r As Long, etait As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FEUILLE_JOURNAL)
    If ws Is Nothing Then Exit Sub
    etait = ws.ProtectContents
    If etait Then ws.Unprotect Password:=""
    r = DerniereLigne(ws, 1) + 1
    If r < 5 Then r = 5
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(r, 2).Value = Texte(NomUtilisateur())
    ws.Cells(r, 3).Value = Texte(macro)
    ws.Cells(r, 4).Value = Texte(Left$(resultat, 30000))
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Font.Name = "Arial"
        .Font.Size = 9
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 217, 217)
    End With
    If etait Then ProtegerFeuille ws
    Err.Clear
End Sub

Public Function NomUtilisateur() As String
    On Error Resume Next
    NomUtilisateur = Application.UserName
    If Len(NomUtilisateur) = 0 Then NomUtilisateur = Environ$("USERNAME")
    If Len(NomUtilisateur) = 0 Then NomUtilisateur = Environ$("USER")
    Err.Clear
End Function

' « jj/mm/aaaa hh:mm » sans dépendre des codes de format régionaux.
Public Function DateHeure(ByVal d As Date) As String
    DateHeure = Format$(d, "dd/mm/yyyy") & " " & Format$(Hour(d), "00") & ":" & Format$(Minute(d), "00")
End Function

' Suffixe de nom de fichier « aaaammjj_hhmm ».
Public Function Horodatage() As String
    Horodatage = Format$(Now, "yyyymmdd") & "_" & Format$(Hour(Now), "00") & Format$(Minute(Now), "00")
End Function

'==============================================================================
' 4. UTILITAIRES
'==============================================================================

' Valeur de cellule -> texte, sans erreur d'exécution sur #N/A, #VALEUR!...
Public Function Txt(ByVal v As Variant) As String
    If IsError(v) Then
        Txt = TexteErreur(v)
    ElseIf IsEmpty(v) Or IsNull(v) Then
        Txt = ""
    ElseIf VarType(v) = vbDate Then
        Txt = Format$(v, "dd/mm/yyyy")
    Else
        Txt = Trim$(CStr(v))
    End If
End Function

' Valeur numérique (nombre ou date) ; False pour le texte, le vide et les erreurs.
Public Function EnNombre(ByVal v As Variant, ByRef x As Double) As Boolean
    If IsError(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate, vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
            x = CDbl(v)
            EnNombre = True
    End Select
End Function

' Lit jj/mm/aaaa, jj-mm-aaaa, jj.mm.aaaa ou aaaa-mm-jj, indépendamment des paramètres régionaux.
Public Function LireDate(ByVal s As String, ByRef d As Date) As Boolean
    Dim p() As String, j As Long, m As Long, a As Long
    s = Trim$(Replace(Replace(s, "-", "/"), ".", "/"))
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
    p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function
    If Not (EstEntier(p(0)) And EstEntier(p(1)) And EstEntier(p(2))) Then Exit Function
    If Len(p(0)) = 4 Then
        a = CLng(p(0)): m = CLng(p(1)): j = CLng(p(2))
    Else
        j = CLng(p(0)): m = CLng(p(1)): a = CLng(p(2))
        If a < 100 Then a = a + 2000
    End If
    If m < 1 Or m > 12 Or j < 1 Or j > 31 Or a < 1990 Or a > 2100 Then Exit Function
    d = DateSerial(a, m, j)
    If Day(d) <> j Then Exit Function
    LireDate = True
End Function

Private Function EstEntier(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 4 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    EstEntier = True
End Function

' Préfixe apostrophe : le texte reste du texte (« 1.1 », « 2024-01 » ou « =x » ne sont pas convertis).
Public Function Texte(ByVal v As Variant) As Variant
    Dim s As String
    s = Txt(v)
    If Len(s) = 0 Then
        Texte = Empty
    Else
        Texte = "'" & s
    End If
End Function

Public Function TexteErreur(ByVal v As Variant) As String
    If v = CVErr(2042) Then
        TexteErreur = "#N/A"
    ElseIf v = CVErr(2015) Then
        TexteErreur = "#VALEUR!"
    ElseIf v = CVErr(2023) Then
        TexteErreur = "#REF!"
    ElseIf v = CVErr(2007) Then
        TexteErreur = "#DIV/0!"
    ElseIf v = CVErr(2029) Then
        TexteErreur = "#NOM?"
    ElseIf v = CVErr(2036) Then
        TexteErreur = "#NOMBRE!"
    ElseIf v = CVErr(2000) Then
        TexteErreur = "#NUL!"
    Else
        TexteErreur = "#ERREUR"
    End If
End Function

Public Function FeuilleExiste(ByVal nom As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not (ws Is Nothing)
End Function

' Numéro de colonne d'une lettre (« AB » -> 28) ; 0 si vide.
Public Function NumCol(ByVal lettre As String) As Long
    Dim i As Long, c As Long
    lettre = UCase$(Trim$(lettre))
    For i = 1 To Len(lettre)
        c = c * 26 + (Asc(Mid$(lettre, i, 1)) - 64)
    Next i
    NumCol = c
End Function

Public Function LettreCol(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(65 + (n - 1) Mod 26) & s
        n = (n - 1) \ 26
    Loop
    LettreCol = s
End Function

Public Function Champ(ByVal regle As String, ByVal idx As Long) As String
    Dim t() As String
    t = Split(regle, "|")
    If idx <= UBound(t) Then Champ = Trim$(t(idx)) Else Champ = ""
End Function

' Ajoute une clé texte unique à une collection (sans erreur si elle existe déjà).
Public Function AjouterCle(ByVal col As Collection, ByVal cle As String) As Boolean
    On Error Resume Next
    col.Add cle, cle
    AjouterCle = (Err.Number = 0)
    Err.Clear
End Function

Public Function HtmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HtmlEsc = s
End Function

' Dossier d'export : dossier du classeur, sinon Documents (classeur non enregistré
' ou ouvert depuis OneDrive / SharePoint). Fonctionne aussi sur Mac.
Public Function DossierExport() As String
    Dim p As String
    p = ThisWorkbook.Path
    If Len(p) = 0 Or LCase$(Left$(p, 4)) = "http" Then
        p = ""
        On Error Resume Next
        p = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
        On Error GoTo 0
        If Len(p) = 0 Then p = Environ$("USERPROFILE")
        If Len(p) = 0 Then p = Environ$("HOME")
        If Len(p) = 0 Then p = CurDir$
    End If
    DossierExport = p
End Function

' Dernière ligne non vide d'une colonne (sans .End(xlUp) : indépendant de la fenêtre active).
Public Function DerniereLigne(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Do While r > 1
        If Len(Txt(ws.Cells(r, col).Value)) > 0 Then Exit Do
        r = r - 1
    Loop
    DerniereLigne = r
End Function

Public Function EstFiltree(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    EstFiltree = ws.FilterMode
End Function

Public Function FeuilleActive() As Object
    On Error Resume Next
    Set FeuilleActive = ActiveSheet
End Function

Public Sub Recalculer()
    On Error Resume Next
    Application.CalculateFull
    If Err.Number <> 0 Then
        Err.Clear
        Application.Calculate
    End If
    On Error GoTo 0
End Sub

' Efface les saisies (cellules sans formule) d'une ligne, colonnes 1 à nbCol.
Public Sub ViderSaisiesLigne(ByVal ws As Worksheet, ByVal r As Long, ByVal nbCol As Long)
    Dim c As Long
    For c = 1 To nbCol
        If Not ws.Cells(r, c).HasFormula Then
            If Not IsEmpty(ws.Cells(r, c).Value) Then ws.Cells(r, c).ClearContents
        End If
    Next c
End Sub
