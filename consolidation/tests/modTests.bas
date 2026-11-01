Option Explicit
'==============================================================================
' modTests - module de recette (ajouté uniquement au classeur de test)
' Pilote les dialogues de modCommun : réponses scriptées, trace des messages.
'==============================================================================

Public Sub T_Init(ByVal reponses As String, ByVal dossier As String)
    Dim p As Variant
    gModeTest = True
    Set gReponses = New Collection
    If Len(reponses) > 0 Then
        For Each p In Split(reponses, "|")
            gReponses.Add CStr(p)
        Next p
    End If
    gTrace = ""
    gDossierTest = dossier
End Sub

Public Function T_Trace() As String
    T_Trace = gTrace
End Function

Public Function T_Restant() As Long
    If gReponses Is Nothing Then Exit Function
    T_Restant = gReponses.Count
End Function

Public Sub T_Activer(ByVal feuille As String, ByVal adresse As String)
    ThisWorkbook.Worksheets(feuille).Activate
    If Len(adresse) > 0 Then ThisWorkbook.Worksheets(feuille).Range(adresse).Select
End Sub

Public Function T_Protegee(ByVal feuille As String) As Boolean
    T_Protegee = ThisWorkbook.Worksheets(feuille).ProtectContents
End Function

Public Function T_Masquee(ByVal feuille As String, ByVal ligne As Long) As Boolean
    T_Masquee = ThisWorkbook.Worksheets(feuille).Rows(ligne).Hidden
End Function

Public Function T_Notification() As String
    T_Init "", ""
    NotificationOuverture
    T_Notification = gTrace
End Function
