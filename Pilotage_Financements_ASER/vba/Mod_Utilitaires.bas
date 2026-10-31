Attribute VB_Name = "Mod_Utilitaires"
'===============================================================================
' MODULE       : Mod_Utilitaires
' CLASSEUR     : Pilotage_Financements_ASER.xlsm
' RÔLE         : Fonctions transverses (optimisation d'exécution, protection sans
'                mot de passe, lecture de DateRef, journalisation d'audit)
' COMPATIBILITÉ: Excel 2016 / 2019 / 2021 / 365 (32 bits et 64 bits)
'===============================================================================
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private mCalcState As XlCalculation

'-------------------------------------------------------------------------------
' Active ou restaure les paramètres de performance d'Excel pendant une macro
'-------------------------------------------------------------------------------
Public Sub ActiverModeMacro(ByVal activer As Boolean)
    On Error Resume Next
    If activer Then
        mCalcState = Application.Calculation
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Protège toutes les feuilles sans mot de passe avec UserInterfaceOnly:=True
'-------------------------------------------------------------------------------
Public Sub InitialiserProtectionClasseur()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                   UserInterfaceOnly:=True, AllowFormattingCells:=True, _
                   AllowSorting:=True, AllowFiltering:=True
    Next ws
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Déprotège une feuille spécifique avant modification structurelle (ListObject)
'-------------------------------------------------------------------------------
Public Sub DeprotegerFeuille(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=""
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Reprotège une feuille spécifique sans mot de passe après modification
'-------------------------------------------------------------------------------
Public Sub ProtegerFeuille(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowSorting:=True, AllowFiltering:=True
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Retourne la valeur actuelle de la plage nommée DateRef (ou Date si invalide)
'-------------------------------------------------------------------------------
Public Function ObtenirDateRef() As Date
    On Error GoTo FallbackDate
    Dim v As Variant
    v = ThisWorkbook.Worksheets("Paramètres").Range("DateRef").Value
    If IsDate(v) Then
        ObtenirDateRef = CDate(v)
    ElseIf IsNumeric(v) And CDbl(v) > 0 Then
        ObtenirDateRef = CDate(CDbl(v))
    Else
        ObtenirDateRef = Date
    End If
    Exit Function
FallbackDate:
    ObtenirDateRef = Date
End Function

'-------------------------------------------------------------------------------
' Ajoute une entrée horodatée dans l'onglet Journal (Tableau structuré T_Journal)
'-------------------------------------------------------------------------------
Public Sub EcrireJournal(ByVal nomMacro As String, ByVal resultat As String)
    On Error GoTo ErrJournal
    
    Dim wsJrn As Worksheet
    Dim loJrn As ListObject
    Dim lr As ListRow
    Dim nomUtilisateur As String
    
    Set wsJrn = ThisWorkbook.Worksheets("Journal")
    Set loJrn = wsJrn.ListObjects("T_Journal")
    
    Call DeprotegerFeuille(wsJrn)
    
    nomUtilisateur = Trim$(Application.UserName)
    If Len(nomUtilisateur) = 0 Then nomUtilisateur = Environ$("USERNAME")
    If Len(nomUtilisateur) = 0 Then nomUtilisateur = "Utilisateur_ASER"
    
    Set lr = loJrn.ListRows.Add
    With lr.Range
        .Cells(1, loJrn.ListColumns("Horodatage").Index).Value = Now
        .Cells(1, loJrn.ListColumns("Horodatage").Index).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(1, loJrn.ListColumns("Utilisateur").Index).Value = nomUtilisateur
        .Cells(1, loJrn.ListColumns("Macro").Index).Value = nomMacro
        .Cells(1, loJrn.ListColumns("Résultat").Index).Value = resultat
    End With
    
    Call ProtegerFeuille(wsJrn)
    Exit Sub

ErrJournal:
    On Error Resume Next
    If Not wsJrn Is Nothing Then Call ProtegerFeuille(wsJrn)
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Recherche l'adresse e-mail d'un détenteur dans T_Detenteurs (onglet Paramètres)
'-------------------------------------------------------------------------------
Public Function ObtenirEmailDetenteur(ByVal idDetenteur As String) As String
    On Error GoTo ErrEmail
    Dim wsParam As Worksheet
    Dim loDet As ListObject
    Dim i As Long
    Dim idxId As Long
    Dim idxMail As Long
    Dim cle As String
    
    ObtenirEmailDetenteur = ""
    Set wsParam = ThisWorkbook.Worksheets("Paramètres")
    Set loDet = wsParam.ListObjects("T_Detenteurs")
    If loDet.DataBodyRange Is Nothing Then Exit Function
    
    idxId = loDet.ListColumns("Identifiant_Detenteur").Index
    idxMail = loDet.ListColumns("Email").Index
    
    For i = 1 To loDet.ListRows.Count
        cle = Trim$(CStr(loDet.DataBodyRange.Cells(i, idxId).Value & ""))
        If StrComp(cle, Trim$(idDetenteur), vbTextCompare) = 0 Then
            ObtenirEmailDetenteur = Trim$(CStr(loDet.DataBodyRange.Cells(i, idxMail).Value & ""))
            Exit Function
        End If
    Next i
    Exit Function

ErrEmail:
    ObtenirEmailDetenteur = ""
End Function
