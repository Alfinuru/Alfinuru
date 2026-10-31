'===============================================================================
' MODULE       : ThisWorkbook
' RÔLE         : Réapplique à l'ouverture la protection sans mot de passe avec
'                UserInterfaceOnly:=True (option non conservée à l'enregistrement)
'===============================================================================
Option Explicit

Private Sub Workbook_Open()
    Call InitialiserProtectionClasseur
End Sub
