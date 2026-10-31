'===============================================================================
' MODULE       : ThisWorkbook
' RÔLE         : Active les raccourcis Ctrl+Shift+A/D/G/P/H/Q/E lorsque ce classeur
'                est actif et les libère lorsqu'il ne l'est plus (ou à sa fermeture)
'===============================================================================
Option Explicit

Private Sub Workbook_Open()
    Call Activer_Raccourcis
End Sub

Private Sub Workbook_Activate()
    Call Activer_Raccourcis
End Sub

Private Sub Workbook_Deactivate()
    Call Desactiver_Raccourcis
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call Desactiver_Raccourcis
End Sub
