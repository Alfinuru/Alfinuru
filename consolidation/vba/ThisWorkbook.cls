Option Explicit

' À l'ouverture : message uniquement s'il existe des alertes critiques (ou une date
' de référence invalide). Le classeur n'est pas modifié : pas d'invite d'enregistrement.
Private Sub Workbook_Open()
    On Error Resume Next
    NotificationOuverture
End Sub
