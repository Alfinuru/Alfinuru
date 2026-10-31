Attribute VB_Name = "Mod_Utilitaires"
'===============================================================================
' MODULE       : Mod_Utilitaires
' CLASSEUR     : Pilotage_Financements_ASER.xlsm
' RÔLE         : Fonctions transverses (optimisation d'exécution, protection sans
'                mot de passe, lecture de DateRef, journalisation d'audit,
'                lecture sécurisée des cellules, ajout de lignes de tableau)
' COMPATIBILITÉ: Excel 2016 / 2019 / 2021 / 365 (32 bits et 64 bits)
'===============================================================================
Option Explicit

Private mCalcState As XlCalculation
Private mModeActif As Boolean

'-------------------------------------------------------------------------------
' Active ou restaure les paramètres de performance d'Excel pendant une macro.
' Le mode de calcul d'origine est mémorisé puis restauré (et non forcé en
' automatique) ; un second appel d'activation n'écrase pas l'état mémorisé.
'-------------------------------------------------------------------------------
Public Sub ActiverModeMacro(ByVal activer As Boolean)
    On Error Resume Next
    If activer Then
        If Not mModeActif Then
            mCalcState = Application.Calculation
            mModeActif = True
        End If
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    Else
        If mModeActif Then
            Application.Calculation = mCalcState
        Else
            Application.Calculation = xlCalculationAutomatic
        End If
        mModeActif = False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Protège toutes les feuilles sans mot de passe avec UserInterfaceOnly:=True
' (appelée à l'ouverture du classeur et en reprise après erreur)
'-------------------------------------------------------------------------------
Public Sub InitialiserProtectionClasseur()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Call ProtegerFeuille(ws)
    Next ws
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
    v = ThisWorkbook.Names("DateRef").RefersToRange.Value
    If IsError(v) Then
        ObtenirDateRef = Date
    ElseIf IsDate(v) Then
        ObtenirDateRef = DateValue(CDate(v))
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        If CDbl(v) > 0 Then
            ObtenirDateRef = CDate(Int(CDbl(v)))
        Else
            ObtenirDateRef = Date
        End If
    Else
        ObtenirDateRef = Date
    End If
    Exit Function
FallbackDate:
    ObtenirDateRef = Date
End Function

'-------------------------------------------------------------------------------
' Lecture sécurisée d'une valeur de cellule en texte (erreurs #N/A, #REF!... -> "")
'-------------------------------------------------------------------------------
Public Function TexteCellule(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or IsArray(v) Then
        TexteCellule = ""
    Else
        TexteCellule = Trim$(CStr(v))
    End If
End Function

'-------------------------------------------------------------------------------
' Lecture sécurisée d'une valeur numérique (texte, vide ou erreur -> valeur défaut)
'-------------------------------------------------------------------------------
Public Function NombreCellule(ByVal v As Variant, Optional ByVal valeurDefaut As Double = 0) As Double
    NombreCellule = valeurDefaut
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or IsArray(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(v)) = 0 Then Exit Function
    End If
    If IsNumeric(v) Then
        NombreCellule = CDbl(v)
    ElseIf IsDate(v) Then
        NombreCellule = CDbl(CDate(v))
    End If
End Function

'-------------------------------------------------------------------------------
' Formatage sécurisé d'une date de cellule (vide/erreur/non-date -> "")
'-------------------------------------------------------------------------------
Public Function DateTexteCellule(ByVal v As Variant) As String
    DateTexteCellule = ""
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or IsArray(v) Then Exit Function
    If IsDate(v) Then
        DateTexteCellule = Format$(CDate(v), "dd/mm/yyyy")
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then DateTexteCellule = Format$(CDate(CDbl(v)), "dd/mm/yyyy")
    End If
End Function

'-------------------------------------------------------------------------------
' Neutralise les caractères spéciaux HTML (injection dans le corps des e-mails)
'-------------------------------------------------------------------------------
Public Function EchapperHTML(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    s = Replace(s, vbCrLf, "<br>")
    s = Replace(s, vbLf, "<br>")
    EchapperHTML = s
End Function

'-------------------------------------------------------------------------------
' Dossier d'export des PDF : dossier du classeur s'il est local, sinon dossier
' par défaut d'Excel (cas des classeurs ouverts depuis SharePoint / OneDrive)
'-------------------------------------------------------------------------------
Public Function DossierExport() As String
    Dim d As String
    d = ThisWorkbook.Path
    If Len(d) = 0 Or LCase$(Left$(d, 4)) = "http" Then d = Application.DefaultFilePath
    If Len(d) = 0 Or LCase$(Left$(d, 4)) = "http" Then d = CurDir$
    If Right$(d, 1) = Application.PathSeparator Then d = Left$(d, Len(d) - 1)
    DossierExport = d
End Function

'-------------------------------------------------------------------------------
' Ajoute une ligne en fin de tableau structuré. Si la ligne située sous le
' tableau est occupée (ex. second tableau empilé dans l'onglet Archive), une
' ligne entière de feuille est d'abord insérée pour éviter le décalage partiel
' de cellules (erreur 1004). Les formules calculées connues sont réappliquées.
'-------------------------------------------------------------------------------
Public Function AjouterLigneTableau(ByVal lo As ListObject) As ListRow
    Dim ws As Worksheet
    Dim ligneSous As Range
    Dim lr As ListRow

    Set ws = lo.Parent
    Set ligneSous = lo.Range.Rows(lo.Range.Rows.Count).Offset(1, 0)
    If Application.WorksheetFunction.CountA(ligneSous) > 0 Then
        ligneSous.EntireRow.Insert Shift:=xlShiftDown
    End If

    Set lr = lo.ListRows.Add
    Call AppliquerFormulesCalculees(lo, lr)
    Set AjouterLigneTableau = lr
End Function

'-------------------------------------------------------------------------------
' Réapplique sur une ligne les formules des colonnes calculées d'un registre
' (indispensable lorsque le tableau a été vidé, par ex. après SupprimerExemples)
'-------------------------------------------------------------------------------
Public Sub AppliquerFormulesCalculees(ByVal lo As ListObject, ByVal lr As ListRow)
    Dim c As Long
    Dim f As String
    For c = 1 To lo.ListColumns.Count
        f = FormuleModele(lo.Name, lo.ListColumns(c).Name)
        If Len(f) > 0 Then
            lr.Range.Cells(1, c).Formula = f
            lr.Range.Cells(1, c).Locked = True
        End If
    Next c
End Sub

'-------------------------------------------------------------------------------
' Référentiel des formules des colonnes calculées (identiques aux formules
' d'origine du classeur)
'-------------------------------------------------------------------------------
Public Function FormuleModele(ByVal nomTable As String, ByVal nomColonne As String) As String
    FormuleModele = ""
    Select Case nomTable
        Case "T_Plan"
            Select Case nomColonne
                Case "Jours de retard"
                    FormuleModele = "=IF(OR(T_Plan[[#This Row],[Statut]]=""Terminé"",ISBLANK(T_Plan[[#This Row],[Échéance]])),0,MAX(0,DateRef-T_Plan[[#This Row],[Échéance]]))"
            End Select
        Case "T_DLM"
            Select Case nomColonne
                Case "Montant accord FCFA"
                    FormuleModele = "=ROUND(T_DLM[[#This Row],[Montant accord]]*INDEX(T_Devises[Taux_FCFA],MATCH(T_DLM[[#This Row],[Devise]],T_Devises[Devise],0)),0)"
                Case "Reste à décaisser FCFA"
                    FormuleModele = "=ROUND((T_DLM[[#This Row],[Montant accord]]-T_DLM[[#This Row],[Montant décaissé]])*INDEX(T_Devises[Taux_FCFA],MATCH(T_DLM[[#This Row],[Devise]],T_Devises[Devise],0)),0)"
                Case "Taux de décaissement"
                    FormuleModele = "=IFERROR(T_DLM[[#This Row],[Montant décaissé]]/T_DLM[[#This Row],[Montant accord]],0)"
                Case "Jours restants"
                    FormuleModele = "=IF(ISBLANK(T_DLM[[#This Row],[DLM en vigueur]]),0,T_DLM[[#This Row],[DLM en vigueur]]-DateRef)"
                Case "Prochaine relance"
                    FormuleModele = "=IF(ISBLANK(T_DLM[[#This Row],[Pièce manquante]]),0,IF(ISBLANK(T_DLM[[#This Row],[Date dernière relance]]),DateRef,T_DLM[[#This Row],[Date dernière relance]]+FrequenceRelance))"
                Case "Niveau d'alerte"
                    FormuleModele = "=IF(T_DLM[[#This Row],[Jours restants]]<=DLM_Rouge,""Rouge"",IF(T_DLM[[#This Row],[Jours restants]]<=DLM_Orange,""Orange"",""Vert""))"
                Case "Escalade DG requise"
                    FormuleModele = "=IF(AND(NOT(ISBLANK(T_DLM[[#This Row],[Pièce manquante]])),OR(T_DLM[[#This Row],[Jours restants]]<=DLM_Rouge,T_DLM[[#This Row],[Nb relances]]>=Escalade_NbRelances)),""Oui"",""Non"")"
            End Select
        Case "T_Garanties"
            Select Case nomColonne
                Case "Montant FCFA"
                    FormuleModele = "=ROUND(T_Garanties[[#This Row],[Montant]]*INDEX(T_Devises[Taux_FCFA],MATCH(T_Garanties[[#This Row],[Devise]],T_Devises[Devise],0)),0)"
                Case "Jours avant échéance"
                    FormuleModele = "=IF(ISBLANK(T_Garanties[[#This Row],[Date d''échéance]]),0,T_Garanties[[#This Row],[Date d''échéance]]-DateRef)"
                Case "Alerte"
                    FormuleModele = "=IF(AND(OR(T_Garanties[[#This Row],[Statut]]=""Valide"",T_Garanties[[#This Row],[Statut]]=""Prorogée""),T_Garanties[[#This Row],[Jours avant échéance]]<0),""Noir"",IF(AND(OR(T_Garanties[[#This Row],[Statut]]=""Valide"",T_Garanties[[#This Row],[Statut]]=""Prorogée""),T_Garanties[[#This Row],[Jours avant échéance]]<=AlerteGarantie),""Rouge"",""Vert""))"
            End Select
        Case "T_Factures"
            Select Case nomColonne
                Case "Reste à payer"
                    FormuleModele = "=MAX(0,T_Factures[[#This Row],[Montant TTC]]-T_Factures[[#This Row],[Montant payé]])"
                Case "Étape en cours"
                    FormuleModele = "=IF(AND(T_Factures[[#This Row],[Montant TTC]]>0,T_Factures[[#This Row],[Montant payé]]>=T_Factures[[#This Row],[Montant TTC]]),""Payée"",IF(T_Factures[[#This Row],[Montant payé]]>0,""Partiellement payée"",IF(NOT(ISBLANK(T_Factures[[#This Row],[Date transmission AC]])),""Transmise AC"",IF(NOT(ISBLANK(T_Factures[[#This Row],[Date visa]])),""Visée"",""Reçue""))))"
                Case "Délai réception" & ChrW(8594) & "visa"
                    FormuleModele = "=IF(OR(ISBLANK(T_Factures[[#This Row],[Date réception]]),ISBLANK(T_Factures[[#This Row],[Date visa]])),0,T_Factures[[#This Row],[Date visa]]-T_Factures[[#This Row],[Date réception]])"
                Case "Délai visa" & ChrW(8594) & "AC"
                    FormuleModele = "=IF(OR(ISBLANK(T_Factures[[#This Row],[Date visa]]),ISBLANK(T_Factures[[#This Row],[Date transmission AC]])),0,T_Factures[[#This Row],[Date transmission AC]]-T_Factures[[#This Row],[Date visa]])"
                Case "Délai AC" & ChrW(8594) & "paiement"
                    FormuleModele = "=IF(OR(ISBLANK(T_Factures[[#This Row],[Date transmission AC]]),ISBLANK(T_Factures[[#This Row],[Date paiement]])),0,T_Factures[[#This Row],[Date paiement]]-T_Factures[[#This Row],[Date transmission AC]])"
                Case "Jours depuis réception"
                    FormuleModele = "=IF(ISBLANK(T_Factures[[#This Row],[Date réception]]),0,IF(T_Factures[[#This Row],[Étape en cours]]=""Payée"",T_Factures[[#This Row],[Date paiement]]-T_Factures[[#This Row],[Date réception]],DateRef-T_Factures[[#This Row],[Date réception]]))"
                Case "Retard contractuel (j)"
                    FormuleModele = "=IF(T_Factures[[#This Row],[Étape en cours]]=""Payée"",0,MAX(0,T_Factures[[#This Row],[Jours depuis réception]]-T_Factures[[#This Row],[Délai contractuel de paiement]]))"
                Case "DLM liée"
                    FormuleModele = "=IFERROR(INDEX(T_DLM[DLM en vigueur],MATCH(T_Factures[[#This Row],[N° accord]],T_DLM[N° accord],0)),0)"
                Case "Jours avant DLM"
                    FormuleModele = "=IF(T_Factures[[#This Row],[DLM liée]]<=0,9999,T_Factures[[#This Row],[DLM liée]]-DateRef)"
            End Select
    End Select
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

    Set lr = AjouterLigneTableau(loJrn)
    With lr.Range
        .Cells(1, loJrn.ListColumns("Horodatage").Index).Value = Now
        .Cells(1, loJrn.ListColumns("Horodatage").Index).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(1, loJrn.ListColumns("Utilisateur").Index).Value = "'" & nomUtilisateur
        .Cells(1, loJrn.ListColumns("Macro").Index).Value = "'" & nomMacro
        .Cells(1, loJrn.ListColumns("Résultat").Index).Value = "'" & Left$(resultat, 32000)
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
        cle = TexteCellule(loDet.DataBodyRange.Cells(i, idxId).Value)
        If StrComp(cle, Trim$(idDetenteur), vbTextCompare) = 0 Then
            ObtenirEmailDetenteur = TexteCellule(loDet.DataBodyRange.Cells(i, idxMail).Value)
            Exit Function
        End If
    Next i
    Exit Function

ErrEmail:
    ObtenirEmailDetenteur = ""
End Function
