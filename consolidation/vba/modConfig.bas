Option Explicit
Option Private Module
'==============================================================================
' modConfig - structure du classeur consolidé (feuilles, colonnes, règles)
' Seul module à adapter si une feuille, une colonne ou un seuil change.
'==============================================================================

Public Const APP_TITRE As String = "Pilotage consolidé DPER / ASER"
Public Const FEUILLE_ACCUEIL As String = "Sommaire"
Public Const FEUILLE_TDB As String = "Tableau_de_bord"
Public Const FEUILLE_REVUE As String = "Plan_Action"
Public Const FEUILLE_PARAM As String = "Paramètres"
Public Const FEUILLE_JOURNAL As String = "Journal"
Public Const FEUILLE_ARCHIVE As String = "Archive"
Public Const FEUILLE_NOTE As String = "Note_Escalade_DG"
Public Const FEUILLE_ALERTES As String = "Alertes"
Public Const FEUILLE_QC As String = "Contrôle_Qualité"
Public Const FEUILLES_PDF As String = "Tableau_de_bord;Plan_Action;Alertes"
Public Const MSG_LIBERER As String = "Bouton « Revenir à AUJOURDHUI() » pour rétablir la date automatique."

' --- DLM : feuille, lignes, colonnes ------------------------------------------
Public Const DLM_FEUILLE As String = "DLM"
Public Const DLM_PREMIERE As Long = 6
Public Const DLM_DERNIERE As Long = 205
Public Const DLM_COL_ID As String = "A"
Public Const DLM_COL_BAILLEUR As String = "B"
Public Const DLM_COL_ACCORD As String = "C"
Public Const DLM_COL_PROJET As String = "D"
Public Const DLM_COL_DLM As String = "H"
Public Const DLM_COL_SOLDE_FCFA As String = "P"
Public Const DLM_COL_JOURS As String = "Q"
Public Const DLM_COL_STATUT As String = "U"
Public Const DLM_COL_PIECE As String = "V"
Public Const DLM_COL_DETENTEUR As String = "W"
Public Const DLM_COL_DERNIERE As String = "X"
Public Const DLM_COL_NB As String = "Y"
Public Const DLM_COL_PROCHAINE As String = "Z"
Public Const DLM_COL_DUE As String = "AA"
Public Const DLM_COL_NIVEAU As String = "AB"
Public Const DLM_COL_ESCALADE As String = "AC"

' --- Annuaire des détenteurs (Paramètres) ------------------------------------
Public Const ANNU_PREMIERE As Long = 24
Public Const ANNU_DERNIERE As Long = 43
Public Const ANNU_COL_NOM As String = "M"
Public Const ANNU_COL_MAIL As String = "AD"

' --- Archive des clôtures ----------------------------------------------------
Public Const GAR_FEUILLE As String = "Garanties"
Public Const GAR_NB_COL As Long = 22            ' A:V
Public Const GAR_COL_STATUT As String = "O"
Public Const GAR_STATUT_CLOS As String = "Mainlevée accordée"
Public Const FAC_FEUILLE As String = "Factures"
Public Const FAC_NB_COL As Long = 32            ' A:AF
Public Const FAC_COL_TTC As String = "I"
Public Const FAC_COL_RESTE As String = "P"
Public Const ARCH_COL_GAR As Long = 1           ' bloc garanties : A..V + date
Public Const ARCH_COL_FAC As Long = 25          ' bloc factures : Y..BD + date
Public Const ARCH_PREMIERE As Long = 5

'------------------------------------------------------------------------------
' Règles d'alerte - une ligne par contrôle :
' Feuille|Ligne en-tête|1re ligne|Dernière ligne|Cols ID|Libellé|Col alerte|Motif|
' Jetons CRITIQUE|Jetons ATTENTION|Col responsable|Col jours
' Jetons séparés par « ; » : préfixe (casse indifférente), « =X » égalité stricte,
' « #>0 » nombre > 0, « #<=REF » date <= date de référence.
' Cols ID : première non vide ; à défaut « ligne n » (une ligne sans identifiant
' reste signalée). Libellé : « C,B » = 1re valeur non vide ; « B+E » = valeurs jointes.
'------------------------------------------------------------------------------
Public Function ReglesAlerte() As Collection
    Dim c As New Collection
    c.Add "Plan_Action|6|7|30|A|C|K|Mesure du plan COMDIR|EN RETARD|ÉCHÉANCE|E|J"
    c.Add "DLM|4|6|205|A|C,D,B|AB|Niveau d'alerte DLM|ÉCHUE;ESCALADE DG|VIGILANCE|W|Q"
    c.Add "DLM|4|6|205|A|C,D,B|AA|Relance due (pièce manquante)||=OUI|W|Q"
    c.Add "Garanties|4|6|205|A,H|B+E|R|Garantie (alerte avant échéance)|EXPIRÉE|ALERTE|T|P"
    c.Add "Factures|4|6|205|A,B|B+F|Z|Priorité de paiement (DLM)|P0|P1||Y"
    c.Add "Factures|4|6|205|A,B|B+F|W|Délai contractuel de paiement|HORS DÉLAI|||V"
    c.Add "Marchés|4|6|205|A|B|U|Avancement du marché|DÉLAI DÉPASSÉ;RETARD|VIGILANCE|S|"
    c.Add "Éval_Entreprises|4|6|205|A|E+D|N|Performance entreprise (classe)|=D|=C|P|"
    c.Add "Fiscal|4|6|205|A|B+F|I|Dossier fiscal incomplet||=NON|K|Q"
    c.Add "DAO_Standards|4|6|105|A|C|L|Document type en retard (jours)|#>0||I|L"
    c.Add "Protocole_MES|4|6|205|A|E+B|V|Mise en service / transfert|RÉSERVES NON LEVÉES;TRANSFERT EN RETARD|||U"
    c.Add "RACI|4|5|34|A|C|P|Matrice des responsabilités|1 SEUL A;R MANQUANT|À RENSEIGNER||"
    Set ReglesAlerte = c
End Function

'------------------------------------------------------------------------------
' Règles de contrôle qualité :
' Feuille|Ligne en-tête|1re ligne|Dernière ligne|Col ID (doublons)|Cols calculées|
' Cols dates saisies|Cols obligatoires|Cols témoins (ligne utilisée ; défaut = obligatoires)
'------------------------------------------------------------------------------
Public Function ReglesControle() As Collection
    Dim c As New Collection
    c.Add "Plan_Action|6|7|30|A|F,J,K|G|B,C,E|C"
    c.Add "DLM|4|6|205|A|L,M,N,O,P,Q,Z,AA,AB,AC|E,F,H,S,T,X,AD|B,H|B,C,H,J,V"
    c.Add "Garanties|4|6|205|A|C,D,E,K,P,Q,R|L,M,U|B,F,J,M|B,F,H,J,M"
    c.Add "Factures|4|6|205|A|D,E,F,G,P,Q,R,S,T,U,V,W,X,Y,Z,AE,AF|J,K,L,M,AA|B,C,I,J|B,C,I,J"
    c.Add "Marchés|4|6|205|A|I,L,N,O,P,Q,U,V,W,X|J,T|B,D,F,G,J,K|B,D,F,G"
    c.Add "Éval_Entreprises|4|6|205|A|E,F,M,N,O|B|B,D|B,D,G"
    c.Add "Fiscal|4|6|205|A|C,D,E,Q|N,O|B,F,I|B,F"
    c.Add "DAO_Standards|4|6|105|A|L|J,K|B,C,I,J|B,C"
    c.Add "Protocole_MES|4|6|205|A|C,D,S,T,U,V|G,H,I,J,K,L,M,N,O,P|B,E,F|B,E"
    c.Add "RACI|4|5|34|C|N,O,P|||D,E,F,G,H,I,J,K,L,M"
    Set ReglesControle = c
End Function

'------------------------------------------------------------------------------
' Registres (nouvelle ligne, effacement des exemples) :
' Feuille|Ligne exemple|1re ligne|Dernière ligne|Dernière colonne|Col ID|Préfixe ID|Chiffres
'------------------------------------------------------------------------------
Public Function Registres() As Collection
    Dim c As New Collection
    c.Add "DLM|5|6|205|AF|A|DLM-|2"
    c.Add "Garanties|5|6|205|V|A|GAR-|3"
    c.Add "Factures|5|6|205|AF|A|FAC-|3"
    c.Add "Marchés|5|6|205|Y|A||0"
    c.Add "Éval_Entreprises|5|6|205|Q|A|EV-|3"
    c.Add "Fiscal|5|6|205|R|A|FIS-|3"
    c.Add "DAO_Standards|5|6|105|N|A|DOC-|2"
    c.Add "Protocole_MES|5|6|205|W|A|MES-|3"
    Set Registres = c
End Function

'------------------------------------------------------------------------------
' Contrôles propres à ce classeur
'------------------------------------------------------------------------------
Public Sub ControlesSpecifiques()
    Dim ws As Worksheet, r As Long, dev As String, x As Double, y As Double, id As String
    Dim vus As New Collection, cle As String, prem As Variant

    If Not DateReferenceValide() Then
        QC_Ajouter FEUILLE_PARAM, "C5", "", "Date de référence invalide", _
            "Paramètres!C5 ne contient pas une date : tous les calculs de délais sont faux. " & MSG_LIBERER, "BLOQUANT"
    End If

    ' Date de lancement du plan : sans elle, aucune échéance du plan n'est calculée
    If Len(Txt(ThisWorkbook.Worksheets("Plan_Action").Range("C4").Value)) = 0 Then
        QC_Ajouter "Plan_Action", "C4", "", "Paramètre manquant", _
            "Date de lancement du plan non renseignée : échéances, jours restants et alertes du plan ne sont pas calculés.", "BLOQUANT"
    End If

    ' DLM : devises sans taux, montants et dates incohérents, e-mail du détenteur
    Set ws = ThisWorkbook.Worksheets(DLM_FEUILLE)
    For r = DLM_PREMIERE To DLM_DERNIERE
        id = IdLigne(ws, r, "A")
        dev = Txt(ws.Cells(r, "I").Value)
        If Len(dev) > 0 Then
            If TauxManquant(dev) Then
                QC_Ajouter ws.Name, "I" & r, id, "Taux de change manquant", _
                    "Devise « " & dev & " » sans taux dans Paramètres (G12:H19) : solde en FCFA non calculé.", "À CORRIGER"
            End If
        End If
        If EnNombre(ws.Cells(r, "K").Value, x) And EnNombre(ws.Cells(r, "J").Value, y) Then
            If x > y Then QC_Ajouter ws.Name, "K" & r, id, "Montant incohérent", _
                "Montant décaissé supérieur au montant accordé.", "À CORRIGER"
        End If
        If EnNombre(ws.Cells(r, "H").Value, x) And EnNombre(ws.Cells(r, "E").Value, y) Then
            If x <= y Then QC_Ajouter ws.Name, "H" & r, id, "Dates incohérentes", _
                "DLM actuelle antérieure ou égale à la date d'entrée en vigueur.", "À CORRIGER"
        End If
        If EnNombre(ws.Cells(r, "F").Value, x) And EnNombre(ws.Cells(r, "H").Value, y) Then
            If x > y Then QC_Ajouter ws.Name, "F" & r, id, "Dates incohérentes", _
                "DLM initiale postérieure à la DLM actuelle.", "À CORRIGER"
        End If
        If Len(Txt(ws.Cells(r, DLM_COL_PIECE).Value)) > 0 Then
            If Len(Txt(ws.Cells(r, DLM_COL_DETENTEUR).Value)) = 0 Then
                QC_Ajouter ws.Name, DLM_COL_DETENTEUR & r, id, "Détenteur non renseigné", _
                    "Pièce manquante sans détenteur : la relance ne peut pas être adressée.", "À COMPLÉTER"
            ElseIf Len(EmailDetenteur(Txt(ws.Cells(r, DLM_COL_DETENTEUR).Value))) = 0 Then
                QC_Ajouter ws.Name, DLM_COL_DETENTEUR & r, id, "E-mail du détenteur manquant", _
                    "Aucune adresse pour « " & Txt(ws.Cells(r, DLM_COL_DETENTEUR).Value) & _
                    " » (Paramètres, colonne AD) : le brouillon de relance sera sans destinataire.", "INFO"
            End If
        End If
    Next r

    ' Garanties : devise sans taux, échéance antérieure à l'émission
    Set ws = ThisWorkbook.Worksheets(GAR_FEUILLE)
    For r = 6 To 205
        dev = Txt(ws.Cells(r, "I").Value)
        If Len(dev) > 0 And UCase$(dev) <> "XOF" Then
            If TauxManquant(dev) Then QC_Ajouter ws.Name, "I" & r, IdLigne(ws, r, "A"), "Taux de change manquant", _
                "Devise « " & dev & " » sans taux dans Paramètres : montant en FCFA non calculé.", "À CORRIGER"
        End If
        If EnNombre(ws.Cells(r, "M").Value, x) And EnNombre(ws.Cells(r, "L").Value, y) Then
            If x < y Then QC_Ajouter ws.Name, "M" & r, IdLigne(ws, r, "A"), "Dates incohérentes", _
                "Date d'échéance antérieure à la date d'émission.", "À CORRIGER"
        End If
    Next r

    ' Factures : enchaînement des dates, montant payé, N° de facture en doublon
    Set ws = ThisWorkbook.Worksheets(FAC_FEUILLE)
    For r = 6 To 205
        id = IdLigne(ws, r, "A")
        OrdreDates ws, r, "J", "K", id, "Date de visa"
        OrdreDates ws, r, "K", "L", id, "Date de transmission à l'AC"
        OrdreDates ws, r, "L", "M", id, "Date de paiement"
        If EnNombre(ws.Cells(r, "N").Value, x) Then
            If EnNombre(ws.Cells(r, "I").Value, y) Then
                If x > y Then QC_Ajouter ws.Name, "N" & r, id, "Montant incohérent", _
                    "Montant payé supérieur au montant TTC.", "À CORRIGER"
            End If
        End If
        cle = UCase$(Txt(ws.Cells(r, "B").Value))
        If Len(cle) > 0 Then
            If Not AjouterCleLigne(vus, cle, r, prem) Then
                QC_Ajouter ws.Name, "B" & r, id, "N° de facture en doublon", _
                    "Même N° qu'en ligne " & prem & " : le plan de paiement et les recherches ne distinguent pas les deux lignes.", "BLOQUANT"
            End If
        End If
    Next r

    ' Pondérations : total attendu = 100 %
    If EnNombre(ThisWorkbook.Worksheets("Éval_Entreprises").Range("M3").Value, x) Then
        If Abs(x - 1) > 0.0001 Then
            QC_Ajouter "Éval_Entreprises", "M3", "", "Pondérations", _
                "Total des pondérations = " & Format$(x, "0%") & " (attendu : 100 %).", "BLOQUANT"
        End If
    End If
End Sub

' Une date d'étape saisie suppose l'étape précédente renseignée et antérieure.
Private Sub OrdreDates(ByVal ws As Worksheet, ByVal r As Long, ByVal colAvant As String, ByVal colApres As String, _
                       ByVal id As String, ByVal etape As String)
    Dim a As Double, b As Double
    If Len(Txt(ws.Cells(r, colApres).Value)) = 0 Then Exit Sub
    If Len(Txt(ws.Cells(r, colAvant).Value)) = 0 Then
        QC_Ajouter ws.Name, colApres & r, id, "Étape précédente manquante", _
            etape & " saisie alors que l'étape précédente (colonne " & colAvant & ") est vide.", "À CORRIGER"
    ElseIf EnNombre(ws.Cells(r, colAvant).Value, a) And EnNombre(ws.Cells(r, colApres).Value, b) Then
        If b < a Then QC_Ajouter ws.Name, colApres & r, id, "Dates incohérentes", _
            etape & " antérieure à l'étape précédente (colonne " & colAvant & ").", "À CORRIGER"
    End If
End Sub

Private Function AjouterCleLigne(ByVal col As Collection, ByVal cle As String, ByVal r As Long, ByRef premiere As Variant) As Boolean
    On Error Resume Next
    col.Add r, cle
    If Err.Number <> 0 Then
        Err.Clear
        premiere = col.Item(cle)
        AjouterCleLigne = False
    Else
        AjouterCleLigne = True
    End If
End Function

Public Function IdLigne(ByVal ws As Worksheet, ByVal r As Long, ByVal colId As String) As String
    IdLigne = Txt(ws.Cells(r, colId).Value)
    If Len(IdLigne) = 0 Then IdLigne = "ligne " & r
End Function

Private Function TauxManquant(ByVal dev As String) As Boolean
    Dim ws As Worksheet, r As Long, x As Double
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PARAM)
    TauxManquant = True
    For r = 12 To 19
        If UCase$(Txt(ws.Cells(r, "G").Value)) = UCase$(dev) Then
            If EnNombre(ws.Cells(r, "H").Value, x) Then TauxManquant = (x <= 0)
            Exit Function
        End If
    Next r
End Function

' E-mail d'un détenteur (Paramètres : nom colonne M, e-mail colonne AD, même ligne).
Public Function EmailDetenteur(ByVal nom As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PARAM)
    For r = ANNU_PREMIERE To ANNU_DERNIERE
        If UCase$(Txt(ws.Cells(r, ANNU_COL_NOM).Value)) = UCase$(Trim$(nom)) And Len(nom) > 0 Then
            EmailDetenteur = Txt(ws.Cells(r, ANNU_COL_MAIL).Value)
            Exit Function
        End If
    Next r
End Function

'------------------------------------------------------------------------------
' Date de référence : Paramètres!C5 (= D_JOUR)
'------------------------------------------------------------------------------
Public Function DateReferenceValide() As Boolean
    Dim v As Variant, x As Double
    v = ThisWorkbook.Worksheets(FEUILLE_PARAM).Range("C5").Value
    If EnNombre(v, x) Then DateReferenceValide = (x >= 1 And x < 2958466)
End Function

' Date de référence ; date du jour si la cellule est invalide (signalé par le contrôle qualité).
Public Function DateReference() As Date
    Dim v As Variant
    If DateReferenceValide() Then
        v = ThisWorkbook.Worksheets(FEUILLE_PARAM).Range("C5").Value
        DateReference = CDate(v)
    Else
        DateReference = Date
    End If
End Function

Public Function DateReferenceTexte() As String
    If DateReferenceValide() Then
        DateReferenceTexte = Format$(DateReference(), "dd/mm/yyyy")
    Else
        DateReferenceTexte = "INVALIDE (Paramètres!C5)"
    End If
End Function

Public Function DateReferenceFigee() As Boolean
    DateReferenceFigee = Not ThisWorkbook.Worksheets(FEUILLE_PARAM).Range("C5").HasFormula
End Function

Public Sub EcrireDateReference(ByVal d As Date)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PARAM)
    Deproteger ws
    ws.Range("C5").Value = d
    ws.Range("C5").NumberFormat = "dd/mm/yyyy"
    ws.Range("E5").Value = "Figée pour la revue (macro)"
End Sub

Public Sub RetablirDateAuto()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(FEUILLE_PARAM)
    Deproteger ws
    ws.Range("C5").Formula = "=TODAY()"
    ws.Range("C5").NumberFormat = "dd/mm/yyyy"
    ws.Range("E5").Value = "Automatique"
End Sub
