# -*- coding: utf-8 -*-
"""
Construit le classeur consolidé (sans VBA) à partir des deux classeurs d'origine :
  - base : Outil_Pilotage_Plan_Amelioration_DPER (feuilles conservées telles quelles
    ou corrigées) ;
  - apports ASER : montants en devise des garanties, paiements partiels, délais par
    étape et retard contractuel des factures, DLM initiale / prorogations / nombre de
    relances, tableau de bord graphique, journal des macros, archive des clôtures,
    note d'escalade DG, annuaire e-mail des détenteurs.

Usage : python construire_classeur.py <DPER.xlsm> <sortie.xlsx>
"""
import sys
import datetime as dt
from copy import copy

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

SRC, OUT = sys.argv[1], sys.argv[2]

# ---------------------------------------------------------------------------
# Styles (conventions du classeur DPER)
# ---------------------------------------------------------------------------
BLEU_NUIT, GRIS_FONCE, GRIS_TXT = 'FF1F3864', 'FF595959', 'FF7F7F7F'
BORD = Border(*(Side(style='thin', color='FFD9D9D9'),) * 4)
F_TITRE = Font(name='Arial', size=14, bold=True, color=BLEU_NUIT)
F_SOUS = Font(name='Arial', size=9, italic=True, color=GRIS_FONCE)
F_SECTION = Font(name='Arial', size=10, bold=True, color=BLEU_NUIT)
F_ENTETE = Font(name='Arial', size=9, bold=True, color='FFFFFFFF')
F_SAISIE = Font(name='Arial', size=9, color='FF0000FF')
F_CALC = Font(name='Arial', size=9, color='FF000000')
F_EX = Font(name='Arial', size=9, italic=True, color=GRIS_TXT)
F_NORMAL = Font(name='Arial', size=9)
FILL_ENT_SAISIE = PatternFill('solid', fgColor=BLEU_NUIT)
FILL_ENT_CALC = PatternFill('solid', fgColor=GRIS_FONCE)
FILL_CALC = PatternFill('solid', fgColor='FFF2F2F2')
FILL_EX = PatternFill('solid', fgColor='FFE7E6E6')
FILL_PARAM = PatternFill('solid', fgColor='FFFFF2CC')
AL_ENT = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL_CELL = Alignment(vertical='center')
AL_WRAP = Alignment(vertical='center', wrap_text=True)
NF_DATE = 'mm-dd-yy'                      # format de date court du poste (jj/mm/aaaa en français)
NF_MONT = '#,##0;\\-#,##0;\\-'
NF_PCT = '0.0%'
NF_J = '0'
DEVERROUILLE = Protection(locked=False)

wb = openpyxl.load_workbook(SRC)          # VBA réinjecté ensuite par assembler.py
WB_DPER = openpyxl.load_workbook(SRC)

# Jeu de règles de mise en forme « statut » du classeur DPER (37 règles par colonne)
_REGLES_STATUT = []
for cf in WB_DPER['DLM'].conditional_formatting:
    if str(cf.sqref).startswith('Y5'):
        for r in cf.rules:
            f = r.formula[0]                       # LEFT(Y5,n)="TOKEN"
            n = int(f.split(',')[1].split(')')[0])
            tok = f.split('="', 1)[1][:-1]
            _REGLES_STATUT.append((n, tok, r.dxf, r.stopIfTrue))
assert len(_REGLES_STATUT) == 37, len(_REGLES_STATUT)


def mfc_statut(ws, plage):
    """Applique le jeu de couleurs de statut DPER sur une plage (1re cellule relative)."""
    premiere = plage.split(':')[0].replace('$', '')
    for n, tok, dxf, stop in _REGLES_STATUT:
        regle = Rule(type='expression', dxf=copy(dxf), stopIfTrue=stop,
                     formula=['LEFT(%s,%d)="%s"' % (premiere, n, tok)])
        ws.conditional_formatting.add(plage, regle)


def mfc_expr(ws, plage, formule, couleur_police=None, fond=None, gras=False):
    dxf = DifferentialStyle(
        font=Font(color=couleur_police, bold=gras) if (couleur_police or gras) else None,
        fill=PatternFill(bgColor=fond, fill_type='solid') if fond else None)
    ws.conditional_formatting.add(plage, Rule(type='expression', dxf=dxf, formula=[formule]))


def titre(ws, t, sous=None):
    ws['A1'] = t
    ws['A1'].font = F_TITRE
    ws.row_dimensions[1].height = 17.65
    if sous:
        ws['A2'] = sous
        ws['A2'].font = F_SOUS


def entete(c, texte, saisie=True):
    c.value = texte
    c.font = F_ENTETE
    c.fill = FILL_ENT_SAISIE if saisie else FILL_ENT_CALC
    c.alignment = AL_ENT
    c.border = BORD


def style_cellule(c, kind, nf=None, exemple=False):
    c.border = BORD
    c.alignment = AL_CELL
    if exemple:
        c.font, c.fill = F_EX, FILL_EX
        if kind == 'in':
            c.protection = DEVERROUILLE
    elif kind == 'in':
        c.font = F_SAISIE
        c.protection = DEVERROUILLE
    else:
        c.font, c.fill = F_CALC, FILL_CALC
    if nf:
        c.number_format = nf


def dv(ws, sqref, **kw):
    v = DataValidation(allow_blank=True, showErrorMessage=True, **kw)
    if v.type == 'list':
        v.errorTitle, v.error = 'Saisie', 'Valeur hors liste (voir feuille Paramètres).'
    elif v.type == 'date':
        v.errorTitle, v.error = 'Date', 'Saisir une date (jj/mm/aaaa).'
    ws.add_data_validation(v)
    for plage in sqref.split():
        v.add(plage)
    return v


def dv_date(ws, sqref):
    return dv(ws, sqref, type='date', operator='between', formula1='36526', formula2='73050')


def dv_liste(ws, sqref, nom):
    return dv(ws, sqref, type='list', formula1=nom)


def dv_entier(ws, sqref):
    v = dv(ws, sqref, type='whole', operator='greaterThanOrEqual', formula1='0')
    v.errorTitle, v.error = 'Nombre', 'Saisir un nombre entier positif ou nul.'
    return v


def dv_montant(ws, sqref):
    v = dv(ws, sqref, type='decimal', operator='greaterThanOrEqual', formula1='0')
    v.errorTitle, v.error = 'Montant', 'Saisir un montant positif ou nul.'
    return v


def dv_unique(ws, col, r1, r2, titre_err, msg):
    v = dv(ws, '%s%d:%s%d' % (col, r1, col, r2), type='custom',
           formula1='COUNTIF($%s$%d:$%s$%d,%s%d)=1' % (col, r1, col, r2, col, r1))
    v.errorTitle, v.error = titre_err, msg
    return v


def note(c, texte):
    cm = Comment(texte, 'Pilotage')
    cm.width, cm.height = 260, 110
    c.comment = cm


# ---------------------------------------------------------------------------
# Registre générique : en-tête ligne 4, exemple ligne 5, données 6..r2
# ---------------------------------------------------------------------------
LEGENDE = ('Police bleue = saisie · fond gris = calcul automatique (protégé) · en-tête bleu = saisie, '
           'en-tête gris = calcul · ligne grisée en italique = exemple de format, exclue des calculs · '
           'seuils et hypothèses : feuille « Paramètres ».')


def registre(ws, t, colonnes, exemple, r2=205, figer='C5', hauteur_ent=48):
    """colonnes : liste de dict l, h, k ('in'|'calc'), w, nf, f (gabarit {r}), note."""
    titre(ws, t, LEGENDE)
    for col in colonnes:
        l = col['l']
        c = ws['%s4' % l]
        entete(c, col['h'], col['k'] == 'in')
        if col.get('note'):
            note(c, col['note'])
        ws.column_dimensions[l].width = col.get('w', 12)
        for r in range(5, r2 + 1):
            cell = ws['%s%d' % (l, r)]
            style_cellule(cell, col['k'], col.get('nf'), exemple=(r == 5))
            if col['k'] == 'calc':
                cell.value = col['f'].format(r=r)
        if l in exemple:
            ws['%s5' % l].value = exemple[l]
    ws.row_dimensions[4].height = hauteur_ent
    ws.row_dimensions[5].height = 23.25
    ws.freeze_panes = figer
    der = colonnes[-1]['l']
    ws.auto_filter.ref = 'A4:%s%d' % (der, r2)
    ws.sheet_view.zoomScale = 90
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = '4:4'


def remplacer_feuille(nom, index=None):
    ancien = wb[nom]
    idx = wb.sheetnames.index(nom)
    wb.remove(ancien)
    ws = wb.create_sheet(nom, idx if index is None else index)
    return ws


def lookup_marche(col_cle, col_marche, r):
    return ('=IF({k}{r}="","",IFERROR(INDEX(Marchés!${m}$5:${m}$205,MATCH({k}{r},Marchés!$A$5:$A$205,0))&"",'
            '"Marché inconnu"))').format(k=col_cle, m=col_marche, r=r)


# ===========================================================================
# 1. DLM (fusion DPER « DLM » + ASER « DLM_Prorogations »)
# ===========================================================================
ws = remplacer_feuille('DLM')
DLM_COLS = [
    dict(l='A', h='ID accord', k='in', w=9, note='Identifiant unique (ex. DLM-05) : clé utilisée par Marchés, Factures et Scénarios_BID.'),
    dict(l='B', h='Bailleur', k='in', w=9),
    dict(l='C', h="N° / intitulé de l'accord de financement", k='in', w=30),
    dict(l='D', h='Projet(s) concerné(s)', k='in', w=24),
    dict(l='E', h="Date d'entrée en vigueur", k='in', w=11, nf=NF_DATE),
    dict(l='F', h='DLM initiale', k='in', w=11, nf=NF_DATE),
    dict(l='G', h='Nb prorogations obtenues', k='in', w=10, nf=NF_J),
    dict(l='H', h='DLM actuelle', k='in', w=11, nf=NF_DATE, note="DLM en vigueur. Prorogation accordée : saisir ici la nouvelle date et incrémenter le nombre de prorogations."),
    dict(l='I', h='Devise', k='in', w=7),
    dict(l='J', h='Montant accordé (devise)', k='in', w=15, nf=NF_MONT),
    dict(l='K', h='Montant décaissé (devise)', k='in', w=15, nf=NF_MONT),
    dict(l='L', h='Solde non décaissé (devise)', k='calc', w=15, nf=NF_MONT,
         f='=IF(J{r}="","",MAX(0,J{r}-N(K{r})))'),
    dict(l='M', h='Taux de décais-sement', k='calc', w=9, nf=NF_PCT,
         f='=IF(N(J{r})=0,"",N(K{r})/J{r})'),
    dict(l='N', h='Délai consommé', k='calc', w=9, nf=NF_PCT,
         f='=IF(OR(E{r}="",H{r}=""),"",IF(N(H{r})<=N(E{r}),"",MAX(0,MIN(1,(D_JOUR-E{r})/(H{r}-E{r})))))',
         note="Part du délai écoulée entre l'entrée en vigueur et la DLM actuelle."),
    dict(l='O', h='Écart décais-sement − délai (pts)', k='calc', w=10, nf=NF_PCT,
         f='=IF(OR(M{r}="",N{r}=""),"",M{r}-N{r})', note='Négatif : décaissement en retard sur le temps écoulé.'),
    dict(l='P', h='Solde non décaissé (FCFA)', k='calc', w=16, nf=NF_MONT,
         f='=IF(OR(L{r}="",I{r}=""),"",IFERROR(IF(INDEX(T_TAUX,MATCH(I{r},L_DEV,0))="","",ROUND(L{r}*INDEX(T_TAUX,MATCH(I{r},L_DEV,0)),0)),""))',
         note='Vide si le taux de la devise manque dans Paramètres (signalé par le contrôle qualité).'),
    dict(l='Q', h='Jours avant DLM', k='calc', w=8, nf=NF_J, f='=IF(H{r}="","",H{r}-D_JOUR)'),
    dict(l='R', h='Proroga-tion requise ?', k='in', w=9),
    dict(l='S', h='Nouvelle DLM demandée', k='in', w=11, nf=NF_DATE),
    dict(l='T', h='Date de dépôt de la demande', k='in', w=11, nf=NF_DATE),
    dict(l='U', h='Statut de la demande', k='in', w=19),
    dict(l='V', h='Pièce manquante', k='in', w=28),
    dict(l='W', h='Détenteur de la pièce', k='in', w=12),
    dict(l='X', h='Dernière relance', k='in', w=11, nf=NF_DATE,
         note='Mise à jour par les macros « Préparer les relances » / « Enregistrer une relance » (ou saisie manuelle).'),
    dict(l='Y', h='Nb relances', k='in', w=8, nf=NF_J, note='Incrémenté par les macros de relance.'),
    dict(l='Z', h='Prochaine relance', k='calc', w=11, nf=NF_DATE,
         f='=IF(OR(V{r}="",U{r}="Accordée",U{r}="Clôturé",U{r}="Non requise"),"",IF(X{r}="",D_JOUR,X{r}+P_FREQ))',
         note='Dernière relance + fréquence (Paramètres) ; date de référence si aucune relance.'),
    dict(l='AA', h='Relance due', k='calc', w=8, f='=IF(Z{r}="","",IF(Z{r}<=D_JOUR,"OUI",""))'),
    dict(l='AB', h="Niveau d'alerte DLM", k='calc', w=15,
         f='=IF(H{r}="","",IF(U{r}="Clôturé","CLÔTURÉ",IF(Q{r}<0,"ÉCHUE",IF(Q{r}<=P_ESC,"ESCALADE DG",'
           'IF(AND(V{r}<>"",N(Y{r})>=P_NBREL),"ESCALADE DG – RELANCES",IF(Q{r}<=P_VIG,"VIGILANCE","OK"))))))',
         note='ÉCHUE · ESCALADE DG (DLM ≤ seuil) · ESCALADE DG – RELANCES (pièce manquante et nb de relances ≥ seuil) · VIGILANCE · OK.'),
    dict(l='AC', h='Escalade DG requise', k='calc', w=9,
         f='=IF(AB{r}="","",IF(OR(AB{r}="ÉCHUE",LEFT(AB{r},11)="ESCALADE DG"),"OUI","NON"))'),
    dict(l='AD', h="Date d'escalade au DG", k='in', w=11, nf=NF_DATE),
    dict(l='AE', h='Instruction / décision du DG', k='in', w=26),
    dict(l='AF', h='Prochaine action / commentaire', k='in', w=32),
]
registre(ws, 'Suivi des DLM et des demandes de prorogation — BID · BOAD · FVC · ADFD', DLM_COLS,
         dict(A='EX', B='BID', C='Accord n° XXXX (exemple)', D='Projet exemple', E=dt.datetime(2022, 1, 15),
              F=dt.datetime(2026, 12, 31), G=1, H=dt.datetime(2027, 6, 30), I='EUR', J=20000000, K=9000000,
              R='Oui', S=dt.datetime(2028, 6, 30), U='À déposer', V='Plan de décaissement actualisé', W='DPER',
              X=dt.datetime(2026, 10, 20), Y=1, AF='Finaliser le dossier et transmettre au bailleur'),
         figer='D5')
for i, b in enumerate(['BID', 'BOAD', 'FVC', 'ADFD']):
    ws['A%d' % (6 + i)] = 'DLM-%02d' % (i + 1)
    ws['B%d' % (6 + i)] = b
dv_unique(ws, 'A', 6, 205, "Identifiant d'accord en doublon",
          'Cet identifiant existe déjà dans la colonne A : les recherches ne liraient que la première ligne.')
dv_liste(ws, 'B6:B205', 'L_BAILLEUR')
dv_date(ws, 'E6:F205 H6:H205 S6:T205 X6:X205 AD6:AD205')
dv_entier(ws, 'G6:G205 Y6:Y205')
dv_liste(ws, 'I6:I205', 'L_DEV')
dv_montant(ws, 'J6:J205')
v = dv(ws, 'K6:K205', type='custom', formula1='OR(K6="",AND(ISNUMBER(K6),K6>=0,OR(J6="",K6<=J6)))')
v.errorTitle, v.error = 'Montant décaissé', 'Le montant décaissé doit être positif et ne pas dépasser le montant accordé.'
dv_liste(ws, 'R6:R205', 'L_OUINON')
dv_liste(ws, 'U6:U205', 'L_STAT_DLM')
dv_liste(ws, 'W6:W205', 'L_DETENTEUR')
for l in ('AA', 'AB', 'AC'):
    mfc_statut(ws, '%s5:%s205' % (l, l))
mfc_expr(ws, 'K6:K205', 'AND(ISNUMBER(K6),ISNUMBER(J6),K6>J6)', 'FFFF0000', gras=True)
mfc_expr(ws, 'H6:H205', 'AND(ISNUMBER(H6),ISNUMBER(E6),H6<=E6)', 'FFFF0000', gras=True)
ws.sheet_properties.tabColor = 'FFC00000'

# ===========================================================================
# 2. Garanties (DPER + devise / montant FCFA ASER)
# ===========================================================================
ws = remplacer_feuille('Garanties')
GAR_COLS = [
    dict(l='A', h='ID', k='in', w=9),
    dict(l='B', h='N° marché', k='in', w=13),
    dict(l='C', h='Projet', k='calc', w=18, f=lookup_marche('B', 'C', '{r}')),
    dict(l='D', h='Bailleur', k='calc', w=9, f=lookup_marche('B', 'D', '{r}')),
    dict(l='E', h='Entreprise titulaire', k='calc', w=20, f=lookup_marche('B', 'F', '{r}')),
    dict(l='F', h='Type de garantie', k='in', w=20),
    dict(l='G', h='Banque / établissement émetteur', k='in', w=18),
    dict(l='H', h='N° de la garantie', k='in', w=13),
    dict(l='I', h='Devise', k='in', w=7, note='Vide = FCFA (XOF).'),
    dict(l='J', h='Montant (devise)', k='in', w=15, nf=NF_MONT),
    dict(l='K', h='Montant (FCFA)', k='calc', w=15, nf=NF_MONT,
         f='=IF(J{r}="","",IF(OR(I{r}="",I{r}="XOF"),J{r},IFERROR(IF(INDEX(T_TAUX,MATCH(I{r},L_DEV,0))="","",ROUND(J{r}*INDEX(T_TAUX,MATCH(I{r},L_DEV,0)),0)),"")))',
         note='Vide si le taux de la devise manque dans Paramètres.'),
    dict(l='L', h="Date d'émission", k='in', w=11, nf=NF_DATE),
    dict(l='M', h="Date d'échéance", k='in', w=11, nf=NF_DATE, note="Garantie prorogée : saisir la nouvelle échéance et le statut « Prorogée »."),
    dict(l='N', h='Événement libératoire', k='in', w=16),
    dict(l='O', h='Statut', k='in', w=16),
    dict(l='P', h='Jours avant échéance', k='calc', w=9, nf=NF_J, f='=IF(M{r}="","",M{r}-D_JOUR)'),
    dict(l='Q', h="Date de déclenche-ment de l'alerte", k='calc', w=11, nf=NF_DATE, f='=IF(M{r}="","",M{r}-P_GAR)'),
    dict(l='R', h='Alerte', k='calc', w=11,
         f='=IF(M{r}="","",IF(OR(O{r}="Mainlevée accordée",O{r}="Appelée"),"CLÔTURÉE",IF(P{r}<0,"EXPIRÉE",IF(P{r}<=P_GAR,"ALERTE","OK"))))',
         note='EXPIRÉE : échéance dépassée, statut encore actif · ALERTE : échéance ≤ seuil (Paramètres) · CLÔTURÉE : mainlevée accordée ou garantie appelée.'),
    dict(l='S', h='Action requise', k='in', w=20),
    dict(l='T', h='Responsable', k='in', w=12),
    dict(l='U', h='Date de la dernière action', k='in', w=11, nf=NF_DATE),
    dict(l='V', h='Commentaires', k='in', w=30),
]
registre(ws, "Registre des garanties — garanties de bonne exécution et cautions d'avance (alerte avant échéance)",
         GAR_COLS, dict(A='EX', B='M-EXEMPLE', F='Garantie de bonne exécution', G='Banque (exemple)', H='GBE-0000',
                        I='XOF', J=25000000, L=dt.datetime(2025, 12, 1), M=dt.datetime(2026, 12, 15),
                        N='Réception provisoire', O='Valide', S='Demander la prorogation', T='DPER'))
dv_unique(ws, 'A', 6, 205, 'Identifiant en doublon', 'Cet identifiant de garantie existe déjà.')
dv_liste(ws, 'B6:B205', 'L_MARCHE')
dv_liste(ws, 'F6:F205', 'L_TYPE_GAR')
dv_liste(ws, 'I6:I205', 'L_DEV')
dv_montant(ws, 'J6:J205')
dv_date(ws, 'L6:M205 U6:U205')
dv_liste(ws, 'O6:O205', 'L_STAT_GAR')
dv_liste(ws, 'S6:S205', 'L_ACT_GAR')
dv_liste(ws, 'T6:T205', 'L_DETENTEUR')
for l in ('C', 'R'):
    mfc_statut(ws, '%s5:%s205' % (l, l))
mfc_expr(ws, 'M6:M205', 'AND(ISNUMBER(M6),ISNUMBER(L6),M6<L6)', 'FFFF0000', gras=True)
ws.sheet_properties.tabColor = 'FFC00000'

# ===========================================================================
# 3. Factures (DPER + paiements partiels, délais par étape, retard ASER)
# ===========================================================================
ws = remplacer_feuille('Factures')
FAC_COLS = [
    dict(l='A', h='ID', k='in', w=9),
    dict(l='B', h='N° facture / décompte', k='in', w=18, note='Clé du plan de paiement : doit être unique.'),
    dict(l='C', h='N° marché', k='in', w=13),
    dict(l='D', h='Projet', k='calc', w=16, f=lookup_marche('C', 'C', '{r}')),
    dict(l='E', h='Bailleur', k='calc', w=9, f=lookup_marche('C', 'D', '{r}')),
    dict(l='F', h='Entreprise', k='calc', w=18, f=lookup_marche('C', 'F', '{r}')),
    dict(l='G', h='Réf. accord (ID DLM)', k='calc', w=9, f=lookup_marche('C', 'E', '{r}')),
    dict(l='H', h='Nature', k='in', w=16),
    dict(l='I', h='Montant TTC (FCFA)', k='in', w=15, nf=NF_MONT),
    dict(l='J', h='Date de réception', k='in', w=11, nf=NF_DATE),
    dict(l='K', h='Date de visa', k='in', w=11, nf=NF_DATE),
    dict(l='L', h="Date de transmission à l'AC", k='in', w=11, nf=NF_DATE),
    dict(l='M', h='Date de paiement', k='in', w=11, nf=NF_DATE, note='Date du (dernier) paiement. Sans montant payé saisi, la facture est réputée payée en totalité.'),
    dict(l='N', h='Montant payé (FCFA)', k='in', w=15, nf=NF_MONT, note='Paiement partiel : saisir le cumul payé. Vide + date de paiement = payé en totalité.'),
    dict(l='O', h='Délai contractuel (j)', k='in', w=9, nf=NF_J, note='Vide = délai par défaut (Paramètres).'),
    dict(l='P', h='Reste à payer (FCFA)', k='calc', w=15, nf=NF_MONT,
         f='=IF(I{r}="","",IF(AND(M{r}<>"",N{r}=""),0,ROUND(MAX(0,I{r}-N(N{r})),2)))'),
    dict(l='Q', h='Étape', k='calc', w=13,
         f='=IF(J{r}="","",IF(AND(N(I{r})>0,P{r}=0),"5-Payé",IF(N(N{r})>0,"4-Payé partiellement",IF(L{r}<>"","3-Transmis AC",IF(K{r}<>"","2-Visé","1-Reçu")))))'),
    dict(l='R', h='Délai réception → visa (j)', k='calc', w=9, nf=NF_J, f='=IF(OR(J{r}="",K{r}=""),"",K{r}-J{r})'),
    dict(l='S', h='Délai visa → AC (j)', k='calc', w=9, nf=NF_J, f='=IF(OR(K{r}="",L{r}=""),"",L{r}-K{r})'),
    dict(l='T', h='Délai AC → paiement (j)', k='calc', w=9, nf=NF_J, f='=IF(OR(L{r}="",M{r}=""),"",M{r}-L{r})'),
    dict(l='U', h='Délai depuis réception (j)', k='calc', w=9, nf=NF_J,
         f='=IF(J{r}="","",IF(Q{r}="5-Payé",IF(M{r}<>"",M{r}-J{r},""),D_JOUR-J{r}))'),
    dict(l='V', h='Retard contractuel (j)', k='calc', w=9, nf=NF_J,
         f='=IF(OR(J{r}="",U{r}=""),"",MAX(0,U{r}-IF(O{r}="",P_PAY,O{r})))'),
    dict(l='W', h='Délai contractuel', k='calc', w=13,
         f='=IF(V{r}="","",IF(V{r}>0,IF(Q{r}="5-Payé","Payé hors délai","HORS DÉLAI"),""))'),
    dict(l='X', h="DLM de l'accord", k='calc', w=11, nf=NF_DATE,
         f='=IF(G{r}="","",IFERROR(IF(INDEX(DLM!$H$5:$H$205,MATCH(G{r},DLM!$A$5:$A$205,0))="","",INDEX(DLM!$H$5:$H$205,MATCH(G{r},DLM!$A$5:$A$205,0))),""))'),
    dict(l='Y', h='Jours avant DLM', k='calc', w=8, nf=NF_J, f='=IF(X{r}="","",X{r}-D_JOUR)'),
    dict(l='Z', h='Priorité de paiement', k='calc', w=9,
         f='=IF(OR(J{r}="",N(P{r})=0),"",IF(X{r}="","PX",IF(Y{r}<0,"P0",IF(Y{r}<=P_ESC,"P1",IF(Y{r}<=P_VIG,"P2","P3")))))',
         note='P0 : DLM échue · P1 : DLM ≤ seuil d’escalade · P2 : ≤ seuil de vigilance · P3 : au-delà · PX : DLM inconnue.'),
    dict(l='AA', h='Date cible de paiement (arrêtée AC/DFP)', k='in', w=11, nf=NF_DATE),
    dict(l='AB', h='Validé AC', k='in', w=8),
    dict(l='AC', h='Validé DFP', k='in', w=8),
    dict(l='AD', h='Commentaires', k='in', w=26),
    dict(l='AE', h='Clé DLM (technique)', k='calc', w=8, nf='0',
         f='=IF(Z{r}="","",IF(X{r}="",99999,X{r}))'),
    dict(l='AF', h='Rang de paiement (technique)', k='calc', w=8, nf='0',
         f='=IF(AE{r}="","",COUNTIFS($AE$6:$AE$205,"<"&AE{r})+COUNTIFS($AE$6:$AE$205,AE{r},$V$6:$V$205,">"&V{r})'
           '+COUNTIFS($AE$6:$AE$205,AE{r},$V$6:$V$205,V{r},$P$6:$P$205,">"&P{r})+COUNTIFS($AE$6:AE{r},AE{r},$V$6:V{r},V{r},$P$6:P{r},P{r}))',
         note='Règle : DLM la plus proche, puis retard contractuel décroissant, puis reste à payer décroissant, puis ordre de saisie.'),
]
registre(ws, "Registre des factures et décomptes — réception · visa · transmission à l'AC · paiement", FAC_COLS,
         dict(A='EX', B='Décompte n° 3 (exemple)', C='M-EXEMPLE', H='Décompte provisoire', I=45000000,
              J=dt.datetime(2026, 9, 15), K=dt.datetime(2026, 9, 25), AA=dt.datetime(2026, 11, 15), AB='Non', AC='Non'))
# la ligne d'exemple n'entre pas dans le classement du plan de paiement
ws['AF5'] = None
dv_unique(ws, 'A', 6, 205, 'Identifiant en doublon', 'Cet identifiant de facture existe déjà.')
dv_unique(ws, 'B', 6, 205, 'N° de facture en doublon', 'Ce numéro de facture / décompte existe déjà : le plan de paiement ne lirait que la première ligne.')
dv_liste(ws, 'C6:C205', 'L_MARCHE')
dv_liste(ws, 'H6:H205', 'L_NAT_FAC')
dv_montant(ws, 'I6:I205')
dv_date(ws, 'J6:M205 AA6:AA205')
v = dv(ws, 'N6:N205', type='custom', formula1='OR(N6="",AND(ISNUMBER(N6),N6>=0,OR(I6="",N6<=I6)))')
v.errorTitle, v.error = 'Montant payé', 'Le montant payé doit être positif et ne pas dépasser le montant TTC.'
dv_entier(ws, 'O6:O205')
dv_liste(ws, 'AB6:AC205', 'L_OUINON')
for l in ('D', 'Q', 'W', 'Z'):
    mfc_statut(ws, '%s5:%s205' % (l, l))
mfc_expr(ws, 'K6:K205', 'AND(K6<>"",OR(J6="",AND(ISNUMBER(J6),K6<J6)))', 'FFFF0000', gras=True)
mfc_expr(ws, 'L6:L205', 'AND(L6<>"",OR(K6="",AND(ISNUMBER(K6),L6<K6)))', 'FFFF0000', gras=True)
mfc_expr(ws, 'M6:M205', 'AND(M6<>"",OR(L6="",AND(ISNUMBER(L6),M6<L6)))', 'FFFF0000', gras=True)
mfc_expr(ws, 'N6:N205', 'AND(ISNUMBER(N6),N6>N(I6))', 'FFFF0000', gras=True)
mfc_expr(ws, 'V6:V205', 'AND(ISNUMBER(V6),V6>0)', 'FF9C0006', 'FFFFC7CE')
for l in ('AE', 'AF'):
    ws.column_dimensions[l].hidden = False
ws.sheet_properties.tabColor = 'FFC00000'

# ===========================================================================
# 4. Plan de paiement (entièrement calculé ; règle de priorité ASER)
# ===========================================================================
ws = remplacer_feuille('Plan_Paiement')
titre(ws, "Plan de paiement priorisé par échéance de DLM (établi avec l'AC et la DFP)",
      "Généré automatiquement depuis « Factures » (factures reçues dont le reste à payer est > 0) : DLM la plus proche, "
      "puis retard contractuel, puis montant. Dates cibles et validations AC / DFP : à saisir dans « Factures ».")
for l, h in zip('CDEF', ['Priorité', 'Définition', 'Nb', 'Reste à payer (FCFA)']):
    entete(ws['%s4' % l], h, False)
defs = [('P0', '="DLM échue"'), ('P1', '="DLM ≤ "&P_ESC&" jours"'), ('P2', '="DLM ≤ "&P_VIG&" jours"'),
        ('P3', '="DLM > "&P_VIG&" jours"'), ('PX', '="DLM non renseignée"')]
for i, (p, d) in enumerate(defs):
    r = 5 + i
    ws['C%d' % r] = p
    ws['C%d' % r].font = Font(name='Arial', size=9, bold=True, color=BLEU_NUIT)
    ws['D%d' % r] = d
    ws['E%d' % r] = '=COUNTIF(Factures!$Z$6:$Z$205,"%s")' % p
    ws['F%d' % r] = '=SUMIFS(Factures!$P$6:$P$205,Factures!$Z$6:$Z$205,"%s")' % p
    for l, nf in (('D', None), ('E', '0'), ('F', NF_MONT)):
        style_cellule(ws['%s%d' % (l, r)], 'calc', nf)
    ws['C%d' % r].border = BORD
ws['C10'] = 'Total'
ws['C10'].font = Font(name='Arial', size=9, bold=True, color=BLEU_NUIT)
ws['E10'], ws['F10'] = '=SUM(E5:E9)', '=SUM(F5:F9)'
for l, nf in (('E', '0'), ('F', NF_MONT)):
    style_cellule(ws['%s10' % l], 'calc', nf)
    ws['%s10' % l].font = Font(name='Arial', size=9, bold=True, color=BLEU_NUIT)
mfc_statut(ws, 'C5:C9')
ws['C12'] = 'Séance de priorisation du'
ws['C13'] = 'Participants AC / DFP / DPER'
for a in ('C12', 'C13'):
    ws[a].font = Font(name='Arial', size=9, bold=True, color=BLEU_NUIT)
for a in ('F12', 'F13', 'G13', 'H13'):
    ws[a].font, ws[a].fill, ws[a].border, ws[a].protection = F_SAISIE, FILL_PARAM, BORD, DEVERROUILLE
ws['F12'].number_format = NF_DATE
dv_date(ws, 'F12')
ws['I13'] = '← noms / visas'
ws['I13'].font = F_SOUS
PP = [('A', 'Rang', 6, '0'), ('B', 'Ligne source (technique)', 6, '0'), ('C', 'ID', 8, None),
      ('D', 'N° facture / décompte', 18, None), ('E', 'N° marché', 13, None), ('F', 'Entreprise', 20, None),
      ('G', 'Bailleur', 9, None), ('H', 'Réf. accord', 9, None), ('I', 'DLM', 11, NF_DATE),
      ('J', 'Jours avant DLM', 8, NF_J), ('K', 'Priorité', 8, None), ('L', 'Étape', 14, None),
      ('M', 'Reste à payer (FCFA)', 15, NF_MONT), ('N', 'Cumul (FCFA)', 16, NF_MONT),
      ('O', 'Délai depuis réception (j)', 9, NF_J), ('P', 'Retard contractuel (j)', 9, NF_J),
      ('Q', 'Date cible de paiement', 11, NF_DATE), ('R', 'Validé AC', 8, None), ('S', 'Validé DFP', 8, None),
      ('T', 'Action attendue de', 24, None)]
SRC_COL = dict(C='A', D='B', E='C', F='F', G='E', H='G', I='X', J='Y', K='Z', L='Q', M='P', O='U', P='V',
               Q='AA', R='AB', S='AC')
TEXTE = set('CDEFGHKLRS')
for l, h, w, nf in PP:
    entete(ws['%s15' % l], h, False)
    ws.column_dimensions[l].width = w
ws.row_dimensions[15].height = 36
for r in range(16, 216):
    k = r - 15
    ws['A%d' % r] = '=IF($B%d="","",%d)' % (r, k)
    ws['B%d' % r] = '=IFERROR(MATCH(%d,Factures!$AF$6:$AF$205,0),"")' % k
    for l, s in SRC_COL.items():
        rng = 'Factures!$%s$6:$%s$205' % (s, s)
        if l in TEXTE:
            ws['%s%d' % (l, r)] = '=IF($B{r}="","",INDEX({g},$B{r})&"")'.format(r=r, g=rng)
        else:
            ws['%s%d' % (l, r)] = '=IF($B{r}="","",IF(INDEX({g},$B{r})="","",INDEX({g},$B{r})))'.format(r=r, g=rng)
    ws['N%d' % r] = '=IF(M%d="","",SUM($M$16:M%d))' % (r, r)
    ws['T%d' % r] = ('=IF(L{r}="","",IF(L{r}="1-Reçu","DPER – visa",IF(L{r}="2-Visé","DPER – transmission à l\'AC",'
                     'IF(OR(L{r}="3-Transmis AC",L{r}="4-Payé partiellement"),"AC / DFP – paiement",""))))').format(r=r)
    for l, _h, _w, nf in PP:
        style_cellule(ws['%s%d' % (l, r)], 'calc', nf)
mfc_statut(ws, 'K16:K215')
mfc_expr(ws, 'P16:P215', 'AND(ISNUMBER(P16),P16>0)', 'FF9C0006', 'FFFFC7CE')
mfc_expr(ws, 'Q16:Q215', 'AND(ISNUMBER(Q16),ISNUMBER(I16),Q16>I16)', 'FFFF0000', gras=True)
ws.freeze_panes = 'C16'
ws.auto_filter.ref = 'A15:T215'
ws.sheet_view.zoomScale = 90
ws.sheet_view.showGridLines = False
ws.page_setup.orientation = 'landscape'
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
ws.print_title_rows = '15:15'
ws.sheet_properties.tabColor = 'FFC00000'

# ===========================================================================
# 5. Feuilles DPER conservées : corrections de références
# ===========================================================================
ws = wb['Scénarios_BID']
CORR = {'$G$5:$G$205': '$I$5:$I$205', '$H$5:$H$205': '$J$5:$J$205', '$I$5:$I$205': '$K$5:$K$205',
        '$J$5:$J$205': '$L$5:$L$205', '$F$5:$F$205': '$H$5:$H$205'}
for row in ws.iter_rows():
    for c in row:
        if isinstance(c.value, str) and 'DLM!' in c.value:
            v = c.value
            for a, b in CORR.items():
                v = v.replace('DLM!' + a, 'DLM!#' + b)    # marqueur : pas de double remplacement
            c.value = v.replace('DLM!#', 'DLM!')

ws = wb['Marchés']
for r in range(5, 206):
    ws['N%d' % r] = ('=IF(A{r}="","",SUMIFS(Factures!$I$6:$I$205,Factures!$C$6:$C$205,A{r})'
                     '-SUMIFS(Factures!$P$6:$P$205,Factures!$C$6:$C$205,A{r}))').format(r=r)
    ws['W%d' % r] = ('=IF(A{r}="","",COUNTIFS(Garanties!$B$6:$B$205,A{r},Garanties!$R$6:$R$205,"ALERTE")'
                     '+COUNTIFS(Garanties!$B$6:$B$205,A{r},Garanties!$R$6:$R$205,"EXPIRÉE"))').format(r=r)
    ws['X%d' % r] = ('=IF(A{r}="","",SUMIFS(Factures!$P$6:$P$205,Factures!$C$6:$C$205,A{r},'
                     'Factures!$J$6:$J$205,"<>"))').format(r=r)

# ===========================================================================
# 6. Paramètres : seuils ASER, listes fusionnées, annuaire e-mail
# ===========================================================================
ws = wb['Paramètres']
ws['B9'] = 'Délai contractuel de paiement par défaut (si non saisi dans Factures)'
for r, (lib, val, unite, src) in zip((14, 15), [
        ('Fréquence de relance des détenteurs de pièces', 7, 'jours', 'Hypothèse à valider'),
        ("Nombre de relances déclenchant l'escalade DG", 2, 'relances', 'Hypothèse à valider')]):
    for col, v_ in zip('BCDE', (lib, val, unite, src)):
        c = ws['%s%d' % (col, r)]
        c.value = v_
        c._style = copy(ws['%s13' % col]._style)
ws['C14'].protection = ws['C15'].protection = DEVERROUILLE
STAT_DLM = ['À préparer', 'En préparation', 'À déposer', 'Déposée – en instruction', 'Accordée', 'Refusée',
            'Non requise', 'Clôturé']
STAT_GAR = ['Valide', 'Prorogée', 'Prorogation demandée', 'Mainlevée demandée', 'Mainlevée accordée', 'Appelée']
for col, liste in (('L', STAT_DLM), ('P', STAT_GAR)):
    for i in range(20):
        c = ws['%s%d' % (col, 24 + i)]
        c.value = liste[i] if i < len(liste) else None
        c._style = copy(ws['K24']._style)
entete(ws['AC23'], 'Détenteur (rappel col. M)', False)
entete(ws['AD23'], 'E-mail du détenteur (relances Outlook)', True)
ws.column_dimensions['AC'].width = 20
ws.column_dimensions['AD'].width = 32
for i in range(20):
    r = 24 + i
    ws['AC%d' % r] = '=IF(M%d="","",M%d)' % (r, r)
    style_cellule(ws['AC%d' % r], 'calc')
    style_cellule(ws['AD%d' % r], 'in')
note(ws['AD23'], "Adresse utilisée par « Préparer les relances » pour le détenteur de la même ligne (colonne M).")

# ===========================================================================
# 7. Nouvelles feuilles issues du classeur ASER
# ===========================================================================
# --- Tableau de bord --------------------------------------------------------
ws = wb.create_sheet('Tableau_de_bord', 1)
ws['B1'] = 'Tableau de bord — Pilotage consolidé DPER / ASER'
ws['B1'].font = F_TITRE
ws['B2'], ws['C2'] = 'Date de référence :', '=D_JOUR'
ws['E2'] = 'Dernière actualisation :'
for a in ('B2', 'E2'):
    ws[a].font = Font(name='Arial', size=9, bold=True, color=BLEU_NUIT)
ws['C2'].number_format = NF_DATE
ws['G2'].number_format = 'dd/mm/yyyy hh:mm'
ws['C2'].font = ws['G2'].font = F_CALC
ws['G2'] = 'à lancer : « Actualiser tout »'
COL_W = dict(A=2, B=16, C=12, D=16, E=12, F=16, G=12, H=16, I=12)
for k, w in COL_W.items():
    ws.column_dimensions[k].width = w


def tuile(ws, ligne, col, libelle, formule, nf='#,##0'):
    c2 = get_column_letter(column_index_from_string(col) + 1)
    ws.merge_cells('%s%d:%s%d' % (col, ligne, c2, ligne))
    ws.merge_cells('%s%d:%s%d' % (col, ligne + 1, c2, ligne + 1))
    a = ws['%s%d' % (col, ligne)]
    a.value = libelle
    a.font = Font(name='Arial', size=8, bold=True, color='FFFFFFFF')
    a.fill = FILL_ENT_SAISIE
    a.alignment = AL_ENT
    b = ws['%s%d' % (col, ligne + 1)]
    b.value = formule
    b.font = Font(name='Arial', size=16, bold=True, color=BLEU_NUIT)
    b.alignment = Alignment(horizontal='center', vertical='center')
    b.number_format = nf
    for cc in (col, c2):
        ws['%s%d' % (cc, ligne)].border = BORD
        ws['%s%d' % (cc, ligne + 1)].border = BORD
        ws['%s%d' % (cc, ligne + 1)].fill = FILL_CALC
    ws.row_dimensions[ligne].height = 30
    ws.row_dimensions[ligne + 1].height = 30


def section(ws, ref, texte):
    ws[ref] = texte
    ws[ref].font = F_SECTION


section(ws, 'B5', 'DLM et prorogations')
tuile(ws, 6, 'B', 'DLM échues', '=COUNTIF(DLM!$AB$6:$AB$205,"ÉCHUE")')
tuile(ws, 6, 'D', 'Escalades DG requises', '=COUNTIF(DLM!$AC$6:$AC$205,"OUI")')
tuile(ws, 6, 'F', 'Relances dues', '=COUNTIF(DLM!$AA$6:$AA$205,"OUI")')
tuile(ws, 6, 'H', 'Solde non décaissé — DLM en escalade (M FCFA)',
      '=SUMIFS(DLM!$P$6:$P$205,DLM!$AC$6:$AC$205,"OUI")/1000000', '#,##0.0')
section(ws, 'B9', 'Garanties et factures')
tuile(ws, 10, 'B', 'Garanties expirées (statut actif)', '=COUNTIF(Garanties!$R$6:$R$205,"EXPIRÉE")')
tuile(ws, 10, 'D', '="Garanties à échéance ≤ "&P_GAR&" j"', '=COUNTIF(Garanties!$R$6:$R$205,"ALERTE")')
tuile(ws, 10, 'F', 'Factures hors délai — reste à payer (M FCFA)',
      '=SUMIFS(Factures!$P$6:$P$205,Factures!$W$6:$W$205,"HORS DÉLAI")/1000000', '#,##0.0')
tuile(ws, 10, 'H', 'Nb factures hors délai', '=COUNTIF(Factures!$W$6:$W$205,"HORS DÉLAI")')
section(ws, 'B13', 'Délais moyens de traitement des factures (jours) et encours')
tuile(ws, 14, 'B', 'Réception → visa', '=IFERROR(AVERAGE(Factures!$R$6:$R$205),"—")', '0.0')
tuile(ws, 14, 'D', 'Visa → transmission AC', '=IFERROR(AVERAGE(Factures!$S$6:$S$205),"—")', '0.0')
tuile(ws, 14, 'F', 'Transmission AC → paiement', '=IFERROR(AVERAGE(Factures!$T$6:$T$205),"—")', '0.0')
tuile(ws, 14, 'H', 'Reste à payer total (M FCFA)', '=SUM(Factures!$P$6:$P$205)/1000000', '#,##0.0')
section(ws, 'B17', "Avancement du plan d'action par horizon")
for l, h in zip('BCDEF', ['Horizon', 'Nb mesures', 'Achevées', 'En retard', 'Avancement moyen']):
    entete(ws['%s18' % l], h, False)
for i in range(3):
    r = 19 + i
    ws['B%d' % r] = '=IFERROR(INDEX(L_HORIZON,%d),"")' % (i + 1)
    ws['C%d' % r] = '=IF(B{r}="","",COUNTIFS(Plan_Action!$B$7:$B$30,B{r},Plan_Action!$C$7:$C$30,"<>"))'.format(r=r)
    ws['D%d' % r] = '=IF(B{r}="","",COUNTIFS(Plan_Action!$B$7:$B$30,B{r},Plan_Action!$H$7:$H$30,"Achevé"))'.format(r=r)
    ws['E%d' % r] = '=IF(B{r}="","",COUNTIFS(Plan_Action!$B$7:$B$30,B{r},Plan_Action!$K$7:$K$30,"EN RETARD"))'.format(r=r)
    ws['F%d' % r] = '=IF(N(C{r})=0,0,SUMIFS(Plan_Action!$I$7:$I$30,Plan_Action!$B$7:$B$30,B{r})/C{r})'.format(r=r)
    for l, nf in zip('BCDEF', (None, '0', '0', '0', NF_PCT)):
        style_cellule(ws['%s%d' % (l, r)], 'calc', nf)
section(ws, 'B23', 'Données des graphiques (calculées)')
for l, h in zip('BC', ['Bailleur', 'Solde non décaissé (M FCFA)']):
    entete(ws['%s24' % l], h, False)
for l, h in zip('EFG', ['Étape', 'Nb factures', 'Reste à payer (M FCFA)']):
    entete(ws['%s24' % l], h, False)
for i in range(10):
    r = 25 + i
    ws['B%d' % r] = '=IFERROR(INDEX(L_BAILLEUR,%d)&"","")' % (i + 1)
    ws['C%d' % r] = '=IF(B{r}="","",SUMIFS(DLM!$P$6:$P$205,DLM!$B$6:$B$205,B{r})/1000000)'.format(r=r)
    style_cellule(ws['B%d' % r], 'calc')
    style_cellule(ws['C%d' % r], 'calc', '#,##0.0')
for i, etape in enumerate(['1-Reçu', '2-Visé', '3-Transmis AC', '4-Payé partiellement']):
    r = 25 + i
    ws['E%d' % r] = etape
    ws['F%d' % r] = '=COUNTIF(Factures!$Q$6:$Q$205,E%d)' % r
    ws['G%d' % r] = '=SUMIFS(Factures!$P$6:$P$205,Factures!$Q$6:$Q$205,E%d)/1000000' % r
    style_cellule(ws['E%d' % r], 'calc')
    style_cellule(ws['F%d' % r], 'calc', '0')
    style_cellule(ws['G%d' % r], 'calc', '#,##0.0')
ws['E30'] = 'Étapes calculées par le registre des factures (ne pas modifier).'
ws['E30'].font = F_SOUS


def graphique(ws, titre_g, cat, val, ancre, nf='#,##0.0'):
    ch = BarChart()
    ch.type = 'col'
    ch.title = titre_g
    ch.style = 10
    ch.legend = None
    ch.height, ch.width = 7.2, 15.5
    ch.add_data(val, titles_from_data=False)
    ch.set_categories(cat)
    ch.y_axis.numFmt = nf
    ch.y_axis.majorGridlines = None
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ws.add_chart(ch, ancre)


graphique(ws, 'Solde non décaissé par bailleur (millions FCFA)',
          Reference(ws, min_col=2, min_row=25, max_row=34), Reference(ws, min_col=3, min_row=25, max_row=34), 'K4')
graphique(ws, 'Factures en souffrance par étape (millions FCFA)',
          Reference(ws, min_col=5, min_row=25, max_row=28), Reference(ws, min_col=7, min_row=25, max_row=28), 'K20')
graphique(ws, "Avancement du plan d'action par horizon",
          Reference(ws, min_col=2, min_row=19, max_row=21), Reference(ws, min_col=6, min_row=19, max_row=21), 'K36', '0%')
ws.sheet_view.showGridLines = False
ws.print_area = 'A1:T52'
ws.page_setup.orientation = 'landscape'
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
ws.sheet_properties.tabColor = BLEU_NUIT

# --- Alertes et Contrôle qualité (générées par les macros) ------------------
for nom, couleur, ents in (('Alertes', 'FFC00000', ['Gravité', 'Feuille', 'Motif', 'ID', 'Libellé', 'Statut / valeur',
                                                    'Jours', 'Responsable / détenteur', 'Lien']),
                           ('Contrôle_Qualité', 'FFED7D31', ['Gravité', 'Feuille', 'Cellule', 'ID', 'Contrôle',
                                                             'Détail', 'Lien'])):
    ws = wb.create_sheet(nom, 2 if nom == 'Alertes' else 3)
    titre(ws, ('Synthèse des alertes' if nom == 'Alertes' else 'Contrôle qualité des saisies') +
          ' — Pilotage consolidé DPER / ASER',
          'Feuille générée par les macros : lancer « Actualiser tout » depuis le Sommaire.')
    for i, h in enumerate(ents):
        entete(ws.cell(row=4, column=i + 1), h, False)
    ws['A5'] = 'Aucune synthèse générée.'
    ws.sheet_properties.tabColor = couleur
    ws.sheet_view.showGridLines = False

# --- Note d'escalade DG ------------------------------------------------------
ws = wb.create_sheet('Note_Escalade_DG')
titre(ws, "Note d'escalade à l'attention du Directeur Général",
      'Feuille générée par la macro « Note d’escalade DG » (ne pas saisir). Export PDF dans le dossier du classeur.')
ws['A4'], ws['A5'], ws['A6'] = 'Date :', 'Objet :', 'Synthèse :'
ws['B5'] = 'Accords dont la DLM nécessite une intervention de la Direction générale'
for a in ('A4', 'A5', 'A6'):
    ws[a].font = Font(name='Arial', size=10, bold=True, color=BLEU_NUIT)
for i, (h, w) in enumerate([('Bailleur', 10), ('ID / N° accord', 26), ('Projet', 22), ('DLM actuelle', 11),
                            ('Jours avant DLM', 9), ('Solde non décaissé (FCFA)', 16), ('Pièce manquante', 26),
                            ('Détenteur', 14), ('Nb relances', 8), ('Dernière relance', 11), ('Motif / historique', 40)]):
    entete(ws.cell(row=8, column=i + 1), h, False)
    ws.column_dimensions[get_column_letter(i + 1)].width = w
ws.row_dimensions[8].height = 36
ws.sheet_properties.tabColor = GRIS_TXT
ws.sheet_view.showGridLines = False

# --- Journal ------------------------------------------------------------------
ws = wb.create_sheet('Journal')
titre(ws, 'Journal des macros', 'Une ligne par exécution de macro (alimenté automatiquement — ne pas modifier).')
for i, (h, w) in enumerate([('Horodatage', 17), ('Utilisateur', 18), ('Macro', 22), ('Résultat', 110)]):
    entete(ws.cell(row=4, column=i + 1), h, False)
    ws.column_dimensions[get_column_letter(i + 1)].width = w
ws['A5'] = dt.datetime(2026, 11, 1, 0, 0)
ws['B5'], ws['C5'] = '—', 'Consolidation'
ws['D5'] = 'Classeur consolidé : fusion « Pilotage DPER — Plan d’amélioration » et « Pilotage des financements ASER ».'
ws['A5'].number_format = 'dd/mm/yyyy hh:mm'
for l in 'ABCD':
    ws['%s5' % l].font = F_NORMAL
    ws['%s5' % l].border = BORD
ws.freeze_panes = 'A5'
ws.auto_filter.ref = 'A4:D5'
ws.sheet_properties.tabColor = GRIS_TXT

# --- Archive ------------------------------------------------------------------
ws = wb.create_sheet('Archive')
titre(ws, 'Archive — lignes clôturées',
      'Garanties en mainlevée (gauche) et factures soldées (droite), déplacées par la macro « Archiver les clôtures ». Valeurs figées.')
ws['A3'] = 'Garanties en mainlevée'
ws['A3'].font = F_SECTION
for i, col in enumerate(GAR_COLS + [dict(h='Date archivage')]):
    entete(ws.cell(row=4, column=1 + i), col['h'], False)
DEB_FAC_ARCH = len(GAR_COLS) + 3          # une colonne vide de séparation
ws.cell(row=3, column=DEB_FAC_ARCH).value = 'Factures soldées'
ws.cell(row=3, column=DEB_FAC_ARCH).font = F_SECTION
for i, col in enumerate(FAC_COLS + [dict(h='Date archivage')]):
    entete(ws.cell(row=4, column=DEB_FAC_ARCH + i), col['h'], False)
ws.row_dimensions[4].height = 48
ws.freeze_panes = 'A5'
ws.sheet_properties.tabColor = GRIS_TXT

# ===========================================================================
# 8. Sommaire (reconstruit)
# ===========================================================================
ws = remplacer_feuille('Sommaire', 0)
titre(ws, "Outil de pilotage consolidé — Plan d'amélioration DPER et financements ASER",
      'Référence : note au COMDIR (horizons 0-30 jours · 1-3 mois · 3-12 mois) — fusion des classeurs « Pilotage DPER » '
      'et « Pilotage des financements ASER ».')
ws['B3'], ws['C3'] = 'Situation au :', '=D_JOUR'
ws['B3'].font = ws['C3'].font = Font(name='Arial', size=10, bold=True, color=BLEU_NUIT)
ws['C3'].number_format = NF_DATE
for l, h in zip('BCDE', ['Horizon', 'Feuille', 'Objet', "Points d'attention"]):
    entete(ws['%s5' % l], h, True)
LIGNES = [
    ('—', 'Tableau_de_bord', 'Indicateurs clés et graphiques : DLM, garanties, factures, avancement du plan.', '="—"'),
    ('—', 'Alertes', 'Synthèse de toutes les alertes (générée par « Actualiser tout »).',
     '=COUNTIF(Alertes!$A$5:$A$5000,"CRITIQUE")'),
    ('—', 'Contrôle_Qualité', 'Anomalies de saisie : doublons, formules écrasées, dates, champs obligatoires.',
     "=COUNTIF('Contrôle_Qualité'!$A$5:$A$5000,\"BLOQUANT\")"),
    ('—', 'Plan_Action', 'Suivi des 15 mesures de la note au COMDIR (responsable, échéance, statut).',
     '=COUNTIF(Plan_Action!$K$7:$K$30,"EN RETARD")'),
    ('0-30 jours', 'DLM', 'DLM et prorogations BID, BOAD, FVC, ADFD : échéance, pièce manquante, détenteur, relances, escalade DG.',
     '=COUNTIF(DLM!$AC$6:$AC$205,"OUI")'),
    ('0-30 jours', 'Garanties', "Garanties de bonne exécution et cautions d'avance ; alerte avant échéance.",
     '=COUNTIF(Garanties!$R$6:$R$205,"ALERTE")+COUNTIF(Garanties!$R$6:$R$205,"EXPIRÉE")'),
    ('0-30 jours', 'Factures', "Factures et décomptes : réception, visa, transmission à l'AC, paiements (partiels), retard.",
     '=COUNTIF(Factures!$W$6:$W$205,"HORS DÉLAI")'),
    ('0-30 jours', 'Plan_Paiement', "Plan de paiement priorisé par échéance de DLM (automatique), à arrêter avec l'AC et la DFP.",
     '=COUNTIF(Factures!$Z$6:$Z$205,"P0")+COUNTIF(Factures!$Z$6:$Z$205,"P1")'),
    ('0-30 jours', 'Scénarios_BID', 'Décaissable avant DLM et montant à risque selon 3 durées de prorogation.', '="—"'),
    ('1-3 mois', 'Marchés', 'Référentiel unique des marchés : alimente tous les registres (projet, bailleur, entreprise, accord).',
     '=COUNTIF(Marchés!$U$6:$U$205,"RETARD")+COUNTIF(Marchés!$U$6:$U$205,"DÉLAI DÉPASSÉ")'),
    ('1-3 mois', 'RACI', 'Matrice des responsabilités des processus du plan.',
     '=COUNTIF(RACI!$P$5:$P$34,"1 SEUL A REQUIS")+COUNTIF(RACI!$P$5:$P$34,"R MANQUANT")'),
    ('1-3 mois', 'Éval_Entreprises', 'Évaluation pondérée de la performance des entreprises (classes A à D).',
     '=COUNTIF(Éval_Entreprises!$N$6:$N$205,"C")+COUNTIF(Éval_Entreprises!$N$6:$N$205,"D")'),
    ('1-3 mois', 'Fiscal', 'Démarche fiscale groupée : dossiers par marché et synthèse par lot.',
     '=COUNTIF(Fiscal!$I$6:$I$205,"Non")'),
    ('3-12 mois', 'DAO_Standards', 'Bibliothèque des études et DAO standardisés et suivi de leur élaboration.',
     '=COUNT(DAO_Standards!$L$6:$L$105)'),
    ('3-12 mois', 'Protocole_MES', 'Protocole de mise en service et de transfert aux exploitants, suivi par site.',
     '=COUNTIF(Protocole_MES!$V$6:$V$205,"RÉSERVES NON LEVÉES")+COUNTIF(Protocole_MES!$V$6:$V$205,"TRANSFERT EN RETARD")'),
    ('—', 'Note_Escalade_DG', 'Note au DG des accords en escalade, exportée en PDF (macro).', '="—"'),
    ('—', 'Journal', 'Traçabilité : une ligne par exécution de macro.', '="—"'),
    ('—', 'Archive', 'Garanties en mainlevée et factures soldées archivées (macro « Archiver les clôtures »).', '="—"'),
    ('—', 'Paramètres', 'Seuils, hypothèses, taux de change, listes déroulantes, e-mails des détenteurs.', '="—"'),
]
for i, (hz, f, objet, pa) in enumerate(LIGNES):
    r = 6 + i
    ws['B%d' % r] = hz
    ws['C%d' % r] = '=HYPERLINK("#\'%s\'!A1","%s")' % (f, f)
    ws['D%d' % r] = objet
    ws['E%d' % r] = pa
    for l in 'BCDE':
        ws['%s%d' % (l, r)].border = BORD
        ws['%s%d' % (l, r)].font = F_NORMAL
    ws['C%d' % r].font = Font(name='Arial', size=9, color='FF0563C1', underline='single')
    ws['D%d' % r].alignment = AL_WRAP
    style_cellule(ws['E%d' % r], 'calc')
    ws['E%d' % r].alignment = Alignment(horizontal='center')
    mfc_expr(ws, 'E%d' % r, 'AND(ISNUMBER(E%d),E%d>0)' % (r, r), 'FF9C0006', 'FFFFC7CE')
fin = 6 + len(LIGNES)
r = fin + 1
ws['B%d' % r] = "Mode d'emploi"
ws['B%d' % r].font = F_SECTION
EMPLOI = [
    "1. Saisir uniquement dans les cellules à police bleue ; les cellules à fond gris sont calculées et protégées "
    "(protection sans mot de passe : Révision > Ôter la protection).",
    "2. La ligne grisée en italique sous chaque en-tête est un exemple de format : elle est exclue des calculs "
    "(bouton « Effacer les exemples »).",
    "3. Ordre de saisie : Paramètres (taux, seuils, e-mails) → DLM (accords) → Marchés (référentiel unique) → registres.",
    "4. Saisir sans ligne vide intermédiaire. « Nouvelle ligne » place le curseur sur la première ligne libre du "
    "registre actif et attribue un identifiant.",
    "5. Prorogation accordée : mettre à jour la « DLM actuelle » et le nombre de prorogations. Garantie prorogée : "
    "nouvelle échéance et statut « Prorogée ».",
    "6. Relances : « Préparer les relances » (brouillons Outlook par détenteur, jamais envoyés automatiquement) ou "
    "« Enregistrer une relance » ; la date et le nombre de relances sont mis à jour après confirmation.",
    "7. Le plan de paiement se recalcule seul ; dates cibles et validations AC / DFP se saisissent dans « Factures ».",
    "8. « Actualiser tout » : recalcul, alertes, contrôle qualité, tableau de bord. Chaque macro est tracée dans « Journal ».",
    "9. Seuils et pondérations en jaune = hypothèses à valider. Capacité : 200 lignes par registre "
    "(100 pour DAO_Standards, 30 pour RACI, 24 mesures pour Plan_Action).",
]
for i, t in enumerate(EMPLOI):
    rr = r + 1 + i
    ws.merge_cells('B%d:E%d' % (rr, rr))
    ws['B%d' % rr] = t
    ws['B%d' % rr].font = F_NORMAL
    ws['B%d' % rr].alignment = AL_WRAP
    ws.row_dimensions[rr].height = 24
r = r + len(EMPLOI) + 2
ws['B%d' % r] = 'Légende'
ws['B%d' % r].font = F_SECTION
for i, (lib, kind, desc) in enumerate([('Saisie', 'in', 'Critique : échue, expirée, en retard, escalade requise'),
                                        ('Calcul automatique', 'calc', 'Attention : alerte, vigilance, relance due'),
                                        ('Hypothèse / paramètre à valider', 'param', 'Conforme'),
                                        ('Exemple de format', 'ex', '')]):
    rr = r + 1 + i
    c = ws['B%d' % rr]
    c.value = 'Abc'
    c.border = BORD
    if kind == 'in':
        c.font = F_SAISIE
    elif kind == 'calc':
        c.font, c.fill = F_CALC, FILL_CALC
    elif kind == 'param':
        c.font, c.fill = F_SAISIE, FILL_PARAM
    else:
        c.font, c.fill = F_EX, FILL_EX
    ws['C%d' % rr] = lib
    ws['D%d' % rr] = desc
    ws['C%d' % rr].font = ws['D%d' % rr].font = F_NORMAL
    if desc:
        d = ws['D%d' % rr]
        d.font = Font(name='Arial', size=9, color=['FF9C0006', 'FF9C5700', 'FF006100'][i])
        d.fill = PatternFill('solid', fgColor=['FFFFC7CE', 'FFFFEB9C', 'FFC6EFCE'][i])
for k, w in dict(A=3, B=12, C=22, D=70, E=14, F=3, G=3).items():
    ws.column_dimensions[k].width = w
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = GRIS_TXT

# ===========================================================================
# 9. Noms définis complémentaires
# ===========================================================================
for nom, ref in (('P_FREQ', 'Paramètres!$C$14'), ('P_NBREL', 'Paramètres!$C$15'),
                 ('DerniereActualisation', 'Tableau_de_bord!$G$2')):
    wb.defined_names[nom] = DefinedName(nom, attr_text=ref)

# ===========================================================================
# 10. Ordre des feuilles, noms de code, protection
# ===========================================================================
ORDRE = [('Sommaire', 'wsSommaire'), ('Tableau_de_bord', 'wsTableauBord'), ('Alertes', 'wsAlertes'),
         ('Contrôle_Qualité', 'wsControle'), ('Plan_Action', 'wsPlanAction'), ('DLM', 'wsDLM'),
         ('Garanties', 'wsGaranties'), ('Factures', 'wsFactures'), ('Plan_Paiement', 'wsPlanPaiement'),
         ('Scénarios_BID', 'wsScenariosBID'), ('Marchés', 'wsMarches'), ('RACI', 'wsRACI'),
         ('Éval_Entreprises', 'wsEvalEntreprises'), ('Fiscal', 'wsFiscal'), ('DAO_Standards', 'wsDAO'),
         ('Protocole_MES', 'wsProtocoleMES'), ('Note_Escalade_DG', 'wsNoteDG'), ('Journal', 'wsJournal'),
         ('Archive', 'wsArchive'), ('Paramètres', 'wsParametres')]
assert sorted(n for n, _ in ORDRE) == sorted(wb.sheetnames), wb.sheetnames
wb._sheets = [wb[n] for n, _ in ORDRE]
for n, code in ORDRE:
    wb[n].sheet_properties.codeName = code
wb.active = 0
for w in wb.worksheets:
    w.sheet_view.tabSelected = (w.title == 'Sommaire')
wb.code_name = 'ThisWorkbook'

# Cellules de saisie des feuilles DPER conservées : police bleue => déverrouillée
for n in ('Plan_Action', 'Scénarios_BID', 'Marchés', 'RACI', 'Éval_Entreprises', 'Fiscal', 'DAO_Standards',
          'Protocole_MES', 'Paramètres'):
    w = wb[n]
    for row in w.iter_rows():
        for c in row:
            col = c.font.color.rgb if c.font is not None and c.font.color is not None else None
            if col == 'FF0000FF' and not (isinstance(c.value, str) and c.value.startswith('=')):
                c.protection = DEVERROUILLE
# Lignes d'exemple (ligne 5) des feuilles conservées : saisies déverrouillées
for n in ('Marchés', 'Éval_Entreprises', 'Fiscal', 'DAO_Standards', 'Protocole_MES'):
    w = wb[n]
    for c in w[5]:
        if c.value is not None and not (isinstance(c.value, str) and c.value.startswith('=')):
            c.protection = DEVERROUILLE

PROTEGEES = [n for n, _ in ORDRE if n not in ('Alertes', 'Contrôle_Qualité')]
for n in PROTEGEES:
    p = wb[n].protection
    p.sheet = True
    p.formatColumns = False      # autorisé
    p.formatRows = False
    p.autoFilter = False
    p.sort = True
    p.insertRows = True
    p.deleteRows = True
    p.objects = True             # boutons et graphiques non déplaçables (les macros restent cliquables)

wb.save(OUT)
print('OK', OUT)
