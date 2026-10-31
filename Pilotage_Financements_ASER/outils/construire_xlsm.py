"""Construit Pilotage_Financements_ASER.xlsm à partir du classeur source (.xlsx)
et des modules VBA du dossier vba/ (encodés en Windows-1252, fins de ligne CRLF).

Usage : python3 outils/construire_xlsm.py
"""
import html
import os
import re
import sys
import zipfile

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(ICI)
sys.path.insert(0, ICI)
from vba_project import build_vba_project  # noqa: E402

SOURCE = os.path.join(RACINE, 'Pilotage_Financements_ASER_source.xlsx')
CIBLE = os.path.join(RACINE, 'Pilotage_Financements_ASER.xlsm')
VBA = os.path.join(RACINE, 'vba')

MODULES_STANDARD = ['Mod_Utilitaires', 'Mod_Pilotage', 'Mod_Communication_Reporting']

# Listes déroulantes : plages figées remplacées par des références de tableau
# (Liste_Accords pointait sur DLM_Prorogations!$C$6:$C$8 et passait en #REF!
# après SupprimerExemples ; les autres ne suivaient pas l'ajout de lignes).
NOMS_DYNAMIQUES = {
    'Liste_Accords': 'T_DLM[N° accord]',
    'Liste_Bailleurs': 'T_Bailleurs[Bailleur]',
    'Liste_Detenteurs': 'T_Detenteurs[Identifiant_Detenteur]',
    'Liste_Devises': 'T_Devises[Devise]',
}

# Boutons du Tableau_de_bord : (macro, libellé, col début, col fin, ligne) (index base 0)
BOUTONS = [
    ('Actualiser', '1. Actualiser Tout', 0, 2, 4),
    ('GenererPlanPaiement', '2. Générer Plan Paiement', 3, 5, 4),
    ('PreparerRelances', '3. Préparer Relances Outlook', 6, 8, 4),
    ('NoteEscaladeDG', '4. Note Escalade DG (PDF)', 9, 11, 4),
    ('RapportMensuel', '5. Rapport Mensuel (PDF)', 0, 2, 5),
    ('ArchiverClotures', '6. Archiver Clôturés', 3, 5, 5),
    ('NouvelleLigne', '7. Nouvelle Ligne', 6, 8, 5),
    ('SupprimerExemples', '8. Supprimer Exemples', 9, 11, 5),
]
EMU_PX = 9525


def lire_module(nom, ext='.bas'):
    with open(os.path.join(VBA, nom + ext), 'rb') as f:
        brut = f.read()
    code = brut.decode('cp1252')
    assert '�' not in code
    return code


def module_document(nom, base_clsid, code_utilisateur=''):
    entete = (
        f'Attribute VB_Name = "{nom}"\r\n'
        f'Attribute VB_Base = "0{base_clsid}"\r\n'
        'Attribute VB_GlobalNameSpace = False\r\n'
        'Attribute VB_Creatable = False\r\n'
        'Attribute VB_PredeclaredId = True\r\n'
        'Attribute VB_Exposed = True\r\n'
        'Attribute VB_TemplateDerived = False\r\n'
        'Attribute VB_Customizable = True\r\n'
    )
    return entete + code_utilisateur


def col_px(largeur_car):
    return int(largeur_car * 7 + 5)


def construire():
    zin = zipfile.ZipFile(SOURCE)
    fichiers = {n: zin.read(n) for n in zin.namelist()}

    # --- Correspondance feuilles -> fichiers
    wb = fichiers['xl/workbook.xml'].decode('utf-8')
    rels = fichiers['xl/_rels/workbook.xml.rels'].decode('utf-8')
    cibles = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    feuilles = re.findall(r'<sheet name="([^"]+)" sheetId="(\d+)"[^>]*r:id="(rId\d+)"', wb)
    feuille_fichier = [(html.unescape(n), 'xl/' + cibles[r]) for n, _, r in feuilles]

    # --- workbook.xml : codeName + noms dynamiques
    wb = wb.replace('<workbookPr ', '<workbookPr codeName="ThisWorkbook" ', 1)
    for nom, ref in NOMS_DYNAMIQUES.items():
        motif = r'(<definedName [^>]*name="' + nom + r'"[^>]*>)[^<]*(</definedName>)'
        wb, n = re.subn(motif, lambda m: m.group(1) + ref + m.group(2), wb)
        assert n == 1, nom
    fichiers['xl/workbook.xml'] = wb.encode('utf-8')

    # --- Feuilles : codeName VBA
    noms_doc = []
    for i, (nom, chemin) in enumerate(feuille_fichier, start=1):
        code_name = f'Feuil{i}'
        noms_doc.append(code_name)
        x = fichiers[chemin].decode('utf-8')
        x, n = re.subn(r'<sheetPr\b', f'<sheetPr codeName="{code_name}"', x, count=1)
        if n == 0:
            x = re.sub(r'(<worksheet\b[^>]*>)', r'\1' + f'<sheetPr codeName="{code_name}"/>', x, count=1)
        fichiers[chemin] = x.encode('utf-8')

    # --- Tableaux : colonnes calculées (calculatedColumnFormula) pour que
    #     Excel propage les formules aux nouvelles lignes
    for nom_t in [n for n in fichiers if n.startswith('xl/tables/')]:
        t = fichiers[nom_t].decode('utf-8')
        nom_table = re.search(r'<table [^>]*name="([^"]+)"', t).group(1)
        if nom_table not in ('T_DLM', 'T_Garanties', 'T_Factures', 'T_Plan'):
            continue
        ref = re.search(r'ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', t)
        premiere = int(ref.group(2)) + 1
        feuille = next(ch for _, ch in feuille_fichier
                       if f'tables/{os.path.basename(nom_t)}' in fichiers.get(
                           ch.replace('worksheets/', 'worksheets/_rels/') + '.rels', b'').decode())
        sx = fichiers[feuille].decode('utf-8')
        colonnes = re.findall(r'<tableColumn [^>]*?/>', t)
        c0 = col_index(ref.group(1))
        for k, tc in enumerate(colonnes):
            lettre = col_lettre(c0 + k)
            nom_col = html.unescape(re.search(r'name="([^"]*)"', tc).group(1))
            if nom_col == 'Délai contractuel de paiement':
                continue  # colonne de saisie (valeur par défaut)
            formules = re.findall(r'<c r="' + lettre + r'(\d+)"[^>]*><f[^>]*>(.*?)</f>', sx)
            formules = [f for r, f in formules if int(r) >= premiere]
            if formules and all(f == formules[0] for f in formules):
                nouveau = tc[:-2] + '><calculatedColumnFormula>' + formules[0] + \
                    '</calculatedColumnFormula></tableColumn>'
                t = t.replace(tc, nouveau, 1)
        fichiers[nom_t] = t.encode('utf-8')

    # --- Plan_Paiement : cumul indépendant de la position (identique à la macro)
    for nom, chemin in feuille_fichier:
        if nom == 'Plan_Paiement':
            x = fichiers[chemin].decode('utf-8')
            x = re.sub(r'(<c r="H\d+"[^>]*><f[^>]*>)SUM\(\$G\$6:G\d+\)(</f>)',
                       r'\1SUM(INDEX(T_PlanPaiement[Reste à payer],1):T_PlanPaiement[[#This Row],[Reste à payer]])\2',
                       x)
            fichiers[chemin] = x.encode('utf-8')

    # --- Tableau_de_bord : « Dernière actualisation » (date + heure) affichée
    #     « ### » faute de largeur -> style dédié avec réduction automatique
    tdb = dict(feuille_fichier)['Tableau_de_bord']
    sx = fichiers[tdb].decode('utf-8')
    styles = fichiers['xl/styles.xml'].decode('utf-8')
    bloc = re.search(r'(<cellXfs count=")(\d+)(">)(.*?)(</cellXfs>)', styles, re.S)
    xfs = re.findall(r'<xf [^>]*?(?:/>|>.*?</xf>)', bloc.group(4), re.S)
    s_d4 = int(re.search(r'<c r="D4" s="(\d+)"', sx).group(1))
    nouveau_xf = xfs[s_d4].replace('wrapText="true"', 'wrapText="false"').replace(
        'shrinkToFit="false"', 'shrinkToFit="true"')
    styles = styles.replace(bloc.group(0), bloc.group(1) + str(len(xfs) + 1) + bloc.group(3) +
                            bloc.group(4) + nouveau_xf + bloc.group(5))
    fichiers['xl/styles.xml'] = styles.encode('utf-8')
    sx = sx.replace(f'<c r="D4" s="{s_d4}"', f'<c r="D4" s="{len(xfs)}"', 1)
    fichiers[tdb] = sx.encode('utf-8')

    # --- Boutons de macros sur le Tableau_de_bord
    largeurs = {int(m.group(1)) - 1: float(m.group(2))
                for m in re.finditer(r'<col [^>]*min="(\d+)"[^>]*width="([\d.]+)"', sx)}
    hauteurs = {int(m.group(1)) - 1: float(m.group(2))
                for m in re.finditer(r'<row r="(\d+)"[^>]*ht="([\d.]+)"', sx)}
    dessin = fichiers['xl/drawings/drawing1.xml'].decode('utf-8')
    formes = []
    marge = 2 * EMU_PX
    for k, (macro, libelle, c1, c2, ligne) in enumerate(BOUTONS):
        larg_fin = col_px(largeurs.get(c2, 8.43)) * EMU_PX
        haut = int(hauteurs.get(ligne, 15) * 12700)
        formes.append(
            '<xdr:twoCellAnchor editAs="absolute">'
            f'<xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>{marge}</xdr:colOff><xdr:row>{ligne}</xdr:row><xdr:rowOff>{marge}</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>{larg_fin - marge}</xdr:colOff><xdr:row>{ligne}</xdr:row><xdr:rowOff>{haut - marge}</xdr:rowOff></xdr:to>'
            f'<xdr:sp macro="[0]!{macro}" textlink="">'
            f'<xdr:nvSpPr><xdr:cNvPr id="{100 + k}" name="btn_{macro}" descr="Lance la macro {macro}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
            '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
            '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="1B365D"/></a:solidFill>'
            '<a:ln w="9525"><a:solidFill><a:srgbClr val="0F2340"/></a:solidFill></a:ln></xdr:spPr>'
            '<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" lIns="36000" rIns="36000" tIns="0" bIns="0" rtlCol="0" anchor="ctr"/>'
            '<a:lstStyle/><a:p><a:pPr algn="ctr"/>'
            '<a:r><a:rPr lang="fr-FR" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            '<a:latin typeface="Calibri"/></a:rPr>'
            f'<a:t>{html.escape(libelle, quote=False)}</a:t></a:r></a:p></xdr:txBody></xdr:sp>'
            '<xdr:clientData fPrintsWithSheet="0"/></xdr:twoCellAnchor>')
    dessin = dessin.replace('</xdr:wsDr>', ''.join(formes) + '</xdr:wsDr>')
    fichiers['xl/drawings/drawing1.xml'] = dessin.encode('utf-8')

    # --- Projet VBA
    modules = [{'name': 'ThisWorkbook', 'document': True,
                'code': module_document('ThisWorkbook', '{00020819-0000-0000-C000-000000000046}',
                                        lire_module('ThisWorkbook', '.cls'))}]
    for nom in noms_doc:
        modules.append({'name': nom, 'document': True,
                        'code': module_document(nom, '{00020820-0000-0000-C000-000000000046}')})
    for nom in MODULES_STANDARD:
        modules.append({'name': nom, 'document': False, 'code': lire_module(nom)})
    fichiers['xl/vbaProject.bin'] = build_vba_project(modules)

    # --- Types de contenu et relations
    ct = fichiers['[Content_Types].xml'].decode('utf-8')
    ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    ct = ct.replace('<Default Extension="xml"',
                    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"', 1)
    fichiers['[Content_Types].xml'] = ct.encode('utf-8')
    ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', rels)]
    rels = rels.replace('</Relationships>',
                        f'<Relationship Id="rId{max(ids) + 1}" '
                        'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                        'Target="vbaProject.bin"/></Relationships>')
    fichiers['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')

    ordre = [n for n in zin.namelist()] + ['xl/vbaProject.bin']
    with zipfile.ZipFile(CIBLE, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in ordre:
            zout.writestr(n, fichiers[n])
    print('Classeur généré :', CIBLE)


def col_index(lettres):
    n = 0
    for ch in lettres:
        n = n * 26 + ord(ch) - 64
    return n


def col_lettre(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


if __name__ == '__main__':
    construire()
