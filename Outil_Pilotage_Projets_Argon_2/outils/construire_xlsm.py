"""Construit Outil_Pilotage_Projets_DLM_Garanties_Factures_Argon_2.xlsm à partir du
classeur source (.xlsx) et des modules du dossier vba/ (Windows-1252, CRLF).

Usage : python3 outils/construire_xlsm.py
"""
import html
import os
import re
import sys
import zipfile

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(ICI)
sys.path.insert(0, os.path.join(os.path.dirname(RACINE), 'Pilotage_Financements_ASER', 'outils'))
from vba_project import build_vba_project  # noqa: E402

NOM = 'Outil_Pilotage_Projets_DLM_Garanties_Factures_Argon_2'
SOURCE = os.path.join(RACINE, NOM + '_source.xlsx')
CIBLE = os.path.join(RACINE, NOM + '.xlsm')
VBA = os.path.join(RACINE, 'vba')
LIGNES = range(9, 25)

# Corrections de formules (texte XML échappé) : feuille -> liste de (colonne, ancien, nouveau)
# Les jours restants étaient calculés par rapport à une date saisie (Date Dernière Action,
# Date Référence, Transmission AC) au lieu de la date du jour : alertes figées ou fausses.
CORRECTIONS = {
    '02_SUIVI_DLM': [
        ('F', 'IF(ISNUMBER(E{r}),E{r}-I{r},&quot;&quot;)',
              'IF(ISNUMBER(E{r}),E{r}-TODAY(),&quot;&quot;)'),
    ],
    '03_REGISTRE_GARANTIES': [
        ('I', 'IF(ISNUMBER(H{r}),H{r}-G{r},&quot;&quot;)',
              'IF(ISNUMBER(H{r}),H{r}-TODAY(),&quot;&quot;)'),
        # Garantie sans échéance : n'est plus affichée « VALIDE (> 60J) »
        ('J', 'IF(B{r}=&quot;&quot;,&quot;&quot;,IF(I{r}&lt;0,',
              'IF(B{r}=&quot;&quot;,&quot;&quot;,IF(I{r}=&quot;&quot;,&quot;DATE À SAISIR&quot;,IF(I{r}&lt;0,'),
        ('K', 'IF(B{r}=&quot;&quot;,&quot;&quot;,IF(I{r}&lt;0,',
              'IF(B{r}=&quot;&quot;,&quot;&quot;,IF(I{r}=&quot;&quot;,&quot;RENSEIGNER L\'ÉCHÉANCE&quot;,IF(I{r}&lt;0,'),
    ],
    '04_FACTURES_DECOMPTES': [
        # Sans date de transmission AC, H-G donnait ~46 000 jours (priorité P3 au lieu de P1)
        ('J', 'IF(ISNUMBER(H{r}),H{r}-G{r},&quot;&quot;)',
              'IF(ISNUMBER(H{r}),H{r}-TODAY(),&quot;&quot;)'),
    ],
}
# Moyennes en #DIV/0! quand aucune valeur n'est saisie
MOYENNES = {
    '00_DASHBOARD': ['A5', 'F10', 'F11', 'F12', 'F13'],
    '01_PLAN_AMELIORATION': ['J5', 'I26'],
    '04_FACTURES_DECOMPTES': ['J5', 'I25'],
}


def lire(nom):
    code = open(os.path.join(VBA, nom), 'rb').read().decode('cp1252')
    assert '�' not in code
    return code


def entete_document(nom, clsid):
    return (f'Attribute VB_Name = "{nom}"\r\nAttribute VB_Base = "0{clsid}"\r\n'
            'Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n'
            'Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n')


def construire():
    zin = zipfile.ZipFile(SOURCE)
    f = {n: zin.read(n) for n in zin.namelist()}
    wb = f['xl/workbook.xml'].decode('utf-8')
    rels = f['xl/_rels/workbook.xml.rels'].decode('utf-8')
    cibles = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    feuilles = [(html.unescape(n), 'xl/' + cibles[r])
                for n, r in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)]
    chemin = dict(feuilles)

    for feuille, regles in CORRECTIONS.items():
        x = f[chemin[feuille]].decode('utf-8')
        for col, ancien, nouveau in regles:
            for r in LIGNES:
                motif = f'<c r="{col}{r}"'
                i = x.index(motif)
                j = x.index('</f>', i)
                a, n = ancien.format(r=r), nouveau.format(r=r)
                bloc = x[i:j]
                assert a in bloc, (feuille, col, r)
                bloc = bloc.replace(a, n, 1)
                if n.count('IF(') > a.count('IF('):  # parenthèse fermante supplémentaire
                    bloc = bloc + ')'
                x = x[:i] + bloc + x[j:]
        f[chemin[feuille]] = x.encode('utf-8')

    for feuille, cellules in MOYENNES.items():
        x = f[chemin[feuille]].decode('utf-8')
        for c in cellules:
            x, n = re.subn(r'(<c r="' + c + r'"[^>]*><f[^>]*>)((?:AVERAGE|AVERAGEIF)\(.*?\))(</f>)',
                           r'\1IFERROR(\2,0)\3', x, count=1)
            assert n == 1, (feuille, c)
        f[chemin[feuille]] = x.encode('utf-8')

    # Références à des feuilles dont le nom commence par un chiffre (« 01_PLAN_AMELIORATION! »)
    # sans apostrophes : Excel exige 'nom'! -> formules, graphiques et listes déroulantes
    # signalés comme contenu illisible à l'ouverture.
    for nom in [n for n in f if re.match(r'xl/(worksheets/sheet\d+|charts/chart\d+)\.xml$', n)]:
        x = f[nom].decode('utf-8')
        f[nom] = re.sub(r"(?<![\w'])(\d\w*)!", r"'\1'!", x).encode('utf-8')

    # Libellé du tableau de bord incohérent avec la règle réelle (45 jours)
    ss = f['xl/sharedStrings.xml'].decode('utf-8')
    assert 'Détecte les DLM &lt;= 30j' in ss
    f['xl/sharedStrings.xml'] = ss.replace('Détecte les DLM &lt;= 30j', 'Détecte les DLM &lt;= 45j').encode('utf-8')

    # Recalcul complet à l'ouverture (formules basées sur AUJOURDHUI())
    wb = wb.replace('<calcPr ', '<calcPr fullCalcOnLoad="1" ', 1)
    wb = wb.replace('<workbookPr ', '<workbookPr codeName="ThisWorkbook" ', 1)
    f['xl/workbook.xml'] = wb.encode('utf-8')

    noms_doc = []
    for i, (_, ch) in enumerate(feuilles, start=1):
        cn = f'Feuil{i}'
        noms_doc.append(cn)
        x = f[ch].decode('utf-8')
        x, n = re.subn(r'<sheetPr\b', f'<sheetPr codeName="{cn}"', x, count=1)
        assert n == 1
        f[ch] = x.encode('utf-8')

    modules = [{'name': 'ThisWorkbook', 'document': True,
                'code': entete_document('ThisWorkbook', '{00020819-0000-0000-C000-000000000046}') + lire('ThisWorkbook.cls')}]
    for cn in noms_doc:
        modules.append({'name': cn, 'document': True,
                        'code': entete_document(cn, '{00020820-0000-0000-C000-000000000046}')})
    code = lire('Macros_Pilotage_Projets_VBA.bas')
    nom_module = re.search(r'Attribute VB_Name = "([^"]+)"', code).group(1)
    modules.append({'name': nom_module, 'document': False, 'code': code})
    f['xl/vbaProject.bin'] = build_vba_project(modules)

    ct = f['[Content_Types].xml'].decode('utf-8')
    ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    ct = ct.replace('<Default Extension="xml"',
                    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"', 1)
    f['[Content_Types].xml'] = ct.encode('utf-8')
    n = max(int(i) for i in re.findall(r'Id="rId(\d+)"', rels)) + 1
    f['xl/_rels/workbook.xml.rels'] = rels.replace(
        '</Relationships>',
        f'<Relationship Id="rId{n}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
        'Target="vbaProject.bin"/></Relationships>').encode('utf-8')

    with zipfile.ZipFile(CIBLE, 'w', zipfile.ZIP_DEFLATED) as z:
        for nom in zin.namelist() + ['xl/vbaProject.bin']:
            z.writestr(nom, f[nom])
    print('Classeur généré :', CIBLE)


if __name__ == '__main__':
    construire()
