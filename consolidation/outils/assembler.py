# -*- coding: utf-8 -*-
"""
Assemble le classeur .xlsm final :
  - projet VBA (source seule) construit depuis consolidation/vba/*.bas|*.cls,
    avec un module de document par feuille (noms de code lus dans le classeur) ;
  - boutons de macros (formes) sur la feuille Sommaire ;
  - types de contenu « classeur prenant en charge les macros ».

Usage : python assembler.py <base.xlsx> <sortie.xlsm> [module_test.bas ...]
"""
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vbaproj import BASE_CLASSEUR, BASE_FEUILLE, construire_vba_project, source_module  # noqa: E402

ICI = os.path.dirname(os.path.abspath(__file__))
DOSSIER_VBA = os.path.join(ICI, '..', 'vba')
MODULES_STD = ['modConfig', 'modCommun', 'modMoteur', 'modRelances', 'modRapports', 'modRegistres']

# Boutons du Sommaire : (section, libellé, macro)
BOUTONS = [
    ('Synthèse et contrôle', None, None),
    (None, 'Actualiser tout', 'ActualiserTout'),
    (None, 'Synthèse des alertes', 'SyntheseAlertes'),
    (None, 'Contrôle qualité des saisies', 'ControleQualite'),
    (None, 'Filtrer les alertes (registres)', 'FiltrerAlertes'),
    (None, 'Retirer les filtres', 'AfficherTout'),
    ('Saisie et relances', None, None),
    (None, 'Nouvelle ligne (registre actif)', 'NouvelleLigne'),
    (None, 'Préparer les relances (Outlook)', 'PreparerRelances'),
    (None, 'Enregistrer une relance DLM', 'EnregistrerRelance'),
    (None, 'Effacer les exemples', 'SupprimerExemples'),
    ('Revue et rapports', None, None),
    (None, "Note d'escalade DG (PDF)", 'NoteEscaladeDG'),
    (None, 'Exporter la revue en PDF', 'ExporterRevuePDF'),
    (None, 'Archiver la revue (instantané)', 'ArchiverRevue'),
    (None, 'Archiver les clôtures', 'ArchiverClotures'),
    (None, 'Figer la date de revue', 'FigerDateReference'),
    (None, 'Revenir à AUJOURDHUI()', 'LibererDateReference'),
]

NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def _esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def dessin_boutons():
    x, y, larg, haut = 8700000, 180000, 2571750, 266700
    parts = []
    nid = 1001

    def zone_texte(texte, sous, y0, taille=1100):
        nonlocal nid
        nid += 1
        p_sous = ''
        if sous:
            p_sous = ('<a:p><a:pPr algn="l"/><a:r><a:rPr lang="fr-FR" sz="800" i="1"><a:solidFill><a:srgbClr val="595959"/>'
                      '</a:solidFill></a:rPr><a:t>%s</a:t></a:r></a:p>' % _esc(sous))
        return ('<xdr:absoluteAnchor><xdr:pos x="%d" y="%d"/><xdr:ext cx="%d" cy="%d"/>'
                '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="%d" name="txtPilotage_%d"/>'
                '<xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
                '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>'
                '<xdr:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="b"/><a:lstStyle/>'
                '<a:p><a:pPr algn="l"/><a:r><a:rPr lang="fr-FR" sz="%d" b="1"><a:solidFill><a:srgbClr val="1F3864"/>'
                '</a:solidFill></a:rPr><a:t>%s</a:t></a:r></a:p>%s</xdr:txBody></xdr:sp>'
                '<xdr:clientData fPrintsWithSheet="0"/></xdr:absoluteAnchor>'
                % (x, y0, larg, 380000 if sous else 230000, nid, nid, x, y0, larg, 380000 if sous else 230000,
                   taille, _esc(texte), p_sous))

    def bouton(texte, macro, y0):
        nonlocal nid
        nid += 1
        return ('<xdr:absoluteAnchor><xdr:pos x="%d" y="%d"/><xdr:ext cx="%d" cy="%d"/>'
                '<xdr:sp macro="[0]!%s" textlink=""><xdr:nvSpPr><xdr:cNvPr id="%d" name="btn_%s" descr="%s - macro %s"/>'
                '<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
                '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>'
                '<a:ln w="9525"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:ln></xdr:spPr>'
                '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="36000" tIns="0" '
                'rIns="36000" bIns="0" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r>'
                '<a:rPr lang="fr-FR" sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>'
                '<a:t>%s</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData fPrintsWithSheet="0"/></xdr:absoluteAnchor>'
                % (x, y0, larg, haut, macro, nid, macro, _esc(texte), macro, x, y0, larg, haut, _esc(texte)))

    parts.append(zone_texte('Macros - pilotage consolidé', "Activer les macros à l'ouverture · liste : Alt+F8", y))
    y += 470000
    for section, libelle, macro in BOUTONS:
        if section:
            y += 60000
            parts.append(zone_texte(section, None, y, 900))
            y += 250000
        else:
            parts.append(bouton(libelle, macro, y))
            y += haut + 70000
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="%s" xmlns:a="%s">%s</xdr:wsDr>'
            % (NS_XDR, NS_A, ''.join(parts)))


def lire_source(nom):
    for ext in ('.bas', '.cls'):
        p = os.path.join(DOSSIER_VBA, nom + ext)
        if os.path.exists(p):
            return open(p, encoding='utf-8').read()
    raise FileNotFoundError(nom)


def assembler(base, sortie, modules_test=()):
    zin = zipfile.ZipFile(base)
    fichiers = {n: zin.read(n) for n in zin.namelist()}
    wbxml = fichiers['xl/workbook.xml'].decode('utf-8')
    rels = fichiers['xl/_rels/workbook.xml.rels'].decode('utf-8')

    # feuilles dans l'ordre, fichier XML et nom de code
    cible = dict(re.findall(r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="/?(?:xl/)?([^"]+)"', rels))
    cible.update({a: b for b, a in re.findall(r'<Relationship[^>]*Target="/?(?:xl/)?([^"]+)"[^>]*Id="(rId\d+)"', rels)})
    feuilles = []
    for nom, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml):
        chemin = 'xl/' + cible[rid]
        xml = fichiers[chemin].decode('utf-8')
        code = re.search(r'<sheetPr[^>]*codeName="([^"]+)"', xml).group(1)
        feuilles.append((nom.replace('&amp;', '&'), chemin, code))
    assert 'codeName="ThisWorkbook"' in wbxml, 'workbookPr codeName manquant'

    # --- projet VBA
    modules = [('ThisWorkbook', 'doc', source_module('ThisWorkbook', 'doc', lire_source('ThisWorkbook'), BASE_CLASSEUR))]
    for _nom, _ch, code in feuilles:
        modules.append((code, 'doc', source_module(code, 'doc', 'Option Explicit\r\n', BASE_FEUILLE)))
    for m in MODULES_STD:
        modules.append((m, 'std', source_module(m, 'std', lire_source(m))))
    for chemin in modules_test:
        nom = os.path.splitext(os.path.basename(chemin))[0]
        modules.append((nom, 'std', source_module(nom, 'std', open(chemin, encoding='utf-8').read())))
    fichiers['xl/vbaProject.bin'] = construire_vba_project(modules)
    if 'vbaProject.bin' not in rels:
        rels = rels.replace('</Relationships>',
                            '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                            'Target="vbaProject.bin"/></Relationships>')

    # --- boutons du Sommaire
    som = [f for f in feuilles if f[0] == 'Sommaire'][0][1]
    sxml = fichiers[som].decode('utf-8')
    assert '<drawing ' not in sxml and '<legacyDrawing' not in sxml
    srels_nom = som.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    srels = fichiers.get(srels_nom, b'').decode('utf-8') or \
        ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    srels = srels.replace('</Relationships>',
                          '<Relationship Id="rIdBtn" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                          'relationships/drawing" Target="../drawings/drawingBoutons.xml"/></Relationships>')
    fichiers[srels_nom] = srels.encode('utf-8')
    if 'xmlns:r=' not in sxml[:500]:
        sxml = sxml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    m = re.search(r'<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b', sxml)
    pos = m.start() if m else sxml.rindex('</worksheet>')
    sxml = sxml[:pos] + '<drawing r:id="rIdBtn"/>' + sxml[pos:]
    fichiers[som] = sxml.encode('utf-8')
    fichiers['xl/drawings/drawingBoutons.xml'] = dessin_boutons().encode('utf-8')

    # --- types de contenu
    ct = fichiers['[Content_Types].xml'].decode('utf-8')
    ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    if 'Extension="bin"' not in ct:
        ct = ct.replace('<Default ', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)
    ct = ct.replace('</Types>', '<Override PartName="/xl/drawings/drawingBoutons.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    fichiers['[Content_Types].xml'] = ct.encode('utf-8')
    fichiers['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')

    ordre = ['[Content_Types].xml'] + [n for n in fichiers if n != '[Content_Types].xml']
    with zipfile.ZipFile(sortie, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in ordre:
            zout.writestr(n, fichiers[n])
    return [m[0] for m in modules]


if __name__ == '__main__':
    mods = assembler(sys.argv[1], sys.argv[2], sys.argv[3:])
    print('OK', sys.argv[2], len(mods), 'modules')
