"""Batterie de tests de l'outil Pilotage_Financements_ASER.

  python3 outils/tester_outil.py                 -> modules de vba/ + classeur .xlsm
  python3 outils/tester_outil.py --modules DIR   -> analyse statique d'autres modules (.bas)
  python3 outils/tester_outil.py --sans-lo       -> sans les tests d'exécution LibreOffice

Tests statiques : encodage, références (feuilles, tableaux, colonnes, noms),
déclarations (Option Explicit), structure des blocs, appels, motifs à risque.
Tests du paquet .xlsm : types de contenu, projet VBA (relu par oletools),
codeName des feuilles, colonnes calculées, listes dynamiques, boutons.
Tests d'exécution (LibreOffice, mode compatibilité VBA) : chargement et
compilation de chaque module, tests unitaires des fonctions utilitaires.
"""
import argparse
import glob
import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(ICI)
XLSX = os.path.join(RACINE, 'Pilotage_Financements_ASER_source.xlsx')
XLSM = os.path.join(RACINE, 'Pilotage_Financements_ASER.xlsm')

RESULTATS = []


def verifier(nom, ok, detail=''):
    RESULTATS.append((nom, ok, detail))
    print(('  [OK]    ' if ok else '  [ECHEC] ') + nom + (('\n            ' + detail) if detail and not ok else ''))


# ---------------------------------------------------------------------------
# Lecture du classeur
# ---------------------------------------------------------------------------
def structure_classeur(chemin):
    z = zipfile.ZipFile(chemin)
    wb = z.read('xl/workbook.xml').decode('utf-8')
    feuilles = [html.unescape(n) for n in re.findall(r'<sheet name="([^"]+)"', wb)]
    noms = {html.unescape(n) for n in re.findall(r'<definedName [^>]*name="([^"]+)"', wb)}
    tables = {}
    for n in z.namelist():
        if n.startswith('xl/tables/'):
            t = z.read(n).decode('utf-8')
            nom = re.search(r'<table [^>]*name="([^"]+)"', t).group(1)
            tables[nom] = [html.unescape(c) for c in re.findall(r'<tableColumn [^>]*name="([^"]*)"', t)]
    return feuilles, noms, tables


# ---------------------------------------------------------------------------
# Analyse des modules VBA
# ---------------------------------------------------------------------------
MOTS_CLES = {'if', 'elseif', 'else', 'case', 'for', 'next', 'do', 'loop', 'while', 'wend', 'with',
             'call', 'set', 'let', 'exit', 'end', 'on', 'goto', 'resume', 'dim', 'redim', 'const',
             'select', 'then', 'msgbox', 'debug', 'err', 'option', 'private', 'public', 'attribute',
             'static', 'application', 'erase', 'me', 'mid', 'mid$'}


def lire_modules(dossier):
    mods = {}
    for chemin in sorted(glob.glob(os.path.join(dossier, '*.bas')) + glob.glob(os.path.join(dossier, '*.cls'))):
        brut = open(chemin, 'rb').read()
        mods[os.path.basename(chemin)] = brut
    return mods


def lignes_logiques(code):
    """Fusionne les continuations ' _' et retire les commentaires/chaînes."""
    out = []
    tampon, debut = '', 0
    for i, l in enumerate(code.split('\n'), start=1):
        l = l.rstrip('\r')
        if not tampon:
            debut = i
        if l.rstrip().endswith(' _'):
            tampon += l.rstrip()[:-1] + ' '
            continue
        tampon += l
        out.append((debut, tampon))
        tampon = ''
    return out


def sans_chaines(l):
    l = re.sub(r'"(?:[^"]|"")*"', '""', l)
    m = re.search(r"(^|\s|:)(?:'|Rem\s)", l)
    if m:
        l = l[:m.start()]
    return l


def decouper_procedures(code):
    procs, cur = [], None
    for num, brute in lignes_logiques(code):
        l = sans_chaines(brute).strip()
        m = re.match(r'(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\((.*)\)', l, re.I)
        if m and cur is None:
            cur = {'nom': m.group(3), 'type': m.group(2).split()[0].capitalize(),
                   'portee': (m.group(1) or 'Public').capitalize(), 'params': m.group(4),
                   'lignes': [], 'debut': num}
            continue
        if cur is not None and re.match(r'End\s+(Sub|Function|Property)\b', l, re.I):
            procs.append(cur)
            cur = None
            continue
        if cur is not None:
            cur['lignes'].append((num, brute, l))
    return procs


def noms_declares(texte):
    noms = []
    for partie in re.split(r',(?![^(]*\))', texte):
        m = re.match(r'\s*(?:ByVal\s+|ByRef\s+|Optional\s+|ParamArray\s+)*(\w+)', partie, re.I)
        if m:
            noms.append(m.group(1))
    return noms


def analyser_module(nom_fichier, code, structure, toutes_procs, modules_globaux):
    feuilles, noms, tables = structure
    erreurs = []
    lignes = lignes_logiques(code)

    # Variables de niveau module
    niveau_module = set()
    dans_proc = False
    for _, brute in lignes:
        l = sans_chaines(brute).strip()
        if re.match(r'(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property)\s', l, re.I):
            dans_proc = True
        if re.match(r'End\s+(Sub|Function|Property)\b', l, re.I):
            dans_proc = False
            continue
        if not dans_proc:
            m = re.match(r'(?:Public|Private|Dim|Global)\s+(?:Const\s+)?(?!Sub|Function|Declare|Enum|Type)(.+)', l, re.I)
            if m:
                niveau_module.update(noms_declares(re.sub(r'\s+As\s+[\w.]+', '', m.group(1))))
            m = re.match(r'(?:Public|Private)?\s*Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+(\w+)', l, re.I)
            if m:
                niveau_module.add(m.group(1))

    if not re.search(r'^\s*Option Explicit', code, re.M):
        erreurs.append(f'{nom_fichier} : Option Explicit absent')

    for p in decouper_procedures(code):
        declares = {}
        for n in noms_declares(re.sub(r'\s+As\s+[\w.]+(?:\s*=\s*[^,]+)?', '', p['params'])):
            declares[n.lower()] = p['debut']
        etiquettes = set()
        for num, brute, l in p['lignes']:
            for m in re.finditer(r'(?:^|:)\s*(?:Dim|Static|Const)\s+(.+)', l, re.I):
                partie = re.sub(r'\s+As\s+[\w.]+(?:\s*=\s*.+)?', '', m.group(1))
                partie = re.sub(r'\([^)]*\)', '', partie)
                for n in noms_declares(partie):
                    if n.lower() in declares:
                        erreurs.append(f"{nom_fichier}:{num} {p['nom']} : déclaration en double de « {n} »")
                    declares[n.lower()] = num
            m = re.match(r'^(\w+):\s*$', l)
            if m:
                etiquettes.add(m.group(1))
        connus = {k for k in declares} | {x.lower() for x in niveau_module} | {p['nom'].lower()} | \
            {x.lower() for x in modules_globaux}

        bloc = []
        for num, brute, l in p['lignes']:
            # cibles d'affectation et variables de boucle
            for instr in l.split(': '):
                instr = instr.strip()
                cibles = []
                m = re.match(r'^(?:Set\s+)?(\w+)\s*(?:\([^=]*\))?\s*=(?!=)', instr, re.I)
                if m and m.group(1).lower() not in MOTS_CLES:
                    cibles.append(m.group(1))
                m = re.match(r'^For\s+(?:Each\s+)?(\w+)', instr, re.I)
                if m:
                    cibles.append(m.group(1))
                m = re.match(r'^ReDim\s+(?:Preserve\s+)?(\w+)', instr, re.I)
                if m:
                    cibles.append(m.group(1))
                for c in cibles:
                    if c.lower() not in connus:
                        erreurs.append(f"{nom_fichier}:{num} {p['nom']} : variable non déclarée « {c} »")
                    elif c.lower() in declares and declares[c.lower()] > num:
                        erreurs.append(f"{nom_fichier}:{num} {p['nom']} : « {c} » utilisée avant sa déclaration")
            # structure des blocs
            ls = l.strip()
            if re.match(r'^If\b.*\bThen\s*$', ls, re.I):
                bloc.append(('If', num))
            elif re.match(r'^End\s+If\b', ls, re.I):
                if not bloc or bloc.pop()[0] != 'If':
                    erreurs.append(f"{nom_fichier}:{num} {p['nom']} : End If sans If")
            elif re.match(r'^For\b', ls, re.I):
                bloc.append(('For', num))
            elif re.match(r'^Next\b', ls, re.I):
                if not bloc or bloc.pop()[0] != 'For':
                    erreurs.append(f"{nom_fichier}:{num} {p['nom']} : Next sans For")
            elif re.match(r'^With\b', ls, re.I):
                bloc.append(('With', num))
            elif re.match(r'^End\s+With\b', ls, re.I):
                if not bloc or bloc.pop()[0] != 'With':
                    erreurs.append(f"{nom_fichier}:{num} {p['nom']} : End With sans With")
            elif re.match(r'^Select\s+Case\b', ls, re.I):
                bloc.append(('Select', num))
            elif re.match(r'^End\s+Select\b', ls, re.I):
                if not bloc or bloc.pop()[0] != 'Select':
                    erreurs.append(f"{nom_fichier}:{num} {p['nom']} : End Select sans Select")
            elif re.match(r'^Do\b', ls, re.I):
                bloc.append(('Do', num))
            elif re.match(r'^Loop\b', ls, re.I):
                if not bloc or bloc.pop()[0] != 'Do':
                    erreurs.append(f"{nom_fichier}:{num} {p['nom']} : Loop sans Do")
            # étiquettes GoTo
            m = re.search(r'\bGoTo\s+(\w+)', ls, re.I)
            if m and m.group(1) != '0' and m.group(1) not in etiquettes:
                erreurs.append(f"{nom_fichier}:{num} {p['nom']} : étiquette « {m.group(1)} » introuvable")
            # appels de procédures
            m = re.match(r'^Call\s+(\w+)', ls, re.I)
            if m and m.group(1).lower() not in toutes_procs:
                erreurs.append(f"{nom_fichier}:{num} {p['nom']} : procédure appelée inconnue « {m.group(1)} »")
        for b, num in bloc:
            erreurs.append(f"{nom_fichier}:{num} {p['nom']} : bloc {b} non fermé")

    # Références au classeur
    vars_tables = {}
    for _, brute in lignes:
        for m in re.finditer(r'Set\s+(\w+)\s*=\s*[\w.()"]*ListObjects\("([^"]+)"\)', brute):
            vars_tables.setdefault(m.group(1), set()).add(m.group(2))
    for num, brute in lignes:
        for m in re.finditer(r'Worksheets\("([^"]+)"\)', brute):
            if m.group(1) not in feuilles:
                erreurs.append(f'{nom_fichier}:{num} : feuille inexistante « {m.group(1)} »')
        for m in re.finditer(r'ListObjects\("([^"]+)"\)', brute):
            if m.group(1) not in tables:
                erreurs.append(f'{nom_fichier}:{num} : tableau inexistant « {m.group(1)} »')
        for m in re.finditer(r'(?:Range|Names)\("([A-Za-z_]\w*)"\)', brute):
            if m.group(1) not in noms:
                erreurs.append(f'{nom_fichier}:{num} : nom défini inexistant « {m.group(1)} »')
        for m in re.finditer(r'(\w+)\.ListColumns\("([^"]+)"\)', brute):
            t = vars_tables.get(m.group(1))
            if t and not any(m.group(2) in tables.get(x, []) for x in t):
                erreurs.append(f'{nom_fichier}:{num} : colonne « {m.group(2)} » absente de {"/".join(sorted(t))}')
            elif not t and not any(m.group(2) in c for c in tables.values()):
                erreurs.append(f'{nom_fichier}:{num} : colonne « {m.group(2)} » absente de tous les tableaux')
        # références structurées dans les formules écrites par le code
        for f in re.findall(r'"(=[^"]*(?:""[^"]*)*)"', brute):
            f = f.replace('""', '"')
            for t, col in re.findall(r'(T_\w+)\[([^\[\]]+)\]', f):
                if t not in tables or (col not in tables[t] and not col.startswith('#')):
                    erreurs.append(f'{nom_fichier}:{num} : référence structurée invalide {t}[{col}]')
            for t, col in re.findall(r'(T_\w+)\[\[#This Row\],\[([^\]]+)\]\]', f):
                col = col.replace("''", "'")
                if col not in tables.get(t, []):
                    erreurs.append(f'{nom_fichier}:{num} : référence structurée invalide {t}[@{col}]')
            for n in re.findall(r'(?<![\w\[])([A-Z][A-Za-z]+_?[A-Za-z]*)(?![\w(\[])', f):
                if n in ('TRUE', 'FALSE'):
                    continue
                if (n in ('DateRef', 'FrequenceRelance', 'DelaiPaiementDefaut', 'AlerteGarantie')
                        or '_' in n) and n not in noms:
                    erreurs.append(f'{nom_fichier}:{num} : nom défini inexistant dans une formule « {n} »')

    # Motifs à risque
    for num, brute in lignes:
        l = sans_chaines(brute)
        if re.search(r'\.Send\b', l):
            erreurs.append(f'{nom_fichier}:{num} : envoi automatique .Send interdit')
        if re.search(r'\.Value\s*&\s*"0"\)', brute):
            erreurs.append(f'{nom_fichier}:{num} : concaténation « & "0" » sur une valeur numérique '
                           '(ex. 2 relances -> "20" -> 21)')
        if re.search(r'CStr\(.*?\.Value\s*&\s*""\)', brute):
            erreurs.append(f'{nom_fichier}:{num} : CStr(cellule & "") plante (erreur 13) sur une valeur '
                           'd\'erreur #N/A/#REF!')
        if re.search(r'IIf\(IsNumeric\(', brute):
            erreurs.append(f'{nom_fichier}:{num} : IIf évalue les deux branches (CDbl sur une erreur -> erreur 13)')
        if re.search(r'"\$G\$6:G"', brute):
            erreurs.append(f'{nom_fichier}:{num} : formule de cumul figée sur $G$6 (dépend de la position du tableau)')
    for p in decouper_procedures(code):
        apres = False
        for num, brute, l in p['lignes']:
            if re.match(r'^ErrHandler:\s*$', l.strip()):
                apres = True
                appel_avant = False
                continue
            if apres:
                if re.match(r'^Call\s+(ActiverModeMacro|ProtegerFeuille|DeprotegerFeuille|InitialiserProtectionClasseur)', l.strip()):
                    appel_avant = True
                if 'Err.Description' in brute and appel_avant:
                    erreurs.append(f"{nom_fichier}:{num} {p['nom']} : Err.Description lu après un appel qui "
                                   'réinitialise Err (message d\'erreur perdu)')
                    break
    # MsgBox HTML : valeurs de cellules insérées sans échappement
    if 'HTMLBody' in code and 'EchapperHTML' not in code:
        erreurs.append(f'{nom_fichier} : contenu des cellules injecté sans échappement dans le corps HTML des e-mails')
    return erreurs


def tests_statiques(dossier, structure):
    print(f'\n== Tests statiques des modules ({dossier})')
    mods = lire_modules(dossier)
    textes = {}
    for nom, brut in mods.items():
        ok_utf8 = True
        try:
            brut.decode('ascii')
            ascii_seul = True
        except UnicodeDecodeError:
            ascii_seul = False
        try:
            brut.decode('utf-8')
            utf8 = not ascii_seul
        except UnicodeDecodeError:
            utf8 = False
        verifier(f'{nom} : encodage Windows-1252 (import VBE sans altération des accents)', not utf8,
                 'fichier encodé en UTF-8 : à l\'import, « Paramètres » devient « ParamÃ¨tres », '
                 '« Payée » devient « PayÃ©e »… (feuilles/colonnes introuvables, erreur 9)')
        verifier(f'{nom} : fins de ligne CRLF', b'\r\n' in brut and brut.count(b'\n') == brut.count(b'\r\n'))
        textes[nom] = brut.decode('cp1252' if utf8 is False else 'utf-8')
        del ok_utf8
    toutes = set()
    globaux = set()
    for nom, code in textes.items():
        for p in decouper_procedures(code):
            toutes.add(p['nom'].lower())
        for m in re.finditer(r'^(?:Public|Global)\s+(?!Sub|Function|Declare)(?:Const\s+)?(\w+)', code, re.M | re.I):
            globaux.add(m.group(1))
    for nom, code in textes.items():
        errs = analyser_module(nom, code, structure, toutes, globaux)
        verifier(f'{nom} : références, déclarations, blocs et motifs à risque', not errs, '\n            '.join(errs))
    return textes


# ---------------------------------------------------------------------------
# Paquet .xlsm
# ---------------------------------------------------------------------------
def tests_paquet(textes):
    print('\n== Tests du classeur .xlsm')
    if not os.path.exists(XLSM):
        verifier('Classeur .xlsm présent', False, XLSM)
        return
    z = zipfile.ZipFile(XLSM)
    noms = z.namelist()
    import xml.etree.ElementTree as ET
    mal_formes = []
    for n in noms:
        if n.endswith('.xml') or n.endswith('.rels'):
            try:
                ET.fromstring(z.read(n))
            except ET.ParseError as e:
                mal_formes.append(f'{n} : {e}')
    verifier('Toutes les parties XML du paquet sont bien formées', not mal_formes, '; '.join(mal_formes))
    ct = z.read('[Content_Types].xml').decode()
    verifier('Type de contenu « macroEnabled »', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' in ct)
    verifier('Type de contenu vbaProject', 'application/vnd.ms-office.vbaProject' in ct)
    rels = z.read('xl/_rels/workbook.xml.rels').decode()
    verifier('Relation vers vbaProject.bin', 'relationships/vbaProject' in rels and 'xl/vbaProject.bin' in noms)
    wb = z.read('xl/workbook.xml').decode()
    verifier('codeName ThisWorkbook', 'codeName="ThisWorkbook"' in wb)

    from oletools.olevba import VBA_Parser
    p = VBA_Parser(XLSM)
    extraits = {n: c for (_, _, n, c) in p.extract_macros()}
    verifier('Projet VBA lisible (oletools)', bool(extraits))
    norm = lambda s: [l for l in s.replace('\r\n', '\n').split('\n') if not l.startswith('Attribute VB_')]  # noqa: E731
    for nom, code in textes.items():
        cle = nom if nom in extraits else None
        verifier(f'Module {nom} injecté à l\'identique', cle is not None and norm(extraits[cle]) == norm(code))
    code_names = set()
    rels_wb = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    for rid in re.findall(r'<sheet [^>]*r:id="(rId\d+)"', wb):
        x = z.read('xl/' + rels_wb[rid]).decode()
        m = re.search(r'<sheetPr [^>]*codeName="([^"]+)"', x)
        code_names.add(m.group(1) if m else None)
    verifier('Chaque feuille a un module document VBA (codeName)',
             None not in code_names and all(c + '.cls' in extraits for c in code_names))

    calc = {}
    for n in noms:
        if n.startswith('xl/tables/'):
            t = z.read(n).decode()
            calc[re.search(r'name="([^"]+)"', t).group(1)] = t.count('<calculatedColumnFormula>')
    verifier('Colonnes calculées déclarées (T_DLM 7, T_Garanties 3, T_Factures 9, T_Plan 1)',
             calc.get('T_DLM') == 7 and calc.get('T_Garanties') == 3 and calc.get('T_Factures') == 9 and calc.get('T_Plan') == 1,
             str(calc))
    verifier('Listes déroulantes dynamiques (Liste_Accords ne passe plus en #REF!)',
             '>T_DLM[N° accord]</definedName>' in wb)
    dessin = z.read('xl/drawings/drawing1.xml').decode()
    macros_boutons = re.findall(r'macro="\[0\]!(\w+)"', dessin)
    publics = set()
    for code in textes.values():
        for pr in decouper_procedures(code):
            if pr['type'] == 'Sub' and pr['portee'] == 'Public' and not pr['params'].strip():
                publics.add(pr['nom'])
    verifier('8 boutons du Tableau_de_bord reliés à des macros publiques existantes',
             len(macros_boutons) == 8 and all(m in publics for m in macros_boutons), str(macros_boutons))
    for i in range(1, len([n for n in noms if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]) + 1):
        x = z.read(f'xl/worksheets/sheet{i}.xml').decode()
        m = re.search(r'<sheetProtection [^>]*password="([0-9A-Fa-f]+)"', x)
        if m and m.group(1).upper() != 'CE4B':
            verifier(f'sheet{i} : protection sans mot de passe (Unprotect "")', False, m.group(1))


# ---------------------------------------------------------------------------
# Exécution LibreOffice
# ---------------------------------------------------------------------------
HARNAIS = '''Option VBASupport 1
Option Explicit
Public Function T_Utilitaires() As String
    Dim r As String
    r = r & IIf(TexteCellule("  ab ") = "ab", "", "TexteCellule;")
    r = r & IIf(TexteCellule(Empty) = "", "", "TexteCellule(Empty);")
    r = r & IIf(TexteCellule(CVErr(2042)) = "", "", "TexteCellule(#N/A);")
    r = r & IIf(NombreCellule("", 9999) = 9999, "", "NombreCellule(vide);")
    r = r & IIf(NombreCellule(CVErr(2042), 9999) = 9999, "", "NombreCellule(#N/A);")
    r = r & IIf(NombreCellule(2, 0) + 1 = 3, "", "NombreCellule(2)+1;")
    r = r & IIf(NombreCellule("abc", 5) = 5, "", "NombreCellule(texte);")
    r = r & IIf(EchapperHTML("<a>&""'") = "&lt;a&gt;&amp;&quot;&#39;", "", "EchapperHTML;")
    r = r & IIf(DateTexteCellule(DateSerial(2026, 12, 15)) = "15/12/2026", "", "DateTexteCellule;")
    r = r & IIf(DateTexteCellule(CVErr(2042)) = "", "", "DateTexteCellule(#N/A);")
    r = r & IIf(ObtenirDateRef() = Date, "", "ObtenirDateRef;")
    r = r & IIf(Len(DossierExport()) > 0, "", "DossierExport;")
    r = r & IIf(Left(FormuleModele("T_DLM", "Jours restants"), 4) = "=IF(", "", "FormuleModele DLM;")
    r = r & IIf(Len(FormuleModele("T_Factures", "Délai visa" & ChrW(8594) & "AC")) > 10, "", "FormuleModele fleche;")
    r = r & IIf(FormuleModele("T_Journal", "Macro") = "", "", "FormuleModele inconnu;")
    T_Utilitaires = r
End Function
Public Function T_RegressionCompteur() As String
    ' Code d'origine : Val(CStr(cellule.Value & "0")) + 1
    T_RegressionCompteur = CStr(Val(CStr(2 & "0")) + 1) & "|" & CStr(CLng(NombreCellule(2, 0)) + 1)
End Function
'''


def tests_libreoffice(textes):
    print('\n== Tests d\'exécution LibreOffice (compatibilité VBA)')
    if not shutil.which('soffice'):
        verifier('LibreOffice disponible', False, 'soffice introuvable : tests ignorés')
        return
    import uno  # noqa: F401
    from com.sun.star.beans import PropertyValue
    from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

    profil = tempfile.mkdtemp(prefix='lo_profil_')
    copie = os.path.join(tempfile.mkdtemp(prefix='lo_doc_'), 'test.xlsm')
    shutil.copy(XLSM, copie)
    port = 2100 + os.getpid() % 500
    proc = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo',
                             '-env:UserInstallation=' + uno.systemPathToFileUrl(profil),
                             f'--accept=socket,host=127.0.0.1,port={port};urp;'],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = None
        for _ in range(90):
            try:
                local = uno.getComponentContext()
                res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
                ctx = res.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')
                break
            except Exception:
                time.sleep(1)
        desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)

        def pv(n, v):
            x = PropertyValue()
            x.Name, x.Value = n, v
            return x
        doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(copie), '_blank', 0,
                                        (pv('Hidden', True), pv('MacroExecutionMode', ALWAYS_EXECUTE_NO_WARN)))
        libs = doc.BasicLibraries
        verifier('Bibliothèque VBAProject chargée', libs.hasByName('VBAProject'))
        libs.loadLibrary('VBAProject')
        lib = libs.getByName('VBAProject')
        attendus = {os.path.splitext(n)[0] for n in textes}
        verifier('Modules présents dans LibreOffice', attendus <= set(lib.getElementNames()),
                 str(sorted(lib.getElementNames())))
        sp = doc.getScriptProvider()

        lib.insertByName('Harnais_Tests', HARNAIS)
        s = sp.getScript('vnd.sun.star.script:VBAProject.Harnais_Tests.T_Utilitaires?language=Basic&location=document')
        r = s.invoke((), (), ())[0]
        verifier('Tests unitaires des fonctions utilitaires (15 cas)', r == '', r)
        s = sp.getScript('vnd.sun.star.script:VBAProject.Harnais_Tests.T_RegressionCompteur?language=Basic&location=document')
        r = s.invoke((), (), ())[0]
        ancien, nouveau = r.split('|')
        print(f'            Compteur « Nb relances » = 2 puis +1 : code d\'origine -> {ancien}, code corrigé -> {nouveau}')
        verifier('Incrément de Nb relances correct (2 -> 3)', nouveau == '3')
        # Compilation : copie du module + fonction sonde ; l'appel de la sonde
        # compile tout le module (retour None = erreur de compilation).
        def compile_ok(nom_mod, code):
            code = '\n'.join(l for l in code.replace('\r\n', '\n').split('\n') if not l.startswith('Attribute VB_'))
            code = 'Option VBASupport 1\n' + code + '\nPublic Function ZSonde() As String\n    ZSonde = "OK"\nEnd Function\n'
            lib.insertByName(nom_mod, code)
            s = sp.getScript(f'vnd.sun.star.script:VBAProject.{nom_mod}.ZSonde?language=Basic&location=document')
            ok = s.invoke((), (), ())[0] is not None
            lib.removeByName(nom_mod)
            return ok
        temoin = 'Option Explicit\nPublic Sub Casse()\n    Dim a As Long\n    If a = 1 Then\n        a = 2\nEnd Sub\n'
        verifier('Témoin : un module volontairement erroné est bien rejeté par le compilateur',
                 not compile_ok('Verif_Temoin', temoin))
        for i, (nom, code) in enumerate(textes.items()):
            verifier(f'{nom} : compilation du module (compilateur Basic LibreOffice, mode VBA)',
                     compile_ok(f'Verif_{i}', code))
        doc.close(True)
    finally:
        try:
            desk.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=30)
        except Exception:
            proc.kill()
        shutil.rmtree(profil, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--modules', default=os.path.join(RACINE, 'vba'))
    ap.add_argument('--sans-lo', action='store_true')
    a = ap.parse_args()
    structure = structure_classeur(XLSX)
    textes = tests_statiques(a.modules, structure)
    if os.path.abspath(a.modules) == os.path.abspath(os.path.join(RACINE, 'vba')):
        tests_paquet(textes)
        if not a.sans_lo:
            tests_libreoffice(textes)
    ko = [r for r in RESULTATS if not r[1]]
    print(f'\nBilan : {len(RESULTATS) - len(ko)}/{len(RESULTATS)} tests réussis')
    sys.exit(1 if ko else 0)


if __name__ == '__main__':
    main()
