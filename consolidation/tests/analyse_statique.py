# -*- coding: utf-8 -*-
"""
Analyse statique du code VBA (complément de la recette LibreOffice, qui ne contrôle
« Option Explicit » qu'à l'exécution alors qu'Excel compile tout le projet) :
  1. noms publics définis dans plusieurs modules (« Nom ambigu » à la compilation) ;
  2. variables locales / paramètres homonymes d'une procédure du projet ;
  3. identifiants affectés ou utilisés sans déclaration (Dim, paramètre, Const,
     variable de module, procédure, mot-clé, membre après un point) ;
  4. caractères hors page de code 1252.
Usage : python analyse_statique.py <dossier_vba> [modules_supplémentaires...]
"""
import glob
import os
import re
import sys

MOTS = set('''
and as boolean byref byval call case const currency date debug dim do double each else elseif empty end enum erase error
exit explicit false for function get global gosub goto if in integer is let lib like long loop me mod module new next not nothing
null object on option optional or preserve print private property public redim rem resume return select set single static step
string sub then to true type until variant wend while with xor attribute private_module base compare binary text
'''.split())
BUILTINS = set(x.lower() for x in '''
Abs Array Asc AscW CBool CByte CCur CDate CDbl CInt CLng CSng CStr CVErr Chr Chr$ ChrW CurDir CurDir$ Date DateSerial Day
Environ Environ$ Err Format Format$ Hour IIf InStr InStrRev InputBox Int IsArray IsDate IsEmpty IsError IsNull IsNumeric
IsObject Join LBound LCase LCase$ Left Left$ Len Mid Mid$ Minute MsgBox Now Replace Right Right$ Round RGB Space Split String$
Trim Trim$ TypeName UBound UCase UCase$ Val VarType Year Month Second Timer Fix Sgn Sqr DoEvents CreateObject GetObject
Application ThisWorkbook ActiveSheet ActiveWindow ActiveCell Selection Worksheets Sheets Range Cells Rows Columns Debug
Worksheet Workbook Collection Range Object
vbCrLf vbLf vbCr vbTab vbNullString vbOKOnly vbOKCancel vbYesNo vbYesNoCancel vbRetryCancel vbYes vbNo vbCancel vbOK
vbInformation vbExclamation vbQuestion vbCritical vbSystemModal vbTextCompare vbBinaryCompare vbObjectError
vbBoolean vbDate vbDouble vbSingle vbInteger vbLong vbCurrency vbDecimal vbByte vbString vbEmpty vbNull vbError
xlContinuous xlCenter xlTop xlLandscape xlSheetVisible xlSheetHidden xlUp xlNone
'''.split())

dossier = sys.argv[1]
fichiers = sorted(glob.glob(os.path.join(dossier, '*.bas')) + glob.glob(os.path.join(dossier, '*.cls'))) + sys.argv[2:]
problemes = []

RE_PROC = re.compile(r'^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+(\w+)\s*(\((.*)\))?', re.I)
RE_END = re.compile(r'^\s*End\s+(Sub|Function|Property)\b', re.I)


def nettoyer(ligne):
    """Retire chaînes et commentaires."""
    out, dans = [], False
    i = 0
    while i < len(ligne):
        ch = ligne[i]
        if ch == '"':
            if dans and i + 1 < len(ligne) and ligne[i + 1] == '"':
                i += 2
                continue
            dans = not dans
            out.append(' ')
        elif ch == "'" and not dans:
            break
        else:
            out.append(' ' if dans else ch)
        i += 1
    return ''.join(out)


def lignes_logiques(texte):
    res, acc = [], ''
    for brute in texte.splitlines():
        l = nettoyer(brute)
        if l.rstrip().endswith(' _'):
            acc += l.rstrip()[:-1] + ' '
            continue
        res.append(acc + l)
        acc = ''
    return res


modules = {}
publics = {}
procs_tous = set()
for f in fichiers:
    nom = os.path.splitext(os.path.basename(f))[0]
    txt = open(f, encoding='utf-8').read()
    try:
        txt.encode('cp1252')
    except UnicodeEncodeError as e:
        problemes.append('%s : caractère hors cp1252 (%s)' % (nom, txt[e.start:e.end]))
    lignes = lignes_logiques(txt)
    modules[nom] = lignes
    niveau = 0
    for l in lignes:
        m = RE_PROC.match(l)
        if m:
            procs_tous.add(m.group(4).lower())
            if not (m.group(1) or '').lower().startswith('private'):
                publics.setdefault(m.group(4).lower(), []).append(nom)
            niveau = 1
            continue
        if RE_END.match(l):
            niveau = 0
            continue
        if niveau == 0:
            m = re.match(r'^\s*(Public|Global)\s+(Const\s+)?(\w+)', l, re.I)
            if m and m.group(3).lower() not in ('sub', 'function', 'property'):
                publics.setdefault(m.group(3).lower(), []).append(nom)

for n, mods in publics.items():
    if len(mods) > 1:
        problemes.append('Nom public ambigu « %s » : %s' % (n, ', '.join(mods)))

globaux = set(publics)
for nom, lignes in modules.items():
    module_vars = set()
    niveau = 0
    for l in lignes:
        if RE_PROC.match(l):
            niveau = 1
        elif RE_END.match(l):
            niveau = 0
        elif niveau == 0:
            m = re.match(r'^\s*(Private|Dim|Public|Global)\s+(Const\s+)?(.*)$', l, re.I)
            if m:
                for d in m.group(3).split(','):
                    mm = re.match(r'\s*(\w+)', d)
                    if mm:
                        module_vars.add(mm.group(1).lower())
    # analyse procédure par procédure
    proc, locaux = None, set()
    for num, l in enumerate(lignes, 1):
        m = RE_PROC.match(l)
        if m:
            proc = m.group(4)
            locaux = {proc.lower()}
            for p in (m.group(6) or '').split(','):
                pm = re.match(r'\s*(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?(\w+)', p, re.I)
                if pm and pm.group(4):
                    locaux.add(pm.group(4).lower())
                    if pm.group(4).lower() in procs_tous:
                        problemes.append('%s.%s : paramètre « %s » homonyme d\'une procédure' % (nom, proc, pm.group(4)))
            continue
        if RE_END.match(l):
            proc = None
            continue
        if proc is None:
            continue
        for instr in l.split(':'):
            dm = re.match(r'^\s*(Dim|Static|Const|ReDim(\s+Preserve)?)\s+(.*)$', instr, re.I)
            if dm and not dm.group(1).lower().startswith('redim'):
                for d in re.split(r',(?![^(]*\))', dm.group(3)):
                    mm = re.match(r'\s*(\w+)', d)
                    if mm:
                        locaux.add(mm.group(1).lower())
                        if mm.group(1).lower() in procs_tous and mm.group(1).lower() != proc.lower():
                            problemes.append('%s.%s : variable « %s » homonyme d\'une procédure' % (nom, proc, mm.group(1)))
                continue
            # étiquettes
            if re.match(r'^\s*\w+:\s*$', instr + ':'):
                pass
        connus = locaux | module_vars | globaux | procs_tous | MOTS | BUILTINS
        code = re.sub(r'^\s*\w+:(?=\s|$)', ' ', l)                     # étiquette en début de ligne
        code = re.sub(r'(?i)\b(GoTo|Resume)\s+\w+', ' ', code)
        code = re.sub(r'\b\d+(\.\d+)?(E[+-]?\d+)?\b', ' ', code)
        code = re.sub(r'&H[0-9A-F]+', ' ', code, flags=re.I)
        code = re.sub(r'\b\w+\s*:=', ' ', code)                          # arguments nommés
        for tok in re.finditer(r'(?<![\w.])([A-Za-z_]\w*\$?)', code):
            t = tok.group(1).lower()
            debut = tok.start()
            if code[:debut].rstrip().endswith('.'):
                continue                                                  # membre d'objet
            if t in connus or t.rstrip('$') in connus:
                continue
            if re.match(r'^\s*\w+\s*:\s*$', l) and t == l.strip().rstrip(':').lower():
                continue
            if re.match(r'(?i)^\s*(ReDim)', l):
                pass
            problemes.append('%s.%s ligne %d : identifiant non déclaré « %s »' % (nom, proc, num, tok.group(1)))

vus = set()
for p in problemes:
    if p not in vus:
        print(p)
        vus.add(p)
print('%d problème(s)' % len(vus))
sys.exit(1 if vus else 0)
