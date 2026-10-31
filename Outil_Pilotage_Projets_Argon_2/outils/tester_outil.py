"""Batterie de tests de Outil_Pilotage_Projets_DLM_Garanties_Factures_Argon_2.xlsm

  python3 outils/tester_outil.py            -> tests statiques, paquet et exécution LibreOffice
  python3 outils/tester_outil.py --original -> mêmes scénarios d'exécution sur le module
                                               et le classeur d'origine (mise en évidence des défauts)
  python3 outils/tester_outil.py --sans-lo  -> sans LibreOffice

Les macros de cet outil n'utilisent pas de tableaux structurés : LibreOffice (mode de
compatibilité VBA) les exécute réellement sur une copie du classeur ; les boîtes de
message n'apparaissent pas en mode sans interface.
"""
import argparse
import datetime
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
OUTILS_ASER = os.path.join(os.path.dirname(RACINE), 'Pilotage_Financements_ASER', 'outils')
sys.path.insert(0, OUTILS_ASER)
import importlib.util  # noqa: E402
_spec = importlib.util.spec_from_file_location('tester_aser', os.path.join(OUTILS_ASER, 'tester_outil.py'))
base = importlib.util.module_from_spec(_spec)  # fonctions d'analyse statique VBA
_spec.loader.exec_module(base)

NOM = 'Outil_Pilotage_Projets_DLM_Garanties_Factures_Argon_2'
XLSM = os.path.join(RACINE, NOM + '.xlsm')
SOURCE = os.path.join(RACINE, NOM + '_source.xlsx')
MODULE = 'Mod_Pilotage_Projets_Turbo'
MACROS_RACCOURCIS = ['Actualiser_Cockpit_Global', 'Verifier_DLM_Et_Escalade_DG', 'Controler_Garanties_60Jours',
                     'Generer_Plan_Paiement_Priorise_AC_DFP', 'Filtrer_Plan_Amelioration_Par_Horizon',
                     'Audit_Qualite_Portefeuille_Complet', 'Exporter_Rapport_Executif_PDF']
verifier = base.verifier
RESULTATS = base.RESULTATS


def feuilles_classeur(chemin):
    wb = zipfile.ZipFile(chemin).read('xl/workbook.xml').decode('utf-8')
    return [html.unescape(n) for n in re.findall(r'<sheet name="([^"]+)"', wb)]


def tests_statiques():
    print('\n== Tests statiques du module')
    textes = base.tests_statiques(os.path.join(RACINE, 'vba'), (feuilles_classeur(SOURCE), set(), {}))
    code = textes['Macros_Pilotage_Projets_VBA.bas']
    verifier('Raccourcis Ctrl+Shift annoncés réellement affectés (OnKey) aux 7 macros',
             all(re.search(r'OnKey "\^\+\w", pref & "' + m + '"', code) for m in MACROS_RACCOURCIS))
    verifier('Aucune plage de registre figée (9 à 24)', not re.search(r'"[A-Z]9:[A-Z]2[45]"|To 24\b', code))
    verifier('« Date Dernière Action » du suivi DLM jamais écrasée par une macro',
             not re.search(r'wsDLM\.Cells\([^)]*,\s*9\)\.Value\s*=', code))
    verifier('Aucun Val() sur des montants ou des jours (décimales perdues, vide = 0)',
             not any('Val(' in base.sans_chaines(l) for l in code.split('\n')))
    for pr in base.decouper_procedures(code):
        if pr['portee'] == 'Public' and pr['type'] == 'Sub' and pr['nom'] in MACROS_RACCOURCIS:
            corps = '\n'.join(b for _, b, _ in pr['lignes'])
            verifier(f'{pr["nom"]} : gestionnaire d\'erreur restaurant le mode de calcul',
                     'On Error GoTo ErrHandler' in corps and 'Desactiver_Mode_Turbo' in corps.split('ErrHandler:')[-1])
    return textes


def tests_paquet(textes):
    print('\n== Tests du classeur .xlsm')
    import xml.etree.ElementTree as ET
    z = zipfile.ZipFile(XLSM)
    mal = []
    for n in z.namelist():
        if n.endswith(('.xml', '.rels')):
            try:
                ET.fromstring(z.read(n))
            except ET.ParseError as e:
                mal.append(f'{n}: {e}')
    verifier('Parties XML bien formées', not mal, '; '.join(mal))
    ct = z.read('[Content_Types].xml').decode()
    verifier('Classeur prenant en charge les macros (.xlsm)',
             'macroEnabled.main+xml' in ct and 'vbaProject' in ct and 'xl/vbaProject.bin' in z.namelist())
    from oletools.olevba import VBA_Parser
    ext = {n: c for (_, _, n, c) in VBA_Parser(XLSM).extract_macros()}
    norm = lambda s: [l for l in s.replace('\r\n', '\n').split('\n') if not l.startswith('Attribute VB_')]  # noqa: E731
    verifier('Module injecté à l\'identique', norm(ext.get(MODULE + '.bas', '')) == norm(textes['Macros_Pilotage_Projets_VBA.bas']))
    verifier('ThisWorkbook injecté (raccourcis à l\'ouverture)', norm(ext.get('ThisWorkbook.cls', '')) == norm(textes['ThisWorkbook.cls']))
    tout = ''.join(z.read(n).decode() for n in z.namelist() if re.match(r'xl/(worksheets|charts)/\w+\.xml$', n))
    verifier('Références aux feuilles « 0x_… » entre apostrophes (formules, graphique, listes)',
             not re.search(r"(?<![\w'])\d\w*!", tout))
    s3 = z.read('xl/worksheets/sheet3.xml').decode()
    s4 = z.read('xl/worksheets/sheet4.xml').decode()
    s5 = z.read('xl/worksheets/sheet5.xml').decode()
    verifier('Jours restants DLM / garanties / factures calculés par rapport à AUJOURDHUI()',
             'E9-TODAY()' in s3 and 'H9-TODAY()' in s4 and 'H9-TODAY()' in s5 and 'E9-I9' not in s3 and 'H9-G9' not in s5)
    verifier('Recalcul complet à l\'ouverture', 'fullCalcOnLoad="1"' in z.read('xl/workbook.xml').decode())


# ---------------------------------------------------------------------------
HARNAIS = '''Option VBASupport 1
Option Explicit
Public Function T_Utilitaires() As String
    Dim r As String
    r = r & IIf(TexteSur(CVErr(2042)) = "", "", "TexteSur(#N/A);")
    r = r & IIf(TexteSur(" ab ") = "ab", "", "TexteSur;")
    r = r & IIf(EstNombre("") = False, "", "EstNombre(vide);")
    r = r & IIf(EstNombre(CVErr(2042)) = False, "", "EstNombre(#N/A);")
    r = r & IIf(EstNombre(12.5) = True, "", "EstNombre(12.5);")
    r = r & IIf(NombreSur("", 7) = 7, "", "NombreSur(vide);")
    r = r & IIf(NombreSur(1234.5, 0) = 1234.5, "", "NombreSur(decimal);")
    r = r & IIf(EstDate(DateSerial(2026, 11, 15)) = True, "", "EstDate(date);")
    r = r & IIf(EstDate("15/11/2026") = False, "", "EstDate(texte);")
    r = r & IIf(TexteJours(15) = "J-15", "", "TexteJours(15);")
    r = r & IIf(Left(TexteJours(-3), 7) = "dépassé", "", "TexteJours(-3);")
    r = r & IIf(TexteJours("") = "date non renseignée", "", "TexteJours(vide);")
    r = r & IIf(Len(MessageTronque(String(2000, "x"))) < 1000, "", "MessageTronque;")
    r = r & IIf(DerniereLigneRegistre(ThisWorkbook.Worksheets("02_SUIVI_DLM")) = 24, "", "DerniereLigne DLM;")
    r = r & IIf(DerniereLigneRegistre(ThisWorkbook.Worksheets("01_PLAN_AMELIORATION")) = 25, "", "DerniereLigne Plan;")
    T_Utilitaires = r
End Function
'''


class LO:
    def __init__(self):
        import uno
        self.uno = uno
        self.profil = tempfile.mkdtemp(prefix='lo_profil_')
        self.port = 2300 + os.getpid() % 500
        self.proc = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo',
                                      '-env:UserInstallation=' + uno.systemPathToFileUrl(self.profil),
                                      f'--accept=socket,host=127.0.0.1,port={self.port};urp;'],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        for _ in range(90):
            try:
                local = uno.getComponentContext()
                res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
                self.ctx = res.resolve(f'uno:socket,host=127.0.0.1,port={self.port};urp;StarOffice.ComponentContext')
                break
            except Exception:
                time.sleep(1)
        self.desk = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)

    def ouvrir(self, chemin):
        from com.sun.star.beans import PropertyValue
        from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

        def pv(n, v):
            x = PropertyValue()
            x.Name, x.Value = n, v
            return x
        copie = os.path.join(tempfile.mkdtemp(prefix='lo_doc_'), 'test.xlsm')
        shutil.copy(chemin, copie)
        doc = self.desk.loadComponentFromURL(self.uno.systemPathToFileUrl(copie), '_blank', 0,
                                             (pv('Hidden', True), pv('MacroExecutionMode', ALWAYS_EXECUTE_NO_WARN)))
        doc.BasicLibraries.loadLibrary('VBAProject')
        return doc

    def fermer(self):
        try:
            self.desk.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=30)
        except Exception:
            self.proc.kill()
        shutil.rmtree(self.profil, ignore_errors=True)


def appeler(doc, module, proc):
    sp = doc.getScriptProvider()
    return sp.getScript(f'vnd.sun.star.script:VBAProject.{module}.{proc}?language=Basic&location=document').invoke((), (), ())[0]


def cell(doc, feuille, ref):
    return doc.Sheets.getByName(feuille).getCellRangeByName(ref)


def date_serie(d):
    return (d - datetime.date(1899, 12, 30)).days


def scenarios(doc, module, prefixe=''):
    """Scénarios d'exécution communs (classeur corrigé ou d'origine)."""
    auj = datetime.date.today()
    fac, dlm, gar = '04_FACTURES_DECOMPTES', '02_SUIVI_DLM', '03_REGISTRE_GARANTIES'

    # 1. Suivi DLM : la date de dernière action saisie doit être conservée
    cell(doc, dlm, 'I10').setValue(date_serie(datetime.date(2026, 9, 1)))
    appeler(doc, module, 'Verifier_DLM_Et_Escalade_DG')
    doc.calculateAll()
    verifier(prefixe + 'Verifier_DLM_Et_Escalade_DG conserve la « Date Dernière Action » saisie',
             cell(doc, dlm, 'I10').getValue() == date_serie(datetime.date(2026, 9, 1)),
             f'I10 réécrite : {cell(doc, dlm, "I10").getString()}')
    e9 = cell(doc, dlm, 'E9').getValue()
    verifier(prefixe + 'Jours restants DLM = échéance - date du jour',
             cell(doc, dlm, 'F10').getValue() == cell(doc, dlm, 'E10').getValue() - date_serie(auj),
             f'F10={cell(doc, dlm, "F10").getValue()} E10={cell(doc, dlm, "E10").getValue()} aujourd\'hui={date_serie(auj)}')
    del e9

    # 2. Garanties : ligne sans date d'échéance
    cell(doc, gar, 'B12').setString('Marché test sans échéance')
    cell(doc, gar, 'F12').setValue(1000)
    appeler(doc, module, 'Controler_Garanties_60Jours')
    doc.calculateAll()
    verifier(prefixe + 'Garantie sans date d\'échéance non affichée « VALIDE »',
             'VALIDE' not in cell(doc, gar, 'J12').getString(), cell(doc, gar, 'J12').getString())

    # 3. Factures : facture urgente saisie plus bas, sans date de transmission AC
    cell(doc, fac, 'B15').setString('Décompte test urgent')
    cell(doc, fac, 'C15').setString('BID')
    cell(doc, fac, 'D15').setValue(1000)
    cell(doc, fac, 'H15').setValue(date_serie(auj + datetime.timedelta(days=10)))
    cell(doc, fac, 'L15').setString('RÉCEPTION')
    doc.calculateAll()
    verifier(prefixe + 'Facture sans transmission AC à J-10 de la DLM classée P1',
             cell(doc, fac, 'K15').getString().startswith('P1'), f'K15={cell(doc, fac, "K15").getString()}')
    appeler(doc, module, 'Generer_Plan_Paiement_Priorise_AC_DFP')
    doc.calculateAll()
    ordre = [cell(doc, fac, f'B{r}').getString() for r in range(9, 25)]
    remplies = [b for b in ordre if b]
    verifier(prefixe + 'Plan de paiement : lignes renseignées en tête, sans ligne vide intercalée',
             ordre[:len(remplies)] == remplies, str(ordre))
    verifier(prefixe + 'Plan de paiement : facture la plus urgente (DLM à J-10) en premier',
             ordre[0] == 'Décompte test urgent', str(ordre[:5]))
    prio = [cell(doc, fac, f'K{r}').getString()[:2] for r in range(9, 9 + len(remplies))]
    verifier(prefixe + 'Plan de paiement : priorités croissantes P1 -> P4', prio == sorted(prio), str(prio))
    formules = all(cell(doc, fac, f'{c}{r}').getFormula().startswith('=') for c in 'IJK' for r in range(9, 25))
    verifier(prefixe + 'Plan de paiement : formules [Auto] (I, J, K) intactes sur toutes les lignes', formules)

    # 4. Audit et actualisation : exécution sans erreur
    for proc in ('Audit_Qualite_Portefeuille_Complet', 'Actualiser_Cockpit_Global'):
        try:
            appeler(doc, module, proc)
            ok = True
        except Exception as e:
            ok = False
            print('            ', str(e)[:200])
        verifier(prefixe + f'{proc} s\'exécute', ok)


def tests_libreoffice(textes, original):
    print('\n== Tests d\'exécution LibreOffice' + (' — VERSION D\'ORIGINE' if original else ''))
    if not shutil.which('soffice'):
        verifier('LibreOffice disponible', False)
        return
    lo = LO()
    try:
        if original:
            chemin = construire_original()
            doc = lo.ouvrir(chemin)
            scenarios(doc, MODULE, '[origine] ')
            doc.close(True)
            return
        doc = lo.ouvrir(XLSM)
        lib = doc.BasicLibraries.getByName('VBAProject')
        verifier('Module chargé dans LibreOffice', MODULE in lib.getElementNames())
        lib.insertByName('Harnais', HARNAIS)
        r = appeler(doc, 'Harnais', 'T_Utilitaires')
        verifier('Tests unitaires des fonctions de lecture sécurisée (15 cas)', r == '', r)
        lib.removeByName('Harnais')

        def compile_ok(nom_mod, code):
            code = '\n'.join(l for l in code.replace('\r\n', '\n').split('\n') if not l.startswith('Attribute VB_'))
            lib.insertByName(nom_mod, 'Option VBASupport 1\n' + code + '\nPublic Function ZSonde() As String\n    ZSonde = "OK"\nEnd Function\n')
            ok = appeler(doc, nom_mod, 'ZSonde') is not None
            lib.removeByName(nom_mod)
            return ok
        verifier('Témoin : un module erroné est rejeté par le compilateur',
                 not compile_ok('VerifTemoin', 'Option Explicit\nPublic Sub X()\n    If 1 = 1 Then\nEnd Sub\n'))
        for i, (nom, code) in enumerate(textes.items()):
            verifier(f'{nom} : compilation du module', compile_ok(f'Verif{i}', code))
        scenarios(doc, MODULE)
        doc.close(True)
    finally:
        lo.fermer()


def construire_original():
    """Classeur d'origine + module d'origine (décodé correctement) pour comparaison."""
    import vba_project
    chemin_rel = os.path.relpath(os.path.join(RACINE, 'vba', 'Macros_Pilotage_Projets_VBA.bas'), os.path.dirname(RACINE))
    depot = os.path.dirname(RACINE)
    commit = subprocess.check_output(['git', 'log', '--diff-filter=A', '--format=%H', '--', chemin_rel],
                                     cwd=depot, text=True).split()[-1]
    code = subprocess.check_output(['git', 'show', f'{commit}:{chemin_rel}'], cwd=depot).decode('utf-8')
    z = zipfile.ZipFile(SOURCE)
    f = {n: z.read(n) for n in z.namelist()}
    f['xl/vbaProject.bin'] = vba_project.build_vba_project([{'name': MODULE, 'document': False, 'code': code}])
    f['[Content_Types].xml'] = f['[Content_Types].xml'].decode().replace(
        'spreadsheetml.sheet.main+xml', 'spreadsheetml.sheet.main+xml').replace(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        'application/vnd.ms-excel.sheet.macroEnabled.main+xml').replace(
        '<Default Extension="xml"', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"', 1).encode()
    rels = f['xl/_rels/workbook.xml.rels'].decode()
    f['xl/_rels/workbook.xml.rels'] = rels.replace('</Relationships>',
        '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
        'Target="vbaProject.bin"/></Relationships>').encode()
    chemin = os.path.join(tempfile.mkdtemp(prefix='origine_'), 'origine.xlsm')
    with zipfile.ZipFile(chemin, 'w', zipfile.ZIP_DEFLATED) as out:
        for n in list(f):
            out.writestr(n, f[n])
    return chemin


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--original', action='store_true')
    ap.add_argument('--sans-lo', action='store_true')
    a = ap.parse_args()
    if a.original:
        tests_libreoffice(None, True)
    else:
        textes = tests_statiques()
        tests_paquet(textes)
        if not a.sans_lo:
            tests_libreoffice(textes, False)
    ko = [r for r in RESULTATS if not r[1]]
    print(f'\nBilan : {len(RESULTATS) - len(ko)}/{len(RESULTATS)} tests réussis')
    sys.exit(1 if ko else 0)


if __name__ == '__main__':
    main()
