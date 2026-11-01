# -*- coding: utf-8 -*-
"""
Recette automatisée des macros du classeur consolidé, exécutée dans LibreOffice
(mode de compatibilité VBA). Les dialogues sont scriptés via modTests.

Usage : python recette.py <classeur_test.xlsm> <dossier_sortie>
Sortie : rapport texte (une ligne par vérification) + code retour 1 si échec.
"""
import datetime as dt
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import uno  # noqa: F401
from com.sun.star.beans import PropertyValue

try:                                    # environnement bac à sable (shim AF_UNIX) si disponible
    sys.path.insert(0, os.environ.get('SOFFICE_SCRIPTS', ''))
    from office.soffice import get_soffice_env  # noqa: E402
except ImportError:
    def get_soffice_env():
        env = os.environ.copy()
        env['SAL_USE_VCLPLUGIN'] = 'svp'
        return env

CLASSEUR, SORTIE = sys.argv[1], sys.argv[2]
os.makedirs(SORTIE, exist_ok=True)
AUJ = dt.date.today()
SERIE0 = dt.date(1899, 12, 30)
RESULTATS = []


def serie(d):
    return (d - SERIE0).days


def J(n):            # date de référence + n jours (numéro de série)
    return serie(AUJ) + n


def pv(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


class Session:
    def __init__(self):
        self.profil = tempfile.mkdtemp(prefix='lo_recette_')
        port = 2010
        self.proc = subprocess.Popen(
            ['soffice', '--headless', '--invisible', '--nologo', '--norestore',
             f'-env:UserInstallation={Path(self.profil).as_uri()}',
             f'--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'],
            env=get_soffice_env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        for _ in range(60):
            try:
                self.ctx = res.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')
                break
            except Exception:
                time.sleep(1)
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        url = Path(CLASSEUR).resolve().as_uri()
        self.doc = self.desktop.loadComponentFromURL(url, '_blank', 0, (pv('MacroExecutionMode', 4),))
        self.sp = self.doc.getScriptProvider()

    def appel(self, nom, *args):
        s = self.sp.getScript(f'vnd.sun.star.script:VBAProject.{nom}?language=Basic&location=document')
        return s.invoke(tuple(args), (), ())[0]

    def macro(self, nom, reponses=''):
        self.appel('modTests.T_Init', reponses, SORTIE)
        t0 = time.time()
        self.appel(nom)
        duree = time.time() - t0
        return self.appel('modTests.T_Trace'), duree

    def f(self, nom):
        return self.doc.Sheets.getByName(nom)

    def c(self, feuille, adr):
        return self.f(feuille).getCellRangeByName(adr)

    def txt(self, feuille, adr):
        return self.c(feuille, adr).getString()

    def num(self, feuille, adr):
        return self.c(feuille, adr).getValue()

    def formule(self, feuille, adr):
        return self.c(feuille, adr).getFormula()

    def ecrire(self, feuille, adr, v):
        c = self.c(feuille, adr)
        if isinstance(v, str):
            c.setString(v)
        else:
            c.setValue(v)

    def fermer(self):
        try:
            self.doc.close(True)
        except Exception:
            pass
        try:
            self.desktop.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(15)
        except Exception:
            self.proc.kill()


def verifier(nom, condition, detail=''):
    RESULTATS.append((bool(condition), nom, detail))
    print(('OK   ' if condition else 'ÉCHEC') + ' | ' + nom + (' | ' + str(detail)[:300] if detail else ''))


def lignes_feuille(s, feuille, col, r1, r2):
    return [s.txt(feuille, '%s%d' % (col, r)) for r in range(r1, r2 + 1)]


# =============================================================================
s = Session()
try:
    calc = s.doc
    for n in calc.Sheets.getElementNames():
        calc.Sheets.getByName(n).unprotect('')

    # ---------------------------------------------------------------- données
    P = 'Paramètres'
    s.ecrire(P, 'H14', 600)                       # taux USD
    s.ecrire(P, 'AD24', 'dper@example.org')       # e-mail du détenteur DPER
    s.ecrire(P, 'AD25', 'dfp@example.org')        # e-mail DFP
    s.ecrire('Plan_Action', 'C4', J(-31))         # lancement du plan
    for col, v in zip('BCDEFGHJKMR', ['Lot 1', 'Projet A', 'BID', 'DLM-01', 'Entreprise A', 5e8, 0, serie(dt.date(2025, 6, 1)), 18, 0.5, 'En cours']):
        s.ecrire('Marchés', col + '6', v)
    s.ecrire('Marchés', 'A6', 'M-001')
    for col, v in zip('ABCDEFGJKMR', ['M-002', 'Lot 2', 'Projet B', 'BOAD', 'DLM-02', 'Entreprise B', 3e8, serie(dt.date(2026, 1, 1)), 12, 0.2, 'En cours']):
        s.ecrire('Marchés', col + '7', v)

    D = 'DLM'
    dlm = {
        6: dict(C='Accord BID 1', D='Projet A', E=serie(dt.date(2022, 1, 15)), F=serie(dt.date(2026, 6, 30)), G=1, H=J(30), I='EUR', J=20e6, K=5e6, V="Rapport d'audit", W='DPER', X=J(-10), Y=2, U='Déposée – en instruction'),
        7: dict(C='Accord BOAD 1', D='Projet B', E=serie(dt.date(2021, 1, 1)), H=J(-5), I='XOF', J=1e9, K=2e8),
        8: dict(C='Accord FVC 1', H=J(150), I='USD', J=1e6, K=0, V='ANO bailleur', W='DFP'),
        9: dict(C='Accord ADFD 1', H=J(400), I='AED', J=5e6, K=1e6, V='Lettre de demande', W='DFP', X=J(-2), Y=2),
        10: dict(A='', B='BID', C='Accord sans identifiant', H=J(20), I='XOF', J=1e6),
        11: dict(A='DLM-06', B='BOAD', C='Accord incohérent', H=J(600), I='XOF', J=100, K=150),
        12: dict(A='DLM-07', B='BID', C='Accord 7', H=J(300), I='XOF', J=1e6, V='Pièce X', W='DPER'),
        13: dict(A='DLM-08', B='FVC', C='Accord 8', H=J(300), I='XOF', J=1e6, V='Pièce Y', W='Entreprise'),
    }
    for r, d in dlm.items():
        for col, v in d.items():
            s.ecrire(D, '%s%d' % (col, r), v)

    G = 'Garanties'
    gar = {
        6: dict(A='GAR-001', B='M-001', F='Garantie de bonne exécution', H='GBE-1', I='XOF', J=25e6, L=serie(dt.date(2025, 12, 1)), M=J(20), O='Valide', T='DPER'),
        7: dict(A='GAR-002', B='M-002', F="Caution d'avance de démarrage", H='CA-2', I='EUR', J=1e5, L=serie(dt.date(2025, 10, 1)), M=J(-3), O='Valide'),
        8: dict(A='GAR-003', B='M-001', F='Garantie de bonne exécution', H='GBE-3', I='XOF', J=1e6, L=serie(dt.date(2025, 1, 1)), M=J(100), O='Mainlevée accordée'),
        9: dict(B='M-002', F='Autre', H='GBE-9', J=2e6, L=serie(dt.date(2025, 1, 1)), M=J(-1), O='Valide'),
    }
    for r, d in gar.items():
        for col, v in d.items():
            s.ecrire(G, '%s%d' % (col, r), v)

    F = 'Factures'
    fac = {
        6: dict(A='FAC-001', B='DEC-1', C='M-001', H='Décompte provisoire', I=100e6, J=J(-100), K=J(-90), L=J(-80)),
        7: dict(A='FAC-002', B='DEC-2', C='M-002', I=50e6, J=J(-10)),
        8: dict(A='FAC-003', B='DEC-3', C='M-001', I=80e6, J=J(-20), K=J(-15), L=J(-10), M=J(-5), N=30e6),
        9: dict(A='FAC-004', B='DEC-4', C='M-002', I=20e6, J=J(-50), K=J(-45), L=J(-40), M=J(-1)),
        10: dict(A='FAC-005', B='DEC-1', C='M-001', I=1e6, J=J(-5)),
        11: dict(A='FAC-006', B='DEC-6', C='M-999', I=5e6, J=J(-3)),
        12: dict(A='FAC-007', B='DEC-7', C='M-002', I=1e6, J=J(-3), K=J(-6)),
    }
    for r, d in fac.items():
        for col, v in d.items():
            s.ecrire(F, '%s%d' % (col, r), v)
    calc.calculateAll()
    for n in calc.Sheets.getElementNames():
        if n not in ('Alertes', 'Contrôle_Qualité'):
            calc.Sheets.getByName(n).protect('')

    # ------------------------------------------------------- formules (contrôle)
    verifier('DLM : niveau ESCALADE DG (DLM à 30 j)', s.txt(D, 'AB6') == 'ESCALADE DG', s.txt(D, 'AB6'))
    verifier('DLM : niveau ÉCHUE', s.txt(D, 'AB7') == 'ÉCHUE', s.txt(D, 'AB7'))
    verifier('DLM : niveau VIGILANCE', s.txt(D, 'AB8') == 'VIGILANCE', s.txt(D, 'AB8'))
    verifier('DLM : escalade par nombre de relances', s.txt(D, 'AB9') == 'ESCALADE DG – RELANCES' and s.txt(D, 'AC9') == 'OUI', s.txt(D, 'AB9'))
    verifier('DLM : relance due sans relance antérieure', s.txt(D, 'AA8') == 'OUI', s.txt(D, 'Z8'))
    verifier('DLM : solde FCFA en USD (taux 600)', abs(s.num(D, 'P8') - 6e8) < 1, s.num(D, 'P8'))
    verifier('DLM : solde FCFA vide si taux manquant (AED)', s.txt(D, 'P9') == '', s.txt(D, 'P9'))
    verifier('DLM : solde jamais négatif (décaissé > accordé)', s.num(D, 'L11') == 0, s.num(D, 'L11'))
    verifier('Garanties : montant EUR converti', abs(s.num(G, 'K7') - 65595700) < 1, s.num(G, 'K7'))
    verifier('Garanties : ALERTE / EXPIRÉE / CLÔTURÉE', [s.txt(G, 'R%d' % r) for r in (6, 7, 8, 9)] == ['ALERTE', 'EXPIRÉE', 'CLÔTURÉE', 'EXPIRÉE'],
             [s.txt(G, 'R%d' % r) for r in (6, 7, 8, 9)])
    verifier('Factures : étapes', [s.txt(F, 'Q%d' % r) for r in (6, 7, 8, 9)] == ['3-Transmis AC', '1-Reçu', '4-Payé partiellement', '5-Payé'],
             [s.txt(F, 'Q%d' % r) for r in (6, 7, 8, 9)])
    verifier('Factures : reste à payer (partiel / soldé)', (s.num(F, 'P8'), s.num(F, 'P9')) == (50e6, 0), (s.num(F, 'P8'), s.num(F, 'P9')))
    verifier('Factures : retard contractuel 40 j et HORS DÉLAI', s.num(F, 'V6') == 40 and s.txt(F, 'W6') == 'HORS DÉLAI', (s.num(F, 'V6'), s.txt(F, 'W6')))
    verifier('Factures : priorités P1 / P0 / PX', [s.txt(F, 'Z%d' % r) for r in (6, 7, 11)] == ['P1', 'P0', 'PX'], [s.txt(F, 'Z%d' % r) for r in (6, 7, 11)])
    ordre = [s.txt('Plan_Paiement', 'D%d' % r) for r in range(16, 23)]
    verifier('Plan de paiement : ordre DLM, puis retard, puis montant', ordre == ['DEC-2', 'DEC-7', 'DEC-1', 'DEC-3', 'DEC-1', 'DEC-6', ''], ordre)
    verifier('Plan de paiement : cumul', s.num('Plan_Paiement', 'N21') == 50e6 + 1e6 + 100e6 + 50e6 + 1e6 + 5e6, s.num('Plan_Paiement', 'N21'))
    verifier('Marchés : montant payé = TTC - reste', s.num('Marchés', 'N6') == 30e6, s.num('Marchés', 'N6'))

    # ------------------------------------------------------- notification
    tr = s.appel('modTests.T_Notification')
    verifier('Ouverture : notification des alertes critiques', 'alerte(s) critique(s)' in tr and 'DLM' in tr, tr)

    # ------------------------------------------------------- ActualiserTout
    tr, d = s.macro('modMoteur.ActualiserTout')
    verifier('ActualiserTout : message de synthèse', 'Classeur recalculé' in tr, tr)
    alertes = [(s.txt('Alertes', 'A%d' % r), s.txt('Alertes', 'B%d' % r), s.txt('Alertes', 'D%d' % r), s.txt('Alertes', 'F%d' % r)) for r in range(5, 60)]
    alertes = [a for a in alertes if a[0]]
    verifier('Alertes : ligne DLM sans identifiant signalée', any(a[1] == 'DLM' and a[2] == '(ligne 10)' for a in alertes), alertes[:5])
    verifier('Alertes : garantie sans ID repérée par son N°', any(a[1] == 'Garanties' and a[2] == 'GBE-9' for a in alertes))
    verifier('Alertes : critiques triées avant attention', [a[0] for a in alertes] == sorted([a[0] for a in alertes], key=lambda g: g != 'CRITIQUE'))
    verifier('Alertes : hors délai facture', any(a[1] == 'Factures' and a[3] == 'HORS DÉLAI' for a in alertes))
    qc = [(s.txt('Contrôle_Qualité', 'A%d' % r), s.txt('Contrôle_Qualité', 'B%d' % r), s.txt('Contrôle_Qualité', 'E%d' % r)) for r in range(5, 120)]
    qc = [q for q in qc if q[0]]
    controles = set(q[2] for q in qc)
    for attendu in ('Identifiant manquant', 'N° de facture en doublon', 'Taux de change manquant', 'Montant incohérent',
                    'Dates incohérentes', 'Référence inconnue', 'E-mail du détenteur manquant'):
        verifier('Contrôle qualité : « %s »' % attendu, attendu in controles, sorted(controles))
    verifier('ActualiserTout : tableau de bord horodaté', s.num('Tableau_de_bord', 'G2') > serie(AUJ), s.txt('Tableau_de_bord', 'G2'))
    verifier('ActualiserTout : feuilles reprotégées', all(s.appel('modTests.T_Protegee', n) for n in ('DLM', 'Tableau_de_bord', 'Journal')))
    verifier('ActualiserTout : durée < 60 s', d < 60, '%.1f s' % d)

    # ------------------------------------------------------- Synthèse / QC seules
    tr, _ = s.macro('modMoteur.SyntheseAlertes')
    verifier('SyntheseAlertes : feuille reconstruite', s.txt('Alertes', 'A4') == 'Gravité' and s.txt('Alertes', 'A5') in ('CRITIQUE', 'ATTENTION'))
    tr, _ = s.macro('modMoteur.ControleQualite')
    verifier('ControleQualite : message', 'anomalie(s) détectée(s)' in tr, tr)

    # ------------------------------------------------------- Filtres
    tr, _ = s.macro('modMoteur.FiltrerAlertes')
    verifier('FiltrerAlertes : message', 'lignes en alerte' in tr, tr)
    verifier('FiltrerAlertes : DLM en alerte visibles (6-10)', not any(s.appel('modTests.T_Masquee', D, r) for r in range(6, 11)))
    verifier('FiltrerAlertes : DLM sans alerte masquée (11)', s.appel('modTests.T_Masquee', D, 11))
    verifier('FiltrerAlertes : 2e règle prise en compte (relance due DLM-08)', not s.appel('modTests.T_Masquee', D, 13))
    verifier('FiltrerAlertes : facture hors délai visible', not s.appel('modTests.T_Masquee', F, 6))
    verifier('FiltrerAlertes : facture sans alerte masquée (PX)', s.appel('modTests.T_Masquee', F, 11))
    verifier('FiltrerAlertes : feuille reprotégée', s.appel('modTests.T_Protegee', D))
    tr, _ = s.macro('modMoteur.AfficherTout')
    verifier('AfficherTout : toutes les lignes réaffichées', not any(s.appel('modTests.T_Masquee', D, r) for r in range(6, 30))
             and not s.appel('modTests.T_Masquee', F, 11), tr)

    # ------------------------------------------------------- Relances
    tr, _ = s.macro('modRelances.EnregistrerRelance', 'DLM-03;XXX|OUI')
    verifier('EnregistrerRelance : identifiant inconnu signalé', 'introuvable' in tr and 'XXX' in tr, tr)
    verifier('EnregistrerRelance : date et compteur', s.num(D, 'X8') == serie(AUJ) and s.num(D, 'Y8') == 1, (s.txt(D, 'X8'), s.num(D, 'Y8')))
    verifier('EnregistrerRelance : prochaine relance recalculée', s.num(D, 'Z8') == serie(AUJ) + 7, s.txt(D, 'Z8'))
    tr, _ = s.macro('modRelances.EnregistrerRelance', '')
    verifier('EnregistrerRelance : annulation sans effet', s.num(D, 'Y8') == 1)
    tr, _ = s.macro('modRelances.EnregistrerRelance', 'DLM-04|NON')
    verifier('EnregistrerRelance : refus de confirmation sans effet', s.num(D, 'Y9') == 2)
    tr, _ = s.macro('modRelances.PreparerRelances', 'OUI|NON')
    verifier('PreparerRelances : un brouillon par détenteur', tr.count('[BROUILLON]') == 2, tr)
    verifier('PreparerRelances : e-mail du détenteur', 'dper@example.org' in tr)
    verifier('PreparerRelances : avertissement e-mail manquant', 'e-mail du détenteur non renseigné' in tr)
    verifier('PreparerRelances : mise à jour après « Oui »', s.num(D, 'Y6') == 3 and s.num(D, 'Y12') == 1, (s.num(D, 'Y6'), s.num(D, 'Y12')))
    verifier('PreparerRelances : aucune mise à jour après « Non »', s.txt(D, 'Y13') == '' and s.txt(D, 'X13') == '')
    tr, _ = s.macro('modRelances.PreparerRelances', 'ANNULER')
    verifier('PreparerRelances : « Annuler » arrête sans mise à jour', s.txt(D, 'Y13') == '', tr)
    journal = [s.txt('Journal', 'D%d' % r) for r in range(5, 200)]
    verifier('Journal : historique « Relance confirmée »', any('Relance confirmée | Accord: DLM-01' in j for j in journal))

    # ------------------------------------------------------- Note d'escalade DG
    tr, _ = s.macro('modRapports.NoteEscaladeDG')
    note = [s.txt('Note_Escalade_DG', 'B%d' % r) for r in range(9, 20)]
    verifier("NoteEscaladeDG : accords en escalade listés", sum(1 for n in note if n) == 4, note)
    verifier("NoteEscaladeDG : historique des relances", 'relances :' in s.txt('Note_Escalade_DG', 'K9') and '/' in s.txt('Note_Escalade_DG', 'K9'), s.txt('Note_Escalade_DG', 'K9'))
    pdfs = [p for p in os.listdir(SORTIE) if p.startswith('Note_Escalade_DG') and p.endswith('.pdf')]
    verifier('NoteEscaladeDG : PDF produit', len(pdfs) >= 1, tr)

    # ------------------------------------------------------- Date de référence
    tr, _ = s.macro('modMoteur.FigerDateReference', '31/02/2026')
    verifier('FigerDateReference : date impossible refusée', 'Date non reconnue' in tr and s.formule(P, 'C5').upper().startswith('=TODAY'), tr)
    tr, _ = s.macro('modMoteur.FigerDateReference', '15/12/2026')
    verifier('FigerDateReference : date figée', s.num(P, 'C5') == serie(dt.date(2026, 12, 15)) and not s.formule(P, 'C5').startswith('='), s.formule(P, 'C5'))
    verifier('FigerDateReference : calculs à la date figée', s.num(D, 'Q6') == J(30) - serie(dt.date(2026, 12, 15)), s.num(D, 'Q6'))
    verifier('FigerDateReference : Paramètres reprotégés', s.appel('modTests.T_Protegee', P))
    tr, _ = s.macro('modMoteur.ControleQualite')
    verifier('Contrôle qualité : date figée signalée', any(s.txt('Contrôle_Qualité', 'E%d' % r) == 'Date de référence figée' for r in range(5, 120)))
    tr, _ = s.macro('modMoteur.LibererDateReference')
    verifier('LibererDateReference : =AUJOURDHUI() rétabli', s.formule(P, 'C5').upper().startswith('=TODAY'), s.formule(P, 'C5'))

    # ------------------------------------------------------- Instantané / PDF
    tr, _ = s.macro('modMoteur.ArchiverRevue', 'OUI')
    nom_hist = 'Hist_' + AUJ.strftime('%Y-%m-%d')
    ok = calc.Sheets.hasByName(nom_hist)
    verifier('ArchiverRevue : instantané créé', ok, tr)
    if ok:
        h = calc.Sheets.getByName(nom_hist)
        formules = sum(1 for r in range(1, 40) for cc in range(0, 12) if h.getCellByPosition(cc, r).getFormula().startswith('='))
        verifier('ArchiverRevue : valeurs sans formules', formules == 0, formules)
        verifier('ArchiverRevue : instantané protégé', s.appel('modTests.T_Protegee', nom_hist))
    tr, _ = s.macro('modMoteur.ArchiverRevue', 'OUI')
    verifier('ArchiverRevue : 2e instantané suffixé', calc.Sheets.hasByName(nom_hist + '_2'), tr)
    tr, _ = s.macro('modMoteur.ArchiverRevue', 'NON')
    verifier('ArchiverRevue : refus sans création', not calc.Sheets.hasByName(nom_hist + '_3'))
    tr, _ = s.macro('modMoteur.ExporterRevuePDF')
    pdf = [p for p in os.listdir(SORTIE) if p.startswith('Revue_Pilotage_')]
    verifier('ExporterRevuePDF : PDF produit', len(pdf) == 1, tr)

    # ------------------------------------------------------- Nouvelle ligne
    s.appel('modTests.T_Activer', G, 'A1')
    tr, _ = s.macro('modRegistres.NouvelleLigne')
    verifier('NouvelleLigne : registre actif, ID suivant', s.txt(G, 'A10') == 'GAR-004', s.txt(G, 'A10'))
    s.appel('modTests.T_Activer', 'Sommaire', 'A1')
    tr, _ = s.macro('modRegistres.NouvelleLigne', '3')
    verifier('NouvelleLigne : choix du registre depuis le Sommaire', s.txt(F, 'A13') == 'FAC-008', s.txt(F, 'A13'))
    s.appel('modTests.T_Activer', 'Sommaire', 'A1')
    tr, _ = s.macro('modRegistres.NouvelleLigne', 'abc')
    verifier('NouvelleLigne : choix invalide refusé', 'Choix non valide' in tr, tr)

    # ------------------------------------------------------- Archivage des clôtures
    tr, _ = s.macro('modRegistres.ArchiverClotures', 'OUI')
    verifier('ArchiverClotures : garantie en mainlevée archivée', s.txt('Archive', 'A5') == 'GAR-003' and s.txt('Archive', 'W5') != '', (s.txt('Archive', 'A5'), s.txt('Archive', 'W5')))
    verifier('ArchiverClotures : facture soldée archivée', s.txt('Archive', 'Y5') == 'FAC-004', s.txt('Archive', 'Y5'))
    verifier('ArchiverClotures : date au format date dans l\'archive', not s.txt('Archive', 'M5').replace('.', '').isdigit(), s.txt('Archive', 'M5'))
    verifier('ArchiverClotures : lignes libérées, formules conservées',
             s.txt(G, 'A8') == '' and s.formule(G, 'C8').startswith('=') and s.txt(F, 'A9') == '' and s.formule(F, 'P9').startswith('='))
    tr, _ = s.macro('modRegistres.ArchiverClotures')
    verifier('ArchiverClotures : rien à archiver ensuite', 'Aucune ligne clôturée' in tr, tr)

    # ------------------------------------------------------- Exemples
    tr, _ = s.macro('modRegistres.SupprimerExemples', 'OUI')
    verifier('SupprimerExemples : lignes d\'exemple effacées', s.txt(D, 'A5') == '' and s.txt(G, 'B5') == '' and s.formule(D, 'L5').startswith('='), tr)
    tr, _ = s.macro('modRegistres.SupprimerExemples')
    verifier('SupprimerExemples : rien à effacer ensuite', 'Aucune ligne' in tr, tr)

    # ------------------------------------------------------- Gestion d'erreur
    calc.Sheets.getByName('Archive').setName('Archive_x')
    g = calc.Sheets.getByName(G)
    g.unprotect('')
    g.getCellRangeByName('O6').setString('Mainlevée accordée')
    g.protect('')
    tr, _ = s.macro('modRegistres.ArchiverClotures', 'OUI')
    verifier("Erreur : message « s'est arrêtée » et journal", "s'est arrêtée" in tr and any('ERREUR' in s.txt('Journal', 'D%d' % r) for r in range(5, 300)), tr)
    verifier('Erreur : registres reprotégés', s.appel('modTests.T_Protegee', G))
    calc.Sheets.getByName('Archive_x').setName('Archive')

    proteg = ['Sommaire', 'Tableau_de_bord', 'Plan_Action', 'DLM', 'Garanties', 'Factures', 'Plan_Paiement', 'Scénarios_BID',
              'Marchés', 'RACI', 'Éval_Entreprises', 'Fiscal', 'DAO_Standards', 'Protocole_MES', 'Note_Escalade_DG',
              'Journal', 'Archive', 'Paramètres']
    nonp = [n for n in proteg if not s.appel('modTests.T_Protegee', n)]
    verifier('Fin de recette : toutes les feuilles de saisie restent protégées', not nonp, nonp)
    tr = s.appel('modMoteur.CompterAlertes')
    verifier('CompterAlertes : format « critiques;attention »', ';' in tr, tr)
    calc.storeToURL(Path(SORTIE, 'apres_recette.xlsm').resolve().as_uri(), (pv('FilterName', 'Calc MS Excel 2007 VBA XML'),))
finally:
    s.fermer()

n_ok = sum(1 for r in RESULTATS if r[0])
print('\n%d/%d vérifications réussies' % (n_ok, len(RESULTATS)))
sys.exit(0 if n_ok == len(RESULTATS) else 1)
