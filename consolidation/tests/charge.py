# -*- coding: utf-8 -*-
"""Test de charge : registres remplis à pleine capacité (200 lignes), durée des macros."""
import datetime as dt, os, sys, time, random
sys.argv = [sys.argv[0], sys.argv[1], sys.argv[2]]
import importlib.util
spec = importlib.util.spec_from_file_location('r', os.path.join(os.path.dirname(__file__), 'recette.py'))
src = open(spec.origin, encoding='utf-8').read()
src = src[:src.index('# =============================================================================\ns = Session()')]
exec(compile(src, 'recette_base', 'exec'))
random.seed(1)
s = Session()
try:
    calc = s.doc
    for n in calc.Sheets.getElementNames():
        calc.Sheets.getByName(n).unprotect('')
    calc.enableAutomaticCalculation(False)
    calc.lockControllers()
    D, G, F, M = 'DLM', 'Garanties', 'Factures', 'Marchés'
    bail = ['BID', 'BOAD', 'FVC', 'ADFD']
    for r in range(6, 206):
        i = r - 5
        for col, v in dict(A='DLM-%03d' % i, B=bail[i % 4], C='Accord %d' % i, E=J(-1500), H=J(random.randint(-30, 500)),
                           I='XOF', J=1e9, K=random.randint(0, 9) * 1e8, V='Pièce %d' % i if i % 3 == 0 else '',
                           W='DPER' if i % 2 else 'DFP', Y=i % 4).items():
            if v != '':
                s.ecrire(D, '%s%d' % (col, r), v)
        for col, v in dict(A='M-%03d' % i, B='Lot %d' % i, C='Projet %d' % (i % 20), D=bail[i % 4], E='DLM-%03d' % i,
                           F='Entreprise %d' % (i % 30), G=5e8, J=J(-400), K=24, M=0.3, R='En cours').items():
            s.ecrire(M, '%s%d' % (col, r), v)
        for col, v in dict(A='GAR-%03d' % i, B='M-%03d' % i, F='Garantie de bonne exécution', H='G%d' % i, J=1e7,
                           L=J(-300), M=J(random.randint(-20, 300)), O='Valide').items():
            s.ecrire(G, '%s%d' % (col, r), v)
        for col, v in dict(A='FAC-%03d' % i, B='DEC-%03d' % i, C='M-%03d' % random.randint(1, 200), I=random.randint(1, 50) * 1e6,
                           J=J(-random.randint(1, 150))).items():
            s.ecrire(F, '%s%d' % (col, r), v)
    calc.unlockControllers()
    calc.enableAutomaticCalculation(True)
    calc.calculateAll()
    for nom in ('modMoteur.ActualiserTout', 'modMoteur.FiltrerAlertes', 'modMoteur.AfficherTout', 'modMoteur.ControleQualite',
                'modRapports.NoteEscaladeDG', 'modRelances.PreparerRelances'):
        tr, d = s.macro(nom, 'NON|NON|NON')
        err = "s'est arrêtée" in tr
        print('%-32s %6.1f s %s' % (nom, d, 'ERREUR ' + tr[:200] if err else 'ok'))
    rangs = sorted(int(s.num(F, 'AF%d' % r)) for r in range(6, 206) if s.txt(F, 'AF%d' % r))
    print('Rangs du plan de paiement uniques et continus :', rangs == list(range(1, len(rangs) + 1)), len(rangs))
    print('Lignes d alerte :', sum(1 for r in range(5, 3000) if s.txt('Alertes', 'A%d' % r) in ('CRITIQUE', 'ATTENTION')))
finally:
    s.fermer()
