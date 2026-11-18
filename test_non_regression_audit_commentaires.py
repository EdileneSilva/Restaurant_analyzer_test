"""
Tests de non-régression (TT-09) pour `audit_general_commentaires.py`, sur les 7 fichiers de
`pret_pour_audit/`. Vérifie le total de cellules marquées par fichier (documenté dans
`catalogue_erreurs.md`, section "Jeu de référence") et l'absence de commentaires en double sur
une même cellule (bug trouvé et corrigé plusieurs fois, voir `catalogue_erreurs.md` n°38).

⚠️ IMPORTANT — ne JAMAIS importer `audit_general` dans le même processus que ce fichier. Voir
l'explication complète en tête de `test_non_regression_audit_general.py` : les deux scripts
partagent et écrasent à l'import les mêmes constantes globales du moteur commun
(`first_scripts/first_attempt_v2_prototype.py`), le second import écrasant silencieusement
celles du premier. Toujours lancer séparément :

    pytest test_non_regression_audit_general.py -v
    pytest test_non_regression_audit_commentaires.py -v
"""

import os

import openpyxl
import pytest

import audit_general_commentaires as audit

DOSSIER_REFERENCE = audit.DOSSIER_ENTREE

TOTAUX_CELLULES = {
    "2026 09 22 Base article groupement 1.xlsx": 756,
    "2026 09 22 Base article groupement 2.xlsx": 827,
    "2026 09 22 Base article groupement 3.xlsx": 987,
    "2026 09 22 Base article groupement 4.xlsx": 955,
    "2026 09 22 Base article groupement 5.xlsx": 864,
    "2026 09 22 Base article groupement commun.xlsx": 956,
    # 744 -> 727 le 2026-11-13, même raison que test_non_regression_audit_general.py (TOTAUX) —
    # voir catalogue_erreurs.md, règle n°61.
    "MDC_IMPORT_MARCHE - Lots frais 2026 exel.xlsx": 727,
}


def _total_cellules_marquees(stats):
    return (stats["marques"] + stats["lignes_dupliquees"]
            + stats["copies_possibles"] + stats["champs_obligatoires_optionnels"])


@pytest.fixture(scope="module")
def indice_prix():
    return audit.construit_indice_prix(DOSSIER_REFERENCE)


def _audite(nom_fichier, indice_prix):
    chemin = os.path.join(DOSSIER_REFERENCE, nom_fichier)
    wb_valeurs = openpyxl.load_workbook(chemin, data_only=True)
    wb = openpyxl.load_workbook(chemin, data_only=False)
    resultats = audit.audite_classeur(wb_valeurs, wb, indice_prix)
    return wb, resultats


@pytest.mark.parametrize("nom_fichier", sorted(TOTAUX_CELLULES))
def test_total_cellules_stable(nom_fichier, indice_prix):
    _, resultats = _audite(nom_fichier, indice_prix)
    # Même formule que `main()` ("Total de cellules marquées") et que
    # `test_non_regression_audit_general.py` — PAS un simple compte de cellules avec
    # commentaire (une cellule peut recevoir 2 signalements indépendants empilés dans le même
    # commentaire, ou être changée sans commentaire pour une correction mécanique silencieuse ;
    # les deux comptages divergent légèrement, vérifié le 2026-10-26).
    cellules = sum(_total_cellules_marquees(r["stats"]) for r in resultats)
    attendu = TOTAUX_CELLULES[nom_fichier]
    assert cellules == attendu, (
        f"{nom_fichier} : {cellules} cellules marquées (attendu {attendu})")


@pytest.mark.parametrize("nom_fichier", sorted(TOTAUX_CELLULES))
def test_sans_doublon_de_commentaire(nom_fichier, indice_prix):
    wb, _ = _audite(nom_fichier, indice_prix)
    doublons = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                texte = cell.comment.text if cell.comment else ""
                if texte.count("[Fournisseur") > 1:
                    doublons.append(cell.coordinate)
    assert not doublons, f"{nom_fichier} : commentaires en double sur {doublons}"


if __name__ == "__main__":
    import sys
    sys.exit(pytest.main([__file__, "-v"]))
