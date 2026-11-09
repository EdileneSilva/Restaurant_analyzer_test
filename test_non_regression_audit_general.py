"""
Tests de non-régression (TT-09) pour `audit_general.py`, sur les 7 fichiers de
`pret_pour_audit/`. Vérifie que les totaux ("lignes avec incohérence", "cellules marquées")
correspondent à ceux documentés dans `catalogue_erreurs.md` (section "Jeu de référence").

⚠️ IMPORTANT — ne JAMAIS importer `audit_general_commentaires` dans le même processus que ce
fichier (ni dans un `conftest.py` partagé, ni en lançant les deux fichiers de test dans la
même invocation sans isolation). Les deux scripts partagent le même moteur
`first_scripts/first_attempt_v2_prototype.py` (`motor`) et chacun écrase à l'import des
constantes globales dessus (`motor.CONTENANT_ABREGE`, `motor.TOLERANCE_RATIO`...) avec SES
PROPRES valeurs — le second import écrase silencieusement celles du premier. Vérifié : importer
les deux dans le même interpréteur change `motor.CONTENANT_ABREGE['colis']` de `'CARTON'`
(valeur d'`audit_general.py`) à `'COLIS'` (valeur d'`audit_general_commentaires.py`), ce qui
fausse silencieusement les résultats de `audit_general.py` sans aucune erreur. C'est pour cette
raison que ce fichier et `test_non_regression_audit_commentaires.py` sont deux fichiers séparés
plutôt qu'un seul paramétré par module : lancer `pytest` sur les deux en même temps démarre un
seul processus qui importerait les deux modules — utiliser `pytest -p no:randomly` n'y change
rien, il faut deux invocations séparées :

    pytest test_non_regression_audit_general.py -v
    pytest test_non_regression_audit_commentaires.py -v

(les lancer dans la même commande, ex. `pytest test_non_regression_audit_general.py
test_non_regression_audit_commentaires.py`, réintroduit le bug ci-dessus — à ne jamais faire).
"""

import os

import openpyxl
import pytest

import audit_general

DOSSIER_REFERENCE = audit_general.DOSSIER_ENTREE

# Totaux de référence mesurés le 2026-10-26 (catalogue_erreurs.md, section "Jeu de référence").
# Si un de ces nombres change sans raison identifiée après une modification du code, c'est une
# régression à investiguer — pas à corriger en silence en changeant juste le chiffre ici.
TOTAUX = {
    "2026 09 22 Base article groupement 1.xlsx": (173, 799),
    "2026 09 22 Base article groupement 2.xlsx": (293, 878),
    "2026 09 22 Base article groupement 3.xlsx": (224, 1024),
    "2026 09 22 Base article groupement 4.xlsx": (191, 1003),
    "2026 09 22 Base article groupement 5.xlsx": (219, 889),
    "2026 09 22 Base article groupement commun.xlsx": (748, 1298),
    "MDC_IMPORT_MARCHE - Lots frais 2026 exel.xlsx": (193, 721),
}


def _total_cellules_marquees(stats):
    return (stats["marques"] + stats["lignes_dupliquees"]
            + stats["copies_possibles"] + stats["champs_obligatoires_optionnels"])


@pytest.fixture(scope="module")
def indice_prix():
    return audit_general.construit_indice_prix(DOSSIER_REFERENCE)


@pytest.mark.parametrize("nom_fichier", sorted(TOTAUX))
def test_totaux_stables(nom_fichier, indice_prix):
    lignes_attendues, cellules_attendues = TOTAUX[nom_fichier]

    chemin = os.path.join(DOSSIER_REFERENCE, nom_fichier)
    wb_valeurs = openpyxl.load_workbook(chemin, data_only=True)
    wb = openpyxl.load_workbook(chemin, data_only=False)
    resultats = audit_general.audite_classeur(wb_valeurs, wb, indice_prix)

    lignes = sum(r["stats"]["lignes_avec_incoherence"] for r in resultats)
    cellules = sum(_total_cellules_marquees(r["stats"]) for r in resultats)

    assert lignes == lignes_attendues, (
        f"{nom_fichier} : {lignes} lignes avec incohérence (attendu {lignes_attendues})")
    assert cellules == cellules_attendues, (
        f"{nom_fichier} : {cellules} cellules marquées (attendu {cellules_attendues})")


if __name__ == "__main__":
    import sys
    sys.exit(pytest.main([__file__, "-v"]))
