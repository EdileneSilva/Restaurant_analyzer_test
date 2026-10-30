"""
Tests de non-régression (US-07 — « Formater un fichier formatable ») pour le prétraitement, sur
les 21 fichiers de `a_formater/` déjà répertoriés dans `pretraitement_a_formater.py` (BPU DQE,
Passion Froid, Episaveurs, BPUR FRESCA, MP ARDECHE 24 — tous les fichiers "formatable" connus à
ce jour ; aucun fichier REGAL/GRADA d'exemple n'étant présent dans ce dépôt, ces deux familles
ne sont pas couvertes ici). Passe par le point d'entrée unique, `pretraitement_auto.
transforme_automatiquement()` — pas directement par `pretraitement_a_formater.py` — pour
couvrir aussi le branchement fait pour US-06 (detecte_format/trouve_entree_connue).

Vérifie, pour chaque fichier, les 3 critères d'acceptation d'US-07 :
1. Le fichier d'origine n'est pas modifié (hash MD5 identique avant/après).
2. Le fichier produit passe la détection avec le statut `deja_au_format`.
3. Le nombre de lignes produites correspond au total de référence (mesuré le 2026-10-30) — si
   ce nombre change sans raison identifiée après une modification du code, c'est une régression
   à investiguer, pas à corriger en silence en changeant juste le chiffre ici.

En bonus (pas un critère d'acceptation d'US-07 à proprement parler, mais directement lié à
US-06) : vérifie aussi que le fichier source brut est bien détecté `formatable` AVANT
transformation.

Contrairement à `test_non_regression_audit_general.py`/`_commentaires.py`, ce fichier n'importe
ni `audit_general` ni `audit_general_commentaires` : aucun risque du bug de constantes
partagées documenté dans ces deux fichiers (`catalogue_erreurs.md` n°41) — peut être lancé seul
ou dans la même invocation que n'importe lequel des deux (ou les deux à la fois, lui excepté).
"""

import hashlib
import os

import pytest

import pretraitement_a_formater as paf
import pretraitement_auto as pa

DOSSIER_REFERENCE = paf.DOSSIER_ENTREE

# Totaux de référence mesurés le 2026-10-30, via pretraitement_auto.transforme_automatiquement()
# (le point d'entrée unique), pas via pretraitement_a_formater.transforme_dossier_a_formater()
# directement — les deux appellent les mêmes fonctions de transformation, mais c'est le premier
# qui est réellement utilisé par app.py.
TOTAUX = {
    "05_BPU DQE Taux de remise_2023_0498_ Lot 1 Viandes et volailles.xlsx": 48,
    "06_BPU DQE Taux de remise_2023_0498_Lot 2 Poissons.xlsx": 36,
    "6-BPU DQE Taux de remise_2023_0498_Lot 3 Legumes V2.xlsx": 54,
    "08_BPU DQE Taux de remise_2023_0498_Lot 4 PC et Entrees.xlsx": 36,
    "09_BPU DQE Taux de remise_2023_0498_Lot 5 Glaces et desserts.xlsx": 38,
    "13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx": 359,
    "Passion froid - Tarif général Mise en place V2 1.xlsx": 412,
    "Episaveurs - Tarifs Catalogue complémentaire Septembre 2024 V2.xls": 3621,
    "1. BPUR FRESCA LOT 1 SURGELES CONV SIQO.xlsx": 438,
    "DADS D&#x27;Ard&#xE8;che et de Saison MP ARDECHE 24 - Lot 06  Crèmerie et produits "
    "frais Du vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx": 49,
    "Episaveurs MP ARDECHE 24 - Lot 1 Epicerie, boisson Du mardi 7 mai 2024 Au mardi 11 "
    "avril 2028.xlsx": 238,
    "LDC MP ARDECHE 24 - Lot 7 Volaille Du mercredi 1 janvier 2025 Au jeudi 15 avril 2027.xlsx": 26,
    "Passion Froid MP ARDECHE 24 - Lot 5 Autres produits de crèmerie Du mercredi 1 janvier "
    "2025 Au mercredi 14 avril 2027.xlsx": 65,
    "Agricourt MP ARDECHE 24 - Lot 2 Epicerie BIO Du mardi 15 avril 2025 Au samedi 10 avril "
    "2032 (1).xlsx": 100,
    "DADS  et de Saison MP ARDECHE 24 - Lot 3 Jus de fruits en bouteilles, produits à base de "
    "chataîgne, compotes et vin Du vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx": 27,
    "DADS D&#x27;Ard&#xE8;che et de Saison MP ARDECHE 24 - Lot 8 Boucherie saucisserie Du "
    "vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx": 35,
    "DADS D&#x27;Ard&#xE8;che et de Saison MP ARDECHE 24 - Lot 11 Poisson frais d'eau douce "
    "Du vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx": 2,
    "Passion Froid MP ARDECHE 24 - Lot 4 Surgelés Du mercredi 5 février 2025 Au jeudi 22 "
    "juillet 2027.xlsx": 142,
    "Passion Froid MP ARDECHE 24 - Lot 9 Charcuterie Traiteur Du lundi 19 août 2024 Au jeudi "
    "13 avril 2028.xlsx": 34,
    "SYSCO FRANCE MP ARDECHE 24 - Lot 10 Produits frais de la mer Du lundi 27 janvier 2025 Au "
    "mardi 13 avril 2027.xlsx": 5,
    "Vivalya MP ARDECHE 24 - Lot 12 Fruits et légumes frais et 4ème et 5ème gamme dont BIO Du "
    "vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx": 136,
}


def _hash_fichier(chemin):
    with open(chemin, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


@pytest.mark.parametrize("nom_fichier", sorted(TOTAUX))
def test_formate_sans_toucher_a_l_original(nom_fichier, tmp_path):
    nb_lignes_attendu = TOTAUX[nom_fichier]
    chemin_source = os.path.join(DOSSIER_REFERENCE, nom_fichier)
    assert os.path.exists(chemin_source), f"Fichier de référence introuvable : {chemin_source}"

    # detecte_statut() (comme detecte_format()) appelle openpyxl.load_workbook() directement,
    # qui ne sait pas lire un .xls brut — seul transforme_automatiquement() sait le faire, en le
    # convertissant d'abord (voir convertit_xls_en_xlsx). Donc ce contrôle-bonus ne s'applique
    # qu'aux fichiers déjà .xlsx ; le seul .xls du jeu de référence (Episaveurs) est quand même
    # couvert par le test lui-même, qui passe bien par transforme_automatiquement() plus bas.
    if not nom_fichier.lower().endswith(".xls"):
        assert pa.detecte_statut(chemin_source, nom_fichier=nom_fichier) == "formatable", (
            f"{nom_fichier} : attendu 'formatable' avant transformation (US-06)")

    hash_avant = _hash_fichier(chemin_source)

    chemin_sortie = str(tmp_path / "sortie.xlsx")
    formato, nb_lignes, causes = pa.transforme_automatiquement(
        chemin_source, chemin_sortie, nom_fichier=nom_fichier
    )

    assert _hash_fichier(chemin_source) == hash_avant, (
        f"{nom_fichier} : le fichier d'origine a été modifié par la transformation")

    assert causes == [], f"{nom_fichier} : cause(s) inattendue(s) : {causes}"
    assert formato == "a_formater", f"{nom_fichier} : format détecté {formato!r} (attendu 'a_formater')"
    assert nb_lignes == nb_lignes_attendu, (
        f"{nom_fichier} : {nb_lignes} lignes produites (attendu {nb_lignes_attendu})")

    assert pa.detecte_statut(chemin_sortie) == "deja_au_format", (
        f"{nom_fichier} : le fichier produit n'est pas reconnu 'deja_au_format'")


if __name__ == "__main__":
    import sys
    sys.exit(pytest.main([__file__, "-v"]))
