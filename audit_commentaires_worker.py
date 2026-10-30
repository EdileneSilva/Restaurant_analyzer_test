"""
Worker exécuté dans un PROCESSUS SÉPARÉ par `audit_pipeline.calcule_resultats_commentaires()` —
jamais importé directement dans le même processus que `audit_general` (voir
`catalogue_erreurs.md` n°41 : les deux scripts d'audit partagent et s'écrasent mutuellement les
mêmes constantes du moteur commun au chargement). N'importe QUE `audit_general_commentaires`.

Usage : python audit_commentaires_worker.py <dossier_entree> <dossier_sortie>

Audite chaque "<i>.xlsx" de <dossier_entree> (déjà au format canonique) avec
`audit_general_commentaires`, et écrit dans <dossier_sortie> :
- "<i>.xlsx" : le classeur audité (couleurs + commentaires + corrections mécaniques) ;
- "<i>.xlsx.json" : [{"nom", "stats", "categories"}, ...] (une entrée par feuille reconnue) ;
- "_indice.json" : {"produit": N, "categorie": N} (taille de l'indice de prix comparables).
"""

import json
import os
import sys

import openpyxl

import audit_general_commentaires as audit


def main():
    dossier_entree, dossier_sortie = sys.argv[1], sys.argv[2]
    noms = sorted(f for f in os.listdir(dossier_entree) if f.endswith(".xlsx"))

    classeurs_valeurs, classeurs = [], []
    for nom in noms:
        chemin = os.path.join(dossier_entree, nom)
        classeurs_valeurs.append(openpyxl.load_workbook(chemin, data_only=True))
        classeurs.append(openpyxl.load_workbook(chemin, data_only=False))

    indice_prix = audit.construit_indice_prix_depuis_classeurs(classeurs_valeurs)
    with open(os.path.join(dossier_sortie, "_indice.json"), "w", encoding="utf-8") as f:
        json.dump(
            {"produit": len(indice_prix["produit"]), "categorie": len(indice_prix["categorie"])}, f
        )

    for nom, wb_valeurs, wb in zip(noms, classeurs_valeurs, classeurs):
        resultats = audit.audite_classeur(wb_valeurs, wb, indice_prix)
        wb.save(os.path.join(dossier_sortie, nom))
        resume = [{"nom": r["nom"], "stats": r["stats"], "categories": r["categories"]}
                  for r in resultats]
        with open(os.path.join(dossier_sortie, nom + ".json"), "w", encoding="utf-8") as f:
            json.dump(resume, f, ensure_ascii=False)


if __name__ == "__main__":
    main()
