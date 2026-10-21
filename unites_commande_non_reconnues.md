# Unités non reconnues par la base de données

Liste des valeurs distinctes trouvées dans les colonnes **Unité de commande**, **Unité de
référence** et **Conditionnement** (mot de récipient détecté) des 7 fichiers sources
(`pret_pour_audit/`), qui ne correspondent à aucune des unités déjà acceptées par la base de
données (comparées aux deux colonnes de la liste fournie : code ET libellé complet).

24 valeurs distinctes ne correspondent à aucune unité reconnue :

| Unité (telle qu'écrite) | Occurrences | Trouvé dans | Exemple de fichier |
|---|---|---|---|
| COLIS | 1098 | Conditionnement, Unité de commande, Unité de référence | 2026 09 22 Base article groupement 1.xlsx |
| CAGEOT | 229 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement 1.xlsx |
| CAISSE | 56 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement 2.xlsx |
| PU | 44 | Unité de commande | MDC_IMPORT_MARCHE - Lots frais 2026 exel.xlsx |
| FILET | 42 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement 1.xlsx |
| ETUI | 21 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement commun.xlsx |
| PLAQUE | 17 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement 1.xlsx |
| PACK | 16 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement 1.xlsx |
| BRIQUE | 15 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement commun.xlsx |
| COFFRET | 14 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement commun.xlsx |
| CLAYETTE | 10 | Conditionnement | 2026 09 22 Base article groupement 3.xlsx |
| BQT | 10 | Unité de commande | MDC_IMPORT_MARCHE - Lots frais 2026 exel.xlsx |
| GOURDE | 9 | Conditionnement | 2026 09 22 Base article groupement 1.xlsx |
| 2 PIECES | 6 | Unité de commande | 2026 09 22 Base article groupement 1.xlsx |
| CARTONS | 6 | Unité de commande | 2026 09 22 Base article groupement 1.xlsx |
| BOCAL | 6 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement commun.xlsx |
| OUTRE | 5 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement commun.xlsx |
| FLACON | 4 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement commun.xlsx |
| BT. | 4 | Unité de commande | MDC_IMPORT_MARCHE - Lots frais 2026 exel.xlsx |
| PLAQUETTE | 3 | Conditionnement | 2026 09 22 Base article groupement commun.xlsx |
| BOUTEILE | 2 | Conditionnement | 2026 09 22 Base article groupement 1.xlsx |
| TUBE | 2 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement commun.xlsx |
| BOMBE | 2 | Conditionnement, Unité de commande | 2026 09 22 Base article groupement commun.xlsx |
| BAC | 1 | Conditionnement | 2026 09 22 Base article groupement commun.xlsx |

## Unités possibles (colonne de droite de la liste fournie)

Liste à utiliser pour remplir la colonne "Unité de remplacement" :

PART, B1/2, B2/1, BOITE 3/1, BOITE 4/1, BOITE 4/4, BOITE 5/1, BARQUETTE, BIDON 10 L, BOITE,
BOTTE, BOUTEILLE, CARTON, CADRE, Gr, KILOGRAMME, LITRE, LOT, PAQUET, PIECE, PLATEAU, Poche,
POT, ROUL, SAC, SACHET, SEAU, Unité

## Notes

- Comparaison faite sans accents et insensible à la casse, en acceptant soit le code court
  (colonne de gauche de la liste fournie), soit le libellé complet (colonne de droite).
- Pour le Conditionnement, seul le MOT DE RÉCIPIENT détecté dans le texte libre est comparé
  (ex: "CARTON DE 12 PLAQUES" → "PLAQUE"), pas la phrase entière.
- "KG", "L", "PC" et "BIDON" trouvés dans les fichiers correspondent bien aux codes courts déjà
  acceptés — ils ne sont donc plus listés ici comme inconnus.
- "2 PIECES" et "CARTONS" (Unité de commande) sont des cas particuliers : texte libre recopié
  tel quel, pas une unité au sens strict — à vérifier au cas par cas.
- "BOUTEILE" (une seule lettre L) est une faute de frappe fréquente pour "BOUTEILLE", déjà
  reconnue par ailleurs.
- "BT." pourrait être une abréviation de "Bouteille" déjà couverte par "BOUT"/"BOUTEILLE", à
  confirmer.
