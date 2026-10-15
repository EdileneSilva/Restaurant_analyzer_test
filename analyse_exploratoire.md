# Analyse exploratoire (TT-03)

Document de travail pour TT-03 — un bloc par fichier, rempli au fur et à mesure.
Base : fichiers listés dans `inventaire_echantillons.csv` (`a_inspecter/`, `a_formater/`).

## a_inspecter/

Les 6 fichiers "Base IANOR AGAPE 16 ..." ont tous exactement les deux mêmes onglets :
- `Feuille 1`
- `lot 3 Produits élaborés` — données réelles, mais hors périmètre (pas utilisées ici)

| Fichier | Onglets |
|---|---|
| Base IANOR AGAPE 16 Charcuterie (1).xlsx | Feuille 1, lot 3 Produits élaborés |
| Base IANOR AGAPE 16 Epicerie.xlsx | Feuille 1, lot 3 Produits élaborés |
| Base IANOR AGAPE 16 Produits laitiers (2).xlsx | Feuille 1, lot 3 Produits élaborés |
| Base IANOR AGAPE 16 Surgelé.xlsx | Feuille 1, lot 3 Produits élaborés |
| Base IANOR AGAPE 16 Viandes fraiches (1).xlsx | Feuille 1, lot 3 Produits élaborés |
| Base IANOR AGAPE 16 Volailles (1) (1).xlsx | Feuille 1, lot 3 Produits élaborés |

## a_formater/

| Fichier | Onglets |
|---|---|
| 05_BPU DQE Taux de remise_2023_0498_ Lot 1 Viandes et volailles.xlsx | Page de Garde, BPU Lot 1, DQE Lot 1, Taux De remise |
| 06_BPU DQE Taux de remise_2023_0498_Lot 2 Poissons.xlsx | Page de Garde, BPU Lot 2, DQE Lot 2, Taux de remise |
| 08_BPU DQE Taux de remise_2023_0498_Lot 4 PC et Entrees.xlsx | Page de Garde, BPU Lot 4, DQE Lot 4, Taux de remise |
| 09_BPU DQE Taux de remise_2023_0498_Lot 5 Glaces et desserts.xlsx | Page de Garde, BPU Lot 5, DQE Lot 5, Taux de remise |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | Page de garde, BPU, DQE, Taux de remise |
| 6-BPU DQE Taux de remise_2023_0498_Lot 3 Legumes V2.xlsx | Page de Garde, BPU Lot 3, DQE Lot 3, Taux de remise |
| Episaveurs - Tarifs Catalogue complémentaire Septembre 2024 V2.xls | Lignes tarifaires |
| Passion froid - Tarif général Mise en place V2 1.xlsx | SANS DOUBLONS, GLOBAL |

## Étape 3 — ligne d'en-tête, nombre de lignes/colonnes, colonnes sans nom

| Fichier | Onglet | Ligne d'en-tête | Lignes x Colonnes | Colonnes sans nom |
|---|---|---|---|---|
| Base IANOR AGAPE 16 Charcuterie (1).xlsx | Feuille 1 | 1 | 52 x 59 | 1 (`Unnamed: 57`, entre "Code RNM/ITAVI" et "ID_AURA") |
| Base IANOR AGAPE 16 Charcuterie (1).xlsx | lot 3 Produits élaborés | 0 | 35 x 13 | 12 |
| Base IANOR AGAPE 16 Epicerie.xlsx | Feuille 1 | 1 | 236 x 59 | 1 |
| Base IANOR AGAPE 16 Epicerie.xlsx | lot 3 Produits élaborés | 0 | 35 x 13 | 12 |
| Base IANOR AGAPE 16 Produits laitiers (2).xlsx | Feuille 1 | 1 | 136 x 59 | 1 |
| Base IANOR AGAPE 16 Produits laitiers (2).xlsx | lot 3 Produits élaborés | 0 | 35 x 13 | 12 |
| Base IANOR AGAPE 16 Surgelé.xlsx | Feuille 1 | 1 | 153 x 59 | 1 |
| Base IANOR AGAPE 16 Surgelé.xlsx | lot 3 Produits élaborés | 0 | 35 x 13 | 12 |
| Base IANOR AGAPE 16 Viandes fraiches (1).xlsx | Feuille 1 | 1 | 36 x 59 | 1 |
| Base IANOR AGAPE 16 Viandes fraiches (1).xlsx | lot 3 Produits élaborés | 0 | 35 x 13 | 12 |
| Base IANOR AGAPE 16 Volailles (1) (1).xlsx | Feuille 1 | 1 | 74 x 59 | 1 |
| Base IANOR AGAPE 16 Volailles (1) (1).xlsx | lot 3 Produits élaborés | 0 | 35 x 13 | 12 |
| 05_BPU...Lot 1 Viandes et volailles.xlsx | Page de Garde | 0 | 48 x 6 | 6 (texte libre, pas tabulaire) |
| 05_BPU...Lot 1 Viandes et volailles.xlsx | BPU Lot 1 | 2 | 53 x 16 | 0 |
| 05_BPU...Lot 1 Viandes et volailles.xlsx | DQE Lot 1 | 2 | 52 x 15 | 0 |
| 05_BPU...Lot 1 Viandes et volailles.xlsx | Taux De remise | 4 | 7 x 3 | 0 |
| 06_BPU...Lot 2 Poissons.xlsx | Page de Garde | 0 | 50 x 6 | 6 |
| 06_BPU...Lot 2 Poissons.xlsx | BPU Lot 2 | 2 | 41 x 17 | 0 |
| 06_BPU...Lot 2 Poissons.xlsx | DQE Lot 2 | 2 | 40 x 15 | 0 |
| 06_BPU...Lot 2 Poissons.xlsx | Taux de remise | 4 | 6 x 3 | 0 |
| 08_BPU...Lot 4 PC et Entrees.xlsx | Page de Garde | 0 | 48 x 6 | 6 |
| 08_BPU...Lot 4 PC et Entrees.xlsx | BPU Lot 4 | 2 | 41 x 15 | 0 |
| 08_BPU...Lot 4 PC et Entrees.xlsx | DQE Lot 4 | 2 | 40 x 14 | 0 |
| 08_BPU...Lot 4 PC et Entrees.xlsx | Taux de remise | 4 | 7 x 3 | 0 |
| 09_BPU...Lot 5 Glaces et desserts.xlsx | Page de Garde | 0 | 34 x 5 | 5 |
| 09_BPU...Lot 5 Glaces et desserts.xlsx | BPU Lot 5 | 2 | 43 x 14 | 0 |
| 09_BPU...Lot 5 Glaces et desserts.xlsx | DQE Lot 5 | 2 | 43 x 13 | 0 |
| 09_BPU...Lot 5 Glaces et desserts.xlsx | Taux de remise | 4 | 7 x 3 | 0 |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | Page de garde | 0 | 36 x 5 | 5 |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | BPU | 3 | 365 x 17 | 0 |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | DQE | 3 | 361 x 17 | 0 |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | Taux de remise | 4 | 19 x 3 | 0 |
| 6-BPU...Lot 3 Legumes V2.xlsx | Page de Garde | 0 | 48 x 6 | 6 |
| 6-BPU...Lot 3 Legumes V2.xlsx | BPU Lot 3 | 2 | 59 x 15 | 0 |
| 6-BPU...Lot 3 Legumes V2.xlsx | DQE Lot 3 | 2 | 56 x 14 | 0 |
| 6-BPU...Lot 3 Legumes V2.xlsx | Taux de remise | 4 | 2 x 3 | 0 |
| Episaveurs - Tarifs Catalogue complémentaire Septembre 2024 V2.xls | Lignes tarifaires | 0 | 3621 x 11 | 0 |
| Passion froid - Tarif général Mise en place V2 1.xlsx | SANS DOUBLONS | 0 | 412 x 8 | 0 |
| Passion froid - Tarif général Mise en place V2 1.xlsx | GLOBAL | 0 | 419 x 8 | 0 |

**Constats clés :**
- `lot 3 Produits élaborés` : taille identique (35x13) dans les 6 fichiers IANOR → probablement le même onglet copié partout, pas spécifique à la catégorie (confirmé réel mais hors périmètre).
- Fichiers BPU DQE : structure en 4 onglets, `Page de Garde` est du texte libre (pas tabulaire), `BPU`/`DQE`/`Taux de remise` ont chacun leur propre ligne d'en-tête (2, 3 ou 4 selon le fichier — pas de position fixe).
- `Episaveurs` et `Passion froid` : plus simples, en-tête dès la ligne 0.

## Étape 4 — cellules fusionnées

| Fichier | Onglet | Plages fusionnées |
|---|---|---|
| Base IANOR AGAPE 16 (les 6 fichiers) | Feuille 1 | 0 |
| Base IANOR AGAPE 16 (les 6 fichiers) | lot 3 Produits élaborés | 0 |
| 05/06/08/09/6-BPU... | Page de Garde | 20-22 |
| 05/06/08/09/6-BPU... | BPU Lot N | 3 |
| 05/06/08/09/6-BPU... | DQE Lot N | 0 |
| 05/06/08/09/6-BPU... | Taux de remise | 2 |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | Page de garde | 24 |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | BPU | 7 |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | DQE | 1 |
| 13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx | Taux de remise | 1 |
| Episaveurs (converti) | Lignes tarifaires | 0 |
| Passion froid | SANS DOUBLONS, GLOBAL | 0 |

**Constats clés :**
- Format canonique (`a_inspecter/`) : **aucune** cellule fusionnée, en-tête simple sur une seule ligne.
- `Page de Garde`/`Page de garde` : confirmé non-tabulaire (beaucoup de fusions, texte libre) — à exclure de l'analyse colonnes/données.
- `BPU`/`DQE`/`Taux de remise` : quelques fusions (1 à 7), probablement des titres de groupe de colonnes au-dessus de la vraie ligne d'en-tête — à vérifier en détail si on construit un script de transformation pour ce format.

## Étape 5 — taux de cellules vides par colonne (format canonique, `a_inspecter/`)

Calculé sur les 6 fichiers IANOR, onglet `Feuille 1` (header ligne 1).

**Toujours vides dans les 6 fichiers (41 colonnes)** — le bloc Fournisseur 2 entier, le bloc
Fournisseur 3 entier, les 14 allergènes, `TAUX TVA`, `Code RNM / ITAVI`, `ID_AURA`,
`Unnamed: 57` : colonnes présentes dans le modèle mais jamais utilisées en pratique par ce
fournisseur de données (à documenter comme "optionnelles, quasi jamais remplies" dans
TT-04, pas comme des erreurs).

**Anomalie trouvée** : `Base IANOR AGAPE 16 Epicerie.xlsx` a une ligne entièrement vide en
fin de tableau (ligne 236, après les 235 lignes de données réelles) — y compris les colonnes
`Marché`/`Lot` qui sont sinon toujours remplies. Artefact Excel classique (ligne incluse dans
la zone utilisée de la feuille sans donnée réelle) → à ajouter au catalogue d'erreurs (TT-05).

**Toujours remplies (hors ligne fantôme ci-dessus)** : `Marché`, `Lot`,
`Désignation de l'article`, `Fournisseur 1`, `Désignation fournisseur`, `Conditionnement`,
`Unité Référence (BPU)`, `PRIX UNITE REFERENCE (BPU)`, `UNITE COMMANDE`,
`PRIX UNITE COMMANDE` — candidats directs aux colonnes "obligatoires" du format cible (TT-04).

**Colonnes partiellement remplies** (varient selon le fichier, ni 0% ni 100%) : `Label`,
`Local`, `Référence produit`, `Origine`, `Indice I0`, `Bio`, `Poids de la PC` — probablement
optionnelles, remplies seulement quand l'information est connue.

## Synthèse pour la suite

Ce document couvre les étapes 2 à 5 de TT-03 pour tous les fichiers de
`inventaire_echantillons.csv`. Sert de base directe à :
- **TT-04** (format cible) : colonnes obligatoires / optionnelles identifiées ci-dessus pour
  le format canonique ; structure à 4 onglets (Page de Garde / BPU / DQE / Taux de remise)
  identifiée pour le format BPU DQE, encore à documenter en détail si un script de
  transformation est construit pour ce format.
- **TT-05** (catalogue d'erreurs) : ligne fantôme vide (Epicerie), en-tête sans nom
  (`Unnamed: 57`), cellules fusionnées dans les pages de garde / certaines zones BPU-DQE.
