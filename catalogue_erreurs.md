# Catalogue d'erreurs (TT-05)

Un bloc par type d'erreur : description, exemple réel (extrait des fichiers canoniques
confirmés dans `pret_pour_audit/`), gravité pressentie, et si une règle existe déjà dans
`audit_geral.py` ou non. Base pour TT-06 (implémentation), TT-08 (seuils) et TT-09 (tests).

## Erreurs déjà couvertes par le code (`audit_geral.py`)

### 1. Lignes strictement dupliquées (`doublons`)
- **Exemple réel** : `2026 09 22 Base article groupement 1.xlsx`, ligne 71, "MELANGE COUSCOUS" —
  identique à la ligne 68. Toute la ligne est surlignée en rouge.
- **Gravité pressentie** : avertissement — à trancher par une personne avant import (laquelle
  garder), le script ne supprime rien.
- **Couverte par le code** : oui.

### 2. Références en double, même prix (`references_doublon`)
- **Exemple réel** : même fichier, ligne 68, "MELANGE COUSCOUS" — même Référence produit
  (000013) et même prix de référence (4.33) que la ligne 71.
- **Gravité pressentie** : avertissement (peut être une simple variante de désignation).
- **Couverte par le code** : oui.

### 3. Unité de référence vide (`unite_ref_vide`)
- **Exemple réel** : ligne 50, "ECHALOTTE EMINCEE" — unité de référence vide, à compléter
  manuellement.
- **Gravité pressentie** : bloquante — sans cette valeur, impossible de comparer prix/unité.
- **Couverte par le code** : oui.

### 4. Unité de référence incohérente (`unite_ref`)
- **Exemple réel** : ligne 369, "ANETH" — unité changée de 'PC' vers 'KG' car le
  conditionnement ("SACHET DE 250 GR") décrit un récipient, pas une unité de mesure directe.
- **Gravité pressentie** : avertissement (changement déjà expliqué et tracé, à confirmer).
- **Couverte par le code** : oui.

### 5. Prix de référence incohérent (`prix_ref`)
- **Exemple réel** : ligne 212, "TRAVERS DE PORC TRANCHE FERMIER" — prix recalculé
  (22.15 ÷ 0.5 kg = 44.3) ET signalé comme très éloigné de la médiane des lignes comparables
  (×5.0 d'écart).
- **Gravité pressentie** : bloquante quand l'écart est fort (facteur ×5 ici) — probable
  erreur de saisie (zéro en trop/en moins).
- **Couverte par le code** : oui.

### 6. Unité de commande incohérente (`unite_cmd`)
- **Exemple réel** : ligne 30, "NAVET ROND VIOLET CAT.1" — unité changée en 'CAGEOT' d'après
  le conditionnement, car 'CARTON' (valeur source) ne correspond pas au contenant décrit.
- **Gravité pressentie** : avertissement.
- **Couverte par le code** : oui.

### 7. Prix de commande incohérent / inexpliqué (`prix_cmd`)
- **Exemple réel** : ligne 104, "FARINE DE MAÏS" — prix de référence cohérent avec la
  catégorie, mais prix de commande (5.9) ne correspond pas à la quantité du conditionnement
  (attendu ~29.5) — le script ne recalcule PAS ce prix, juste signale.
- **Gravité pressentie** : bloquante — prix probablement faux (zéro oublié), à corriger avant
  import.
- **Couverte par le code** : oui.

### 8. Poids à l'unité de référence rempli automatiquement (`poids_rempli`)
- **Exemple réel** : ligne 107, "HUILE D'OLIVE VIERGE EXTRA" — poids déduit du
  conditionnement (0.75 L par pièce), confirmé par le ratio des prix (×12 pièces).
- **Gravité pressentie** : information, pas une erreur — seule valeur que le script écrit
  automatiquement (toujours avec confirmation par ratio de prix).
- **Couverte par le code** : oui.

### 9. Valeurs manquantes (`valeurs_manquantes`)
- **Exemple réel** : ligne 200, "EMINCE DE PORC FERMIER" — prix de commande manquant.
- **Gravité pressentie** : bloquante.
- **Couverte par le code** : oui.

### 10. Contenant à double imbrication, unité de référence 'PC' imprécise (CAS 3)
- **Exemple réel** : `Base article groupement 1.xlsx` et `2.xlsx`, lignes "GOURDE
  POMME/FRAISE/POIRE/PECHE" (8 lignes au total, 4 par fichier) — conditionnement "CARTON DE 8
  PACKS DE 4 GOURDES DE 90 GR" : unité de référence source 'PC' générique, alors que le vrai
  niveau vendu est le pack (contenant intermédiaire). Règle métier documentée dans
  `useful_stuff/Exemples unités et prix 6.xlsx` ("CAS 3" : colisage défini → le nom du
  contenant devient l'unité de référence/commande, avec poids et nombre de pièces
  obligatoires).
- **Correction appliquée** (uniquement quand le ratio Prix commande ÷ Prix référence confirme
  exactement le nombre de contenants intermédiaires, ex. 8 packs/carton) : Unité de référence
  → 'PACK', Poids à l'unité de référence → 0.36 (kg), Nombre de pièces dans conditionnement →
  4, Unité de commande → 'CARTON'.
- **Gravité pressentie** : information (correction automatique fiable, confirmée par le prix),
  pas une erreur à signaler à une personne.
- **Couverte par le code** : oui (ajouté le 2026-10-15 — motif volontairement très strict,
  limité aux 8 lignes exactes ci-dessus sur l'ensemble des 11 fichiers testés, voir
  `piece_carton_candidats` / `_traite_ligne_etendue_nucleo` dans `audit_geral.py`).

### 15. Bug de détection de colonne sur en-tête au texte différent (corrigé)
- **Découvert en creusant le cas des œufs ci-dessous** : `localise_feuille_et_colonnes`
  repérait la colonne "Unité de référence" par une recherche en substring
  (`"unité de référence" in texte`). Les 6 fichiers `Base IANOR AGAPE 16 ...` utilisent un
  texte différent pour leurs colonnes 4 ("Unité Référence", sans "de") et 5 ("Poids de la PC
  (si Unité de référence = PC)") — cette dernière contient, à l'intérieur de sa parenthèse,
  la phrase exacte "unité de référence", ce qui faisait pointer `col_unite_ref` à TORT sur la
  colonne Poids (toujours vide) dans les 6 fichiers, au lieu de la vraie colonne 4. Résultat
  avant correction : TOUTES les lignes de ces 6 fichiers étaient signalées "Unité de référence
  vide", quelle que soit leur vraie valeur.
- **Gravité** : bloquante (cassait l'audit entier de 6 fichiers sur 11, pas juste quelques
  lignes).
- **Corrigé le 2026-10-15** : détection par préfixe (`texte.startswith("poids")` /
  `texte.startswith("unité") and "référence" in texte`) au lieu d'une recherche en substring —
  couvre les deux formulations sans ambiguïté. Zéro impact sur les 5 fichiers de référence
  (vérifié, mêmes totaux avant/après).

### 16. Unité de référence = contenant + quantité, sans poids (ex: œufs "plateau de 60")
- **Exemple réel** : `Base IANOR AGAPE 16 Produits laitiers (2).xlsx`, lignes 104-106 — unité
  de référence directement écrite "plateau de 60" / "caisse de 180" / "caisse de 360", prix de
  référence toujours égal au prix de commande (jamais divisé par le nombre de pièces).
- **Besoin métier** : pouvoir comparer le prix à l'unité (par œuf) entre fournisseurs, quel
  que soit le conditionnement (un plateau de 60 n'est pas comparable à une caisse de 360 sans
  ramener au prix par pièce).
- **Correction appliquée** : Unité de référence → 'PC', Prix de référence recalculé par pièce
  (prix de commande ÷ N), Unité de commande → nom du contenant seul ('PLATEAU'/'CAISSE'),
  Nombre de pièces dans conditionnement → N (seulement si la colonne existe dans le fichier —
  absente des fichiers IANOR AGAPE, voir erreur n°13).
- **Gravité pressentie** : bloquante avant correction (prix non comparable / prix de référence
  copié du prix de commande sans division) ; information une fois corrigé.
- **Couverte par le code** : oui (ajouté le 2026-10-15, motif strict limité à "CONTENANT DE N"
  sans aucun poids — 3 lignes concernées sur les 11 fichiers testés, voir
  `RE_CONTENANT_SEUL_SANS_POIDS` dans `audit_geral.py`).

## Erreurs observées (TT-03/TT-04) mais PAS ENCORE couvertes par le code

### 11. Ligne entièrement vide en fin de tableau (artefact Excel)
- **Exemple réel** : `Base IANOR AGAPE 16 Epicerie.xlsx`, ligne 236 — toutes les colonnes
  vides, y compris Marché/Lot qui sont sinon toujours remplis.
- **Gravité pressentie** : avertissement (ligne à ignorer/supprimer, pas une vraie donnée).
- **Couverte par le code** : non — pas de règle dédiée pour détecter une ligne 100% vide au
  milieu/fin de la zone utilisée.

### 12. Colonne d'en-tête sans nom
- **Exemple réel** : colonne 61 du modèle (en-tête vide, entre "Code RNM / ITAVI" et
  "ID_AURA"), détectée par pandas comme `Unnamed: 57` dans les fichiers IANOR AGAPE.
- **Gravité pressentie** : avertissement — à clarifier avec le métier (usage réel ou
  colonne à supprimer du modèle).
- **Couverte par le code** : non.

### 13. Colonne obligatoire manquante dans le fichier
- **Exemple réel** : les 6 fichiers `Base IANOR AGAPE 16 ...` n'ont pas la colonne 6
  ("Nombre de Pièces dans conditionnement"), présente dans le modèle de référence.
- **Gravité pressentie** : bloquante — fichier incomplet par rapport au format cible.
- **Couverte par le code** : non — `pretraitement_auto.detecte_format` ne vérifie
  actuellement ni le nombre de colonnes ni leur contenu exact (voir note dans
  `format_cible.md`).

### 14. Cellules fusionnées dans la zone d'en-tête (format BPU DQE)
- **Exemple réel** : onglet "Page de Garde" des fichiers `a_formater/`, 20 à 24 plages
  fusionnées — texte libre, pas de données tabulaires.
- **Gravité pressentie** : sans objet pour l'audit (onglet à exclure), mais bloquant pour
  un futur script de transformation de ce format s'il n'est pas anticipé.
- **Couverte par le code** : non applicable (format pas encore pris en charge).

## Jeu de référence (pour TT-09)

Les 5 fichiers de `pret_pour_audit/` (4x "Base article groupement" + "MDC_IMPORT_MARCHE")
ont des résultats connus et stables (vérifiés à plusieurs reprises) :

| Fichier | Lignes avec incohérence | Cellules marquées |
|---|---|---|
| Base article groupement 1 | 124 | 255 |
| Base article groupement 2 | 247 | 427 |
| Base article groupement 3 | 179 | 325 |
| Base article groupement 4 | 147 | 301 |
| MDC_IMPORT_MARCHE | 213 | 351 |

Ces chiffres servent de résultat attendu pour TT-09 (tests de non-régression) : si une
future modification du code change ces totaux sans raison identifiée, c'est un signal
de régression à investiguer.

> Historique : les totaux de cellules marquées de groupement 1 et 2 sont passés de 247→255
> et 419→427 le 2026-10-15, suite à l'ajout de la règle n°10 ci-dessus (contenant à double
> imbrication, 8 lignes concernées, +1 cellule par ligne pour la nouvelle colonne "Nombre de
> pièces" + 1 cellule pour la correction de l'unité de commande). "Lignes avec incohérence"
> n'a pas changé (ces lignes étaient déjà signalées avant, juste avec une valeur moins
> précise) — seul le nombre de cellules marquées a augmenté. Groupement 3, 4 et
> MDC_IMPORT_MARCHE sont restés strictement identiques (aucune ligne concernée par cette
> règle), confirmant qu'elle n'a aucun effet de bord ailleurs.
