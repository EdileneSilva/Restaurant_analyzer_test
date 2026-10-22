# Catalogue d'erreurs (TT-05)

Un bloc par type d'erreur : description, exemple réel (extrait des fichiers canoniques
confirmés dans `pret_pour_audit/`), gravité pressentie, et si une règle existe déjà dans
`audit_general.py` ou non. Base pour TT-06 (implémentation), TT-08 (seuils) et TT-09 (tests).

## Erreurs déjà couvertes par le code (`audit_general.py`)

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
- **Mise à jour 2026-10-22** : cette catégorie de rapport (`unite_ref_vide`) n'existe plus en
  tant que telle — fusionnée dans la nouvelle catégorie n°27 ("Champs obligatoires vides"),
  marquée en orange avec un commentaire explicite au lieu d'une simple couleur grise muette.

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
- **Mise à jour 2026-10-22** : en réalité, cette catégorie de rapport n'a jamais été remplie
  (aucun code ne produisait le mot "manquant" dans un commentaire — un prix vide était
  seulement colorié en gris, sans texte). Catégorie supprimée, remplacée par la n°27 ("Champs
  obligatoires vides"), qui elle couvre vraiment ce cas.

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
  `piece_carton_candidats` / `_traite_ligne_etendue_nucleo` dans `audit_general.py`).

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
  `RE_CONTENANT_SEUL_SANS_POIDS` dans `audit_general.py`).

### 17. Conditionnement vide pour un produit vendu au poids (CAS 2, "VRAC")
- **Exemple réel** : `Base IANOR AGAPE 16 Viandes fraiches (1).xlsx` (29 lignes, 5-38) et
  `Base IANOR AGAPE 16 Volailles (1) (1).xlsx` (28 lignes, 3-28) — Conditionnement vide, mais
  Unité de référence = Unité de commande (KG), et prix de référence = prix de commande
  exactement : pas une donnée manquante, le produit est vendu au poids sans récipient fixe
  (CAS 2 documenté dans `useful_stuff/Exemples unités et prix 6.xlsx` — viandes, poissons,
  fromages, fruits et légumes).
- **Correction appliquée** : Conditionnement rempli à `'VRAC'`, avec commentaire expliquant
  pourquoi (pas une anomalie, juste l'absence de récipient).
- **Gravité pressentie** : information, pas une erreur.
- **Couverte par le code** : oui (ajouté le 2026-10-16).

### 18. Désignation fournisseur vide (prix renseigné malgré tout)
- **Exemple réel** : `MDC_IMPORT_MARCHE - Lots frais 2026 exel.xlsx`, 260 occurrences — semble
  être une convention de ce fichier (colonne jugée redondante avec Désignation de l'article),
  pas un oubli systématique.
- **Décision provisoire** : copier la valeur de Désignation de l'article en attendant une
  clarification avec le métier — **pas encore implémenté**, à faire une fois la discussion
  métier tranchée.
- **Gravité pressentie** : avertissement.
- **Couverte par le code** : non (décision en attente).

### 19. Origine non reconnue ou ambiguë
- **Exemple réel** : 88 occurrences sur les 11 fichiers testés — régions au lieu d'un pays
  ("Charente", "Sud-Ouest"), valeurs combinées ("FRANCE, UE", "UE/non UE", "Islande/Iles
  féroé"), faute de frappe ("Esspagne"), qualificatif non standard ("transformé en France").
- **Gravité pressentie** : avertissement (ambiguïté à clarifier par une personne, jamais
  corrigée automatiquement — contrairement à `pretraitement.traite_origine`, qui lui réécrit
  en 'France').
- **Couverte par le code** : oui (ajouté le 2026-10-16, `verifie_origine` dans
  `audit_general.py` — liste de pays volontairement élargie par rapport à celle de
  `pretraitement.py`, observée directement dans les fichiers réels).

### 20. Poids à l'unité de référence écrit en texte libre au lieu d'un nombre
- **Exemple réel** : 219 lignes sur les fichiers `Base IANOR AGAPE 16 Epicerie/Produits
  laitiers/Surgelé.xlsx` (+ 1 ligne dans `MDC_IMPORT_MARCHE`, ligne 91, "unité") — au lieu
  d'un nombre, la cellule contient du texte comme `'seau de 5 kg'`, `'100 g'`, `'boîte 3/1'`.
  Avant correction, ce texte était traité comme "déjà rempli" (il n'est pas vide), masquant
  silencieusement une vraie incohérence de type.
- **Correction appliquée** : réutilisation du même parseur que pour le Conditionnement (même
  vocabulaire de récipients) — 152 valeurs converties automatiquement en nombre (kg/L), 67
  non convertibles (codes de contenance "boîte N/N" sans équivalent poids connu, ou "unité")
  signalées pour correction manuelle.
- **Bug corrigé au passage** : "boîte" (avec accent) ne correspondait à aucun mot de récipient
  connu (liste écrite sans accent, "boite") — corrigé, les deux formes sont maintenant
  reconnues.
- **Gravité pressentie** : bloquante avant correction (problème de type non détecté) ;
  information pour les 152 converties, avertissement pour les 67 restantes.
- **Couverte par le code** : oui (ajouté le 2026-10-16).

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

## Erreurs et règles ajoutées depuis (2026-10-20 à 2026-10-22)

Le jeu de fichiers de référence est passé de 5 à **7 fichiers** (ajout de
`2026 09 22 Base article groupement 5.xlsx` et `2026 09 22 Base article groupement
commun.xlsx`). Les règles ci-dessous ont été ajoutées ou corrigées en les auditant.

### 21. Palette de couleurs consolidée (8 → 6, + 1 nouvelle)

Les couleurs historiques, qui se recoupaient sémantiquement, ont été fusionnées en 6, plus
une 7e ajoutée depuis pour un besoin différent (complétude, pas cohérence) :

| Catégorie | Couleur (ARGB) |
|---|---|
| Changement confirmé (D/L recalculés avec confirmation par le ratio des prix) | `FF93CDDC` (bleu) |
| Cellule vide / aucune conversion fiable / champ optionnel vide | `FFBFBFBF` (gris) |
| Incohérence (prix suggéré non appliqué, unité de commande suspecte...) | `FFD5A6BD` (rose/mauve) |
| Prix sans aucune explication dans le texte | `FF92D050` (vert) |
| Copie possible (même Référence+prix, ou même Désignation) | `FFFF6969` (rouge clair) |
| Ligne strictement dupliquée (ligne entière) | `FFFF0000` (rouge vif) |
| **Champ obligatoire vide** (nouveau, n°27 ci-dessous) | `FFFABF8F` (orange) |

### 22. Le prix de référence (K) n'est JAMAIS écrit automatiquement

Avant : quand le script était confiant dans un recalcul de prix de référence, il écrivait la
nouvelle valeur (couleur bleue). **Décision changée** : même un recalcul confiant n'est plus
jamais appliqué automatiquement — seulement surligné en rose + commenté ("Prix de référence
suggéré : ... — à vérifier, non modifié automatiquement"), exactement comme une incohérence.
Une personne valide toujours elle-même un changement de prix. Modifié à la fois dans
`audit_general.py` et dans le moteur partagé `first_scripts/first_attempt_v2_prototype.py`.

### 23. Vocabulaire de récipients très étendu

`MOTS_CONTENANT` est passé de ~22 à 32 mots, au fil des fichiers réels rencontrés : ajout de
`bac`, `coffret`, `plaquette`, `paquet`, `etui`/`étui`, `bombe` (+ `bouteile`, faute de frappe
fréquente à un seul L, déjà tolérée). Une seconde liste séparée (`mots_types` dans
`verifie_unite_commande_etendue`) doit être tenue à jour EN PARALLÈLE à chaque ajout — sinon
l'unité de commande fraîchement confirmée est re-signalée à tort comme incohérente juste
après (bug trouvé et corrigé plusieurs fois ce lot-ci).

### 24. Formats de conditionnement supplémentaires reconnus

- `CARTON 6 PIECES DE 300 GR` (sans "DE" entre le récipient et le nombre).
- `COLIS 30 G X100` / `COLIS X960 5 KG` (poids et compte dans n'importe quel ordre, séparés
  par "X" collé au nombre).
- `CARTON X 5 ST X 1 KG` (sachet abrégé en "ST", reconnu UNIQUEMENT à cette position précise
  — pas ajouté au vocabulaire général des récipients, trop court et ambigu ailleurs).
- `PIECE DE 8 PORTIONS` (le mot de comptage final n'est plus limité à une liste fermée :
  "pièces"/"u"/"portions"/"tranches", ou n'importe quel autre mot qui n'est pas lui-même une
  unité de poids/volume reconnue — ex: "CAGEOT BOIS 12 U", "CARTON 14 SACHETS").
- Qualificatifs de matière à un ou plusieurs mots entre le récipient et sa quantité
  ("CAISSE BOIS 10 KG", "CAGEOT BOIS OU CARTON 5KG").
- `CONTENANT DE N X QTÉ UNITÉ` dont le ratio confirme N exactement (pas la quantité totale) :
  réinterprété comme "prix par sous-paquet", Unité de référence devient `PC`, le poids d'un
  sous-paquet renseigne la colonne 5, prix inchangés (ex: "CARTON DE 9 X 500 GR" avec ratio=9).

### 25. "LOT DE N" : l'unité de commande ne garde que "LOT"

Avant : l'unité de commande pouvait devenir "LOT DE 4". **Décision changée** : elle ne garde
que `LOT` (le nombre n'est pas vraiment une unité), et le nombre est reporté en préfixe du
Conditionnement ("LOT DE 4, 4 x 125g") — sauf si le mot "lot" y figure déjà sous une forme ou
une autre, pour ne pas le répéter. Centralisé au point d'écriture unique de la colonne Unité
de commande dans `audite_fournisseur`.

### 26. Lignes et colonnes masquées (résidu de filtre Excel) réaffichées automatiquement

Une personne qui remplit le fichier applique parfois un filtre automatique pour ne voir qu'un
sous-ensemble de produits, puis enregistre le fichier dans cet état : les lignes (voire une
colonne) masquées contiennent souvent un produit bien réel, invisible à la relecture si elles
ne sont pas réaffichées. `desocculte_lignes_et_colonnes` les réaffiche systématiquement, pour
tous les fichiers — purement cosmétique, ne touche aucune donnée. Trouvé sur
`2026 09 22 Base article groupement 5.xlsx` (577 lignes masquées) et
`MDC_IMPORT_MARCHE...xlsx` (1 colonne, Désignation fournisseur du Fournisseur 1).

### 27. Champs obligatoires / optionnels vides, signalés explicitement

Nouvelle règle (`signale_champs_obligatoires_optionnels`), appliquée uniquement sur une ligne
qui contient déjà au moins une donnée (une ligne totalement vide reste gérée par la règle n°11
ci-dessous, candidate à la suppression) :
- **Obligatoires** (orange, "bloque l'intégration") : Marché, Lot, Désignation de l'article,
  Unité de référence, et par bloc Fournisseur : son nom, l'Unité de commande, le Prix de
  référence, le Prix de commande.
- **Optionnels** (gris, "conseillé") : par bloc Fournisseur : Conditionnement, Désignation
  fournisseur, Référence produit.
- Un bloc Fournisseur N entièrement vide n'est jamais signalé (il n'a simplement pas
  participé à cette ligne — le bloc Fournisseur 2/3 n'existe que pour comparaison avec le 1).

### 28. Même Désignation de l'article, mais Lot différent

Deux lots distincts ne devraient pas décrire le même produit avec un texte de désignation
rigoureusement identique — signe qu'il manque une information distinctive (variété, poids,
origine, label...) dans au moins l'une des deux lignes. Fusionné dans
`signale_designation_dupliquee` (même fonction que la règle n°2, juste un second signalement
indépendant sur la même colonne — une ligne peut recevoir les deux à la fois).

### 29. Rapport `.md` : catégories manquantes ou jamais remplies, corrigées

En relisant les nouveaux signalements ci-dessus, plusieurs trous ont été trouvés dans
`calcule_categories_rapport` (qui classe les commentaires posés sur la feuille par catégorie
pour le rapport `.md` et l'app Streamlit) : le signalement "même Désignation, même
fournisseur" (règle n°2 originale) n'apparaissait dans AUCUNE catégorie du rapport (invisible,
bien que marqué dans le fichier) ; "Unité de référence vide" et "Valeurs manquantes"
n'étaient, elles, jamais remplies du tout (aucun code ne produit le mot "manquant" dans un
commentaire) — supprimées, remplacées par les nouvelles catégories n°27/28 ci-dessus, plus
cohérentes. Un oubli similaire existait dans `app.py` (total de cellules marquées affiché dans
l'app, pas mis à jour avec la nouvelle règle n°27) — corrigé aussi.

Le rapport `.md` liste maintenant, tout en haut, les valeurs UNIQUES de Lot (triées par ordre
croissant du numéro de lot) et de chaque Fournisseur (espaces de début/fin comptés, pour
repérer une variante qui devrait être identique à une autre mais ne l'est pas exactement).

### 30. Tolérance de ratio resserrée de 5 % à 2 %

`TOLERANCE_RATIO` (utilisée partout pour confirmer qu'un ratio Prix commande ÷ Prix référence
correspond à une quantité attendue) est passée de 5 % à 2 %, jugée trop permissive. Conséquence
mesurée : 5 cellules en plus signalées sur `groupement 1` (des cas à 2-3,5 % d'écart,
auparavant acceptés en silence), aucune perte de confirmation déjà posée ailleurs.

### 31. Bug corrigé — doublon complet qui ignorait la Désignation

`detecte_lignes_dupliquees` comparait toutes les colonnes SAUF la Désignation de l'article
pour décider qu'une ligne entière est un doublon strict — deux produits réellement différents
(ex: "SAUTE VEAU BIO" vs "SAUTE VEAU LR", même référence/conditionnement/prix) étaient peints
en rouge vif à tort. Corrigé : la Désignation fait maintenant partie de la clé de
comparaison — seules les VRAIES duplicatas restent en rouge vif ; la fonction séparée qui
servait de redondance partielle à ce cas (`detecte_doublons_par_reference`) a été supprimée.

### 32. Bug corrigé — "pièce"/"pièce" jamais utilisé comme nom de récipient réel

Plusieurs motifs de conditionnement utilisent "pièce" comme marqueur de comptage générique
("CARTON 6 PIECES DE 300 GR", "CARTON DE 40 PIECES") — ce n'est jamais un vrai récipient à
reprendre tel quel dans l'unité de commande ou de référence. Corrigé à trois endroits
différents du code (trouvé un par un, le même bug reproduit par plusieurs règles
indépendantes) : la substitution directe de l'unité de commande, le cas "contenant à double
imbrication" (CAS 3, règle n°10), et la détection de récipient imbriqué de premier niveau.

## Jeu de référence (pour TT-09)

Les 7 fichiers de `pret_pour_audit/` (5x "Base article groupement" + "Base article
groupement commun" + "MDC_IMPORT_MARCHE") ont des résultats connus et stables, mesurés au
2026-10-22 (après toutes les règles ci-dessus) :

| Fichier | Lignes avec incohérence | Cellules marquées |
|---|---|---|
| Base article groupement 1 | 146 | 797 |
| Base article groupement 2 | 256 | 826 |
| Base article groupement 3 | 195 | 991 |
| Base article groupement 4 | 154 | 959 |
| Base article groupement 5 | 196 | 870 |
| Base article groupement commun | 690 | 1260 |
| MDC_IMPORT_MARCHE | 210 | 643 |

Ces chiffres servent de résultat attendu pour TT-09 (tests de non-régression) : si une
future modification du code change ces totaux sans raison identifiée, c'est un signal
de régression à investiguer. Les écarts importants avec les anciens chiffres (ex: groupement
1 : 256→797 cellules) reflètent l'accumulation des nombreuses règles n°21-32 ci-dessus
(notamment n°27, qui à elle seule ajoute plusieurs centaines de signalements par fichier),
pas une régression — voir le détail de chaque règle pour la justification.

> Historique : les totaux de cellules marquées de groupement 1 et 2 sont passés de 247→255
> et 419→427 le 2026-10-15, suite à l'ajout de la règle n°10 ci-dessus (contenant à double
> imbrication, 8 lignes concernées, +1 cellule par ligne pour la nouvelle colonne "Nombre de
> pièces" + 1 cellule pour la correction de l'unité de commande). "Lignes avec incohérence"
> n'a pas changé (ces lignes étaient déjà signalées avant, juste avec une valeur moins
> précise) — seul le nombre de cellules marquées a augmenté. Groupement 3, 4 et
> MDC_IMPORT_MARCHE sont restés strictement identiques (aucune ligne concernée par cette
> règle), confirmant qu'elle n'a aucun effet de bord ailleurs.

> Historique (2026-10-16) : groupement 1 (255→256, 124→125) et groupement 4 (301→302,
> 147→148) suite à l'ajout de la règle n°19 (Origine non reconnue/ambiguë) : chacun a 1 ligne
> avec Origine "FRANCE, UE" (valeur combinée). MDC_IMPORT_MARCHE (351→352, 213→214) suite à
> la règle n°20 (Poids en texte libre) : ligne 91, "unité". Groupement 2 et 3 inchangés. La
> règle n°17 (VRAC) ne touche aucun des 5 fichiers de référence — elle ne concerne que les
> fichiers IANOR AGAPE (Viandes fraîches, Volailles), donc pas de ligne dans ce tableau pour
> elle.
