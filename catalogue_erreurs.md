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

**Mise à jour (2026-11-09)** : `audit_general_commentaires.py` avait une exception volontaire
qui n'existe plus — Unité de référence et Unité de commande vides n'y étaient jamais signalées
comme "champ obligatoire vide" (contrairement à `audit_general.py`), au motif que choisir la
bonne unité serait la responsabilité de l'équipe d'audit, pas de la personne qui remplit le
fichier. Retiré à la demande de l'utilisateur : la règle doit être identique dans les deux
moteurs, peu importe la raison pour laquelle une colonne est vide — voir les totaux mis à jour
dans « Jeu de référence » plus bas.

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

### 30. Tolérance de ratio resserrée de 5 % à 2 %, puis réélargie à 3 %

`TOLERANCE_RATIO` (utilisée partout pour confirmer qu'un ratio Prix commande ÷ Prix référence
correspond à une quantité attendue) est passée de 5 % à 2 % le 2026-10-22, jugée trop
permissive. Conséquence mesurée : 5 cellules en plus signalées sur `groupement 1` (des cas à
2-3,5 % d'écart, auparavant acceptés en silence), aucune perte de confirmation déjà posée
ailleurs.

**Mise à jour (2026-10-22, plus tard le même jour)** : 2 % s'est révélé trop strict à l'usage
inverse — plusieurs cas réels avec un écart de 2-3 % (arrondis de pesée, conditionnement
légèrement variable) étaient signalés à tort comme incohérents alors qu'ils sont normaux côté
métier. Remontée à **3 %**, valeur actuelle de `motor.TOLERANCE_RATIO`, utilisée de façon
identique dans `audit_general.py` et `audit_general_commentaires.py`. (Deux commentaires de
code mentionnant encore "2%" à titre d'exemple historique subsistent dans `audit_general.py` —
lignes ~1268 et ~1439 — ils documentent le raisonnement d'origine, pas la valeur actuellement
appliquée ; à corriger si le code est retouché pour une autre raison.)

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

## Erreurs et règles ajoutées depuis (2026-10-22 à 2026-10-26)

### 33. Mot de récipient non reconnu par la base : remplacement automatique (`audit_general.py` seulement)

Avant, un mot comme "COLIS" ou "CAGEOT" (absent de la liste fermée acceptée par la base
d'intégration, voir `unites_commande_non_reconnues.md`) était seulement confirmé comme cohérent
avec le conditionnement, jamais corrigé. **Décision métier appliquée** : dans
`audit_general.py` uniquement, l'Unité de commande (et, le cas échéant, l'Unité de référence)
est maintenant réécrite directement par son équivalent officiel (`motor.CONTENANT_ABREGE`,
construit à partir de la colonne "Unité de remplacement" de `unites_commande_non_reconnues.xlsx`
— ex. `colis → CARTON`, `cageot → PLATEAU`, `caisse → CARTON`, `pack → LOT`), avec un
commentaire explicite ("'COLIS' n'étant pas reconnu par la base, remplacé par son équivalent
officiel 'CARTON'"). Une abréviation d'un récipient déjà correct (ex. `bq → BARQUETTE`) est
traitée pareil, dans le même dictionnaire.
- **Choix volontaire et explicite** : cette règle ne s'applique **que** dans
  `audit_general.py`. Dans `audit_general_commentaires.py` (version destinée à une personne non
  technique), elle est désactivée : "la responsabilité de laisser l'unité correcte est celle de
  l'équipe d'audit, pas de la personne qui remplit le fichier" — voir le docstring de ce
  fichier et `format_cible.md`. La seule chose que `audit_general_commentaires.py` reconnaît,
  c'est quand la personne a DÉJÀ écrit le mot officiellement correct (ex. "CARTON" alors que le
  conditionnement dit "colis") : dans ce cas, aucune incohérence n'est signalée — c'est déjà le
  résultat recherché.
- **Gravité pressentie** : information (correction fiable, motivée par l'intégration, pas une
  erreur de saisie).
- **Couverte par le code** : oui, `audit_general.py` seulement (ajoutée le 2026-10-26).

### 34. Format "BOITE N/M" reconnu comme catégorie de récipient à part entière

- **Exemple réel** : conditionnement du type "BOITE 4/1", "BOITE 5/1" — avant, ce format
  n'était rattaché à aucune catégorie connue et pouvait déclencher une incohérence Unité de
  commande à tort, en plus du signalement légitime posé ailleurs.
- **Correction appliquée** : `RE_FORMAT_BOITE` reconnaît ce motif et le rattache à la catégorie
  "boite", qui est désormais acceptée comme cohérente avec une Unité de commande "BOITE".
- **Gravité pressentie** : correction technique (évite un double-commentaire contradictoire sur
  la même cellule), pas une erreur en soi.
- **Couverte par le code** : oui (ajoutée le 2026-10-22, `reconsidere_format_boite`).

### 35. "LOT DE N" : affinage et garde-fou contre la répétition

Suite de la règle n°25 : l'Unité de commande "N PIECES" écrite telle quelle par le fournisseur
(ex. "12 PIECES") est maintenant reconnue et convertie en "LOT DE N" **sauf si** l'Unité de
commande actuelle est déjà "LOT" ou "LOT DE N" sous une autre forme — dans ce cas, aucun
changement n'est proposé (évite de transformer une valeur déjà correcte en une formulation
légèrement différente, qui serait un changement sans réel apport). Détection des comptages de
pièces élargie à des formats libres ("6 BOTTES", "8U", "X20 U"), pas seulement "N pièces"/"N
pieces" au sens strict.
- **Gravité pressentie** : information (correction fiable).
- **Couverte par le code** : oui.

### 36. Validation d'un "Poids à l'unité de référence" déjà rempli (nouveau — avant, seule la case VIDE était vérifiée)

- **Avant** : le script ne remplissait le "Poids à l'unité de référence" que lorsqu'il était
  vide ; une valeur déjà présente n'était jamais recoupée avec le conditionnement ou la
  désignation.
- **Correction appliquée** : `reconsidere_poids_piece_existant` recoupe maintenant une valeur
  déjà présente avec le poids déductible du conditionnement/de la désignation ; en cas d'écart,
  la cellule est surlignée avec un commentaire ("le poids attendu semble plutôt être...") — la
  valeur n'est JAMAIS réécrite automatiquement dans ce cas (contrairement au remplissage d'une
  case vide, qui lui reste une écriture directe si confirmé par le ratio de prix).
- **Gravité pressentie** : avertissement — à vérifier par une personne.
- **Couverte par le code** : oui (ajoutée le 2026-10-22).

### 37. Bug corrigé — mauvaise suggestion "SAC" par défaut sur légumes/fruits quand un vrai récipient figurait dans un texte "sale"

- **Exemple réel** : `2026 09 22 Base article groupement commun.xlsx`, lignes 58-60 et 63 —
  conditionnement du type "COLIS BOIS 4 KG 12 PIECES". `applique_defaut_legume` ne regardait que
  le motif STRICT (`motor.mot_type_conditionnement`), qui échoue sur un texte aussi irrégulier,
  concluait à tort "aucun récipient trouvé" et proposait "SAC" par défaut — alors qu'un
  récipient ("COLIS") était bien nommé, juste dans un texte au format inhabituel.
- **Correction appliquée** : le garde-fou vérifie maintenant AUSSI le motif souple
  (`mot_contenant_dans_conditionnement`, recherche n'importe où dans le texte) avant de
  conclure à l'absence de récipient — corrigé à la fois dans `audit_general.py` et
  `audit_general_commentaires.py`.
- **Gravité** : le changement erroné écrasait une Unité de commande déjà correcte
  ("COLIS"/devenu "CARTON" une fois la règle n°33 appliquée) par "SAC", avec un commentaire
  contradictoire — corrigé.
- **Couverte par le code** : oui (corrigée le 2026-10-26).

### 38. Bugs de double-commentaire corrigés (plusieurs cas indépendants)

Motif récurrent ce lot-ci : deux mécanismes indépendants (le chemin principal qui décide d'un
changement, et le garde-fou de fin de fonction `verifie_unite_commande_etendue`/
`audite_fournisseur` qui revérifie la cohérence) arrivaient chacun à une conclusion différente
sur la MÊME cellule, produisant deux commentaires empilés et contradictoires. Cas trouvés et
corrigés un par un, par enrichissement du garde-fou (reconnaît maintenant aussi l'abréviation
d'un récipient déjà correct, en plus du remplacement officiel pour un récipient non reconnu) :
format BOITE N/M (règle n°34), BOCAL/POT (comparaison sur le texte littéral de l'Unité de
commande, pas sur la catégorie regroupée qui aurait fait rater la comparaison), LOT/LOT DE N
(règle n°35), et COLIS/CARTON/BARQUETTE (règles n°33 et 37 combinées — dernier cas trouvé :
`audit_general_commentaires.py`, groupement commun ligne 7, "BQ DE 125 GR X 12" confirmé en
"BARQUETTE", mais revérifié à tort contre une liste qui ne couvrait que les remplacements
officiels, pas les abréviations).
- **Couverte par le code** : oui, dans les deux fichiers, au fil des cas trouvés par tests de
  non-régression sur les 7 fichiers (aucune duplication de commentaire restante au 2026-10-26,
  vérifié par balayage automatique de tous les commentaires posés).

### 39. Performance — propriétés `max_column`/`max_row` d'openpyxl recalculées à chaque ligne

Bug de performance pur (zéro impact sur le résultat) : plusieurs fonctions
(`marque_lignes_dupliquees`, `applique_formatage_uniforme`, `calcule_categories_rapport`, et le
repérage des lignes vides dans `traite_feuille`) relisaient `ws.max_column`/`ws.max_row` à
CHAQUE itération de leur boucle sur les lignes — ces propriétés ne sont pas mises en cache par
openpyxl, chaque accès reparcourt toutes les cellules de la feuille. Sur les plus gros fichiers,
le temps total des 7 fichiers combinés était passé d'environ 20 secondes à près de 2 minutes.
Corrigé en mettant `max_column`/`max_row` dans une variable locale une seule fois avant chaque
boucle. Ce même correctif technique a été appliqué dans les deux fichiers.

### 40. `audit_general_commentaires.py` : liste d'exceptions mécaniques étendue

Au-delà des corrections mécaniques déjà listées dans son propre docstring (voir le fichier
lui-même, en-tête), deux écritures silencieuses supplémentaires ont été ajoutées ce lot-ci,
dans le même esprit ("aucun jugement de contenu requis") : le remplissage du "Poids à l'unité
de référence" vide (déjà silencieux dans `audit_general.py`, désormais avec un commentaire
explicatif ici, par transparence envers un lecteur non technique) et la normalisation d'un
synonyme pur d'unité de référence déjà dans la liste acceptée (kilo/kilogramme → KG, litre → L,
pièce/piece → PC — jamais "Unité", qui reste un code à part entière). Le remplacement
"mot non reconnu → équivalent officiel de la base" (règle n°33) reste volontairement exclu de
ce fichier — seule son exception inverse (reconnaître que le mot déjà écrit EST l'équivalent
officiel, et donc ne rien signaler) y est présente.

### 41. ⚠️ Danger découvert en écrivant les tests (TT-09) — importer les deux scripts dans le même processus corrompt silencieusement leurs résultats

`audit_general.py` et `audit_general_commentaires.py` partagent le même moteur
(`first_scripts/first_attempt_v2_prototype.py`, importé comme `motor`) et **chacun écrase, dès
son import, des constantes globales dessus** (`motor.CONTENANT_ABREGE`, `motor.TOLERANCE_RATIO`,
etc. — voir en tête de chaque fichier, juste après les imports) avec SES PROPRES valeurs.
Problème : comme `motor` n'est chargé qu'une seule fois par processus Python (mécanisme normal
de `sys.modules`), si un même programme importe les DEUX scripts, le second import écrase
silencieusement les constantes posées par le premier — sans aucune erreur, aucun avertissement.

- **Repéré comment** : en écrivant un premier jet de tests de non-régression qui importait les
  deux scripts dans le même fichier de test. Les totaux de `audit_general.py` tombaient de 234 à
  204 lignes avec incohérence sur `groupement 3` (et des écarts similaires sur 3 autres
  fichiers) selon que le test tournait seul ou après avoir importé
  `audit_general_commentaires`.
- **Confirmé** : `motor.CONTENANT_ABREGE['colis']` vaut `'CARTON'` juste après `import
  audit_general`, mais devient `'COLIS'` dès qu'`audit_general_commentaires` est importé
  ensuite dans le même interpréteur — alors que la variable manipulée est littéralement le même
  objet Python (`ag.motor is agc.motor` → `True`).
- **Pas un bug dans la logique métier des deux scripts** : chacun, utilisé seul (en CLI, ou
  importé seul comme le fait `app.py`), se comporte correctement. Le risque n'existe que si un
  même processus Python importe les deux.
- **Mitigation appliquée** : `test_non_regression_audit_general.py` et
  `test_non_regression_audit_commentaires.py` sont deux fichiers de test séparés, chacun
  n'important qu'un seul des deux scripts, à lancer dans deux invocations `pytest` distinctes —
  jamais ensemble.
- **Point de vigilance pour la suite** : tout futur code qui aurait besoin des deux en même
  temps (ex. un agent qui orchestrerait les deux versions, ou un test qui les comparerait) devra
  isoler chaque appel dans son propre processus (`subprocess`, ou deux process `multiprocessing`
  séparés) plutôt que de les importer tous les deux dans le même interpréteur.

### 42. Même unité des deux côtés → mêmes prix, sans exception (`verifie_coherence_unite_prix_identiques`)

Nouvelle règle inconditionnelle (les deux moteurs) : quand l'Unité de référence et l'Unité de
commande finales sont le même texte, les deux prix (référence et commande) DOIVENT être
égaux — signalé systématiquement, même quand le Conditionnement semble expliquer
mathématiquement l'écart (contrairement au reste du moteur, qui acceptait ce cas en silence).
Garde-fou contre le doublon : si `resultat.couleur_m` est déjà posé (le moteur partagé a déjà
détecté ce même cas via son propre canal "même unité, ratio ≠ 1, rien ne l'explique"), la
nouvelle règle ne rajoute rien, pour ne pas signaler deux fois la même cellule (trouvé via
`test_sans_doublon_de_commentaire`). N'affecte que "MDC_IMPORT_MARCHE" parmi les 7 fichiers de
référence — exemple réel : ligne 163 "pâté de foie supérieur", KG des deux côtés, 0,6667€
contre 0,95€.

**Bug corrigé (2026-11-10)** : la comparaison utilisait `resultat.nouveau_k` (une simple
SUGGESTION de recalcul, jamais écrite dans la cellule — voir règle n°22, "le prix de référence
n'est JAMAIS écrit automatiquement") à la place du Prix de référence RÉEL, ce qui pouvait
produire un message contradictoire : "les deux prix devraient être égaux... mais Prix de
référence (0.225) ≠ Prix de commande (2.25)" alors que le Prix de référence RÉEL affiché dans
la cellule valait en fait aussi 2.25€ (donc déjà égal). Trouvé par l'utilisateur en testant le
fichier Vivalya MP ARDECHE 24 (ligne 92, "Pomme de terre crue épluchée cube... 5 kg", Unité de
référence = Unité de commande = 'KG', Prix de référence = Prix de commande = 2,25€,
Conditionnement = "10KG"). Corrigé en séparant deux cas : (1) les prix RÉELS diffèrent → message
d'origine, avec les vraies valeurs ; (2) les prix réels sont déjà égaux, mais une suggestion de
recalcul (`nouveau_k`) existe quand même, parce que le Conditionnement indique une quantité
réelle différente de 1 (ex: "10KG") → nouveau message, qui ne prétend plus que les prix
diffèrent : "cette relation est correcte. Mais le Conditionnement... indique en fait une
quantité différente de 1... vérifiez si l'Unité de commande et/ou le Prix de commande
reflètent bien tout le conditionnement, ou seulement une partie." Aucun changement de total
(même cellule marquée avant/après, seul le texte était faux).

### 43. Conditionnement réduit à un nombre nu, sans unité ni récipient (`reconsidere_nombre_sans_unite`)

Un Conditionnement recopié tel quel depuis la source quand aucun code/mot n'a permis de
l'interpréter (ex: "6,000", vu sur `pretraitement_a_formater.transforme_ardeche` quand le code
source "UC" n'est ni KG/L/UN ni un mot de contenant reconnu dans la désignation) passait
inaperçu : aucune règle ne signalait ce nombre, alors qu'un nombre différent de 1 peut être un
multiplicateur de prix jamais vérifié. Nouvelle règle, les deux moteurs : un nombre nu ≠ 1 dans
le Conditionnement est désormais signalé explicitement sur sa propre cellule ("ce nombre peut
être un multiplicateur du prix qui n'a pas été pris en compte"). Un nombre = 1 reste muet (rien
à multiplier). Exemple réel : DADS MP ARDECHE 24 Lot 06, "Lait 1/2 écrémé UHT BIO", UC="BT"
(aucun mot de contenant dans la désignation), Conditionnement="6,000".

### 44. Faux signalement de prix sur un contenant pluriel d'unités déjà vendues à l'unité

Quand le Conditionnement nomme un récipient trouvé dans la désignation ET que l'Unité de
commande dit déjà qu'on vend à l'unité individuelle ("UNITÉ", "PIECE", "PC"), le nombre lu
devant ce récipient (ex: "48" dans "48 POTS DE 125G") n'est qu'un compte de récipients dans un
emballage plus grand, PAS un multiplicateur du prix de commande (déjà au récipient). Avant
cette règle, `determine_base_et_quantite` traitait ce nombre comme un total à diviser
(48 × 125 g = 6 kg), suggérant à tort un prix de référence 48 fois trop petit. Corrigé en
retrouvant le nombre N séparément (même regex que celle qui l'a reconnu) et en divisant la
quantité par N avant le calcul du ratio, uniquement quand l'Unité de commande indique déjà une
vente à l'unité. La suggestion "l'unité de commande semble plutôt être '<récipient>'" (utile,
correcte) n'est pas affectée, seule la suggestion de prix erronée disparaît. Exemple réel :
DADS MP ARDECHE 24 Lot 06, "Yaourt 1/2 écrémé nature pot carton (x48) 125 g", Unité de
commande="UNITÉ", Conditionnement="48 POTS DE 125G", Prix de référence=3,04€/kg (correct) —
suggérait à tort 0,0633€/kg avant correction.

### 45. Récipient d'une seule unité de référence : aucune correction proposée, dans aucun sens

Quand l'Unité de référence et l'Unité de commande sont déjà identiques, que le Prix de
référence est déjà égal au Prix de commande (rien à recalculer), et que le Conditionnement ne
décrit qu'UNE seule unité de référence (ex: "BRIQUE DE 1 L" ou "SACHET DE 1 KG", quantité = 1),
plus aucune suggestion de changer l'Unité de commande n'est faite — ni vers le mot du récipient
(le nommer n'apporte aucune information de plus que l'unité de mesure déjà écrite), ni dans
l'autre sens. Règle symétrique demandée explicitement par l'utilisateur ("vale pra os dois
lados"), confirmée après lui avoir montré, comme contre-exemple, le cas réel "SACHET DE 1 KG"
(`2026 09 22 Base article groupement 1.xlsx`, lignes 183-191, ex: STEAK HACHE DE BOEUF) —
structurellement identique au cas "BRIQUE DE 1 L" qu'elle avait signalé (DADS MP ARDECHE 24
Lot 06) : la même règle s'applique aux deux, pas de garder un cas en exception. Pour une
quantité ≠ 1 (ex: "SEAU DE 10 L"), la suggestion continue d'apparaître normalement — voir les
totaux de non-régression mis à jour ci-dessous.

### 46. Prix à 0€ : alarme dédiée, et analyse de cohérence unité/prix désactivée pour la ligne

Un Prix de référence ou un Prix de commande valant exactement 0€ est désormais signalé
explicitement sur sa propre cellule ("Un prix de 0€ est impossible pour un article réellement
vendu — vérifier la saisie..."), au lieu de passer inaperçu ou de produire un signalement
inadapté (un ratio/calcul de prix basé sur une division par 0 n'a aucun sens). Quand l'un des
deux prix vaut 0, toute la suite de l'analyse de cohérence unité/prix pour cette ligne
(`traite_ligne_etendue`, `verifie_coherence_unite_prix_identiques`, etc.) est sautée : demandé
explicitement par l'utilisateur, qui a jugé que "não fará muito sentido" de comparer des
unités/prix quand le prix de départ est déjà 0. Trouvé sur le fichier Vivalya MP ARDECHE 24
(beaucoup de lignes à "0,000 €" côté source, probablement des prix non communiqués par le
fournisseur plutôt que des cases vides) — n'affecte aucun des 7 fichiers de référence (aucun
prix à 0€ exact dedans), donc aucun total de non-régression à mettre à jour.

## Jeu de référence (pour TT-09)

Les 7 fichiers de `pret_pour_audit/` (5x "Base article groupement" + "Base article
groupement commun" + "MDC_IMPORT_MARCHE") ont des résultats connus et stables. **Valeurs
actuelles (2026-11-09, après toutes les règles ci-dessus, y compris la dernière du
2026-11-09)** — ce sont exactement celles vérifiées par `test_non_regression_audit_general.py`
(`TOTAUX`) aujourd'hui ; l'historique des changements successifs (règle par règle, avec la
raison et l'exemple réel qui l'a motivée) est détaillé dans les paragraphes "Mise à jour"
ci-dessous, du plus ancien au plus récent :

| Fichier | Lignes avec incohérence | Cellules marquées |
|---|---|---|
| Base article groupement 1 | 173 | 799 |
| Base article groupement 2 | 293 | 878 |
| Base article groupement 3 | 224 | 1024 |
| Base article groupement 4 | 191 | 1003 |
| Base article groupement 5 | 219 | 889 |
| Base article groupement commun | 748 | 1298 |
| MDC_IMPORT_MARCHE | 193 | 721 |

Ces chiffres (`audit_general.py`, sortie dans `transformed_files/`) servent de résultat attendu
pour TT-09 (tests de non-régression) : si une future modification du code change ces totaux
sans raison identifiée, c'est un signal de régression à investiguer — voir
`test_non_regression_audit_general.py`, qui vérifie ces totaux automatiquement.

> **Mise à jour (2026-10-30)** : "Base article groupement commun" baisse de 772→760 lignes et
> 1348→1310 cellules suite à l'élargissement de `PAYS_ORIGINE_CONNUS` (synchronisé avec
> `PAYS_UE`/`PAYS_HORS_UE` de `pretraitement_a_formater.py`, demandé le même jour) : 14 valeurs
> d'Origine réellement présentes dans ce fichier (AUTRICHE, BULGARIE, COTE D'IVOIRE, EQUATEUR,
> LITUANIE, NOUVELLE-ZÉLANDE, PAYS BAS, ROUMANIE, SUEDE, THAILANDE, TUNISIE, TURQUIE, VIETNAM)
> étaient signalées à tort comme "non reconnue" — ce n'était pas une régression du code
> d'audit, juste une liste de pays trop courte. Vérifié en comparant ligne par ligne les
> résultats avant/après sur ce même fichier (seules ces 14 valeurs changent de statut).

> `audit_general_commentaires.py` (sortie dans `commented_files/`) a ses propres totaux,
> structurellement différents puisque cette version ne prend presque jamais de décision de
> contenu (voir son docstring). **Valeurs actuelles (2026-11-09)**, vérifiées par
> `test_non_regression_audit_commentaires.py` (`TOTAUX_CELLULES`) : 756 / 827 / 987 / 955 /
> 864 / 955 / 744 cellules marquées, dans le même ordre de fichiers que ci-dessus. Historique :
> au 2026-10-26, 664 / 737 / 855 / 811 / 754 / 965 / 738
> cellules marquées, dans le même ordre de fichiers que ci-dessus (total 5271 sur les 7
> fichiers). Zéro doublon de commentaire et zéro changement sans effet réel (no-op) vérifiés
> par balayage automatique sur les 7 fichiers à cette date. **Mise à jour (2026-10-30)** :
> "Base article groupement commun" passe de 965 à 927 cellules marquées, même cause que
> ci-dessus (liste de pays élargie, identique dans les deux moteurs).
>
> **Mise à jour (2026-11-09)** : `signale_champs_obligatoires_optionnels` (dans
> `audit_general_commentaires.py`) avait une exception volontaire, documentée dans son propre
> code mais jamais dans ce fichier : Unité de référence et Unité de commande vides n'étaient
> JAMAIS signalées comme "champ obligatoire vide" dans ce moteur (contrairement à
> `audit_general.py`), au motif que "laisser l'unité correcte est notre responsabilité
> (l'équipe d'audit), pas celle de la personne qui remplit le fichier". Retiré à la demande de
> l'utilisateur, qui a établi une règle contraire : la règle obligatoire/optionnel doit être
> identique partout, peu importe la raison pour laquelle une colonne est vide. Conséquence
> attendue : tous les totaux de cellules montent, puisque bien plus de lignes ont ces deux
> champs vides par construction (ex: BPUR FRESCA, qui n'a jamais d'Unité de commande). Nouveaux
> totaux, même ordre de fichiers : 766 / 837 / 997 / 962 / 871 / 967 / 743.

> **Mise à jour (2026-11-09, plus tard le même jour)** : nouvelle règle
> `verifie_coherence_unite_prix_identiques` (les deux moteurs), demandée explicitement par
> l'utilisateur : quand l'Unité de référence et l'Unité de commande finales sont le même
> texte, les deux prix DOIVENT être égaux — signalé systématiquement, même quand le
> Conditionnement semble expliquer mathématiquement l'écart (contrairement au reste du moteur,
> qui acceptait ce cas en silence). N'affecte que "MDC_IMPORT_MARCHE" parmi les 7 fichiers de
> référence : 194→198 lignes et 720→726 cellules (`audit_general.py`), 743→749 cellules
> (`audit_general_commentaires.py`). Les 6 autres fichiers restent strictement identiques.
> Vérifié ligne par ligne : les nouvelles lignes signalées sont de vrais cas où Prix de
> référence ≠ Prix de commande malgré une unité identique (ex: ligne 163 "pâté de foie
> supérieur", KG des deux côtés, 0,6667€ contre 0,95€). Garde-fou ajouté dans la même règle :
> si `resultat.couleur_m` est déjà posé (le moteur partagé a déjà détecté ce même cas via son
> propre canal "même unité, ratio ≠ 1, rien ne l'explique" — ex: lignes 244-259, calibres
> approximatifs comme "100g et plus"), la nouvelle règle ne rajoute rien, pour ne pas dupliquer
> le même constat deux fois sur la même cellule (trouvé via `test_sans_doublon_de_commentaire`,
> qui a immédiatement attrapé le problème).

> **Mise à jour (2026-11-09, encore plus tard le même jour)** : nouvelle règle symétrique
> (les deux moteurs) — quand l'Unité de référence et l'Unité de commande sont déjà identiques,
> que le Prix de référence est déjà égal au Prix de commande (rien à recalculer), et que le
> Conditionnement ne décrit qu'UNE seule unité de référence (ex: "SACHET DE 1 KG" ou "BRIQUE DE
> 1 L"), la suggestion "l'unité de commande semble plutôt être '<mot du conditionnement>'"
> n'est plus faite, dans AUCUN sens — un contenant d'une seule unité ne groupe rien, nommer le
> récipient n'apporte aucune information de plus que l'unité de mesure déjà écrite. Demandée
> explicitement par l'utilisateur ("vale pra os dois lados"), confirmée après lui avoir montré
> le cas réel "SACHET DE 1 KG" (groupement 1, lignes 183-191, ex: STEAK HACHE DE BOEUF) comme
> contre-exemple structurellement identique au cas qu'elle avait signalé ("BRIQUE DE 1 L",
> vu sur DADS MP ARDECHE 24 Lot 06, "Lait fermenté ... brique 1 lt") : elle a choisi d'appliquer
> la même règle aux deux, pas de garder un cas en exception. Nouveaux totaux (`audit_general.py`
> "lignes avec incohérence" / "cellules marquées", même ordre de fichiers) : 173/799, 293/878,
> 224/1024, 191/1003, 219/889, 748/1298, 193/721 — baisse partout sauf groupement 5 (inchangé,
> aucune ligne concernée). `audit_general_commentaires.py` (cellules marquées seulement) :
> 756, 827, 987, 955, 864, 955, 744. Vérifié que les lignes qui perdent leur signalement sont
> bien toutes des cas "contenant d'1 seule unité, prix déjà cohérent" (ex: groupement 1
> lignes 183-191, "SACHET DE 1 KG" avec Unité de commande déjà 'KG' et Prix identique des deux
> côtés) — pas un effet de bord sur un autre cas.

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
