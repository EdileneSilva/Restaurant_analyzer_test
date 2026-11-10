# Format cible (TT-04)

Document de référence du format canonique ("Base article groupement" / "Feuille 1"),
basé sur l'analyse exploratoire (TT-03, `analyse_exploratoire.md`) et croisé avec les
règles déjà codées (`audit_general.py`, `pretraitement.py`). **À faire relire et valider
par le métier** (critère d'acceptation de l'US-02) avant de considérer ce document comme
définitif.

> **Mis à jour le 2026-10-26** pour refléter les règles effectivement codées à ce jour
> (`audit_general.py` et `audit_general_commentaires.py`) : classification obligatoire/optionnel
> par colonne alignée sur `signale_champs_obligatoires_optionnels`, vocabulaire de récipients
> très étendu pour Unité de référence et Unité de commande, tolérance de ratio à 3 % (et non
> plus 2 %, voir `catalogue_erreurs.md` règle n°30), et le remplacement automatique — devenu
> effectif, pas seulement détecté — des mots non reconnus par la base (règle n°33, uniquement
> dans `audit_general.py`).

>Les fichiers `Base IANOR AGAPE 16 ...` sont
> incomplets : colonne 6 absente, et texte légèrement différent sur les colonnes 4 et 5.
> **Action à prévoir** : la colonne 6 ("Nombre de Pièces dans conditionnement") doit être
> ajoutée/complétée dans ces fichiers avant qu'ils ne soient considérés réellement au format
> cible.

## Bloc commun (colonnes 1 à 6)

| # | Colonne | Type | Obligatoire | Valeurs autorisées | Remarque |
|---|---|---|---|---|---|
| 1 | Marché | Texte libre | Oui (100% rempli, TT-03 ; vide → orange, bloquant pour l'intégration) | — | Catégorie de marché (ex. "Epicerie") |
| 2 | Lot | Texte libre | Oui (100% rempli, TT-03 ; vide → orange, bloquant pour l'intégration) | — | Numéro de lot, zéro ajouté devant 1-9 (règle codée dans `audit_general.py`). Attention : deux lignes avec la même Désignation de l'article mais un Lot différent sont signalées (rose) — un lot distinct ne devrait pas décrire un produit avec un texte rigoureusement identique à un autre lot sans information distinctive supplémentaire |
| 3 | Désignation de l'article | Texte libre | Oui (100% rempli, TT-03 ; vide → orange, bloquant pour l'intégration) | — | |
| 4 | Unité de référence (BPU) | Choix limité | **Oui** — champ bloquant pour l'intégration si vide (`signale_champs_obligatoires_optionnels`, marqué en orange) | `KG`, `L`, `PC`, ou le nom d'un récipient intermédiaire quand le conditionnement décrit un double niveau d'emballage confirmé par le ratio des prix (ex: `BOUTEILLE`, `PACK`, `SACHET`, `CAGEOT`... — "CAS 3", voir `catalogue_erreurs.md`) | En-tête exact confirmé par `2026 09 22 Base article groupement 1.xlsx` : "Unité de référence\n(BPU)" — c'est le texte du code qui fait foi. Absent/différent dans les fichiers IANOR AGAPE, à corriger dans ces fichiers |
| 5 | Poids à l'unité de référence | Nombre | Non (rempli seulement si Unité de référence = PC, cf. règle TT-03) — rempli AUTOMATIQUEMENT par le script quand le conditionnement et/ou la désignation fournisseur/article le confirment sans ambiguïté (couleur bleue, sans commentaire) | — | Idem : en-tête confirmé par le fichier de référence, différent dans les fichiers IANOR AGAPE ("Poids de la PC..."), à corriger dans ces fichiers |
| 6 | Nombre de pièces dans conditionnement | Nombre | Non | — | **Confirmé présente dans le modèle de référence** (`2026 09 22 Base article groupement 1.xlsx`) — **absente** des 6 fichiers IANOR AGAPE, à ajouter dans ces fichiers avant de les considérer réellement au format cible. **Définition précisée** : nombre de pièces individuelles contenues dans UNE unité de référence (colonne 4) — indépendant de l'unité de commande. Exemple concret : "GOURDE POMME" (groupement 1, ligne 118), conditionnement "CARTON DE 8 PACKS DE 4 GOURDES DE 90 GR" — si Unité de référence = PACK (poids 0.36 kg = 4×90g), alors colonne 6 = 4 (gourdes par pack), PAS 8 (packs par carton, qui est juste le ratio implicite entre unité de référence et unité de commande, dérivable du prix, non stocké ici) |

## Bloc Fournisseur 1, 2, 3 (12 colonnes chacun, sauf Fournisseur 3 sans "Circuit court" — particularité déjà connue du modèle)

**Obligatoire/optionnel ci-dessous = ce que vérifie `signale_champs_obligatoires_optionnels`
(`audit_general.py`)** : un bloc Fournisseur N entièrement vide n'est jamais signalé (il n'a
simplement pas participé) ; dès qu'au moins une de ses 7 colonnes ci-dessous (nom,
Conditionnement, Désignation fournisseur, Référence, Prix de référence, Unité de commande,
Prix de commande) est renseignée, les champs obligatoires encore vides sont marqués en orange
("bloque l'intégration"), les optionnels encore vides en gris ("conseillé").

| Colonne (répétée x3) | Type | Obligatoire | Valeurs autorisées | Remarque |
|---|---|---|---|---|
| Fournisseur N (nom) | Texte libre | **Oui** si le bloc a au moins une donnée (F1 : 100% ; F2/F3 : non remplis dans l'échantillon de référence, TT-03) | — | |
| Conditionnement | Texte libre | **Optionnel** (conseillé) si le bloc a au moins une donnée | — | Format "N x quantité unité" ou avec nom de récipient (sac, seau, barquette, bac, coffret, plaquette, paquet, étui, bombe...) — vocabulaire complet dans `MOTS_CONTENANT` (`audit_general.py`), régulièrement étendu au fil des fichiers réels |
| Désignation fournisseur | Texte libre | **Optionnel** (conseillé) si le bloc a au moins une donnée | — | Peut rester vide par convention de fichier (ex: `MDC_IMPORT_MARCHE`, colonne entièrement vide) sans bloquer l'intégration |
| Référence produit | Texte libre | **Optionnel** (conseillé) si le bloc a au moins une donnée (variable selon fichier, TT-03) | — | |
| PRIX UNITE REFERENCE (BPU) KG/L/PC | Nombre (prix) | **Oui** si le bloc a au moins une donnée | — | Le script ne le recalcule JAMAIS automatiquement, même quand il est confiant dans son calcul : uniquement surligné en rose + commenté ("Prix de référence suggéré...") — une personne valide toujours elle-même un changement de prix |
| UNITE COMMANDE | Choix limité | **Oui** si le bloc a au moins une donnée | `KG`, `L`, `PC`, ou nom de récipient (`SAC`, `BARQUETTE`, `CAGEOT`, `COFFRET`, `PLAQUETTE`, `ETUI`, `BOMBE`...), ou `LOT` (voir remarque) | Quand le conditionnement confirme un lot de N pièces sans récipient nommé, l'unité de commande ne garde que `LOT` — le nombre N est reporté en préfixe du Conditionnement ("LOT DE N, ...") plutôt que répété dans l'unité elle-même |
| PRIX UNITE COMMANDE | Nombre (prix) | **Oui** si le bloc a au moins une donnée | — | Toujours considéré comme le prix réel par le script, jamais recalculé |
| Local (X si local) | Choix limité | Non | `X` ou vide | |
| Bio (X si bio) | Choix limité | Non | `X` ou vide | |
| Circuit court | Choix limité | Non — **absent du bloc Fournisseur 3** (particularité confirmée du modèle) | `X` ou vide | |
| Label | Texte libre | Non | — | Ex. observé : "LABEL ROUGE", "Le Porc Français VPF" |
| Origine | Choix limité + exceptions | Non | `France`, `Union Européenne`, `Hors Union Européenne`, `Monde`, ou un pays reconnu (liste `PAYS_ETRANGERS_CONNUS` dans `pretraitement.py` : Espagne, Italie, Belgique, Chine, Royaume-Uni, Chili, Pérou, Irlande, Égypte, Asie, Allemagne, Portugal, Pays-Bas, Maroc) | Toute autre valeur est remplacée par "France" + commentaire de vérification (règle déjà codée) |
| Indice I0 | Nombre | Non | — | |

## ⚠️ Unité de commande acceptée par l'outil ≠ unité reconnue par la base de données

L'outil tolère en lecture un vocabulaire de récipients volontairement très large pour Unité de
référence et Unité de commande (voir `MOTS_CONTENANT`), construit au fil des fichiers réels
rencontrés. **Ce n'est pas la même chose que la liste des unités que la base de données
d'intégration reconnaît réellement** : d'après l'échange avec l'équipe d'intégration, celle-ci
n'accepte qu'une liste fermée de codes courts (ex. `C` pour Carton, `PAQ` pour Paquet, `U` pour
Unité...), et toute valeur hors de cette liste devient `PIECE` par défaut lors de
l'intégration — silencieusement, sans erreur bloquante.

Une comparaison complète entre ce que les 7 fichiers de référence contiennent (Unité de
commande, Unité de référence, et les mots de récipient détectés dans le Conditionnement) et
cette liste fermée est tenue à jour dans
[`unites_commande_non_reconnues.md`](unites_commande_non_reconnues.md) — 24 valeurs
distinctes ne correspondent aujourd'hui à aucune unité reconnue (ex: `COLIS`, `CAGEOT`, `PU`,
`FILET`, `BQT`, `PACK`, `COFFRET`...), avec en face la colonne "Unité de remplacement" qui donne
l'équivalent officiel pour chacune.

**Mise à jour 2026-10-26 — ce mapping est maintenant appliqué, mais différemment selon le
script** (`catalogue_erreurs.md`, règle n°33) :
- `audit_general.py` **réécrit directement** le mot non reconnu par son équivalent officiel
  (ex. `COLIS → CARTON`, `CAGEOT → PLATEAU`), avec un commentaire expliquant pourquoi — cette
  version est celle utilisée pour préparer l'intégration en base, donc responsable du résultat
  final.
- `audit_general_commentaires.py` **ne le fait volontairement pas** : laisser la bonne unité
  est jugé être la responsabilité de l'équipe d'audit, pas de la personne qui remplit le
  fichier et qui lit cette version. Il reconnaît seulement l'exception inverse (le mot déjà
  écrit EST l'équivalent officiel attendu) pour ne pas signaler une incohérence qui n'en est
  pas une.

## Bloc Allergènes (14 colonnes, une par allergène)

| Colonne | Type | Obligatoire | Valeurs autorisées |
|---|---|---|---|
| Gluten, Crustacés, Oeufs, Poissons, Arachides, Soja, Lait, Fruits à coques, Céleri, Moutarde, Sésame, Sulfites, Lupins, Mollusques | Choix limité | Non (100% vides dans les 6 fichiers analysés, TT-03) | Marqueur si présent, vide sinon — **convention exacte à confirmer avec le métier** : l'en-tête de "Gluten" dit "o si oui, vide si non" (lettre `o`), différent de la convention `X` utilisée pour Local/Bio/Circuit court — vérifier si c'est voulu ou une incohérence historique du modèle |

## Divers (fin de ligne)

| Colonne | Type | Obligatoire | Valeurs autorisées | Remarque |
|---|---|---|---|---|
| TAUX TVA | Nombre (%) | Non observé rempli dans l'échantillon | — | |
| Code RNM / ITAVI | Texte libre | Non | — | |
| (colonne sans nom) | — | — | — | Colonne 61, en-tête vide dans le modèle (`""` dans `ENTETES`) — à clarifier avec le métier si elle a un usage |
| ID_AURA | Texte libre | Non | — | |

## Anomalies à verser au catalogue d'erreurs (TT-05)

- Ligne entièrement vide en fin de tableau dans `Base IANOR AGAPE 16 Epicerie.xlsx` (artefact Excel, zone utilisée dépassant les données réelles).
- En-tête sans nom détecté par pandas (`Unnamed: 57`) — correspond à la colonne 61 sans nom du modèle, voir ci-dessus.
- Cellules fusionnées présentes uniquement dans le format BPU DQE (`a_formater/`), jamais dans le format canonique — à anticiper si un script de transformation est écrit pour ce format.

## Points ouverts à valider avec le métier avant de clore le TT-04

1. Convention `o`/vide pour Gluten vs `X`/vide pour Local/Bio/Circuit court — incohérence voulue ou erreur historique du modèle ?
2. Colonne 61 sans nom — a-t-elle un usage ou peut-elle être supprimée du modèle ?

## Note technique (hors TT-04, à traiter séparément)

**Résolu le 2026-11-10** (voir `catalogue_erreurs.md`, règle n°55) : `pretraitement_auto.
detecte_format` considérait un fichier "canonique" dès qu'il trouvait les en-têtes
"désignation de l'article" + "unité de référence" en ligne 1-5, sans vérifier le nombre de
colonnes ni leur contenu exact — les fichiers `Base IANOR AGAPE 16 ...` passaient donc ce test
alors qu'il leur manque la colonne 6, sans que personne n'en soit informé. Un 4e statut
externe, `incomplet` (`canonique_incomplet` en interne), signale désormais ce cas — et
(demandé le jour même, un premier essai se contentant de signaler ayant été jugé insuffisant :
tous les fichiers canoniques doivent partager la même structure) `transforme_automatiquement`
**insère réellement** la colonne manquante (en-tête + cellules vides, rien à inventer comme
contenu) à la bonne position, décalant nativement tout le reste (`ws.insert_cols`) — la
lecture des colonnes se faisant par texte d'en-tête, pas par position, ce décalage ne perturbe
jamais l'audit. Le fichier produit est un vrai fichier canonique complet, identique en
structure aux autres.
