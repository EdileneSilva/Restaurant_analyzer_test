# Format cible (TT-04)

Document de référence du format canonique ("Base article groupement" / "Feuille 1"),
basé sur l'analyse exploratoire (TT-03, `analyse_exploratoire.md`) et croisé avec les
règles déjà codées (`audit_geral.py`, `pretraitement.py`). **À faire relire et valider
par le métier** (critère d'acceptation de l'US-02) avant de considérer ce document comme
définitif.

>Les fichiers `Base IANOR AGAPE 16 ...` sont
> incomplets : colonne 6 absente, et texte légèrement différent sur les colonnes 4 et 5.
> **Action à prévoir** : la colonne 6 ("Nombre de Pièces dans conditionnement") doit être
> ajoutée/complétée dans ces fichiers avant qu'ils ne soient considérés réellement au format
> cible.

## Bloc commun (colonnes 1 à 6)

| # | Colonne | Type | Obligatoire | Valeurs autorisées | Remarque |
|---|---|---|---|---|---|
| 1 | Marché | Texte libre | Oui (100% rempli, TT-03) | — | Catégorie de marché (ex. "Epicerie") |
| 2 | Lot | Texte libre | Oui (100% rempli, TT-03) | — | Numéro de lot, zéro ajouté devant 1-9 (règle codée dans `audit_geral.py`) |
| 3 | Désignation de l'article | Texte libre | Oui (100% rempli, TT-03) | — | |
| 4 | Unité de référence (BPU) | Choix limité | Oui | `KG`, `L`, `PC` | En-tête exact confirmé par `2026 09 22 Base article groupement 1.xlsx` : "Unité de référence\n(BPU)" — c'est le texte du code qui fait foi. Absent/différent dans les fichiers IANOR AGAPE, à corriger dans ces fichiers |
| 5 | Poids à l'unité de référence | Nombre | Non (rempli seulement si Unité de référence = PC, cf. règle TT-03) | — | Idem : en-tête confirmé par le fichier de référence, différent dans les fichiers IANOR AGAPE ("Poids de la PC..."), à corriger dans ces fichiers |
| 6 | Nombre de pièces dans conditionnement | Nombre | Non | — | **Confirmé présente dans le modèle de référence** (`2026 09 22 Base article groupement 1.xlsx`) — **absente** des 6 fichiers IANOR AGAPE, à ajouter dans ces fichiers avant de les considérer réellement au format cible. **Définition précisée** : nombre de pièces individuelles contenues dans UNE unité de référence (colonne 4) — indépendant de l'unité de commande. Exemple concret : "GOURDE POMME" (groupement 1, ligne 118), conditionnement "CARTON DE 8 PACKS DE 4 GOURDES DE 90 GR" — si Unité de référence = PACK (poids 0.36 kg = 4×90g), alors colonne 6 = 4 (gourdes par pack), PAS 8 (packs par carton, qui est juste le ratio implicite entre unité de référence et unité de commande, dérivable du prix, non stocké ici) |

## Bloc Fournisseur 1, 2, 3 (12 colonnes chacun, sauf Fournisseur 3 sans "Circuit court" — particularité déjà connue du modèle)

| Colonne (répétée x3) | Type | Obligatoire | Valeurs autorisées | Remarque |
|---|---|---|---|---|
| Fournisseur N (nom) | Texte libre | F1 : oui (100%) ; F2/F3 : non (jamais remplis dans l'échantillon, TT-03) | — | |
| Conditionnement | Texte libre | Oui si Fournisseur N rempli | — | Format "N x quantité unité" ou avec nom de récipient (sac, seau, barquette...) — vocabulaire complet dans `MOTS_CONTENANT` (`audit_geral.py`) |
| Désignation fournisseur | Texte libre | Oui si Fournisseur N rempli | — | |
| Référence produit | Texte libre | Non (variable selon fichier, TT-03) | — | |
| PRIX UNITE REFERENCE (BPU) KG/L/PC | Nombre (prix) | Oui si Fournisseur N rempli | — | |
| UNITE COMMANDE | Choix limité | Oui si Fournisseur N rempli | `KG`, `L`, `PC`, ou nom de récipient (SAC, BARQUETTE...) | |
| PRIX UNITE COMMANDE | Nombre (prix) | Oui si Fournisseur N rempli | — | |
| Local (X si local) | Choix limité | Non | `X` ou vide | |
| Bio (X si bio) | Choix limité | Non | `X` ou vide | |
| Circuit court | Choix limité | Non — **absent du bloc Fournisseur 3** (particularité confirmée du modèle) | `X` ou vide | |
| Label | Texte libre | Non | — | Ex. observé : "LABEL ROUGE", "Le Porc Français VPF" |
| Origine | Choix limité + exceptions | Non | `France`, `Union Européenne`, `Hors Union Européenne`, `Monde`, ou un pays reconnu (liste `PAYS_ETRANGERS_CONNUS` dans `pretraitement.py` : Espagne, Italie, Belgique, Chine, Royaume-Uni, Chili, Pérou, Irlande, Égypte, Asie, Allemagne, Portugal, Pays-Bas, Maroc) | Toute autre valeur est remplacée par "France" + commentaire de vérification (règle déjà codée) |
| Indice I0 | Nombre | Non | — | |

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

`pretraitement_auto.detecte_format` considère un fichier "canonique" dès qu'il trouve les
en-têtes "désignation de l'article" + "unité de référence" en ligne 1-5, sans vérifier le
nombre de colonnes ni leur contenu exact. Les fichiers `Base IANOR AGAPE 16 ...` passent donc
ce test alors qu'il leur manque la colonne 6 — à corriger dans un ticket à part (détection
plus stricte, ou étape de complétion automatique de cette colonne).
