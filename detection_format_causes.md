# Détection du format d'entrée — statuts et causes (US-06)

Écrit le 2026-10-29 comme spec pure (vocabulaire de statut + catalogue de causes, construits à
partir d'échecs **déjà codés**, pas inventés), **implémenté le 2026-10-30** tel que décrit
ci-dessous, avec l'accord explicite du métier pour toucher les scripts protégés
(`pretraitement.py`, `pretraitement_grada.py`, `pretraitement_auto.py`) que ça impliquait. Voir
[`suivi_user_stories_01_08.md`](suivi_user_stories_01_08.md) pour le statut des 3 critères.

## 1. Statuts (US-06, critère 1) — implémenté

| Valeur interne (`pretraitement_auto.detecte_format()`) | Statut externe (`detecte_statut()`) | Remarque |
|---|---|---|
| `'canonique'` | `deja_au_format` | |
| `'canonique_incomplet'` | `incomplet` | Ajouté le 2026-11-10 : en-têtes ancres trouvées ("désignation de l'article" + "unité de référence"), mais une colonne structurelle du bloc commun manque (pour l'instant, seule "Nombre de pièces dans conditionnement", famille "Base IANOR AGAPE 16") — voir section 4. |
| `'regal'` | `formatable` | `pretraitement.transforme_catalogue_regal`. |
| `'grada'` | `formatable` | `transforme_fichier_grada_auto`. |
| `'a_formater'` | `formatable` | BPU DQE, Passion Froid, Episaveurs, BPUR FRESCA, MP ARDECHE 24 — reconnus par nom de fichier exact via `pretraitement_a_formater.trouve_entree_connue` (voir limite ci-dessous). |
| `'bpu_webgerest_cd45'` | `formatable` | Ajouté le 2026-11-13 : famille "TR__BPU_pour_intégration_Webgerest_CD45" (13 fichiers reçus le 2026-11-12, un par fournisseur/lot). Contrairement aux 5 familles ci-dessus, **reconnue par CONTENU** (feuille "BPU - LOT N" + en-tête "FAMILLE DE PRODUITS"/"UNITE MINIMUM"/"UNITE DE VENTE", `pretraitement_a_formater.feuille_bpu_webgerest_cd45`), pas par nom de fichier : le nom change à chaque envoi (un fichier par fournisseur), aucun registre possible. Distinguée de l'autre famille "BPU DQE" déjà connue (`transforme_bpu_dqe`, feuille "Page de Garde" + en-têtes "Familles"/"PU HT") malgré un nom d'onglet qui peut lui aussi commencer par "BPU". Voir `transforme_bpu_dqe_webgerest_cd45` pour le détail des deux anomalies déjà rencontrées (inversion Prix/Unité de vente sur un fichier, multiplicateur nu sans nom d'unité sur deux autres) et le siglaire développé en toutes lettres sans encore de substitution officielle (permission métier en attente). |
| `None` | `non_formatable` | Format vraiment inconnu, ou famille de `a_formater/` jamais vue (fichier pas répertorié). |

`detecte_format()` garde son vocabulaire interne (utile pour savoir QUEL traitement appeler) ;
`detecte_statut()`, nouvelle fonction, l'enveloppe dans le vocabulaire externe ci-dessus.

**Limite assumée** : la reconnaissance de `a_formater/` est par nom de fichier exact, pas par
en-tête — 3 des 5 familles (Passion Froid, Episaveurs, BPUR FRESCA) lisent des colonnes à
position fixe, sans aucune validation de texte d'en-tête, donc inventer une détection de
contenu pour elles aurait été un risque non validé plutôt qu'une vraie amélioration. Un fichier
de ces familles jamais vu avant reste `non_formatable` tant qu'il n'est pas ajouté à un
registre (`FICHIERS_BPU_DQE`, `FICHIER_PASSION_FROID`, etc., dans `pretraitement_a_formater.py`).

## 2. Catalogue de causes (US-06, critères 2 et 3) — implémenté

Les 6 `raise ValueError(...)` identifiés le 2026-10-29 lèvent désormais tous une
`causes_format.CauseFormatInvalide` (nouveau fichier `causes_format.py`, importé par les 4
scripts concernés) — une `ValueError` qui porte en plus `code`/`onglet`/`colonne`/`cellule`,
récupérable via `.cause()` en `{code, message, onglet, colonne, cellule}` :

| Code | Catégorie demandée | Où (fonction / fichier) | Onglet | Colonne | Cellule |
|---|---|---|---|---|---|
| `ONGLET_MANQUANT` | onglet manquant | `_choisit_feuille_grada`, `pretraitement_auto.py` | — | — | — |
| `ONGLET_MANQUANT` | onglet manquant | `transforme_passion_froid`, `pretraitement_a_formater.py` | — | — | — (ni « SANS DOUBLONS » ni « GLOBAL » trouvés — voir note ci-dessous) |
| `ENTETE_MODIFIE` | en-tête modifié | `_trouve_ligne_entete_bpu`, `pretraitement_a_formater.py` | `ws.title` | — | `"lignes 1 à N"` (zone cherchée) |
| `COLONNE_OBLIGATOIRE_ABSENTE` | colonne obligatoire absente | `transforme_catalogue_regal`, `pretraitement.py` | `ws.title` | **toutes** les colonnes manquantes, pas juste la 1ère | `"ligne 1"` (REGAL : en-tête toujours ligne 1) |
| `COLONNE_OBLIGATOIRE_ABSENTE` | colonne obligatoire absente | `transforme_bpu_dqe`, `pretraitement_a_formater.py` | `ws.title` | **toutes** les colonnes essentielles manquantes | `"ligne N"` (ligne d'en-tête réellement trouvée) |
| `FORMAT_NON_SUPPORTE` *(au-delà du minimum)* | format de fichier non supporté | `convertit_xls_en_xlsx`, `pretraitement_auto.py` | — | — | — (message LibreOffice brut, `stderr`) |
| `SECTION_INTROUVABLE` *(au-delà du minimum)* | — | `_trouve_ligne_tags`, `pretraitement_grada.py` | `ws.title` | — | `"lignes 1 à N"` (zone cherchée) |
| `COLONNE_OPTIONNELLE_ABSENTE` *(au-delà du minimum, 2026-11-10)* | colonne optionnelle absente — **jamais bloquante** (voir section 4) | `transforme_automatiquement`, `pretraitement_auto.py` | — | `"Nombre de pièces dans conditionnement"` | — |

`transforme_passion_froid` (fichier "Passion froid - Tarif général...") cherche d'abord
l'onglet « SANS DOUBLONS » (déjà dédupliqué, préféré) puis, s'il est absent, se rabat sur
« GLOBAL » (même disposition de colonnes, vérifié le 2026-11-10) — demandé par l'utilisateur
pour que l'outil reste utilisable si un futur envoi de ce fichier ne contenait plus l'onglet
dédupliqué. `ONGLET_MANQUANT` n'est levée que si NI l'un NI l'autre n'est trouvé.

`pretraitement_auto.transforme_automatiquement()` attrape `CauseFormatInvalide` autour de
chaque appel de transformation (REGAL/GRADA/a_formater) et renvoie `(None, None, [cause])` au
lieu de laisser l'exception remonter jusqu'à l'appelant (`app.py`, qui affiche désormais ces
causes au lieu d'un message générique). Testé le 2026-10-30 en corrompant volontairement
l'en-tête d'un vrai fichier BPU DQE : la cause `ENTETE_MODIFIE` remonte proprement, avec la
zone de lignes cherchée en `cellule`.

**Les deux lacunas relevées après la 1ère implémentation (2026-10-30, résolues) :**
1. *"Cellule" toujours `None`.* Résolu là où une cellule/zone a un sens réel : `"ligne N"` pour
   les deux `COLONNE_OBLIGATOIRE_ABSENTE` (la ligne d'en-tête où chercher), `"lignes 1 à N"`
   pour `ENTETE_MODIFIE`/`SECTION_INTROUVABLE` (la zone balayée sans rien trouver). Reste
   `None`, **honnêtement**, pour `ONGLET_MANQUANT` (aucune feuille ne correspond — il n'y a pas
   de cellule à montrer sur un onglet qui n'existe pas) et `FORMAT_NON_SUPPORTE` (échec AVANT
   même l'ouverture du classeur par openpyxl — aucune cellule n'est accessible à ce stade).
   Remplir ces deux-là aurait voulu dire inventer une valeur sans rapport avec la vraie cause.
2. *Une seule cause par échec, jamais plusieurs simultanées.* Résolu pour le seul cas où
   plusieurs causes DE MÊME TYPE pouvaient coexister (plusieurs colonnes manquantes à la fois)
   — les deux `COLONNE_OBLIGATOIRE_ABSENTE` listent maintenant TOUTES les colonnes manquantes
   en une seule cause, pas juste la première trouvée (REGAL ne le faisait pas encore ; BPU DQE
   le faisait déjà depuis la 1ère implémentation). **Reste une limite assumée** : un fichier ne
   peut remonter qu'UNE SEULE cause par tentative de transformation (ex: pas à la fois
   `ONGLET_MANQUANT` ET `COLONNE_OBLIGATOIRE_ABSENTE` dans la même liste) — chaque famille
   (REGAL/GRADA/a_formater) est un chemin de code séparé et s'arrête à la première exception
   rencontrée ; détecter plusieurs problèmes de nature différente en une seule passe voudrait
   dire transformer chaque `raise` en une collecte progressive dans toute la chaîne de
   validation (changement d'architecture plus large, pas fait ici, risque de régression plus
   élevé pour un bénéfice marginal : dans la pratique, un fichier mal formé échoue presque
   toujours sur UN SEUL problème structurel à la fois).

## 3. Ce qui reste ouvert

- Les deux `COLONNE_OBLIGATOIRE_ABSENTE` sont restées propres à chaque famille (REGAL vs BPU
  DQE ont des colonnes différentes) plutôt que fusionnées en un seul format de message.
- `FORMAT_NON_SUPPORTE`/`SECTION_INTROUVABLE` ont été gardés (au-delà du minimum demandé).
- La détection de `a_formater/` reste par nom de fichier exact (voir §1, limite assumée) :
  n'importe quel NOUVEAU fichier (même d'une famille déjà couverte) reste `non_formatable`
  tant qu'il n'est pas ajouté à un registre dans `pretraitement_a_formater.py`. Les 8 fichiers
  arrivés le 2026-10-30 (Vivalya, SYSCO FRANCE, Agricourt, et des DADS/Passion Froid MP
  ARDECHE 24 supplémentaires) ont déjà été traités le même jour — voir
  `suivi_user_stories_01_08.md`, section "Ce qui bloque le plus".

## 4. Statut `incomplet` (`canonique_incomplet`) — ajouté le 2026-11-10

Avant cet ajout, un fichier canonique auquel il manquait une colonne structurelle (ex: les 6
fichiers `Base IANOR AGAPE 16 ...`, sans colonne "Nombre de pièces dans conditionnement") était
traité exactement comme `deja_au_format` — rien n'en informait la personne qui dépose le
fichier, alors que ce n'est techniquement pas un format complet. Demandé par l'utilisateur, qui
a choisi d'en faire un statut externe à part entière plutôt qu'une simple note.

- `canonique_incomplet` est renvoyé par `detecte_format()` quand les en-têtes ancres sont
  trouvées (`_ligne_entete_canonique`) mais que `_colonne_nb_pieces_absente()` confirme
  l'absence de la colonne "Nombre de pièces dans conditionnement" sur cette même ligne
  d'en-tête.
- **Mise à jour (2026-11-10, le jour même)** : `transforme_automatiquement()` ne se contente
  plus de signaler sans rien changer — l'utilisateur a demandé une vraie correction, pour que
  tous les fichiers canoniques partagent la même structure. Appelle désormais
  `complete_colonnes_canonique_incomplet()`, qui INSÈRE la colonne manquante (en-tête +
  cellules vides, via `ws.insert_cols` — décale nativement tout le contenu/style des colonnes
  suivantes) à la bonne position (juste après "Poids à l'unité de référence", repérée par
  texte d'en-tête). `nb_lignes` est donc un vrai compte (plus `None`/`0`), comme pour les
  autres familles transformées (REGAL/GRADA/a_formater) ; `causes` continue de porter une
  cause INFORMATIVE, `COLONNE_OPTIONNELLE_ABSENTE`, qui ne bloque jamais rien (contrairement
  aux autres causes du tableau ci-dessus, toujours liées à un `non_formatable`).
- `audit_pipeline._pretraite_fichiers()` affiche cette cause comme une note de transformation
  réelle (pas une erreur) : "reconnu comme format canonique, mais incomplet — corrigé
  automatiquement avant l'audit : ...".
- Aucun impact sur la lecture des autres colonnes, ni avant ni après l'insertion :
  `localise_feuille_et_colonnes` (`audit_general.py`) les repère par texte d'en-tête, jamais
  par position fixe — un décalage de colonnes (qu'il soit dû à l'absence d'origine, ou à
  l'insertion qui la corrige) ne perturbe donc jamais la lecture des autres champs.
- **Mise en forme de la colonne insérée (2026-11-10, même jour)** : `ws.insert_cols` ne donne
  aucun style à la colonne vide qu'il crée — signalé par l'utilisateur après relecture visuelle
  d'un fichier AGAPE transformé (police/fond/bordure ne correspondant pas au reste du document).
  `complete_colonnes_canonique_incomplet()` copie désormais explicitement le style de la colonne
  "Poids à l'unité de référence" (juste à gauche, non touchée par l'insertion) vers la nouvelle
  colonne, ligne par ligne — avec un repli sur la colonne encore à gauche pour la bordure
  spécifiquement, quand celle de "Poids" est elle-même dépourvue de bordure sur une ligne donnée
  (défaut déjà présent dans les fichiers AGAPE source sur cette colonne précise, pas introduit
  ici). Voir `catalogue_erreurs.md`, règle n°57.
