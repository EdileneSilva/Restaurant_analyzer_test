# Détection du format d'entrée — statuts et causes (US-06)

Écrit le 2026-10-29 comme spec pure (vocabulaire de statut + catalogue de causes, construits à
partir d'échecs **déjà codés**, pas inventés), **implémenté le 2026-10-30** tel que décrit
ci-dessous, avec l'accord explicite du métier pour toucher les scripts protégés
(`pretraitement.py`, `pretraitement_grada.py`, `pretraitement_auto.py`) que ça impliquait. Voir
[`suivi_user_stories_01_06.md`](suivi_user_stories_01_06.md) pour le statut des 3 critères.

## 1. Statuts (US-06, critère 1) — implémenté

| Valeur interne (`pretraitement_auto.detecte_format()`) | Statut externe (`detecte_statut()`) | Remarque |
|---|---|---|
| `'canonique'` | `deja_au_format` | |
| `'regal'` | `formatable` | `pretraitement.transforme_catalogue_regal`. |
| `'grada'` | `formatable` | `transforme_fichier_grada_auto`. |
| `'a_formater'` | `formatable` | BPU DQE, Passion Froid, Episaveurs, BPUR FRESCA, MP ARDECHE 24 — reconnus par nom de fichier exact via `pretraitement_a_formater.trouve_entree_connue` (voir limite ci-dessous). |
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
| `ENTETE_MODIFIE` | en-tête modifié | `_trouve_ligne_entete_bpu`, `pretraitement_a_formater.py` | `ws.title` | — | `"lignes 1 à N"` (zone cherchée) |
| `COLONNE_OBLIGATOIRE_ABSENTE` | colonne obligatoire absente | `transforme_catalogue_regal`, `pretraitement.py` | `ws.title` | **toutes** les colonnes manquantes, pas juste la 1ère | `"ligne 1"` (REGAL : en-tête toujours ligne 1) |
| `COLONNE_OBLIGATOIRE_ABSENTE` | colonne obligatoire absente | `transforme_bpu_dqe`, `pretraitement_a_formater.py` | `ws.title` | **toutes** les colonnes essentielles manquantes | `"ligne N"` (ligne d'en-tête réellement trouvée) |
| `FORMAT_NON_SUPPORTE` *(au-delà du minimum)* | format de fichier non supporté | `convertit_xls_en_xlsx`, `pretraitement_auto.py` | — | — | — (message LibreOffice brut, `stderr`) |
| `SECTION_INTROUVABLE` *(au-delà du minimum)* | — | `_trouve_ligne_tags`, `pretraitement_grada.py` | `ws.title` | — | `"lignes 1 à N"` (zone cherchée) |

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
  `suivi_user_stories_01_06.md`, section "Ce qui bloque le plus".
