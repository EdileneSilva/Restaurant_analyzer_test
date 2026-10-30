# Restaurant Analyzer — Audit et prétraitement des fichiers du marché alimentaire

Outil qui audite la cohérence unités/prix des fichiers d'achats (format "Base article
groupement"), et transforme automatiquement vers ce format des fichiers bruts de
fournisseurs qui ne sont pas déjà dans ce format.

## Scripts

- `app.py` — application Streamlit : dépôt de fichier(s), détection + transformation
  automatique si besoin, puis audit — au choix entre deux boutons/moteurs (voir
  `audit_pipeline.py`). Propose 3 téléchargements par fichier : le fichier formaté (juste après
  prétraitement, avant audit), et le fichier audité (dans le mode choisi). Rien n'est écrit sur
  le serveur, tout se fait en mémoire.
- `audit_pipeline.py` — chaîne complète détection + formatage + audit (US-08), réutilisable
  telle quelle par `app.py` ou par un futur agent (US-09) : n'importe pas `streamlit`, prend une
  liste de `(nom, contenu_bytes)` (pas un objet d'interface). Deux fonctions, un moteur chacune :
  `calcule_resultats` (`audit_general`, en mémoire) et `calcule_resultats_commentaires`
  (`audit_general_commentaires`, dans un sous-processus séparé via `audit_commentaires_worker.py`
  — les deux moteurs d'audit ne peuvent jamais être importés dans le même processus, voir plus
  bas et `catalogue_erreurs.md` n°41).
- `audit_general.py` — le script d'audit lui-même (logique de cohérence unités/prix), moteur
  "complet" (corrections mécaniques + surlignage) utilisé par `app.py` via `audit_pipeline.py`.
  Peut aussi être lancé en CLI (`python audit_general.py`) : lit tous les `.xlsx` de
  `pret_pour_audit/` et écrit les fichiers audités + rapports dans `transformed_files/`.
- `audit_general_commentaires.py` — variante de `audit_general.py` pour une personne qui ne
  travaille pas sur ce fichier au quotidien : toute décision de contenu n'est que surlignée
  avec un commentaire en langage courant, rien n'est changé (sauf une liste précise de
  corrections purement mécaniques — voir le docstring en tête du fichier). Utilisée par `app.py`
  (bouton "commentaires seulement") via `audit_pipeline.calcule_resultats_commentaires`, mais
  TOUJOURS dans un sous-processus séparé (`audit_commentaires_worker.py`), jamais importée
  directement dans le même processus que `audit_general`. Lancable aussi en CLI (`python
  audit_general_commentaires.py`), lit le même `pret_pour_audit/`, écrit dans `commented_files/`.
- `audit_commentaires_worker.py` — petit script CLI (`python audit_commentaires_worker.py
  <dossier_entree> <dossier_sortie>`), le sous-processus isolé ci-dessus : n'importe QUE
  `audit_general_commentaires`, jamais appelé directement par une personne.
- `test_non_regression_audit_general.py` / `test_non_regression_audit_commentaires.py` — tests
  `pytest` (TT-09) qui vérifient que les totaux (lignes avec incohérence, cellules marquées)
  produits par chacun des deux scripts ci-dessus, sur chacun des 7 fichiers de
  `pret_pour_audit/`, correspondent aux totaux de référence documentés dans
  `catalogue_erreurs.md`. **Toujours lancer séparément** (`pytest
  test_non_regression_audit_general.py` puis, dans une autre commande,
  `test_non_regression_audit_commentaires.py`) — jamais les deux dans la même invocation
  pytest : voir `catalogue_erreurs.md` n°41 pour pourquoi (les deux scripts partagent et
  écrasent au chargement les mêmes constantes du moteur commun).
- `pretraitement.py` — transforme les catalogues au format REGAL vers le format canonique.
- `pretraitement_grada.py` — transforme les catalogues de la famille GRADA (AGRICOURT, VDR,
  Boucherie) vers le format canonique.
- `pretraitement_auto.py` — point d'entrée unique utilisé par `app.py` : détecte
  automatiquement le format d'un fichier déposé (canonique / REGAL / GRADA / une des 5
  familles de `a_formater/` / `.xls` ancien à convertir / non reconnu) et applique le bon
  traitement. `detecte_statut()` expose le vocabulaire US-06 (`deja_au_format`/`formatable`/
  `non_formatable`) ; `transforme_automatiquement()` renvoie en plus une liste de causes
  structurées (`{code, message, onglet, colonne, cellule}`) quand le résultat est
  `non_formatable` — voir US-06 dans `suivi_user_stories_01_08.md` et
  `detection_format_causes.md`.
- `pretraitement_a_formater.py` — transforme les fichiers bruts de `a_formater/` qui ne sont
  pas déjà pris en charge par les deux scripts ci-dessus (familles BPU DQE, Passion Froid,
  Episaveurs, BPUR FRESCA, MP ARDECHE 24). Utilisable en CLI (`python
  pretraitement_a_formater.py`, écrit dans `a_formater_transforme/` — jamais directement dans
  `pret_pour_audit/`, pour ne pas fausser l'indice de prix partagé utilisé par
  `audit_general.construit_indice_prix()`) ou via `pretraitement_auto.py` (fichier par
  fichier, reconnu par nom de fichier exact — voir `trouve_entree_connue`).
- `causes_format.py` — `CauseFormatInvalide`, l'exception structurée (code + emplacement)
  utilisée par `pretraitement.py`, `pretraitement_grada.py`, `pretraitement_a_formater.py` et
  `pretraitement_auto.py` pour signaler un format invalide sans juste planter l'appelant
  (US-06.2/.3). Fichier séparé pour éviter un import circulaire entre ces scripts.
- `test_non_regression_pretraitement.py` — test `pytest` (US-07) qui vérifie, sur les 21
  fichiers connus de `a_formater/` (paramétré), que `pretraitement_auto.
  transforme_automatiquement()` ne modifie jamais le fichier d'origine, que le fichier produit
  est reconnu `deja_au_format`, et que son nombre de lignes correspond au total de référence.
  N'importe ni `audit_general` ni `audit_general_commentaires` : peut être lancé seul ou avec
  n'importe lequel des deux tests d'audit ci-dessus (voir US-07 dans `suivi_user_stories_01_08.md`).
- `first_scripts/first_attempt_v2_prototype.py` — règles de cohérence déjà validées,
  réutilisées telles quelles par `audit_general.py` (import direct, ne pas supprimer).

## Dossiers de données

- `a_inspecter/` — fichiers reçus du métier dont le statut n'est pas encore déterminé.
  C'est ici qu'arrivent les nouveaux fichiers, le temps de vérifier leur format. Ceux déjà
  au format canonique sont ensuite déplacés tels quels vers `pret_pour_audit/`.
- `a_formater/` — fichiers confirmés comme n'étant pas encore pris en charge
  automatiquement (nécessitent un nouveau script de transformation, ou une confirmation
  métier sur le mapping des colonnes avant de l'écrire — voir
  `mapping_colonnes_a_formater.xlsx`). Un script existe pour les 21 fichiers reçus à ce jour
  (`pretraitement_a_formater.py`), y compris les 12 fichiers "MP ARDECHE 24" (4 premiers
  mappés le 2026-10-29, 8 de plus le 2026-10-30 — même en-tête partagé, donc pas de nouvelle
  décision de mapping colonne par colonne, juste un vocabulaire Origine/Signes étendu à de
  nouveaux départements français et de nouveaux sigles de label — voir
  `mapping_colonnes_a_formater.xlsx`).
- `a_formater_transforme/` — sortie CLI de `pretraitement_a_formater.py`. Volontairement
  séparé de `pret_pour_audit/` : à relire fichier par fichier avant de déplacer à la main vers
  `pret_pour_audit/`.
- `pret_pour_audit/` — dossier d'entrée du mode CLI de `audit_general.py`
  (`DOSSIER_ENTREE`) : uniquement des fichiers déjà au format canonique, prêts à être
  audités directement (`python audit_general.py`).
- `transformed_files/` — sortie du mode CLI de `audit_general.py` (fichiers audités +
  rapports).
- `commented_files/` — sortie du mode CLI de `audit_general_commentaires.py` (fichiers
  commentés + rapports, même contenu de base que `transformed_files/` mais rédigé pour un
  public non technique).
- `useful_stuff/` — matériel de référence (pas des échantillons à auditer) : documents et
  modèles utilisés par l'ancienne référente du traitement manuel, pouvant être réutilisés.

Le détail fichier par fichier (source, date de réception, statut, anonymisation) est dans
[`inventaire_echantillons.csv`](inventaire_echantillons.csv).

## Documentation de référence

- [`format_cible.md`](format_cible.md) — schéma détaillé du format canonique, colonne par
  colonne (obligatoire/optionnel, valeurs autorisées).
- [`catalogue_erreurs.md`](catalogue_erreurs.md) — catalogue des règles de cohérence codées
  dans `audit_general.py`, avec un exemple réel par règle, et le jeu de résultats de référence
  (non-régression) sur les 7 fichiers de `pret_pour_audit/`.
- [`unites_commande_non_reconnues.md`](unites_commande_non_reconnues.md) — écart entre les
  unités tolérées par l'outil et celles réellement acceptées par la base de données
  d'intégration.
- [`mapping_colonnes_a_formater.md`](mapping_colonnes_a_formater.md) — version Markdown
  (suivie par git) de `mapping_colonnes_a_formater.xlsx` (non suivi : `.xlsx`, ignoré par
  `.gitignore`) : correspondance colonne par colonne entre le format canonique et chacun des
  10 formats fournisseurs de `a_formater/`, avec les décisions de mapping déjà tranchées en
  notes. À relire et mettre à jour manuellement si le `.xlsx` change.
- [`analyse_exploratoire.md`](analyse_exploratoire.md) — analyse ponctuelle (TT-03) de
  l'échantillon de fichiers reçus avant la mise en place du format canonique ; photo d'un
  instant donné, pas mise à jour au fil de l'eau.
- [`suivi_user_stories_01_08.md`](suivi_user_stories_01_08.md) — état d'avancement détaillé des
  US-01 à US-08 (critères d'acceptation un par un : fait / partiel / pas encore fait), avec
  renvoi vers le document ou la ligne de code concerné pour chacun.
- [`detection_format_causes.md`](detection_format_causes.md) — US-06 : vocabulaire de statut
  (`deja_au_format`/`formatable`/`non_formatable`) et catalogue de causes, à partir des échecs
  déjà codés dans les scripts de prétraitement — implémenté le 2026-10-30 (`causes_format.py`).
- [`test_manuel_streamlit.md`](test_manuel_streamlit.md) — guide de test MANUEL (US-08.3/US-05.3 :
  validation par le métier sur des fichiers réels, jamais tracée jusqu'ici) : user story +
  checklist pour déposer un échantillon de fichiers dans l'app et vérifier à l'œil que le
  formatage ET les deux modes d'audit sont cohérents avec les fichiers source.

## Lancer l'app

```
docker compose up -d --build
```
Puis ouvrir `http://localhost:8501`.
