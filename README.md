# Restaurant Analyzer — Audit et prétraitement des fichiers du marché alimentaire

Outil qui audite la cohérence unités/prix des fichiers d'achats (format "Base article
groupement"), et transforme automatiquement vers ce format des fichiers bruts de
fournisseurs qui ne sont pas déjà dans ce format.

## Scripts

- `app.py` — application Streamlit : dépôt de fichier(s), détection + transformation
  automatique si besoin, puis audit. Rien n'est écrit sur le serveur, tout se fait en
  mémoire.
- `audit_general.py` — le script d'audit lui-même (logique de cohérence unités/prix), celui
  utilisé par `app.py`. Peut aussi être lancé en CLI (`python audit_general.py`) : lit tous les
  `.xlsx` de `pret_pour_audit/` et écrit les fichiers audités + rapports dans
  `transformed_files/`.
- `audit_general_commentaires.py` — variante de `audit_general.py` pour une personne qui ne
  travaille pas sur ce fichier au quotidien : toute décision de contenu n'est que surlignée
  avec un commentaire en langage courant, rien n'est changé (sauf une liste précise de
  corrections purement mécaniques — voir le docstring en tête du fichier). N'est PAS importée
  par `app.py` — uniquement en CLI (`python audit_general_commentaires.py`), lit le même
  `pret_pour_audit/`, écrit dans `commented_files/`.
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
  `non_formatable` — voir US-06 dans `suivi_user_stories_01_06.md` et
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
- [`analyse_exploratoire.md`](analyse_exploratoire.md) — analyse ponctuelle (TT-03) de
  l'échantillon de fichiers reçus avant la mise en place du format canonique ; photo d'un
  instant donné, pas mise à jour au fil de l'eau.
- [`suivi_user_stories_01_06.md`](suivi_user_stories_01_06.md) — état d'avancement détaillé des
  US-01 à US-06 (critères d'acceptation un par un : fait / partiel / pas encore fait), avec
  renvoi vers le document ou la ligne de code concerné pour chacun.
- [`detection_format_causes.md`](detection_format_causes.md) — préparation d'US-06 : vocabulaire
  de statut (`deja_au_format`/`formatable`/`non_formatable`) et catalogue de causes proposés,
  à partir des échecs déjà codés dans les scripts de prétraitement. Spec, rien d'implémenté.

## Lancer l'app

```
docker compose up -d --build
```
Puis ouvrir `http://localhost:8501`.
