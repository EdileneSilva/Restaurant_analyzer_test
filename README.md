# Restaurant Analyzer — Audit et prétraitement des fichiers du marché alimentaire

Outil qui audite la cohérence unités/prix des fichiers d'achats (format "Base article
groupement"), et transforme automatiquement vers ce format des fichiers bruts de
fournisseurs qui ne sont pas déjà dans ce format.

## Scripts

- `app.py` — application Streamlit : dépôt de fichier(s), détection + transformation
  automatique si besoin, puis audit. Rien n'est écrit sur le serveur, tout se fait en
  mémoire.
- `audit_geral.py` — le script d'audit lui-même (logique de cohérence unités/prix). Peut
  aussi être lancé en CLI (`python audit_geral.py`) : lit tous les `.xlsx` de `raw_data/`
  et écrit les fichiers audités + rapports dans `transformed_files/`.
- `pretraitement.py` — transforme les catalogues au format REGAL vers le format canonique.
- `pretraitement_grada.py` — transforme les catalogues de la famille GRADA (AGRICOURT, VDR,
  Boucherie) vers le format canonique.
- `pretraitement_auto.py` — point d'entrée unique utilisé par `app.py` : détecte
  automatiquement le format d'un fichier déposé (canonique / REGAL / GRADA / `.xls`
  ancien à convertir / non reconnu) et applique le bon traitement.
- `first_scripts/first_attempt_v2_prototype.py` — règles de cohérence déjà validées,
  réutilisées telles quelles par `audit_geral.py` (import direct, ne pas supprimer).

## Dossiers de données

- `raw_data/`, `transformed_files/` — entrée/sortie du mode CLI de `audit_geral.py`.
  Vides par défaut ; à remplir uniquement si vous lancez le script directement en ligne
  de commande plutôt que via l'app Streamlit.
- `to_format/` — échantillons reçus du métier dont le format n'est pas encore pris en
  charge automatiquement (nécessitent un nouveau script de transformation).
- `to_inspect/` — échantillons reçus du métier dont le statut n'était pas encore déterminé
  au moment de la réception.
- `useful_stuff/` — matériel de référence (pas des échantillons à auditer) : documents et
  modèles utilisés par l'ancienne référente du traitement manuel, pouvant être réutilisés.

Le détail fichier par fichier (source, date de réception, statut, anonymisation) est dans
[`inventaire_echantillons.csv`](inventaire_echantillons.csv) — ne pas dupliquer cette
information ici, seulement la convention de chaque dossier.

## Lancer l'app

```
docker compose up -d --build
```
Puis ouvrir `http://localhost:8501`.
