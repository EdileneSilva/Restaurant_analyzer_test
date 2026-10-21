# Restaurant Analyzer — Audit et prétraitement des fichiers du marché alimentaire

Outil qui audite la cohérence unités/prix des fichiers d'achats (format "Base article
groupement"), et transforme automatiquement vers ce format des fichiers bruts de
fournisseurs qui ne sont pas déjà dans ce format.

## Scripts

- `app.py` — application Streamlit : dépôt de fichier(s), détection + transformation
  automatique si besoin, puis audit. Rien n'est écrit sur le serveur, tout se fait en
  mémoire.
- `audit_geral.py` — le script d'audit lui-même (logique de cohérence unités/prix). Peut

  aussi être lancé en CLI (`python audit_geral.py`) : lit tous les `.xlsx` de
  `pret_pour_audit/` et écrit les fichiers audités + rapports dans `transformed_files/`.
- `pretraitement.py` — transforme les catalogues au format REGAL vers le format canonique.
- `pretraitement_grada.py` — transforme les catalogues de la famille GRADA (AGRICOURT, VDR,
  Boucherie) vers le format canonique.
- `pretraitement_auto.py` — point d'entrée unique utilisé par `app.py` : détecte
  automatiquement le format d'un fichier déposé (canonique / REGAL / GRADA / `.xls`
  ancien à convertir / non reconnu) et applique le bon traitement.
- `first_scripts/first_attempt_v2_prototype.py` — règles de cohérence déjà validées,
  réutilisées telles quelles par `audit_geral.py` (import direct, ne pas supprimer).

## Dossiers de données

- `a_inspecter/` — fichiers reçus du métier dont le statut n'est pas encore déterminé.
  C'est ici qu'arrivent les nouveaux fichiers, le temps de vérifier leur format. Ceux déjà
  au format canonique sont ensuite déplacés tels quels vers `pret_pour_audit/`.
- `a_formater/` — fichiers confirmés comme n'étant pas encore pris en charge
  automatiquement (nécessitent un nouveau script de transformation). Une fois le script
  écrit (ou via `pretraitement_auto.py`), le résultat transformé va dans
  `pret_pour_audit/`.
- `pret_pour_audit/` — dossier d'entrée du mode CLI de `audit_geral.py`
  (`DOSSIER_ENTREE`) : uniquement des fichiers déjà au format canonique, prêts à être
  audités directement (`python audit_geral.py`).
- `transformed_files/` — sortie du mode CLI de `audit_geral.py` (fichiers audités +
  rapports).
- `useful_stuff/` — matériel de référence (pas des échantillons à auditer) : documents et
  modèles utilisés par l'ancienne référente du traitement manuel, pouvant être réutilisés.

Le détail fichier par fichier (source, date de réception, statut, anonymisation) est dans
[`inventaire_echantillons.csv`](inventaire_echantillons.csv).

## Lancer l'app

```
docker compose up -d --build
```
Puis ouvrir `http://localhost:8501`.
