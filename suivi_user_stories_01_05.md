# Suivi des User Stories US-01 à US-05

État d'avancement détaillé, critère d'acceptation par critère d'acceptation, par rapport au
backlog (`backlog_jira_import_v3.csv`, Epic "Assistant IA de contrôle et de préparation des
fichiers du marché alimentaire"). Légende : ✅ fait · ⚠️ partiel / à finir · ❌ pas fait.

Photo prise le **2026-10-26**. Comme pour `analyse_exploratoire.md`, ce document n'est pas
forcément remis à jour à chaque commit — à relire à chaque fois qu'on reprend une de ces US.

## US-01 — Recevoir et inventorier les fichiers réels

> En tant que développeuse, je veux recevoir un échantillon de fichiers réels (bruts, corrigés
> et déjà au bon format), afin de construire les scripts sur des cas concrets.

1. ✅ **Fichiers stockés dans un espace sécurisé avec inventaire (source, date, version du
   modèle).** Dépôt Gitea privé + [`inventaire_echantillons.csv`](inventaire_echantillons.csv),
   une ligne par fichier (dossier, source, date de réception, statut, anonymisation). Mis à
   jour le 2026-10-26 pour inclure les 7 fichiers de `pret_pour_audit/`, qui manquaient depuis
   leur ajout comme jeu de référence.
2. ✅ **L'échantillon couvre au moins un fichier déjà au format, un fichier à formater et un
   fichier hors format.** `pret_pour_audit/` + `a_inspecter/` (déjà au format), `a_formater/`
   (hors format, nouveaux formats fournisseurs BPU DQE et autres).
3. ⚠️ **Données personnelles identifiées et anonymisées.** Fait pour le fichier où un vrai
   problème avait été trouvé (6 e-mails + 2 téléphones nominatifs dans
   `useful_stuff/Copie de Copie de Demande...xlsx`, anonymisés le 2026-10-15). Les autres
   fichiers (dont les 7 de `pret_pour_audit/`, scannés le 2026-10-26) ne contiennent aucun
   e-mail ni numéro de téléphone détecté. **Limite à connaître** : le scan est un filtre
   automatique par motif (e-mail, téléphone français) sur toutes les cellules — il ne
   détecterait pas un nom de personne isolé dans du texte libre (ex. dans un commentaire ou une
   désignation). Pas de preuve d'un tel cas dans l'échantillon actuel, mais la méthode ne le
   garantit pas à 100 %.

## US-02 — Documenter le format cible et les erreurs types

> En tant que développeuse, je veux un document qui décrit le format attendu et les erreurs
> fréquentes, afin d'en déduire les règles d'audit et de formatage.

1. ✅ **Format cible : chaque colonne (nom, type, obligatoire, valeurs autorisées).**
   [`format_cible.md`](format_cible.md), mis à jour le 2026-10-26.
2. ✅ **Catalogue d'erreurs avec un exemple par erreur.** [`catalogue_erreurs.md`](catalogue_erreurs.md),
   40 règles/cas documentés à ce jour, chacun avec un exemple réel tiré des fichiers.
3. ❌ **Document relu et validé par le métier.** Pas encore fait — `format_cible.md` liste
   encore 2 points ouverts explicitement en attente d'arbitrage métier (convention `o`/vide
   pour Gluten vs `X`/vide pour Local/Bio, et l'usage de la colonne 61 sans nom). Tant que ces
   deux points ne sont pas tranchés, le document ne peut pas être considéré comme définitif.

## US-03 — Auditer un fichier au format cible

> En tant que chargé(e) de marché, je veux un rapport des lignes en erreur sur un fichier au
> format cible, afin de savoir précisément quoi faire corriger.

1. ✅ **Le script d'audit fonctionne seul, sans passer par le formatage.**
   `audit_general.py` / `audit_general_commentaires.py` lisent directement `pret_pour_audit/`
   (fichiers déjà canoniques) en CLI, sans aucune étape de transformation.
2. ⚠️ **Rapport listant, par anomalie : ligne, colonne, valeur trouvée, règle violée, gravité.**
   Partiellement : chaque anomalie est bien rattachée à une ligne et une désignation (le
   commentaire posé sur la cellule concernée identifie implicitement la colonne), et le texte
   du commentaire explique la règle et la valeur attendue en langage libre. La gravité est
   codée par la couleur (7 couleurs, `catalogue_erreurs.md` règle n°21) plutôt que par un champ
   explicite "bloquant/avertissement" associé à chaque ligne du rapport `.md`. Ce n'est donc
   pas encore une structure de données strictement exploitable champ par champ (utile surtout
   pour un futur agent, US-09+).
3. ❌ **Rapport disponible en JSON (réutilisable par l'agent) ET en version lisible.** Seule la
   version lisible existe (rapport `.md` + fichier `.xlsx` surligné/commenté). Aucune sortie
   JSON structurée n'est produite par `audit_general.py` ni `audit_general_commentaires.py` à
   ce jour — à construire avant de pouvoir brancher un agent dessus (dépendance pour US-09).
4. ❌ **Score global et statut (à corriger / quasi prêt) avec seuils configurables.** Pas
   implémenté. Ce qui existe : des compteurs bruts ("lignes avec incohérence", "cellules
   marquées", détail par catégorie) — ni score normalisé, ni statut à deux états, ni notion de
   seuil configurable. Dépendance directe pour US-13 ("Signaler un fichier quasi prêt").

## US-04 — Repérer visuellement les lignes à corriger

> En tant que chargé(e) de marché, je veux que les lignes en erreur soient surlignées dans
> l'Excel, afin de les relire directement dans le fichier.

1. ⚠️ **Une couleur par type d'anomalie, avec une légende dans le fichier.** La palette de
   couleurs existe et est stable (7 couleurs, voir `catalogue_erreurs.md` règle n°21), mais la
   légende n'existe aujourd'hui que dans ce document externe — **aucune feuille de légende
   n'est ajoutée dans le fichier `.xlsx` lui-même** (vérifié : les classeurs produits n'ont que
   la feuille de données, pas de second onglet "Légende"). À faire, mais nécessite de toucher
   `audit_general.py` — hors périmètre de cette mise à jour documentaire.
2. ✅ **Mise en forme d'origine conservée (cellule par cellule, pas de réécriture complète).**
   Confirmé : toutes les modifications passent par openpyxl, cellule par cellule, sur le
   classeur original chargé tel quel.
3. ✅ **Le surlignage se retire simplement dans Excel.** Remplissage de cellule standard
   (`PatternFill`) + commentaire standard — supprimables via les fonctions natives d'Excel
   ("Sans remplissage", suppression de commentaire), rien de propriétaire ou de verrouillé.

## US-05 — Tester l'audit dans Streamlit (POC 1)

> En tant que testeur, je veux déposer un fichier au format cible dans une app Streamlit et
> voir le résultat de l'audit, afin de valider les règles avec le métier avant d'aller plus
> loin.

1. ⚠️ **Le testeur dépose un fichier, voit le statut, le score et la liste des anomalies.**
   `app.py` permet bien le dépôt et affiche la liste des anomalies (groupées par catégorie,
   dépliables), mais "statut" et "score" au sens de l'US-03 (point 4 ci-dessus) n'existent pas
   encore — seuls des compteurs bruts sont affichés (lignes avec incohérence, cellules
   marquées...). Bloqué par la même dépendance que US-03.4.
2. ⚠️ **Il télécharge le fichier surligné et le rapport.** Seul le fichier `.xlsx` audité est
   proposé au téléchargement (`st.download_button`) ; le rapport en tant que fichier séparé
   n'est pas proposé au téléchargement dans l'app (son contenu est affiché en ligne via les
   listes dépliables, ce qui couvre l'usage mais pas littéralement "télécharger le rapport").
3. ❓ **Une session de test avec le métier a lieu et ses retours sont notés dans Jira.** Aucune
   trace dans ce dépôt d'une session de test POC1 déjà tenue avec le métier — probablement
   encore à organiser une fois les points 1 et 2 ci-dessus comblés.

## Ce qui bloque le plus, pour la suite

Les gaps des US-03/04/05 ci-dessus partagent presque tous la même racine : **pas de structure
de résultat normalisée** (JSON + score/statut + seuils). Tant que ça n'existe pas, impossible de
cocher honnêtement US-03.3/.4, US-04.1 (légende — moins directement, mais la légende devrait
elle aussi être généré à partir d'une liste de catégories centralisée, qui n'existe pas encore
sous forme de données), ni US-05.1. À traiter ensemble plutôt qu'un par un, dans une future
session qui touchera le code (hors périmètre de cette mise à jour, volontairement limitée aux
fichiers `.md`/`.csv` et à un nouveau script de tests).
