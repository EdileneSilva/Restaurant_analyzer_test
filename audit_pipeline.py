"""
Chaîne complète détection + formatage + audit (US-08) : n'importe PAS streamlit, pour pouvoir
être réutilisée TELLE QUELLE par n'importe quel appelant — `app.py` (POC 2 Streamlit), un futur
agent (US-09), un script CLI, ou un test — sans wrapper ni adaptation (voir
`suivi_user_stories_01_08.md`, US-08.2 : avant ce module, la fonction équivalente vivait dans
`app.py`, qui exécute du code Streamlit au niveau module dès l'import — la réutiliser ailleurs
aurait donc exécuté cette interface au passage, en plus d'exiger des objets ayant la forme
précise de `st.file_uploader` (`.getvalue()`/`.name`)).

Deux moteurs d'audit, au choix de l'appelant :
- `calcule_resultats` : moteur `audit_general` (corrections mécaniques silencieuses + reste
  surligné/commenté), en mémoire dans CE processus.
- `calcule_resultats_commentaires` : moteur `audit_general_commentaires` (quasiment tout
  surligné/commenté, rien changé sauf la même petite liste mécanique), dans un SOUS-PROCESSUS
  séparé (`audit_commentaires_worker.py`) — jamais importé directement ici : `audit_general` et
  `audit_general_commentaires` partagent et s'écrasent mutuellement les mêmes constantes du
  moteur commun au chargement (voir `catalogue_erreurs.md` n°41) ; le seul moyen sûr de
  proposer les deux dans la même appli (qui a déjà `audit_general` chargé) est d'isoler le
  second dans son propre interpréteur Python.

Les deux prennent une liste de (nom, contenu_bytes) — un couple simple, pas un objet
d'interface — et renvoient une structure de données simple (dict/liste/bytes), sans aucun objet
openpyxl ni widget d'interface dedans. Les deux renvoient aussi `xlsx_transforme_bytes` par
fichier : le contenu juste après prétraitement, AVANT l'audit (quel que soit le moteur choisi
ensuite) — identique au contenu déposé quand le fichier était déjà au format canonique.
"""

import io
import json
import os
import subprocess
import sys
import tempfile

import audit_general as audit
import pretraitement_auto

_WORKER_COMMENTAIRES = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     "audit_commentaires_worker.py")


def _pretraite_fichiers(fichiers):
    """Partie commune aux deux moteurs d'audit : détection + formatage (pretraitement_auto),
    qui n'importe ni `audit_general` ni `audit_general_commentaires` et ne pose donc aucun
    risque de collision entre les deux. Renvoie (items, erreurs, notes_pretraitement) où
    `items` est une liste de (nom, contenu_bytes_pretraite)."""
    items = []
    erreurs = []
    notes_pretraitement = []

    for nom_fichier, contenu in fichiers:
        nom_base = nom_fichier.rsplit(".", 1)[0]

        try:
            tampon_sortie = io.BytesIO()
            formato, nb_lignes, causas = pretraitement_auto.transforme_automatiquement(
                io.BytesIO(contenu), tampon_sortie, nom_fichier=nom_base
            )
        except Exception as exc:
            erreurs.append(f"Erreur lors de la détection/du prétraitement de **{nom_fichier}** : {exc}")
            continue

        if formato is None:
            if causas:
                detail = " ; ".join(
                    f"[{c['code']}]" + (f" (onglet « {c['onglet']} »)" if c.get("onglet") else "")
                    + f" {c['message']}"
                    for c in causas
                )
                erreurs.append(f"**{nom_fichier}** : format non formatable — {detail}")
            else:
                erreurs.append(
                    f"**{nom_fichier}** : format non reconnu (ni REGAL, ni famille GRADA, ni "
                    "une famille connue de a_formater/, ni déjà au format attendu) — ignoré."
                )
            continue

        if nb_lignes is not None:
            contenu = tampon_sortie.getvalue()
            if formato != "canonique":
                notes_pretraitement.append(
                    f"**{nom_fichier}** reconnu comme format « {formato} » et transformé "
                    f"automatiquement ({nb_lignes} lignes) avant l'audit."
                )
            else:
                notes_pretraitement.append(
                    f"**{nom_fichier}** était en .xls (ancien format), converti en .xlsx "
                    "automatiquement avant l'audit."
                )
        # À ce stade, `contenu` est toujours le contenu APRÈS prétraitement (transformé, ou
        # identique au contenu déposé si rien n'avait besoin de l'être).
        items.append((nom_fichier, contenu))

    return items, erreurs, notes_pretraitement


def calcule_resultats(fichiers):
    """fichiers : liste de (nom, contenu_bytes). Renvoie un dict :
    {erreurs, notes_pretraitement, indice_prix, nb_fichiers, fichiers}, où chaque élément de
    "fichiers" est {nom, resultats, xlsx_bytes (après audit), xlsx_transforme_bytes (avant
    audit)}. Moteur `audit_general` (corrections mécaniques + surlignage/commentaires)."""
    items, erreurs, notes_pretraitement = _pretraite_fichiers(fichiers)

    classeurs_valeurs, classeurs, noms, contenus_transformes = [], [], [], []
    for nom_fichier, contenu in items:
        try:
            wb_valeurs = audit.openpyxl.load_workbook(io.BytesIO(contenu), data_only=True)
            wb = audit.openpyxl.load_workbook(io.BytesIO(contenu), data_only=False)
        except Exception as exc:
            erreurs.append(f"Impossible d'ouvrir **{nom_fichier}** (après prétraitement) : {exc}")
            continue
        classeurs_valeurs.append(wb_valeurs)
        classeurs.append(wb)
        noms.append(nom_fichier)
        contenus_transformes.append(contenu)

    if not classeurs:
        return {
            "erreurs": erreurs, "notes_pretraitement": notes_pretraitement,
            "indice_prix": None, "fichiers": [],
        }

    indice_prix = audit.construit_indice_prix_depuis_classeurs(classeurs_valeurs)

    fichiers_resultat = []
    for nom_fichier, wb_valeurs, wb, contenu_transforme in zip(
        noms, classeurs_valeurs, classeurs, contenus_transformes
    ):
        resultats = audit.audite_classeur(wb_valeurs, wb, indice_prix)
        tampon_xlsx = io.BytesIO()
        wb.save(tampon_xlsx)
        fichiers_resultat.append({
            "nom": nom_fichier,
            "resultats": resultats,
            "xlsx_bytes": tampon_xlsx.getvalue(),
            "xlsx_transforme_bytes": contenu_transforme,
        })

    return {
        "erreurs": erreurs,
        "notes_pretraitement": notes_pretraitement,
        "indice_prix": {"produit": len(indice_prix["produit"]), "categorie": len(indice_prix["categorie"])},
        "nb_fichiers": len(classeurs),
        "fichiers": fichiers_resultat,
    }


def calcule_resultats_commentaires(fichiers):
    """Même contrat que calcule_resultats (même structure de retour), mais moteur
    `audit_general_commentaires` — exécuté dans un sous-processus isolé (voir docstring de
    module). Une erreur du sous-processus (ex: fichier corrompu d'une façon qui fait planter
    l'audit) revient comme une entrée de `erreurs`, pas comme une exception."""
    items, erreurs, notes_pretraitement = _pretraite_fichiers(fichiers)
    if not items:
        return {
            "erreurs": erreurs, "notes_pretraitement": notes_pretraitement,
            "indice_prix": None, "fichiers": [],
        }

    with tempfile.TemporaryDirectory() as dossier_entree, \
         tempfile.TemporaryDirectory() as dossier_sortie:
        noms_temp = [f"{i}.xlsx" for i in range(len(items))]
        for (_, contenu), nom_temp in zip(items, noms_temp):
            with open(os.path.join(dossier_entree, nom_temp), "wb") as f:
                f.write(contenu)

        resultat_proc = subprocess.run(
            [sys.executable, _WORKER_COMMENTAIRES, dossier_entree, dossier_sortie],
            capture_output=True, text=True, timeout=300,
        )
        if resultat_proc.returncode != 0:
            erreurs.append(
                "Échec de l'audit « commentaires seulement » : "
                f"{resultat_proc.stderr.strip() or resultat_proc.stdout.strip()}"
            )
            return {
                "erreurs": erreurs, "notes_pretraitement": notes_pretraitement,
                "indice_prix": None, "fichiers": [],
            }

        with open(os.path.join(dossier_sortie, "_indice.json"), encoding="utf-8") as f:
            indice_prix = json.load(f)

        fichiers_resultat = []
        for (nom_fichier, contenu_transforme), nom_temp in zip(items, noms_temp):
            with open(os.path.join(dossier_sortie, nom_temp), "rb") as f:
                xlsx_bytes = f.read()
            with open(os.path.join(dossier_sortie, nom_temp + ".json"), encoding="utf-8") as f:
                resultats = json.load(f)
            fichiers_resultat.append({
                "nom": nom_fichier,
                "resultats": resultats,
                "xlsx_bytes": xlsx_bytes,
                "xlsx_transforme_bytes": contenu_transforme,
            })

    return {
        "erreurs": erreurs,
        "notes_pretraitement": notes_pretraitement,
        "indice_prix": indice_prix,
        "nb_fichiers": len(items),
        "fichiers": fichiers_resultat,
    }
