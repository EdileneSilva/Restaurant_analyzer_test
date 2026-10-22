"""
Application de démonstration (Streamlit) — Audit de cohérence unités/prix
==========================================================================

Dépose un ou plusieurs fichiers .xlsx — déjà au format canonique ("Base article
groupement N", "MDC_IMPORT_MARCHE"...), ou dans un format reconnu (REGAL, famille
GRADA, ancien .xls) qui sera converti automatiquement — lance l'audit
(exactement la même logique que audit_general.py), et permet de télécharger les
fichiers audités + les rapports — sans rien écrire sur le serveur, tout se fait
en mémoire.
"""

import io

import streamlit as st

import audit_general as audit
import pretraitement_auto

st.set_page_config(page_title="Audit BPU — démo", page_icon="📋", layout="wide")

st.title("📋 Audit de cohérence unités/prix")
st.caption(
    "Dépose un ou plusieurs fichiers .xlsx : déjà au format canonique "
    "(« Base article groupement N », « MDC_IMPORT_MARCHE »...), ou dans un format "
    "reconnu (REGAL, famille GRADA, ancien .xls) — converti automatiquement avant "
    "l'audit. Rien n'est écrit sur le serveur : tout se passe en mémoire, pour cette "
    "session seulement."
)

fichiers_deposes = st.file_uploader(
    "Fichiers à auditer",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

lancer = st.button("Lancer l'audit", type="primary", disabled=not fichiers_deposes)


def calcule_resultats(fichiers_deposes):
    """Fait tout le calcul UNE SEULE FOIS (lecture, indice de prix, audit, sérialisation en
    bytes) et renvoie une structure simple (listes/dicts/bytes) — aucun objet openpyxl ni
    aucun widget Streamlit dedans, pour pouvoir être gardée telle quelle dans
    st.session_state et réutilisée à chaque nouveau rendu (ex: après un clic sur un bouton
    de téléchargement, qui relance tout le script mais ne doit pas refaire tout le calcul ni
    faire disparaître les résultats déjà affichés)."""
    classeurs_valeurs = []
    classeurs = []
    noms = []
    erreurs = []
    notes_pretraitement = []

    for fichier in fichiers_deposes:
        contenu = fichier.getvalue()
        nom_base = fichier.name.rsplit(".", 1)[0]

        try:
            tampon_sortie = io.BytesIO()
            formato, nb_lignes = pretraitement_auto.transforme_automatiquement(
                io.BytesIO(contenu), tampon_sortie, nom_fichier=nom_base
            )
        except Exception as exc:
            erreurs.append(f"Erreur lors de la détection/du prétraitement de **{fichier.name}** : {exc}")
            continue

        if formato is None:
            erreurs.append(
                f"**{fichier.name}** : format non reconnu (ni REGAL, ni famille GRADA, ni déjà "
                "au format attendu) — ignoré."
            )
            continue

        if nb_lignes is not None:
            contenu = tampon_sortie.getvalue()
            if formato != "canonique":
                notes_pretraitement.append(
                    f"**{fichier.name}** reconnu comme format « {formato} » et transformé "
                    f"automatiquement ({nb_lignes} lignes) avant l'audit."
                )
            else:
                notes_pretraitement.append(
                    f"**{fichier.name}** était en .xls (ancien format), converti en .xlsx "
                    "automatiquement avant l'audit."
                )

        try:
            wb_valeurs = audit.openpyxl.load_workbook(io.BytesIO(contenu), data_only=True)
            wb = audit.openpyxl.load_workbook(io.BytesIO(contenu), data_only=False)
        except Exception as exc:
            erreurs.append(f"Impossible d'ouvrir **{fichier.name}** (après prétraitement) : {exc}")
            continue
        classeurs_valeurs.append(wb_valeurs)
        classeurs.append(wb)
        noms.append(fichier.name)

    if not classeurs:
        return {
            "erreurs": erreurs, "notes_pretraitement": notes_pretraitement,
            "indice_prix": None, "fichiers": [],
        }

    indice_prix = audit.construit_indice_prix_depuis_classeurs(classeurs_valeurs)

    fichiers_resultat = []
    for nom_fichier, wb_valeurs, wb in zip(noms, classeurs_valeurs, classeurs):
        resultats = audit.audite_classeur(wb_valeurs, wb, indice_prix)
        tampon_xlsx = io.BytesIO()
        wb.save(tampon_xlsx)
        fichiers_resultat.append({
            "nom": nom_fichier,
            "resultats": resultats,
            "xlsx_bytes": tampon_xlsx.getvalue(),
        })

    return {
        "erreurs": erreurs,
        "notes_pretraitement": notes_pretraitement,
        "indice_prix": {"produit": len(indice_prix["produit"]), "categorie": len(indice_prix["categorie"])},
        "nb_fichiers": len(classeurs),
        "fichiers": fichiers_resultat,
    }


if lancer and fichiers_deposes:
    with st.spinner("Audit en cours..."):
        st.session_state["audit_resultats"] = calcule_resultats(fichiers_deposes)

donnees = st.session_state.get("audit_resultats")

if donnees:
    for message in donnees["erreurs"]:
        st.error(message)

    for message in donnees["notes_pretraitement"]:
        st.info(message, icon="🔄")

    if donnees["fichiers"]:
        st.info(
            f"{donnees['indice_prix']['produit']} produits comparables (≥ 3 lignes) et "
            f"{donnees['indice_prix']['categorie']} catégories comparables (≥ 8 lignes) repérés "
            f"parmi les {donnees['nb_fichiers']} fichier(s) déposé(s)."
        )

        for fichier in donnees["fichiers"]:
            st.divider()
            st.subheader(fichier["nom"])

            if not fichier["resultats"]:
                st.warning("Aucune feuille reconnue dans ce fichier (format inattendu) — ignoré.")
                continue

            for resultat in fichier["resultats"]:
                stats = resultat["stats"]
                total_marques = (stats["marques"] + stats["lignes_dupliquees"]
                                  + stats["copies_possibles"] + stats["champs_obligatoires_optionnels"])

                with st.container(border=True):
                    st.markdown(
                        f"**Feuille « {resultat['nom']} » — "
                        f"{stats['lignes_avec_incoherence']} ligne(s) avec incohérence, "
                        f"{total_marques} cellule(s) marquée(s)**"
                    )
                    # "Lignes à supprimer" = toutes les lignes peintes en rouge (candidates à
                    # la suppression) : les doublons ET les lignes vides/résiduelles — pas
                    # seulement ces dernières.
                    lignes_a_supprimer = stats["lignes_dupliquees"] + stats["lignes_vides_a_supprimer"]
                    colonnes = st.columns(4)
                    colonnes[0].metric("Lignes avec incohérence", stats["lignes_avec_incoherence"])
                    colonnes[1].metric("Cellules marquées", total_marques)
                    colonnes[2].metric("Lignes à supprimer (en rouge)", lignes_a_supprimer)
                    colonnes[3].metric("Références en double (à comparer)", stats["copies_possibles"])

                    # Juste les titres + compteurs par défaut ; les détails (ligne par ligne,
                    # avec le commentaire) ne s'affichent que si on clique pour les ouvrir.
                    for cle, titre in audit.TITRES_CATEGORIES_RAPPORT.items():
                        elements = resultat["categories"][cle]
                        if not elements:
                            st.caption(f"{titre} (0)")
                            continue
                        with st.expander(f"{titre} ({len(elements)})"):
                            for ligne, designation, texte in elements:
                                desig_courte = (designation or "").strip()
                                st.markdown(f"**Ligne {ligne}** — {desig_courte}  \n{texte}")

            nom_base = fichier["nom"].rsplit(".", 1)[0]
            st.download_button(
                f"⬇️ Télécharger {nom_base} - audit.xlsx",
                data=fichier["xlsx_bytes"],
                file_name=f"{nom_base} - audit.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"xlsx-{fichier['nom']}",
            )

        st.divider()
        st.success(
            "Terminé. Aucun prix ni unité n'a été modifié dans les fichiers déposés — seules "
            "la colonne Lot et, quand elle était vide et confirmée de façon fiable, la colonne "
            "Poids à l'unité de référence ont été renseignées. Tout le reste n'est que "
            "surlignage + commentaire, à valider par une personne."
        )
    elif not donnees["erreurs"]:
        st.warning("Aucun fichier n'a pu être ouvert.")
