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

import streamlit as st

import audit_general as audit
import audit_pipeline

# Les clés de audit_general.TITRES_CATEGORIES_RAPPORT sont les mêmes dans
# audit_general_commentaires (vérifié), mais PAS tous les libellés : "poids_rempli" y décrit un
# comportement différent (suggestion commentée, rien n'est rempli, contrairement au mode
# "complet"). audit_general_commentaires n'est jamais importé ici (isolé dans un sous-processus,
# voir audit_pipeline.py) — ce libellé est donc recopié tel quel plutôt qu'importé.
TITRE_POIDS_COMMENTAIRES = "Poids à l'unité de référence — suggestion (rien n'est rempli)"

st.set_page_config(page_title="Audit BPU — démo", page_icon="📋", layout="wide")

# Charte graphique alignée sur ianord.fr (couleurs + police "Inter" relevées sur le site) —
# complète le thème défini dans .streamlit/config.toml, que Streamlit ne permet pas d'étendre
# davantage (polices/arrondis/badges) sans un peu de CSS.
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    h1, h2, h3 { color: #34415A; }

    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 10px !important;
        box-shadow: 0 1px 4px rgba(52, 65, 90, 0.08);
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 8px;
        font-weight: 600;
    }

    .mode-badge {
        display: inline-block;
        padding: 0.15rem 0.75rem;
        border-radius: 999px;
        font-size: 0.85rem;
        font-weight: 700;
        background: #F6F3EA;
        color: #8A6D00;
        border: 1px solid #D9A605;
    }
    .mode-badge.commentaires {
        background: #EEF1F4;
        color: #34415A;
        border-color: #34415A;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

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

st.caption(
    "Deux façons d'auditer, au choix : **« complet »** corrige directement une petite liste de "
    "détails mécaniques (Bio, Label, numéro de Lot, synonymes d'unité...) et surligne/commente "
    "tout le reste ; **« commentaires seulement »** ne corrige quasiment rien — tout est "
    "surligné/commenté en langage courant, à valider par une personne (pensé pour quelqu'un qui "
    "ne travaille pas avec ce fichier au quotidien)."
)
col_btn_complet, col_btn_commentaires = st.columns(2)
lancer_complet = col_btn_complet.button(
    "Lancer l'audit complet (avec corrections)", type="primary", disabled=not fichiers_deposes
)
lancer_commentaires = col_btn_commentaires.button(
    "Lancer l'audit (commentaires seulement)", disabled=not fichiers_deposes
)

if (lancer_complet or lancer_commentaires) and fichiers_deposes:
    with st.spinner("Audit en cours..."):
        # audit_pipeline.calcule_resultats* ne connaît pas l'objet UploadedFile de Streamlit
        # (US-08.2 : réutilisable tel quel par un futur agent) — juste des (nom, contenu).
        arquivos = [(f.name, f.getvalue()) for f in fichiers_deposes]
        if lancer_complet:
            st.session_state["audit_resultats"] = audit_pipeline.calcule_resultats(arquivos)
            st.session_state["audit_mode"] = "complet"
        else:
            st.session_state["audit_resultats"] = audit_pipeline.calcule_resultats_commentaires(arquivos)
            st.session_state["audit_mode"] = "commentaires"

donnees = st.session_state.get("audit_resultats")
mode = st.session_state.get("audit_mode")

if donnees:
    if mode:
        libelle_mode = ("complet (corrections + surlignage)" if mode == "complet"
                         else "commentaires seulement (rien n'est changé, sauf le mécanique)")
        classe_badge = "" if mode == "complet" else "commentaires"
        st.markdown(
            f"Mode d'audit utilisé pour ce résultat : "
            f'<span class="mode-badge {classe_badge}">{libelle_mode}</span>',
            unsafe_allow_html=True,
        )

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
            st.subheader(f"📄 {fichier['nom']}")

            if not fichier["resultats"]:
                st.warning("Aucune feuille reconnue dans ce fichier (format inattendu) — ignoré.")
                continue

            for resultat in fichier["resultats"]:
                stats = resultat["stats"]
                total_marques = (stats["marques"] + stats["lignes_dupliquees"]
                                  + stats["copies_possibles"] + stats["champs_obligatoires_optionnels"])

                with st.container(border=True):
                    icone_statut = "✅" if stats["lignes_avec_incoherence"] == 0 else "⚠️"
                    st.markdown(
                        f"{icone_statut} **Feuille « {resultat['nom']} » — "
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

                    # Marché/Lot (comble_marche_lot) : pas dans TITRES_CATEGORIES_RAPPORT, donc
                    # pas de détail ligne par ligne possible (pas de commentaire posé, juste une
                    # couleur) — affiché en légende plutôt qu'en repli déroulant.
                    st.caption(
                        f"Marché/Lot recopiés automatiquement (lignes voisines identiques) : "
                        f"{stats['marche_lot_rempli']} · à confirmer manuellement (surlignés) : "
                        f"{stats['marche_lot_signale']}"
                    )

                    # Doublons Lot/Fournisseur dus à un espace en trop ou une casse différente :
                    # déjà corrigés automatiquement par corrige_doublons_texte (audit_general*.py)
                    # avant d'arriver ici — affiché en légende, comme Marché/Lot ci-dessus. Le cas
                    # "même description, numéro de Lot différent" n'est lui que signalé (jamais
                    # corrigé seul), dans un repli déroulant séparé.
                    avertissements_lot = stats["doublons_texte_a_verifier"]
                    st.caption(
                        f"Doublons Lot/Fournisseur (espace ou casse) corrigés automatiquement : "
                        f"{stats['doublons_texte_corriges']}"
                        + ("" if avertissements_lot else " · rien d'autre à vérifier.")
                    )
                    if avertissements_lot:
                        with st.expander(
                            f"⚠️ Lots à vérifier manuellement (même description, numéro "
                            f"différent) ({len(avertissements_lot)})"
                        ):
                            for message in avertissements_lot:
                                st.markdown(f"- {message}")

                    # Juste les titres + compteurs par défaut ; les détails (ligne par ligne,
                    # avec le commentaire) ne s'affichent que si on clique pour les ouvrir.
                    titres_categories = dict(audit.TITRES_CATEGORIES_RAPPORT)
                    if mode == "commentaires":
                        titres_categories["poids_rempli"] = TITRE_POIDS_COMMENTAIRES
                    for cle, titre in titres_categories.items():
                        elements = resultat["categories"][cle]
                        if not elements:
                            st.caption(f"{titre} (0)")
                            continue
                        with st.expander(f"{titre} ({len(elements)})"):
                            for ligne, designation, texte in elements:
                                desig_courte = (designation or "").strip()
                                st.markdown(f"**Ligne {ligne}** — {desig_courte}  \n{texte}")

            nom_base = fichier["nom"].rsplit(".", 1)[0]
            col_transforme, col_audit = st.columns(2)
            if fichier.get("a_ete_transforme"):
                col_transforme.download_button(
                    f"⬇️ Télécharger {nom_base} - formaté (avant audit).xlsx",
                    data=fichier["xlsx_transforme_bytes"],
                    file_name=f"{nom_base} - formaté.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"xlsx-transforme-{fichier['nom']}",
                    help="Le fichier tel qu'il était juste avant l'audit (déjà au format cible, "
                         "mais sans le surlignage ni les commentaires de l'audit) — utile pour "
                         "vérifier le formatage indépendamment du résultat de l'audit.",
                )
            else:
                # Rien n'a été transformé (fichier déjà canonique) : proposer quand même ce
                # téléchargement redonnerait un contenu identique à celui déposé, ce qui prête
                # à confusion plutôt que d'aider — demandé le 2026-11-09.
                col_transforme.caption(
                    "Déjà au format canonique — rien n'a été transformé avant l'audit."
                )
            # Nom de fichier distinct selon le mode : les deux modes donnent des contenus
            # différents pour le même fichier déposé — un nom identique risquerait que l'un
            # écrase l'autre dans le dossier de téléchargement si la personne récupère les
            # deux (demandé le 2026-11-09).
            suffixe_mode = "commentaires" if mode == "commentaires" else "complet"
            col_audit.download_button(
                f"⬇️ Télécharger {nom_base} - audit ({suffixe_mode}).xlsx",
                data=fichier["xlsx_bytes"],
                file_name=f"{nom_base} - audit ({suffixe_mode}).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"xlsx-{fichier['nom']}",
            )

        st.divider()
        if mode == "commentaires":
            st.success(
                "Terminé. Aucun prix ni unité n'a été modifié dans les fichiers déposés — seule "
                "une petite liste de détails mécaniques (Bio, Label, numéro de Lot, synonymes "
                "d'unité...) a été corrigée directement. Même « Poids à l'unité de référence » "
                "n'est ici qu'une suggestion commentée, pas un remplissage silencieux. Tout le "
                "reste n'est que surlignage + commentaire, à valider par une personne."
            )
        else:
            st.success(
                "Terminé. Aucun prix ni unité n'a été modifié dans les fichiers déposés — seules "
                "la colonne Lot et, quand elle était vide et confirmée de façon fiable, la colonne "
                "Poids à l'unité de référence ont été renseignées. Tout le reste n'est que "
                "surlignage + commentaire, à valider par une personne."
            )
    elif not donnees["erreurs"]:
        st.warning("Aucun fichier n'a pu être ouvert.")
