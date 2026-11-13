"""
Point d'entrée unique de prétraitement : détecte automatiquement le format
d'un fichier déposé (REGAL, famille GRADA, une des 5 familles de a_formater/,
ou déjà au format canonique de raw_data/) et le transforme en conséquence —
pensé pour être appelé une seule fois par fichier, aussi bien en CLI qu'depuis
l'app Streamlit, juste avant de passer le résultat à audit_general.py.

Ne modifie pas la LOGIQUE de pretraitement.py (REGAL), pretraitement_grada.py
(AGRICOURT/VDR/Boucherie) ni pretraitement_a_formater.py (BPU DQE/Passion
Froid/Episaveurs/BPUR FRESCA/MP ARDECHE 24) : réutilise leurs fonctions. La
partie GRADA est en plus généralisée ici en UNE SEULE fonction
(transforme_fichier_grada_auto) qui couvre les trois variantes déjà écrites
séparément dans pretraitement_grada.py, en repérant tout dynamiquement (tag,
texte d'en-tête, feuille) — ces trois fonctions d'origine restent telles
quelles, inchangées, au cas où il faille revenir dessus.

US-06 (voir detection_format_causes.md) : `detecte_format` garde son
vocabulaire interne ('canonique'/'regal'/'grada'/'a_formater'/None, utile
pour savoir QUEL traitement appeler) ; `detecte_statut` l'enveloppe dans le
vocabulaire externe demandé par le backlog (deja_au_format/formatable/
non_formatable). `transforme_automatiquement` renvoie en plus une liste de
causes structurées ({code, message, onglet, colonne, cellule}, voir
causes_format.CauseFormatInvalide) au lieu de laisser une exception de
format planter l'appelant.
"""

import io
import os
import shutil
import subprocess
import tempfile
from copy import copy

import openpyxl

import pretraitement as pt
import pretraitement_grada as pg
import pretraitement_a_formater as paf
from causes_format import CauseFormatInvalide

_SIGNATURE_OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _est_ancien_xls(chemin_source):
    """Détecte un .xls (binaire OLE2, Excel 97-2003) à partir de sa signature,
    indépendamment du nom de fichier (fiable même sur un BytesIO sans nom)."""
    _reseek(chemin_source)
    if hasattr(chemin_source, "read"):
        debut = chemin_source.read(8)
        _reseek(chemin_source)
    else:
        with open(chemin_source, "rb") as f:
            debut = f.read(8)
    return debut == _SIGNATURE_OLE2


def convertit_xls_en_xlsx(chemin_source):
    """Convertit un .xls ancien en .xlsx via LibreOffice en mode headless
    (préserve la mise en forme et les cellules fusionnées, contrairement à
    une conversion xlrd/openpyxl qui ne recopierait que les valeurs). Renvoie
    un BytesIO positionné au début, prêt à être relu par openpyxl."""
    _reseek(chemin_source)
    with tempfile.TemporaryDirectory() as dossier_tmp:
        chemin_xls = os.path.join(dossier_tmp, "source.xls")
        if hasattr(chemin_source, "read"):
            with open(chemin_xls, "wb") as f:
                f.write(chemin_source.read())
        else:
            shutil.copy(chemin_source, chemin_xls)

        resultat = subprocess.run(
            ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", dossier_tmp, chemin_xls],
            capture_output=True, text=True, timeout=120,
        )
        chemin_xlsx = os.path.join(dossier_tmp, "source.xlsx")
        if resultat.returncode != 0 or not os.path.exists(chemin_xlsx):
            raise CauseFormatInvalide(
                "FORMAT_NON_SUPPORTE",
                f"Échec de la conversion .xls -> .xlsx via LibreOffice : {resultat.stderr.strip()}",
            )
        with open(chemin_xlsx, "rb") as f:
            return io.BytesIO(f.read())


def _ligne_entete_canonique(ws):
    """Repère si une feuille a déjà le format canonique de raw_data/ (même
    logique qu'audit_general.localise_feuille_et_colonnes, en version
    minimale) : une ligne (parmi les 5 premières) avec une cellule contenant
    « désignation de l'article » et une autre « unité de référence »."""
    for r in range(1, min(5, ws.max_row) + 1):
        textes = [pt._normalise(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        a_designation = any("designation de l'article" in t for t in textes)
        # Même piège que localise_feuille_et_colonnes (audit_general.py), trouvé le 2026-11-09
        # en vérifiant les 6 fichiers "Base IANOR AGAPE 16" de a_inspecter/ : leur en-tête
        # "Poids de la PC (si Unité de référence = PC)" contient la phrase "unité de
        # référence" entre parenthèses — une recherche en simple substring matchait À TORT
        # cette colonne Poids (jamais la vraie colonne 4, dont l'en-tête, mal orthographié
        # "Unité Référence" sans "de", ne matchait justement PAS la recherche d'origine). Le
        # résultat final ("canonique") restait bon par coïncidence — ce n'était donc pas
        # observable sur ces 6 fichiers précis — mais la détection elle-même était fragile :
        # corrigé en excluant les cellules commençant par "poids", comme côté audit_general.
        a_unite_ref = any(
            t.startswith("unite") and "reference" in t
            for t in textes if not t.startswith("poids")
        )
        if a_designation and a_unite_ref:
            return r
    return None


def _colonne_nb_pieces_absente(ws, ligne_entete):
    """Vrai si la colonne "Nombre de pièces dans conditionnement" (bloc commun, voir
    format_cible.md) n'est pas trouvée sur la ligne d'en-tête déjà repérée par
    _ligne_entete_canonique. Connu comme manquant dans la famille "Base IANOR AGAPE 16"
    (confirmé sur les 6 fichiers de a_inspecter/ le 2026-11-10) — un fichier canonique sans
    cette colonne reste utilisable tel quel pour l'audit (la colonne est optionnelle côté
    contenu, voir format_cible.md ligne 32 ; et sa lecture se fait par texte d'en-tête, pas par
    position fixe, donc son absence ne décale aucune autre colonne), mais mérite un
    avertissement plutôt qu'un silence total."""
    textes = [pt._normalise(ws.cell(row=ligne_entete, column=c).value)
              for c in range(1, ws.max_column + 1)]
    return not any("nombre de" in t and "pieces" in t for t in textes)


def _colonne_insertion_nb_pieces(ws, ligne_entete):
    """Position où insérer la colonne "Nombre de pièces dans conditionnement" dans une feuille
    canonique qui ne l'a pas : juste après "Poids à l'unité de référence" (sa position dans
    format_cible.md, colonne 6, juste avant le 1er bloc Fournisseur) — repérée par texte
    d'en-tête, pas par position fixe, pour s'adapter à n'importe quel décalage déjà présent.
    Si aucune colonne "Poids..." n'est trouvée (cas non vu en pratique), se rabat juste après
    "Unité de référence"."""
    col_poids = col_unite_ref = None
    for c in range(1, ws.max_column + 1):
        texte = pt._normalise(ws.cell(row=ligne_entete, column=c).value)
        if texte.startswith("poids") and col_poids is None:
            col_poids = c
        elif texte.startswith("unite") and "reference" in texte and col_unite_ref is None:
            col_unite_ref = c
    return (col_poids + 1) if col_poids else (col_unite_ref + 1 if col_unite_ref else 6)


def complete_colonnes_canonique_incomplet(chemin_source, chemin_sortie):
    """Complète un fichier déjà "canonique" auquel il manque la colonne "Nombre de pièces dans
    conditionnement" (voir detecte_format, 'canonique_incomplet') : insère la colonne à la
    bonne position (juste après "Poids à l'unité de référence"), avec son en-tête, cellules
    vides (aucune donnée n'existe pour ce champ dans ces fichiers — rien à inventer, juste la
    structure attendue). `ws.insert_cols` décale nativement tout le contenu/style des colonnes
    suivantes, rien d'autre n'est modifié. Toutes les feuilles canoniques concernées de ce
    classeur sont complétées (une famille comme "Base IANOR AGAPE 16" n'a qu'une seule feuille
    de données, mais rien n'empêche plusieurs onglets dans un futur fichier similaire).
    Renvoie le nombre de lignes de données (Désignation non vide) sous l'en-tête, dans la
    PREMIÈRE feuille canonique trouvée. Demandé par l'utilisateur le 2026-11-10 : plutôt que de
    seulement signaler l'absence de cette colonne (voir 'incomplet' dans detecte_statut), pour
    que tous les fichiers canoniques partagent exactement la même structure."""
    wb = openpyxl.load_workbook(chemin_source)  # pas data_only : préserve formules/styles
    nb_lignes = 0
    premiere_feuille_traitee = False
    for ws in wb.worksheets:
        ligne_entete = _ligne_entete_canonique(ws)
        if not ligne_entete or not _colonne_nb_pieces_absente(ws, ligne_entete):
            continue
        col_designation = next(
            (c for c in range(1, ws.max_column + 1)
             if "designation de l'article" in pt._normalise(ws.cell(row=ligne_entete, column=c).value)),
            None)
        col_insertion = _colonne_insertion_nb_pieces(ws, ligne_entete)
        ws.insert_cols(col_insertion)
        ws.cell(row=ligne_entete, column=col_insertion).value = (
            "Nombre de pièces dans conditionnement")
        # ws.insert_cols ne copie aucun style dans la colonne vide créée (contrairement à ce que
        # dit la docstring ci-dessus, qui ne parlait QUE du contenu/style des colonnes SUIVANTES,
        # décalées) : trouvé le 2026-11-10 par l'utilisateur sur un fichier AGAPE traité (police,
        # fond et bordures manquants sur toute la colonne F). Vérifié sur le modèle de référence
        # ("2026 09 22 Base article groupement 1.xlsx", où cette colonne existe déjà) : son style
        # (en-tête ET données) est identique à celui de la colonne juste à gauche ("Poids à
        # l'unité de référence") sur chaque ligne — donc on le reproduit en copiant ce style
        # depuis la colonne immédiatement à gauche (non touchée par l'insertion), ligne par ligne.
        # Exception trouvée en validant sur les 6 fichiers AGAPE : sur ces fichiers précis, cette
        # colonne "Poids" a elle-même perdu sa bordure sur les lignes de données (contrairement à
        # toutes les colonnes autour, qui l'ont) — un défaut déjà présent dans le fichier source,
        # pas introduit ici. Copier aveuglément aurait reproduit ce défaut sur la nouvelle colonne
        # au lieu de l'habillage "comme le reste du document" qu'elle doit avoir : si la bordure de
        # la colonne modèle est entièrement absente sur une ligne donnée, on se rabat sur la
        # bordure de la colonne encore à gauche ("Unité de référence"), qui l'a de façon fiable.
        col_modele = col_insertion - 1
        col_secours = col_modele - 1 if col_modele > 1 else None
        for r in range(ligne_entete, ws.max_row + 1):
            cel_modele = ws.cell(row=r, column=col_modele)
            cel_neuve = ws.cell(row=r, column=col_insertion)
            cel_neuve.font = copy(cel_modele.font)
            cel_neuve.fill = copy(cel_modele.fill)
            cel_neuve.alignment = copy(cel_modele.alignment)
            cel_neuve.number_format = cel_modele.number_format
            bordure = cel_modele.border
            if col_secours and not any((
                    bordure.left.style, bordure.right.style,
                    bordure.top.style, bordure.bottom.style)):
                bordure = ws.cell(row=r, column=col_secours).border
            cel_neuve.border = copy(bordure)
        if not premiere_feuille_traitee:
            if col_designation and col_designation >= col_insertion:
                col_designation += 1  # décalée par l'insertion ci-dessus
            nb_lignes = sum(
                1 for r in range(ligne_entete + 1, ws.max_row + 1)
                if col_designation and ws.cell(row=r, column=col_designation).value)
            premiere_feuille_traitee = True
    wb.save(chemin_sortie)
    return nb_lignes


def _est_regal(ws):
    textes_entete = {pt._normalise(ws.cell(row=1, column=c).value) for c in range(1, ws.max_column + 1)}
    return "designation produit" in textes_entete and "cond rendu dans magento" in textes_entete


def _feuille_grada(ws):
    try:
        pg._trouve_ligne_tags(ws)
        return True
    except ValueError:
        return False


def detecte_format(chemin_source, nom_fichier=None):
    """Renvoie 'canonique', 'canonique_incomplet', 'regal', 'grada', 'a_formater',
    'bpu_webgerest_cd45' ou None (format non reconnu).

    'bpu_webgerest_cd45' : famille "TR__BPU_pour_intégration_Webgerest_CD45" (13 fichiers reçus
    le 2026-11-12, un par fournisseur/lot) — détectée par CONTENU (feuille "BPU - LOT N" +
    en-tête), PAS par nom de fichier comme les 5 familles de a_formater/ ci-dessous : ici, le
    nom change à chaque envoi (un fichier par fournisseur), pas de registre possible. Voir
    pretraitement_a_formater.feuille_bpu_webgerest_cd45/transforme_bpu_dqe_webgerest_cd45.

    'canonique_incomplet' : les deux en-têtes ancres sont bien là (Désignation de l'article +
    Unité de référence), mais une colonne du bloc commun attendue par format_cible.md manque
    structurellement (pour l'instant, seule "Nombre de pièces dans conditionnement" est
    connue comme manquante, famille "Base IANOR AGAPE 16") — le fichier reste utilisable tel
    quel pour l'audit (voir transforme_automatiquement), juste signalé différemment de
    'canonique' pour ne pas prétendre silencieusement qu'il est complet. Demandé par
    l'utilisateur le 2026-11-10.

    `nom_fichier` sert UNIQUEMENT à reconnaître une des 5 familles de a_formater/ (BPU DQE,
    Passion Froid, Episaveurs, BPUR FRESCA, MP ARDECHE 24) : contrairement à REGAL/GRADA/
    canonique, détectés par le contenu (texte d'en-tête), ces familles sont reconnues par nom
    de fichier exact (voir pretraitement_a_formater.trouve_entree_connue et son commentaire :
    3 des 5 n'ont aucune validation de colonnes par texte, lire par position fixe)."""
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    for ws in wb.worksheets:
        ligne_entete = _ligne_entete_canonique(ws)
        if ligne_entete:
            if _colonne_nb_pieces_absente(ws, ligne_entete):
                return "canonique_incomplet"
            return "canonique"
    for ws in wb.worksheets:
        if _est_regal(ws):
            return "regal"
    for ws in wb.worksheets:
        if _feuille_grada(ws):
            return "grada"
    for ws in wb.worksheets:
        if paf.feuille_bpu_webgerest_cd45(ws):
            return "bpu_webgerest_cd45"
    if nom_fichier and paf.trouve_entree_connue(nom_fichier):
        return "a_formater"
    return None


# Vocabulaire externe demandé par US-06, critère 1 (deja_au_format/formatable/non_formatable)
# — construit au-dessus de detecte_format, qui garde son vocabulaire interne plus détaillé
# (nécessaire pour savoir QUEL traitement appeler). 'regal'/'grada'/'a_formater' sont bien les
# trois valeurs regroupées sous 'formatable' que le backlog demande comme statut unique, pas
# trois statuts distincts exposés tels quels à l'extérieur. 'incomplet' est un 4e statut
# ajouté le 2026-11-10 (demandé par l'utilisateur) : un fichier canonique auquel il manque une
# colonne structurelle (voir detecte_format) reste utilisable pour l'audit, mais ce n'est plus
# honnête de le confondre avec 'deja_au_format' (qui sous-entend "complet").
_STATUT_EXTERNE = {
    "canonique": "deja_au_format",
    "canonique_incomplet": "incomplet",
    "regal": "formatable",
    "grada": "formatable",
    "a_formater": "formatable",
    "bpu_webgerest_cd45": "formatable",
    None: "non_formatable",
}


def detecte_statut(chemin_source, nom_fichier=None):
    """Vocabulaire US-06 (deja_au_format/formatable/non_formatable). Attention : ne fait que la
    détection rapide (comme detecte_format) — un fichier dit 'formatable' ici peut encore
    échouer plus loin pendant la transformation réelle (ex: en-tête partiellement modifié,
    repéré seulement à ce moment-là) ; voir transforme_automatiquement pour les causes dans ce
    cas, qui bascule alors le résultat final en non_formatable."""
    return _STATUT_EXTERNE[detecte_format(chemin_source, nom_fichier=nom_fichier)]


def _choisit_feuille_grada(wb):
    """Choisit la feuille GRADA à utiliser. S'il y en a plusieurs (ex: VDR,
    un onglet par trimestre), prend la DERNIÈRE qui correspond — en supposant
    que les onglets sont dans l'ordre chronologique, comme observé dans VDR."""
    candidates = [ws for ws in wb.worksheets if _feuille_grada(ws)]
    if not candidates:
        raise CauseFormatInvalide(
            "ONGLET_MANQUANT", "Aucune feuille au format GRADA (tag _REFERENCE introuvable)."
        )
    return candidates[-1]


def transforme_fichier_grada_auto(chemin_source, chemin_sortie, nom_fournisseur=None):
    """Généralise transforme_agricourt / transforme_vdr / transforme_boucherie
    de pretraitement_grada.py en une seule fonction : toutes les colonnes
    utiles sont repérées dynamiquement (tag, ou à défaut texte d'en-tête),
    donc ça couvre les trois variantes sans avoir besoin de savoir laquelle
    c'est. Les particularités de chacune restent gérées : Libellé/Désignation
    fournisseur par tag ou par texte selon ce qui existe, Local (oui/non)
    seulement si la colonne existe (AGRICOURT), prix par pièce lu dans les
    commentaires si l'unité de gestion le dit (VDR) et que ça s'y trouve."""
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    ws = _choisit_feuille_grada(wb)

    marche = (pg.extrait_titre_lot(ws.cell(row=2, column=1).value)
              or pg.extrait_titre_lot(ws.cell(row=2, column=2).value))

    ligne_tags = pg._trouve_ligne_tags(ws)
    tags = pg._tags_de_la_ligne(ws, ligne_tags)

    col_libelle = (tags.get("_DESIGNATION_CONSULTATION", [None])[-1]
                   or pg._trouve_colonne_texte(ws, ligne_tags, "libellé"))
    col_desig_fourn = tags.get("_DESIGNATION", [None])[-1] or tags.get("_NOM_PUBLIC", [None])[-1]
    col_reference = tags["_REFERENCE"][-1]
    col_prix = tags["_PRIX"][-1]
    col_ug = tags["_UG"][-1]
    col_cond = pg._trouve_colonne_texte(ws, ligne_tags, "conditionnement minimum de vente")
    col_signe = tags.get("_SIGNE_QUALITE", [None])[-1] or tags.get("_SIGNES", [None])[-1]
    col_origine = (pg._trouve_colonne_texte(ws, ligne_tags, "origine matière première")
                   or pg._trouve_colonne_texte(ws, ligne_tags, "origine"))
    col_local_oui_non = pg._trouve_colonne_texte(ws, ligne_tags, "oui ou non")
    col_commentaires = pg._trouve_colonne_texte(ws, ligne_tags, "commentaires")

    ligne_debut = ligne_tags + 1
    ligne_fin = pg._derniere_ligne_avec_reference(ws, ligne_debut, col_reference)
    lots = pg._lots_par_section(ws, ligne_debut, ligne_fin, col_libelle, col_reference)

    fournisseur = nom_fournisseur or ws.title.strip()

    lignes = []
    for r, section in lots.items():
        libelle = ws.cell(row=r, column=col_libelle).value
        desig_fourn = ws.cell(row=r, column=col_desig_fourn).value
        reference = ws.cell(row=r, column=col_reference).value
        prix = pg._nombre(ws.cell(row=r, column=col_prix).value)
        ug = ws.cell(row=r, column=col_ug).value
        cond = ws.cell(row=r, column=col_cond).value if col_cond else None
        signe = ws.cell(row=r, column=col_signe).value if col_signe else None
        origine_brute = ws.cell(row=r, column=col_origine).value if col_origine else None
        oui_non = ws.cell(row=r, column=col_local_oui_non).value if col_local_oui_non else None
        commentaires = ws.cell(row=r, column=col_commentaires).value if col_commentaires else None

        if not prix:
            continue

        ug_norm = pt._normalise(ug)
        prix_piece = pg.prix_depuis_commentaire(commentaires)
        if ug_norm in ("eur/kg", "kg", "€/kg"):
            unite_ref, unite_cmd, prix_cmd, commentaire_unite = pg._traite_prix_kg(prix, ug_norm, cond)
        elif ug_norm in ("piece", "pièce") and prix_piece is not None:
            unite_ref, unite_cmd, prix_cmd, commentaire_unite = "KG", "PIECE", prix_piece, None
        else:
            unite_ref, unite_cmd, prix_cmd, commentaire_unite = pg._traite_prix_exception(prix, ug, cond)

        local_x = "X" if pt._normalise(oui_non) == "oui" else None
        bio_x, label = pg._bio_et_label(signe)
        valeur_origine, _, commentaire_origine = pt.traite_origine(origine_brute)

        lignes.append(pg._ligne_vide_de_base(
            fournisseur=fournisseur, lot=pg._combine_lot(marche, section), designation=libelle,
            unite_ref=unite_ref, poids=None, commentaire_unite=commentaire_unite, cond=cond,
            desig_fourn=desig_fourn, ref=reference, prix_ref=prix, unite_cmd=unite_cmd,
            prix_cmd=prix_cmd, local=local_x, bio=bio_x, label=label, origine=valeur_origine,
            commentaire_origine=commentaire_origine, marche=None,
        ))

    lignes.sort(key=lambda l: (pt._normalise(l["lot"]), pt._normalise(l["designation"])))
    return pt.ecrit_classeur_canonique(lignes, chemin_sortie, "GRADA")


def _reseek(chemin):
    """chemin_source peut être un chemin (str) ou un fichier en mémoire
    (BytesIO, ex: dans l'app Streamlit) — dans ce dernier cas, il faut
    revenir au début avant chaque nouvelle lecture avec openpyxl."""
    if hasattr(chemin, "seek"):
        chemin.seek(0)


def transforme_automatiquement(chemin_source, chemin_sortie, nom_fichier=None):
    """Point d'entrée unique : détecte le format et transforme en
    conséquence. Renvoie (format_detecte, nb_lignes, causes) — nb_lignes est
    None quand le fichier était déjà au format canonique (rien à transformer :
    le fichier déposé peut être utilisé tel quel pour l'audit) ou quand le
    format n'a pas été reconnu (rien n'a été écrit dans chemin_sortie).
    `causes` (US-06.2/.3) est une liste de dicts {code, message, onglet,
    colonne, cellule}. Non vide quand format_detecte et nb_lignes valent
    tous les deux None (non_formatable) : soit le format n'a été reconnu par
    aucune des 4 familles, soit il a été reconnu mais la transformation a
    quand même échoué sur une cause structurée (en-tête partiellement
    modifié, colonne manquante...) au lieu de planter. Également non vide,
    mais de façon NON bloquante cette fois (format_detecte et nb_lignes
    restent renseignés normalement), quand format_detecte vaut
    'canonique_incomplet' : une seule cause informative, "COLONNE_OPTIONNELLE_
    ABSENTE", qui ne fait que signaler une colonne structurelle manquante
    mais utilisable telle quelle (voir detecte_format).

    chemin_source accepte aussi bien un chemin sur disque qu'un fichier en
    mémoire (BytesIO) ; dans ce dernier cas, nom_fichier sert à nommer le
    fournisseur (impossible à déduire d'un BytesIO) ET, pour les familles de
    a_formater/, à les reconnaître par nom de fichier (voir detecte_format).

    Si le fichier déposé est un ancien .xls, il est d'abord converti en
    .xlsx (LibreOffice headless) avant toute détection de format — dans ce
    cas, même si le résultat est déjà au format canonique, chemin_sortie
    reçoit quand même le .xlsx converti (nb_lignes vaut alors 0, pas None,
    pour signaler à l'appelant qu'il doit relire depuis chemin_sortie)."""
    _reseek(chemin_source)
    nom_fichier_registre = nom_fichier or (
        os.path.basename(chemin_source) if isinstance(chemin_source, str) else None
    )
    nom_fournisseur = nom_fichier or (
        os.path.splitext(os.path.basename(chemin_source))[0]
        if isinstance(chemin_source, str) else "Fournisseur"
    )

    foi_convertido_de_xls = _est_ancien_xls(chemin_source)
    if foi_convertido_de_xls:
        try:
            chemin_source = convertit_xls_en_xlsx(chemin_source)
        except CauseFormatInvalide as exc:
            return None, None, [exc.cause()]

    format_detecte = detecte_format(chemin_source, nom_fichier=nom_fichier_registre)

    if format_detecte == "canonique_incomplet":
        # Contrairement à 'canonique' (bloc suivant, simple passage tel quel) : insère pour de
        # bon la colonne manquante (voir complete_colonnes_canonique_incomplet), pour que TOUS
        # les fichiers canoniques partagent la même structure, plutôt que de seulement
        # signaler l'absence sans rien corriger — demandé par l'utilisateur le 2026-11-10,
        # après avoir d'abord essayé un simple signalement ('incomplet' dans detecte_statut,
        # toujours renvoyé pour que l'appelant sache qu'une correction a eu lieu).
        _reseek(chemin_source)
        nb_lignes = complete_colonnes_canonique_incomplet(chemin_source, chemin_sortie)
        return "canonique_incomplet", nb_lignes, [{
            "code": "COLONNE_OPTIONNELLE_ABSENTE",
            "message": (
                "Colonne « Nombre de pièces dans conditionnement » absente de l'en-tête — "
                "ajoutée automatiquement (vide, aucune donnée disponible pour ce champ) avant "
                "l'audit."
            ),
            "onglet": None, "colonne": "Nombre de pièces dans conditionnement", "cellule": None,
        }]
    if format_detecte == "canonique":
        if foi_convertido_de_xls:
            _reseek(chemin_source)
            conteudo = chemin_source.read()
            if hasattr(chemin_sortie, "write"):
                chemin_sortie.write(conteudo)
            else:
                with open(chemin_sortie, "wb") as f:
                    f.write(conteudo)
            return "canonique", 0, []
        return "canonique", None, []
    if format_detecte == "regal":
        _reseek(chemin_source)
        try:
            return "regal", pt.transforme_catalogue_regal(chemin_source, chemin_sortie), []
        except CauseFormatInvalide as exc:
            return None, None, [exc.cause()]
    if format_detecte == "grada":
        _reseek(chemin_source)
        try:
            return ("grada",
                    transforme_fichier_grada_auto(chemin_source, chemin_sortie, nom_fournisseur), [])
        except CauseFormatInvalide as exc:
            return None, None, [exc.cause()]
    if format_detecte == "a_formater":
        _reseek(chemin_source)
        cle_registre = paf.trouve_entree_connue(nom_fichier_registre)
        try:
            return "a_formater", paf.transforme_fichier_connu(chemin_source, chemin_sortie, cle_registre), []
        except CauseFormatInvalide as exc:
            return None, None, [exc.cause()]
    if format_detecte == "bpu_webgerest_cd45":
        _reseek(chemin_source)
        try:
            nb_lignes, _avertissements = paf.transforme_bpu_dqe_webgerest_cd45(
                chemin_source, chemin_sortie, nom_fichier_pour_fournisseur=nom_fichier_registre)
            return "bpu_webgerest_cd45", nb_lignes, []
        except CauseFormatInvalide as exc:
            return None, None, [exc.cause()]
    return None, None, [{
        "code": "FORMAT_NON_SUPPORTE",
        "message": (
            "Format non reconnu : ni canonique, ni REGAL, ni famille GRADA, ni une famille "
            "connue de a_formater/."
        ),
        "onglet": None, "colonne": None, "cellule": None,
    }]
