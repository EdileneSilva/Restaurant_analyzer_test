"""
Point d'entrée unique de prétraitement : détecte automatiquement le format
d'un fichier déposé (REGAL, famille GRADA, ou déjà au format canonique de
raw_data/) et le transforme en conséquence — pensé pour être appelé une
seule fois par fichier, aussi bien en CLI qu'depuis l'app Streamlit, juste
avant de passer le résultat à audit_geral.py.

Ne modifie ni pretraitement.py (REGAL) ni pretraitement_grada.py
(AGRICOURT/VDR/Boucherie) : réutilise leurs fonctions. La partie GRADA est
en plus généralisée ici en UNE SEULE fonction (transforme_fichier_grada_auto)
qui couvre les trois variantes déjà écrites séparément dans
pretraitement_grada.py, en repérant tout dynamiquement (tag, texte d'en-tête,
feuille) — ces trois fonctions d'origine restent telles quelles, inchangées,
au cas où il faille revenir dessus.
"""

import io
import os
import shutil
import subprocess
import tempfile

import openpyxl

import pretraitement as pt
import pretraitement_grada as pg

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
            raise ValueError(
                f"Échec de la conversion .xls -> .xlsx via LibreOffice : {resultat.stderr.strip()}"
            )
        with open(chemin_xlsx, "rb") as f:
            return io.BytesIO(f.read())


def _ligne_entete_canonique(ws):
    """Repère si une feuille a déjà le format canonique de raw_data/ (même
    logique qu'audit_geral.localise_feuille_et_colonnes, en version
    minimale) : une ligne (parmi les 5 premières) avec une cellule contenant
    « désignation de l'article » et une autre « unité de référence »."""
    for r in range(1, min(5, ws.max_row) + 1):
        textes = [pt._normalise(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        a_designation = any("designation de l'article" in t for t in textes)
        a_unite_ref = any("unite de reference" in t for t in textes)
        if a_designation and a_unite_ref:
            return r
    return None


def _est_regal(ws):
    textes_entete = {pt._normalise(ws.cell(row=1, column=c).value) for c in range(1, ws.max_column + 1)}
    return "designation produit" in textes_entete and "cond rendu dans magento" in textes_entete


def _feuille_grada(ws):
    try:
        pg._trouve_ligne_tags(ws)
        return True
    except ValueError:
        return False


def detecte_format(chemin_source):
    """Renvoie 'canonique', 'regal', 'grada' ou None (format non reconnu)."""
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    for ws in wb.worksheets:
        if _ligne_entete_canonique(ws):
            return "canonique"
    for ws in wb.worksheets:
        if _est_regal(ws):
            return "regal"
    for ws in wb.worksheets:
        if _feuille_grada(ws):
            return "grada"
    return None


def _choisit_feuille_grada(wb):
    """Choisit la feuille GRADA à utiliser. S'il y en a plusieurs (ex: VDR,
    un onglet par trimestre), prend la DERNIÈRE qui correspond — en supposant
    que les onglets sont dans l'ordre chronologique, comme observé dans VDR."""
    candidates = [ws for ws in wb.worksheets if _feuille_grada(ws)]
    if not candidates:
        raise ValueError("Aucune feuille au format GRADA (tag _REFERENCE introuvable).")
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
    conséquence. Renvoie (format_detecte, nb_lignes) — nb_lignes est None
    quand le fichier était déjà au format canonique (rien à transformer : le
    fichier déposé peut être utilisé tel quel pour l'audit) ou quand le
    format n'a pas été reconnu (rien n'a été écrit dans chemin_sortie).

    chemin_source accepte aussi bien un chemin sur disque qu'un fichier en
    mémoire (BytesIO) ; dans ce dernier cas, nom_fichier sert à nommer le
    fournisseur (impossible à déduire d'un BytesIO).

    Si le fichier déposé est un ancien .xls, il est d'abord converti en
    .xlsx (LibreOffice headless) avant toute détection de format — dans ce
    cas, même si le résultat est déjà au format canonique, chemin_sortie
    reçoit quand même le .xlsx converti (nb_lignes vaut alors 0, pas None,
    pour signaler à l'appelant qu'il doit relire depuis chemin_sortie)."""
    _reseek(chemin_source)
    nom_fournisseur = nom_fichier or (
        os.path.splitext(os.path.basename(chemin_source))[0]
        if isinstance(chemin_source, str) else "Fournisseur"
    )

    foi_convertido_de_xls = _est_ancien_xls(chemin_source)
    if foi_convertido_de_xls:
        chemin_source = convertit_xls_en_xlsx(chemin_source)

    format_detecte = detecte_format(chemin_source)

    if format_detecte == "canonique":
        if foi_convertido_de_xls:
            _reseek(chemin_source)
            conteudo = chemin_source.read()
            if hasattr(chemin_sortie, "write"):
                chemin_sortie.write(conteudo)
            else:
                with open(chemin_sortie, "wb") as f:
                    f.write(conteudo)
            return "canonique", 0
        return "canonique", None
    if format_detecte == "regal":
        _reseek(chemin_source)
        return "regal", pt.transforme_catalogue_regal(chemin_source, chemin_sortie)
    if format_detecte == "grada":
        _reseek(chemin_source)
        return "grada", transforme_fichier_grada_auto(chemin_source, chemin_sortie, nom_fournisseur)
    return None, None
