"""
Corrige trois limitations d'openpyxl sur l'APPARENCE des bulles de commentaire Excel : taille,
ancrage Excel desktop, et police.

1. openpyxl ne conserve PAS `comment.width`/`comment.height` quand il RELIT un fichier déjà
   sauvegardé : vérifié directement (sauvegarder une bulle avec width=320/height=150, la
   relire, retombe sur 144/79 — les valeurs par défaut). Dans le pipeline réel, pretraitement
   sauvegarde le classeur (bulles dimensionnées), puis audit_general / audit_general_commentaires
   RELIT ce même fichier avant de le ré-sauvegarder : à ce moment-là, toute bulle posée par
   pretraitement a DÉJÀ perdu sa taille, que le code de l'audit touche cette cellule ou non.
   `recalcule_toutes_les_notes()` recalcule la taille de TOUTES les bulles du classeur juste
   avant la sauvegarde finale, pas seulement celles que le script courant vient de modifier.

2. openpyxl écrit bien la taille voulue dans le style VML (`width:...px;height:...px`), que
   LibreOffice et Excel Online utilisent pour l'affichage — mais PAS Excel desktop, qui se base
   sur l'élément `<x:Anchor>` du VML pour la taille réelle de la bulle. openpyxl n'écrit jamais
   cet élément (vérifié dans `openpyxl.comments.shape_writer._shape_factory` : seuls
   MoveWithCells, SizeWithCells, AutoFill, Row, Column sont écrits) — faute de lui, Excel
   desktop retombe sur une taille par défaut quel que soit le style demandé.
   `corrige_ancrage_apres_sauvegarde()` relit le .xlsx déjà sauvegardé et injecte un `<x:Anchor>`
   calculé pour chaque bulle, avec la même formule (largeur de colonne / hauteur de ligne, en
   pixels) que la bibliothèque XlsxWriter utilise pour ses propres bulles de commentaire
   (`xlsxwriter/worksheet.py`, `_position_object_pixels`/`_comment_vertices` ;
   `xlsxwriter/vml.py`, `_write_anchor`). Essayé d'abord avec des largeurs de colonne/hauteurs
   de ligne PAR DÉFAUT uniformes (pour rester simple) — signalé par l'utilisatrice le
   2026-11-18 comme insuffisant : plusieurs colonnes réelles (ex. D/E/J/K/M/R du format
   canonique) ont des largeurs très différentes du défaut Excel (8.43 caractères), donc le
   nombre de colonnes réellement couvertes par l'ancre était faux, rendant la bulle bien plus
   grande ou plus petite que prévu selon la colonne. Corrigé en lisant les largeurs/hauteurs
   RÉELLES du classeur (`ws.column_dimensions`/`ws.row_dimensions`, sinon la valeur par défaut
   de la feuille, sinon le défaut Excel) au moment de `recalcule_toutes_les_notes()` (le seul
   endroit qui a encore accès à l'objet `Workbook` avant sa sauvegarde finale), puis en les
   passant à `corrige_ancrage_apres_sauvegarde()` pour que l'ancre de chaque bulle utilise les
   vraies dimensions de SA feuille.

3. openpyxl n'offre aucun moyen de choisir la police du TEXTE d'un commentaire (`Comment`
   n'écrit jamais de run enrichi `<r><rPr>...` dans `xl/comments/*.xml`, seulement un `<t>` nu
   — l'affichage dépend alors du réglage "police des commentaires" de l'Excel de la personne
   qui ouvre le fichier). Demandé par l'utilisatrice le 2026-11-18 : Calibri 8 pt, fixe, quel
   que soit le réglage de qui ouvre le fichier — et une bulle plus compacte en conséquence.
   `corrige_ancrage_apres_sauvegarde()` réécrit aussi chaque `xl/comments/*.xml` pour
   envelopper le texte dans un run avec cette police.

Utilisation (dans le script qui sauvegarde le classeur en dernier) :

    import notes_excel

    dimensions_feuilles = notes_excel.recalcule_toutes_les_notes(wb, _dimensions_commentaire)
    tampon = io.BytesIO()
    wb.save(tampon)
    contenu = notes_excel.corrige_ancrage_apres_sauvegarde(tampon.getvalue(), dimensions_feuilles)
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET  # stdlib — évite d'ajouter lxml à requirements.txt/Docker

_VMLNS = "urn:schemas-microsoft-com:vml"
_EXCELNS = "urn:schemas-microsoft-com:office:excel"
_SMLNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XMLNS_ESPACE = "{http://www.w3.org/XML/1998/namespace}space"
_TYPE_BULLE_COMMENTAIRE = "#_x0000_t202"

ET.register_namespace("", _SMLNS)  # réécrit `xl/comments/*.xml` sans préfixe "ns0:", comme Excel

# Mêmes valeurs par défaut qu'Excel pour une colonne/ligne non redimensionnée à la main
# (confirmé par le calcul de XlsxWriter : colonne = largeur par défaut 8.43 caractères ->
# int(8.43*7+0.5)+5 = 64px ; ligne = hauteur par défaut 15pt -> int(4/3*15) = 20px).
_LARGEUR_COLONNE_DEFAUT_PX = 64
_HAUTEUR_LIGNE_DEFAUT_PX = 20
_DECALAGE_X_DEFAUT_PX = 15
_DECALAGE_Y_DEFAUT_PX = 10

_RE_TAILLE_STYLE = re.compile(r"width:(\d+)px;height:(\d+)px")

# Police fixe des commentaires — demandé par l'utilisatrice le 2026-11-18 (au lieu du réglage
# par défaut d'Excel, variable d'un poste à l'autre). 8 pt jugé trop petit à l'usage, repassé à
# 9 pt le même jour (confirmé bon). Largeur FIXE, seule la hauteur varie selon la longueur du
# texte (demande explicite : voir dimensions_commentaire() plus bas).
#
# Taille de bulle : calibration revue une 2e fois le 2026-11-18 (même jour), l'utilisatrice
# ayant signalé dans Excel (PC Windows) que la hauteur était partout "au moins le double" de
# ce qu'il faudrait (exemple réel vérifié : `BPF - LOT 5`, ligne 11, commentaire de 337
# caractères — bulle à 162 px alors qu'il en faudrait ~81) — alors que la 1ère calibration
# (320 px / 58 caractères à 11 pt, confirmée correcte à l'époque pour l'ANCIENNE bulle, bien
# plus grande) donnait, mise à l'échelle pour 9 pt, une hauteur par ligne proche de celle
# utilisée ici. Les 2 constantes de hauteur sont donc désormais calées sur ce retour direct
# (division par ~2) plutôt que sur la mise à l'échelle théorique, qui s'est avérée insuffisante
# une fois testée dans le vrai Excel. Largeur réduite à 150 px (demande explicite), caractères/
# ligne recalculés pour la même largeur de caractère qu'avant (~4,5 px/caractère à 9 pt) —
# seule la hauteur était en cause, pas la largeur de caractère.
NOM_POLICE = "Calibri"
TAILLE_POLICE_PT = 9
# Largeur : +25% le 2026-11-18 (150 -> 188, règle n°75) jugée encore trop étroite le même jour
# -> +25% de plus (188 -> 235) — caractères/ligne recalculés dans la même proportion à chaque
# fois (même ~4,5 px/caractère à 9 pt).
_LARGEUR_BULLE_PX = 235
_CARACTERES_PAR_LIGNE = 51  # 235 px / ~4,5 px par caractère à 9 pt
# Hauteur par ligne/de base : voir règle n°74 (passée de 7 à 13 px/ligne sur une base physique).
# +10% le 2026-11-18 (même jour), demande explicite de l'utilisatrice.
_HAUTEUR_BASE_PX = 66
_HAUTEUR_PAR_LIGNE_PX = 14
_HAUTEUR_SURCHARGE_PX = 14


def dimensions_commentaire(texte):
    """Taille de bulle (largeur, hauteur) adaptée à la longueur du texte, pour la police/taille
    fixées ci-dessus (`NOM_POLICE`/`TAILLE_POLICE_PT`) — seule source de ces constantes, reprise
    par les `_dimensions_commentaire()` locaux de `pretraitement.py`, `audit_general.py` et
    `audit_general_commentaires.py`."""
    nb_lignes = sum(max(1, -(-len(p) // _CARACTERES_PAR_LIGNE)) for p in texte.split("\n"))
    hauteur = max(_HAUTEUR_BASE_PX, nb_lignes * _HAUTEUR_PAR_LIGNE_PX + _HAUTEUR_SURCHARGE_PX)
    return _LARGEUR_BULLE_PX, hauteur


def _largeur_colonne_px(ws, col_idx):
    """Largeur réelle (en pixels) de la colonne `col_idx` (1-indexée) de `ws` — largeur propre
    à la colonne si définie, sinon celle par défaut de la feuille, sinon le défaut Excel
    (8.43 caractères). Même formule que XlsxWriter (`_size_col`)."""
    from openpyxl.utils import get_column_letter
    dim = ws.column_dimensions.get(get_column_letter(col_idx))
    largeur = dim.width if dim and dim.width else (ws.sheet_format.defaultColWidth or 8.43)
    if largeur < 1:
        return int(largeur * 12 + 0.5)
    return int(largeur * 7 + 0.5) + 5


def _hauteur_ligne_px(ws, row_idx):
    """Hauteur réelle (en pixels) de la ligne `row_idx` (1-indexée) de `ws` — même logique que
    `_largeur_colonne_px`, pour les lignes."""
    dim = ws.row_dimensions.get(row_idx)
    hauteur_pt = dim.height if dim and dim.height else (ws.sheet_format.defaultRowHeight or 15)
    return int(4.0 / 3.0 * hauteur_pt)


def recalcule_toutes_les_notes(wb, dimensionne):
    """`dimensionne(texte) -> (largeur, hauteur)` : la fonction `_dimensions_commentaire` déjà
    utilisée par l'appelant. Passe sur TOUTES les cellules commentées du classeur, pas
    seulement celles que le script courant vient de poser/modifier — nécessaire parce
    qu'openpyxl ne conserve pas width/height d'un commentaire au rechargement (voir docstring
    du module). Renvoie aussi, par nom de feuille, les largeurs de colonnes et hauteurs de
    lignes réellement utilisées par ses commentaires — à passer tel quel à
    `corrige_ancrage_apres_sauvegarde()`, seul moyen pour l'ancre Excel desktop de refléter la
    vraie taille des colonnes/lignes du fichier plutôt qu'une taille par défaut uniforme (voir
    règle n°73 de `catalogue_erreurs.md`)."""
    dimensions_feuilles = {}
    for ws in wb.worksheets:
        colonnes, lignes = {}, {}
        for ligne in ws.iter_rows():
            for cel in ligne:
                if cel.comment:
                    cel.comment.width, cel.comment.height = dimensionne(cel.comment.text)
                    colonnes.setdefault(cel.column, _largeur_colonne_px(ws, cel.column))
                    lignes.setdefault(cel.row, _hauteur_ligne_px(ws, cel.row))
        if colonnes or lignes:
            dimensions_feuilles[ws.title] = {"colonnes": colonnes, "lignes": lignes}
    return dimensions_feuilles


def _vertices_ancrage(ligne0, colonne0, largeur_px, hauteur_px, largeurs_colonnes, hauteurs_lignes):
    """(col_debut, x1, ligne_debut, y1, col_fin, x2, ligne_fin, y2) — même ordre que l'élément
    `<x:Anchor>` (8 valeurs en pixels, décalages à l'intérieur de la cellule de début/fin).
    `largeurs_colonnes`/`hauteurs_lignes` : largeur/hauteur réelle (px) par index de colonne/
    ligne (0-indexé, comme `<x:Column>`/`<x:Row>`), au-delà desquels la valeur par défaut
    (`_LARGEUR_COLONNE_DEFAUT_PX`/`_HAUTEUR_LIGNE_DEFAUT_PX`) sert de repli."""
    x1, y1 = _DECALAGE_X_DEFAUT_PX, _DECALAGE_Y_DEFAUT_PX

    col_fin, reste_largeur = colonne0, largeur_px + x1
    while True:
        largeur_col = largeurs_colonnes.get(col_fin, _LARGEUR_COLONNE_DEFAUT_PX)
        if reste_largeur < largeur_col:
            break
        reste_largeur -= largeur_col
        col_fin += 1

    ligne_fin, reste_hauteur = ligne0, hauteur_px + y1
    while True:
        hauteur_ligne = hauteurs_lignes.get(ligne_fin, _HAUTEUR_LIGNE_DEFAUT_PX)
        if reste_hauteur < hauteur_ligne:
            break
        reste_hauteur -= hauteur_ligne
        ligne_fin += 1

    return colonne0, x1, ligne0, y1, col_fin, reste_largeur, ligne_fin, reste_hauteur


def _injecte_ancrage_vml(donnees_vml, largeurs_colonnes, hauteurs_lignes):
    racine = ET.fromstring(donnees_vml)
    for forme in racine.findall(f".//{{{_VMLNS}}}shape"):
        if forme.get("type") != _TYPE_BULLE_COMMENTAIRE:
            continue
        donnees_client = forme.find(f"{{{_EXCELNS}}}ClientData")
        if donnees_client is None or donnees_client.find(f"{{{_EXCELNS}}}Anchor") is not None:
            continue  # pas une bulle de commentaire, ou <x:Anchor> déjà présent

        style = forme.get("style") or ""
        correspondance = _RE_TAILLE_STYLE.search(style)
        ligne_el = donnees_client.find(f"{{{_EXCELNS}}}Row")
        colonne_el = donnees_client.find(f"{{{_EXCELNS}}}Column")
        if not correspondance or ligne_el is None or colonne_el is None:
            continue

        largeur_px, hauteur_px = int(correspondance.group(1)), int(correspondance.group(2))
        ligne0, colonne0 = int(ligne_el.text), int(colonne_el.text)
        # colonne0/ligne0 sont 0-indexées (comme <x:Column>/<x:Row>) ; largeurs_colonnes/
        # hauteurs_lignes (venus de recalcule_toutes_les_notes) sont indexées par cel.column/
        # cel.row d'openpyxl, 1-indexées — d'où le +1 ci-dessous.
        vertices = _vertices_ancrage(
            ligne0, colonne0, largeur_px, hauteur_px,
            {c - 1: px for c, px in largeurs_colonnes.items()},
            {r - 1: px for r, px in hauteurs_lignes.items()},
        )

        ancrage = ET.SubElement(donnees_client, f"{{{_EXCELNS}}}Anchor")
        ancrage.text = ", ".join(str(v) for v in vertices)
        # Excel attend Anchor juste après SizeWithCells, avant AutoFill/Row/Column.
        taille_el = donnees_client.find(f"{{{_EXCELNS}}}SizeWithCells")
        donnees_client.remove(ancrage)
        donnees_client.insert(list(donnees_client).index(taille_el) + 1, ancrage)

    return ET.tostring(racine, xml_declaration=False)


def _applique_police_comments_xml(donnees_xml):
    """Enveloppe le texte nu (`<t>`) de chaque commentaire dans un run (`<r><rPr>...`) avec la
    police fixée par `NOM_POLICE`/`TAILLE_POLICE_PT`, pour qu'elle ne dépende plus du réglage
    de l'Excel de qui ouvre le fichier (openpyxl n'écrit jamais ce run : voir docstring du
    module)."""
    racine = ET.fromstring(donnees_xml)
    for texte_el in racine.findall(f".//{{{_SMLNS}}}comment/{{{_SMLNS}}}text"):
        t_el = texte_el.find(f"{{{_SMLNS}}}t")
        if t_el is None:
            continue  # déjà en run(s), rien à faire
        texte = t_el.text or ""
        texte_el.remove(t_el)
        r_el = ET.SubElement(texte_el, f"{{{_SMLNS}}}r")
        rpr_el = ET.SubElement(r_el, f"{{{_SMLNS}}}rPr")
        ET.SubElement(rpr_el, f"{{{_SMLNS}}}sz").set("val", str(TAILLE_POLICE_PT))
        ET.SubElement(rpr_el, f"{{{_SMLNS}}}rFont").set("val", NOM_POLICE)
        ET.SubElement(rpr_el, f"{{{_SMLNS}}}family").set("val", "2")
        nouveau_t = ET.SubElement(r_el, f"{{{_SMLNS}}}t")
        nouveau_t.text = texte
        nouveau_t.set(_XMLNS_ESPACE, "preserve")
    return ET.tostring(racine, xml_declaration=False)


_RNS = "http://schemas.openxmlformats.org/package/2006/relationships"
_ROFFNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _vml_par_feuille(zin):
    """{nom_de_feuille: chemin_du_vmlDrawing} — suit la chaîne de relations OOXML (workbook.xml
    -> workbook.xml.rels -> sheetN.xml -> sheetN.xml.rels), seul moyen fiable de savoir à
    quelle feuille appartient tel `xl/drawings/*.vml` (l'ordre de numérotation des fichiers ne
    correspond pas forcément à l'ordre des feuilles, par ex. quand une feuille n'a pas de
    commentaire)."""
    racine_wb = ET.fromstring(zin.read("xl/workbook.xml"))
    nom_par_rid = {
        feuille.get("r:id") or feuille.get(f"{{{_ROFFNS}}}id"): feuille.get("name")
        for feuille in racine_wb.findall(f".//{{{_SMLNS}}}sheet")
    }

    racine_rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    chemin_sheet_par_rid = {
        rel.get("Id"): rel.get("Target").lstrip("/")
        for rel in racine_rels.findall(f"{{{_RNS}}}Relationship")
        if rel.get("Target", "").split("/")[-1].startswith("sheet")
    }

    resultat = {}
    for rid, nom_feuille in nom_par_rid.items():
        chemin_sheet = chemin_sheet_par_rid.get(rid)
        if not chemin_sheet:
            continue
        chemin_sheet = chemin_sheet if chemin_sheet.startswith("xl/") else f"xl/{chemin_sheet}"
        dossier, fichier = chemin_sheet.rsplit("/", 1)
        chemin_rels = f"{dossier}/_rels/{fichier}.rels"
        if chemin_rels not in zin.namelist():
            continue
        racine_sheet_rels = ET.fromstring(zin.read(chemin_rels))
        for rel in racine_sheet_rels.findall(f"{{{_RNS}}}Relationship"):
            if rel.get("Type", "").endswith("/vmlDrawing"):
                cible = rel.get("Target").lstrip("/")
                resultat[nom_feuille] = cible if cible.startswith("xl/") else f"xl/{cible}"
    return resultat


def corrige_ancrage_apres_sauvegarde(contenu_xlsx, dimensions_feuilles=None):
    """contenu_xlsx : bytes d'un .xlsx déjà sauvegardé (par openpyxl, dernière sauvegarde du
    pipeline). `dimensions_feuilles` : renvoyé par `recalcule_toutes_les_notes()` (largeurs de
    colonnes/hauteurs de lignes réelles, par feuille) — `None` ou `{}` revient aux valeurs par
    défaut d'Excel pour toutes les colonnes/lignes (comportement d'origine, approximatif si le
    fichier a des largeurs personnalisées). Renvoie de nouveaux bytes : `<x:Anchor>` injecté
    dans chaque `xl/drawings/*.vml` (taille de bulle respectée par Excel desktop, avec les
    vraies dimensions de colonnes/lignes de SA feuille) ET police fixée dans chaque
    `xl/comments/*.xml` (voir docstring du module) — openpyxl n'offre aucune API pour écrire ni
    l'un ni l'autre directement. Si le classeur n'a aucun commentaire (donc aucun .vml), renvoie
    `contenu_xlsx` inchangé."""
    dimensions_feuilles = dimensions_feuilles or {}
    with zipfile.ZipFile(io.BytesIO(contenu_xlsx), "r") as zin:
        noms_vml = [n for n in zin.namelist() if n.startswith("xl/drawings/") and n.endswith(".vml")]
        if not noms_vml:
            return contenu_xlsx
        noms_comments = [n for n in zin.namelist() if n.startswith("xl/comments/") and n.endswith(".xml")]

        try:
            vml_par_feuille = _vml_par_feuille(zin)
        except (KeyError, ET.ParseError):
            vml_par_feuille = {}  # structure inattendue : repli sur les valeurs par défaut
        dimensions_par_vml = {
            chemin_vml: dimensions_feuilles.get(nom_feuille, {})
            for nom_feuille, chemin_vml in vml_par_feuille.items()
        }

        tampon = io.BytesIO()
        with zipfile.ZipFile(tampon, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                donnees = zin.read(item.filename)
                if item.filename in noms_vml:
                    dims = dimensions_par_vml.get(item.filename, {})
                    donnees = _injecte_ancrage_vml(
                        donnees, dims.get("colonnes", {}), dims.get("lignes", {}))
                elif item.filename in noms_comments:
                    donnees = _applique_police_comments_xml(donnees)
                zout.writestr(item, donnees)
    return tampon.getvalue()
