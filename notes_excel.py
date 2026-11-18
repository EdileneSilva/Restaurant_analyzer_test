"""
Corrige deux limitations d'openpyxl sur la TAILLE des bulles de commentaire Excel (jamais la
police ni le texte lui-même — hors périmètre ici).

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
   calculé pour chaque bulle, avec la même formule (largeur de colonne / hauteur de ligne PAR
   DÉFAUT d'Excel, en pixels) que la bibliothèque XlsxWriter utilise pour ses propres bulles de
   commentaire (`xlsxwriter/worksheet.py`, `_position_object_pixels`/`_comment_vertices` ;
   `xlsxwriter/vml.py`, `_write_anchor`) — approximatif si le fichier a des largeurs de colonne
   personnalisées (non prises en compte ici, pour rester simple), mais largement suffisant pour
   que le texte ne soit plus coupé, contre une bulle fixe quelle que soit la taille demandée
   actuellement.

Utilisation (dans le script qui sauvegarde le classeur en dernier) :

    import notes_excel

    notes_excel.recalcule_toutes_les_notes(wb, _dimensions_commentaire)
    tampon = io.BytesIO()
    wb.save(tampon)
    contenu = notes_excel.corrige_ancrage_apres_sauvegarde(tampon.getvalue())
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET  # stdlib — évite d'ajouter lxml à requirements.txt/Docker

_VMLNS = "urn:schemas-microsoft-com:vml"
_EXCELNS = "urn:schemas-microsoft-com:office:excel"
_TYPE_BULLE_COMMENTAIRE = "#_x0000_t202"

# Mêmes valeurs par défaut qu'Excel pour une colonne/ligne non redimensionnée à la main
# (confirmé par le calcul de XlsxWriter : colonne = largeur par défaut 8.43 caractères ->
# int(8.43*7+0.5)+5 = 64px ; ligne = hauteur par défaut 15pt -> int(4/3*15) = 20px).
_LARGEUR_COLONNE_DEFAUT_PX = 64
_HAUTEUR_LIGNE_DEFAUT_PX = 20
_DECALAGE_X_DEFAUT_PX = 15
_DECALAGE_Y_DEFAUT_PX = 10

_RE_TAILLE_STYLE = re.compile(r"width:(\d+)px;height:(\d+)px")


def recalcule_toutes_les_notes(wb, dimensionne):
    """`dimensionne(texte) -> (largeur, hauteur)` : la fonction `_dimensions_commentaire` déjà
    utilisée par l'appelant. Passe sur TOUTES les cellules commentées du classeur, pas
    seulement celles que le script courant vient de poser/modifier — nécessaire parce
    qu'openpyxl ne conserve pas width/height d'un commentaire au rechargement (voir docstring
    du module)."""
    for ws in wb.worksheets:
        for ligne in ws.iter_rows():
            for cel in ligne:
                if cel.comment:
                    cel.comment.width, cel.comment.height = dimensionne(cel.comment.text)


def _vertices_ancrage(ligne0, colonne0, largeur_px, hauteur_px):
    """(col_debut, x1, ligne_debut, y1, col_fin, x2, ligne_fin, y2) — même ordre que l'élément
    `<x:Anchor>` (8 valeurs en pixels, décalages à l'intérieur de la cellule de début/fin)."""
    x1, y1 = _DECALAGE_X_DEFAUT_PX, _DECALAGE_Y_DEFAUT_PX

    col_fin, reste_largeur = colonne0, largeur_px + x1
    while reste_largeur >= _LARGEUR_COLONNE_DEFAUT_PX:
        reste_largeur -= _LARGEUR_COLONNE_DEFAUT_PX
        col_fin += 1

    ligne_fin, reste_hauteur = ligne0, hauteur_px + y1
    while reste_hauteur >= _HAUTEUR_LIGNE_DEFAUT_PX:
        reste_hauteur -= _HAUTEUR_LIGNE_DEFAUT_PX
        ligne_fin += 1

    return colonne0, x1, ligne0, y1, col_fin, reste_largeur, ligne_fin, reste_hauteur


def _injecte_ancrage_vml(donnees_vml):
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
        vertices = _vertices_ancrage(int(ligne_el.text), int(colonne_el.text), largeur_px, hauteur_px)

        ancrage = ET.SubElement(donnees_client, f"{{{_EXCELNS}}}Anchor")
        ancrage.text = ", ".join(str(v) for v in vertices)
        # Excel attend Anchor juste après SizeWithCells, avant AutoFill/Row/Column.
        taille_el = donnees_client.find(f"{{{_EXCELNS}}}SizeWithCells")
        donnees_client.remove(ancrage)
        donnees_client.insert(list(donnees_client).index(taille_el) + 1, ancrage)

    return ET.tostring(racine, xml_declaration=False)


def corrige_ancrage_apres_sauvegarde(contenu_xlsx):
    """contenu_xlsx : bytes d'un .xlsx déjà sauvegardé (par openpyxl, dernière sauvegarde du
    pipeline). Renvoie de nouveaux bytes, avec `<x:Anchor>` injecté dans chaque
    `xl/drawings/*.vml` (bulles de commentaire) — seule façon de faire respecter la taille des
    bulles par Excel desktop (voir docstring du module) ; openpyxl n'offre aucune API pour
    écrire cet élément directement. Si le classeur n'a aucun commentaire (donc aucun .vml),
    renvoie `contenu_xlsx` inchangé."""
    with zipfile.ZipFile(io.BytesIO(contenu_xlsx), "r") as zin:
        noms_vml = [n for n in zin.namelist() if n.startswith("xl/drawings/") and n.endswith(".vml")]
        if not noms_vml:
            return contenu_xlsx

        tampon = io.BytesIO()
        with zipfile.ZipFile(tampon, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                donnees = zin.read(item.filename)
                if item.filename in noms_vml:
                    donnees = _injecte_ancrage_vml(donnees)
                zout.writestr(item, donnees)
    return tampon.getvalue()
