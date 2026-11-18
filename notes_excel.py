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
   calculé pour chaque bulle, avec la même formule (largeur de colonne / hauteur de ligne PAR
   DÉFAUT d'Excel, en pixels) que la bibliothèque XlsxWriter utilise pour ses propres bulles de
   commentaire (`xlsxwriter/worksheet.py`, `_position_object_pixels`/`_comment_vertices` ;
   `xlsxwriter/vml.py`, `_write_anchor`) — approximatif si le fichier a des largeurs de colonne
   personnalisées (non prises en compte ici, pour rester simple), mais largement suffisant pour
   que le texte ne soit plus coupé, contre une bulle fixe quelle que soit la taille demandée
   actuellement.

3. openpyxl n'offre aucun moyen de choisir la police du TEXTE d'un commentaire (`Comment`
   n'écrit jamais de run enrichi `<r><rPr>...` dans `xl/comments/*.xml`, seulement un `<t>` nu
   — l'affichage dépend alors du réglage "police des commentaires" de l'Excel de la personne
   qui ouvre le fichier). Demandé par l'utilisatrice le 2026-11-18 : Calibri 8 pt, fixe, quel
   que soit le réglage de qui ouvre le fichier — et une bulle plus compacte en conséquence.
   `corrige_ancrage_apres_sauvegarde()` réécrit aussi chaque `xl/comments/*.xml` pour
   envelopper le texte dans un run avec cette police.

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
# par défaut d'Excel, variable d'un poste à l'autre). Taille de bulle recalculée pour cette
# police/taille à partir d'un point de calibration réel, pas d'une estimation générique :
# l'utilisatrice a confirmé que l'ANCIENNE bulle (320 px, 58 caractères/ligne, voir règle n°60)
# s'affichait jusque-là en Calibri 11 — donc ~320/58 ≈ 5,52 px/caractère à 11 pt. Mis à
# l'échelle linéairement pour 8 pt : ~5,52 × 8/11 ≈ 4,01 px/caractère (PAS 7 px/caractère, la
# largeur du chiffre '0' utilisée ailleurs dans ce module pour l'ancrage Excel desktop — les
# lettres d'un texte normal sont en moyenne plus étroites qu'un chiffre). Proportion ajustée le
# 2026-11-18 (même jour) à la demande de l'utilisatrice : bulle plus étroite et plus haute
# (moins large, plus de lignes visibles) plutôt qu'une réduction uniforme — à valider
# visuellement dans Excel (elle peut redemander d'ajuster ces 4 constantes si besoin).
NOM_POLICE = "Calibri"
TAILLE_POLICE_PT = 8
_LARGEUR_BULLE_PX = 200
_CARACTERES_PAR_LIGNE = 50  # 200 px / ~4,01 px par caractère à 8 pt
_HAUTEUR_BASE_PX = 75
_HAUTEUR_PAR_LIGNE_PX = 16


def dimensions_commentaire(texte):
    """Taille de bulle (largeur, hauteur) adaptée à la longueur du texte, pour la police/taille
    fixées ci-dessus (`NOM_POLICE`/`TAILLE_POLICE_PT`) — seule source de ces 4 constantes,
    reprise par les `_dimensions_commentaire()` locaux de `pretraitement.py`,
    `audit_general.py` et `audit_general_commentaires.py`."""
    nb_lignes = sum(max(1, -(-len(p) // _CARACTERES_PAR_LIGNE)) for p in texte.split("\n"))
    hauteur = max(_HAUTEUR_BASE_PX, nb_lignes * _HAUTEUR_PAR_LIGNE_PX + 18)
    return _LARGEUR_BULLE_PX, hauteur


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


def corrige_ancrage_apres_sauvegarde(contenu_xlsx):
    """contenu_xlsx : bytes d'un .xlsx déjà sauvegardé (par openpyxl, dernière sauvegarde du
    pipeline). Renvoie de nouveaux bytes : `<x:Anchor>` injecté dans chaque `xl/drawings/*.vml`
    (taille de bulle respectée par Excel desktop) ET police fixée dans chaque
    `xl/comments/*.xml` (voir docstring du module) — openpyxl n'offre aucune API pour écrire ni
    l'un ni l'autre directement. Si le classeur n'a aucun commentaire (donc aucun .vml), renvoie
    `contenu_xlsx` inchangé."""
    with zipfile.ZipFile(io.BytesIO(contenu_xlsx), "r") as zin:
        noms_vml = [n for n in zin.namelist() if n.startswith("xl/drawings/") and n.endswith(".vml")]
        if not noms_vml:
            return contenu_xlsx
        noms_comments = [n for n in zin.namelist() if n.startswith("xl/comments/") and n.endswith(".xml")]

        tampon = io.BytesIO()
        with zipfile.ZipFile(tampon, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                donnees = zin.read(item.filename)
                if item.filename in noms_vml:
                    donnees = _injecte_ancrage_vml(donnees)
                elif item.filename in noms_comments:
                    donnees = _applique_police_comments_xml(donnees)
                zout.writestr(item, donnees)
    return tampon.getvalue()
