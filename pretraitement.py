"""
Prétraitement — transforme des fichiers "en vrac" (messy_files/) dans le même
format que ceux de raw_data/, pour qu'ils puissent ensuite être audités par
audit_geral.py (qui ne dépend que du contenu des en-têtes, pas du nom de la
feuille ni de la position absolue des colonnes).

Pour l'instant : seul le catalogue au format REGAL est pris en charge
(transforme_catalogue_regal).
"""

import re
import unicodedata

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

COULEUR_ORIGINE_INCERTAINE = PatternFill(
    start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
)

# Aspect copié du modèle raw_data/ (fichiers groupement) : ligne 1 = avertissement,
# ligne 2 = en-têtes (fond bleu pâle, gras), ligne 3+ = données, toutes avec
# bordure fine sur les cellules remplies.
LIGNE1_COL1 = "NE PAS TOUCHER_1001"
LIGNE1_COL2 = "NE PAS INSERER DE LIGNE CI-DESSUS, NE PAS MODIFIER L'ORDRE DES COLONNES"
FONT_LIGNE1_COL1 = Font(name="Calibri", size=11)
FONT_LIGNE1_COL2 = Font(name="Calibri", size=28)

_BORDURE_FINE = Side(style="thin", color="FF000000")
BORDURE_ENTETE = Border(left=_BORDURE_FINE, right=_BORDURE_FINE, bottom=_BORDURE_FINE)
BORDURE_DONNEE = Border(
    left=_BORDURE_FINE, right=_BORDURE_FINE, top=_BORDURE_FINE, bottom=_BORDURE_FINE
)
FONT_DONNEE = Font(name="Calibri", size=11)
ALIGNEMENT = Alignment(horizontal="left", vertical="center")

# Largeurs de colonnes (1 à 62) copiées telles quelles du modèle raw_data/
# 2026 09 22 Base article groupement 1.xlsx.
LARGEURS_COLONNES = {
    1: 31.0, 2: 44.26953125, 3: 49.26953125, 4: 15.26953125, 5: 13.0, 6: 17.7265625,
    7: 31.81640625, 8: 31.81640625, 9: 30.81640625, 10: 11.81640625, 11: 14.54296875,
    12: 13.1796875, 13: 13.1796875, 14: 14.453125, 15: 13.0, 16: 13.0, 17: 14.453125,
    18: 14.453125, 19: 13.0, 20: 28.0, 21: 13.0, 22: 13.0, 23: 10.81640625, 24: 14.453125,
    25: 13.0, 26: 13.0, 27: 13.0, 28: 30.1796875, 29: 13.0, 30: 13.0, 31: 13.0, 32: 8.54296875,
    33: 16.453125, 34: 12.81640625, 35: 20.1796875, 36: 14.453125, 37: 13.0, 38: 23.7265625,
    39: 21.54296875, 40: 14.54296875, 41: 19.0, 42: 13.0, 43: 13.0, 44: 13.0, 45: 13.0,
    46: 13.0, 47: 13.0, 48: 13.0, 49: 13.0, 50: 13.0, 51: 13.0, 52: 13.81640625, 53: 13.0,
    54: 13.0, 55: 13.0, 56: 13.0, 57: 13.0, 58: 13.0, 59: 13.0, 60: 13.0, 61: 13.0, 62: 13.0,
}


def _taille_police_bloc_fournisseur(col, debut):
    offset = col - debut
    if offset <= 2:
        return 16
    if offset == 3:
        return 9
    return 8


def _taille_police_entete(col):
    """Reproduit les tailles de police (par colonne) du modèle raw_data/ —
    pas une taille uniforme : chaque colonne du cabeçalho a la sienne."""
    if col == 1:
        return 12
    if col in (2, 3, 4):
        return 11
    if col in (5, 6):
        return 8
    if 7 <= col <= 19:
        return _taille_police_bloc_fournisseur(col, 7)
    if 20 <= col <= 32:
        return _taille_police_bloc_fournisseur(col, 20)
    if 33 <= col <= 44:
        return _taille_police_bloc_fournisseur(col, 33)
    return 11

# Couleurs exactes extraites du thème du fichier raw_data/2026 09 22 Base
# article groupement 1.xlsx (combinaison thème de couleur + tint, résolue en
# RGB) : bleu pâle (1-6), orange/corail pour Fournisseur 1 (7-19), vert pour
# Fournisseur 2 (20-32), violet pour Fournisseur 3 (33-44), vert très pâle
# pour les allergènes (45-58), bleu-cyan pour TVA (59), rien après.
FILL_MARCHE = PatternFill(start_color="C6D9F1", end_color="C6D9F1", fill_type="solid")
FILL_FOURNISSEUR_1 = PatternFill(start_color="FCD5B5", end_color="FCD5B5", fill_type="solid")
FILL_FOURNISSEUR_2 = PatternFill(start_color="D7E4BD", end_color="D7E4BD", fill_type="solid")
FILL_FOURNISSEUR_3 = PatternFill(start_color="CCC1DA", end_color="CCC1DA", fill_type="solid")
FILL_ALLERGENES = PatternFill(start_color="EBF1DE", end_color="EBF1DE", fill_type="solid")
FILL_TVA = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")

# En-têtes d'un bloc "Fournisseur N" — identique pour 1 et 2 ; le bloc 3, dans
# le modèle original, n'a pas de "Circuit court" (particularité du modèle,
# reproduite telle quelle).
_ENTETES_BLOC_FOURNISSEUR = [
    "Conditionnement", "Désignation fournisseur", "Référence produit",
    "PRIX UNITE REFERENCE (BPU)\nKG / L / PC", "UNITE COMMANDE", "PRIX UNITE COMMANDE",
    "Local\n(X si local)", "Bio\n(X si bio)", "Circuit court", "Label", "Origine", "Indice I0",
]
_ENTETES_BLOC_FOURNISSEUR_3 = [
    "Conditionnement", "Désignation fournisseur", "Référence produit",
    "PRIX UNITE REFERENCE (BPU)\nKG / L / PC", "UNITE COMMANDE", "PRIX UNITE COMMANDE",
    "Local\n(X si local)", "Bio\n(X si bio)", "Label", "Origine", "Indice I0",
]

ALLERGENES = [
    "Gluten",
    "Crustacés",
    "Oeufs",
    "Poissons",
    "Arachides",
    "Soja",
    "Lait",
    "Fruits à coques",
    "Céleri",
    "Moutarde",
    "Sésame",
    "Sulfites",
    "Lupins",
    "Mollusques",
]

ENTETES = (
    ["Marché", "Lot", "Désignation de l'article", "Unité de référence\n(BPU)",
     "Poids à l'unité de référence\n", "Nombre de \nPièces dans \nconditionnement",
     "Fournisseur 1"] + _ENTETES_BLOC_FOURNISSEUR
    + ["Fournisseur 2"] + _ENTETES_BLOC_FOURNISSEUR
    + ["Fournisseur 3"] + _ENTETES_BLOC_FOURNISSEUR_3
    + ["Gluten (o si oui, vide si non)", "Crustacés", "Oeufs", "Poissons", "Arachides", "Soja",
       "Lait", " Fruits à coques", "Céleri", "Moutarde", "Sésame", "Sulfites", "Lupins",
       "Mollusques", "TAUX TVA", "Code RNM / ITAVI", "", "ID_AURA"]
)
NB_COLONNES = len(ENTETES)  # 62, comme le modèle

COL_MARCHE = 1
COL_LOT = 2
COL_DESIGNATION = 3
COL_UNITE_REF = 4
COL_POIDS_REF = 5

COL_FOURNISSEUR = 7  # Fournisseur 1
COL_COND = COL_FOURNISSEUR + 1
COL_DESIG_FOURN = COL_FOURNISSEUR + 2
COL_REF = COL_FOURNISSEUR + 3
COL_PRIX_REF = COL_FOURNISSEUR + 4
COL_UNITE_CMD = COL_FOURNISSEUR + 5
COL_PRIX_CMD = COL_FOURNISSEUR + 6
COL_LOCAL = COL_FOURNISSEUR + 7
COL_BIO = COL_FOURNISSEUR + 8
COL_LABEL = COL_FOURNISSEUR + 10
COL_ORIGINE = COL_FOURNISSEUR + 11

COL_GLUTEN = 45
COL_TVA = 59

ORIGINES_LARGES = {"france", "union europeenne", "hors union europeenne", "monde"}
PHRASES_REGIONALES = {"produit regional", "fabrique en region"}
PAYS_ETRANGERS_CONNUS = {
    "espagne", "italie", "belgique", "chine", "royaume-uni", "chili", "perou",
    "irlande", "egypte", "asie", "allemagne", "portugal", "pays-bas", "maroc",
}


def _sans_accents(texte):
    return "".join(
        c for c in unicodedata.normalize("NFD", texte) if unicodedata.category(c) != "Mn"
    )


def _normalise(texte):
    if texte is None:
        return ""
    return _sans_accents(str(texte)).strip().lower()


def _trouve_colonnes(ws, noms_recherches):
    """Associe chaque nom logique (clé) au numéro de colonne du premier en-tête
    (ligne 1) qui correspond, en comparant les textes normalisés."""
    entetes = {}
    for c in range(1, ws.max_column + 1):
        texte = _normalise(ws.cell(row=1, column=c).value)
        if texte:
            entetes[texte] = c

    resultat = {}
    for cle, recherche in noms_recherches.items():
        colonne = entetes.get(_normalise(recherche))
        resultat[cle] = colonne
    return resultat


def traite_origine(valeur_source):
    """Décide quoi écrire dans la colonne Origine à partir d'une valeur brute
    de la source, et si le produit doit aussi être marqué Local.

    Renvoie (valeur_a_ecrire, est_regional_confirme, commentaire_ou_none).
    """
    norm = _normalise(valeur_source)
    if not norm:
        return None, False, None
    if norm in ORIGINES_LARGES:
        return str(valeur_source).strip(), False, None
    if norm in PHRASES_REGIONALES:
        return "France", True, None
    if norm in PAYS_ETRANGERS_CONNUS:
        return str(valeur_source).strip(), False, None
    commentaire = (
        f"Donnée source : « {valeur_source} » — remplacé par « France » (valeur non "
        "reconnue comme un pays ni comme une des catégories larges France/UE/Hors UE/Monde) ; "
        "vérifier si ce produit doit aussi être marqué Local."
    )
    return "France", False, commentaire


UNITES_FACTURATION_DIRECTES = {"kg": "KG", "litre": "L", "piece": "PC"}
UNITES_CONTENANT = {"barquette", "colis", "botte"}
RE_N_FOIS_QUANTITE = re.compile(r"^\s*(\d+)\s*[xX]\s*(\d+[.,]?\d*)\s*(kg|g|l|ml|cl)\b", re.I)


def _n_cond_et_quantite(cond):
    """Lit « N x {quantité}{unité} » (masse OU volume) et renvoie
    (n, valeur_numerique, unite_normalisee) ou (None, None, None)."""
    if not cond:
        return None, None, None
    m = RE_N_FOIS_QUANTITE.match(str(cond))
    if not m:
        return None, None, None
    return int(m.group(1)), float(m.group(2).replace(",", ".")), m.group(3).lower()


def _n_cond_et_poids(cond):
    """Lit « N x {poids}{kg|g} » et renvoie (n, poids_en_kg) ou (None, None) —
    uniquement pour les unités de masse (pas de sens pour un volume)."""
    n, valeur, unite = _n_cond_et_quantite(cond)
    if n is None or unite not in ("kg", "g"):
        return None, None
    return n, (valeur / 1000 if unite == "g" else valeur)


def _formate_poids(poids_kg):
    if poids_kg < 1:
        return f"{round(poids_kg * 1000)}G"
    texte = f"{poids_kg:.3f}".rstrip("0").rstrip(".")
    return f"{texte.replace('.', ',')}KG"


def _formate_quantite(valeur, unite):
    if unite in ("kg", "g"):
        return _formate_poids(valeur / 1000 if unite == "g" else valeur)
    if unite == "l":
        texte = f"{valeur:.3f}".rstrip("0").rstrip(".")
        return f"{texte.replace('.', ',')}L"
    return f"{round(valeur)}{unite.upper()}"  # ml, cl : toujours des entiers en pratique


def reecrit_conditionnement(cond, unite_commande, unite_ref_finale=None):
    """Réécrit le Conditionnement en deux temps, appliqués à TOUTES les
    lignes, indépendamment de ce qui a été décidé pour Unité de référence :

    1. Si Unité de référence est finalement « L » mais que le texte est en
       masse (kg/g) — ex: un liquide dont le conditionnement REGAL dit
       '1 x 5000g' au lieu de volume — on convertit en volume (1g = 1mL),
       directement en L si ça atteint 1000 mL ou plus.
    2. Insère le mot de l'Unité de commande (ex: COLIS, mais aussi PIECE —
       même générique, il est inséré pareil) dans un Conditionnement du type
       « N x quantité » qui n'en a pas. Important : c'est ce texte de
       Conditionnement, pas le champ Unité de commande d'origine (qui
       disparaît une fois dans le format canonique), qui alimente l'audit ;
       sans aucun mot dedans, un texte nu du type « 4 x 840g » est pris pour
       un « LOT DE 4 » générique. Avec « PIECE DE N X ... », quand N > 1 ça
       reste PC avec le poids rempli (testé, sans effet secondaire) ; quand
       N == 1 ça bascule proprement en KG/L (pas de nom de pièce isolée à
       garder) — l'un comme l'autre sans l'étiquette « LOT DE 1 » bizarre
       qu'on avait en laissant le texte nu.
    """
    n, valeur, unite = _n_cond_et_quantite(cond)
    if n is None:
        return cond

    conversion_effectuee = False
    if unite_ref_finale == "L" and unite in ("kg", "g"):
        valeur_ml = valeur * 1000 if unite == "kg" else valeur
        valeur, unite = (valeur_ml / 1000, "l") if valeur_ml >= 1000 else (valeur_ml, "ml")
        conversion_effectuee = True

    quantite_texte = _formate_quantite(valeur, unite)

    conteneur = str(unite_commande or "").strip()
    if not conteneur:
        if not conversion_effectuee:
            return cond  # rien n'a changé (ni conversion, ni contenant à ajouter)
        return f"{n} X {quantite_texte}" if n > 1 else quantite_texte

    if not conversion_effectuee and _normalise(conteneur) in _normalise(cond):
        return cond
    return (
        f"{conteneur} DE {n} X {quantite_texte}" if n > 1
        else f"{conteneur} DE {quantite_texte}"
    )


def traite_unite_reference(unite_facturation, unite_commande, cond, prix_ref, prix_cmd, poids_piece_g):
    """Décide quoi écrire dans Unité de référence et Poids à l'unité de
    référence, à partir de l'« Unité de facturation » du REGAL (le
    Conditionnement, lui, est réécrit séparément par reecrit_conditionnement,
    sur TOUTES les lignes — pas seulement celles-ci).

    - KG/LITRE se traduisent directement en KG/L ; pas de poids par pièce
      (ça n'a pas de sens pour une unité déjà en kg/L).
    - PIECE (déjà la bonne base) et BARQUETTE/COLIS/BOTTE (des noms de
      contenant, qu'il faut ramener à PC) sont traités pareil : on cherche à
      confirmer un nombre de pièces de référence dans le conditionnement
      (« N x poids ») et, si confirmé, on remplit le poids — lu directement
      dans le conditionnement, qui est garanti cohérent avec ce qui est
      facturé, contrairement à la colonne « Poids pièce » qui référence
      parfois une sous-unité encore plus fine (ex: le poids d'1 tranche dans
      une barquette de plusieurs tranches, pas le poids de la barquette).
      Confirmation : soit il n'y a qu'une seule pièce, soit le ratio prix
      commande/prix facturation retombe sur ce nombre, soit — si le prix est
      nul/absent et ne permet pas de calculer ce ratio — l'unité de
      facturation et l'unité de commande sont déjà deux mots différents (le
      REGAL distingue alors lui-même le contenant détaillé de l'unité
      d'achat, ce qui suffit à confirmer sans avoir besoin du prix).
      Sinon (ambigu — rien ne confirme le nombre de pièces), on laisse tel
      quel pour l'instant — cette partie sera traitée plus tard, côté audit.

    Renvoie (valeur_unite, poids_kg_ou_none, commentaire_ou_none).
    """
    norm = _normalise(unite_facturation)
    if norm in ("kg", "litre"):
        return UNITES_FACTURATION_DIRECTES[norm], None, None

    est_piece_directe = norm == "piece"
    if not est_piece_directe and norm not in UNITES_CONTENANT:
        return unite_facturation, None, None

    n_cond, poids_cond = _n_cond_et_poids(cond)
    ratio = (prix_cmd or 0) / prix_ref if prix_ref else None
    uc_norm = _normalise(unite_commande)
    uf_differe_uc = uc_norm and uc_norm != norm

    cas_confirme = n_cond == 1 or (
        n_cond and n_cond > 1 and (
            (ratio and abs(ratio - n_cond) < 0.05 * n_cond)
            or (ratio is None and uf_differe_uc)
        )
    )
    if not cas_confirme:
        return unite_facturation, None, None

    poids_kg = poids_cond if poids_cond is not None else (
        (poids_piece_g / 1000) if poids_piece_g else None
    )
    commentaire = None if est_piece_directe else (
        f"Unité de référence d'origine : « {unite_facturation} » (remplacée par PC)."
    )
    return "PC", poids_kg, commentaire


def _ecrit_ligne1_info(ws):
    ws.cell(row=1, column=1, value=LIGNE1_COL1).font = FONT_LIGNE1_COL1
    ws.cell(row=1, column=2, value=LIGNE1_COL2).font = FONT_LIGNE1_COL2


def _applique_largeurs_colonnes(ws):
    for c, largeur in LARGEURS_COLONNES.items():
        ws.column_dimensions[get_column_letter(c)].width = largeur


def _remplissage_entete(col):
    if col <= 6:
        return FILL_MARCHE, True
    if 7 <= col <= 19:
        return FILL_FOURNISSEUR_1, True
    if 20 <= col <= 32:
        return FILL_FOURNISSEUR_2, True
    if 33 <= col <= 44:
        return FILL_FOURNISSEUR_3, True
    if 45 <= col <= 58:
        return FILL_ALLERGENES, False
    if col == 59:
        return FILL_TVA, False
    return None, False


def _ecrit_entetes(ws):
    for c, texte in enumerate(ENTETES, start=1):
        cel = ws.cell(row=2, column=c, value=texte or None)
        fill, gras = _remplissage_entete(c)
        cel.font = Font(name="Calibri", size=_taille_police_entete(c), bold=gras)
        if fill:
            cel.fill = fill
        cel.border = BORDURE_ENTETE
        cel.alignment = ALIGNEMENT


def _ecrit(ws, r, c, valeur):
    return ws.cell(row=r, column=c, value=valeur)


def _applique_grille_complete(ws, premiere_ligne, derniere_ligne):
    """Applique police + bordure + alignement à TOUTES les cellules de la
    zone de données (remplies ou non), comme dans les fichiers raw_data/ où
    la grille entière est mise en forme, indépendamment du contenu."""
    for r in range(premiere_ligne, derniere_ligne + 1):
        for c in range(1, NB_COLONNES + 1):
            cel = ws.cell(row=r, column=c)
            cel.font = FONT_DONNEE
            cel.border = BORDURE_DONNEE
            cel.alignment = ALIGNEMENT


def transforme_catalogue_regal(chemin_source, chemin_sortie):
    wb_source = openpyxl.load_workbook(chemin_source, data_only=True)
    ws_source = wb_source["Feuil1"] if "Feuil1" in wb_source.sheetnames else wb_source.worksheets[0]

    colonnes = _trouve_colonnes(ws_source, {
        "designation": "Désignation produit",
        "cond": "Cond rendu dans Magento",
        "prix_cmd": "Prix commande (HT)",
        "unite_cmd": "Unité de commande",
        "prix_ref": "Prix facturation (HT)",
        "unite_ref": "Unité de facturation",
        "ref": "Code produit fournisseur",
        "lot": "Catégorie 2",
        "local": "Produit Local ?",
        "egalim_type": "Egalim type",
        "origine": "Origine",
        "fournisseur": "Code fournisseur",
        "tva": "Code TVA",
        "poids_piece": "Poids pièce",
    })
    for cle, colonne in colonnes.items():
        if colonne is None:
            raise ValueError(f"Colonne source introuvable pour « {cle} » dans {chemin_source}")

    colonnes_allergenes = _trouve_colonnes(ws_source, {nom: nom for nom in ALLERGENES})

    lignes = []
    for r in range(2, ws_source.max_row + 1):
        designation = ws_source.cell(row=r, column=colonnes["designation"]).value
        if not designation:
            continue

        fournisseur = ws_source.cell(row=r, column=colonnes["fournisseur"]).value
        lot = ws_source.cell(row=r, column=colonnes["lot"]).value

        valeur_origine, regional_confirme, commentaire_origine = traite_origine(
            ws_source.cell(row=r, column=colonnes["origine"]).value
        )

        local_x = None
        if _normalise(ws_source.cell(row=r, column=colonnes["local"]).value) == "o":
            local_x = "X"
        if regional_confirme:
            local_x = "X"

        egalim = ws_source.cell(row=r, column=colonnes["egalim_type"]).value
        bio_x = "X" if _normalise(egalim) == "bio" else None
        label = egalim if egalim and _normalise(egalim) != "bio" else None

        cond_source = ws_source.cell(row=r, column=colonnes["cond"]).value
        unite_cmd_source = ws_source.cell(row=r, column=colonnes["unite_cmd"]).value

        valeur_unite_ref, poids_kg, commentaire_unite = traite_unite_reference(
            ws_source.cell(row=r, column=colonnes["unite_ref"]).value,
            unite_cmd_source,
            cond_source,
            ws_source.cell(row=r, column=colonnes["prix_ref"]).value,
            ws_source.cell(row=r, column=colonnes["prix_cmd"]).value,
            ws_source.cell(row=r, column=colonnes["poids_piece"]).value,
        )
        cond_a_ecrire = reecrit_conditionnement(cond_source, unite_cmd_source, valeur_unite_ref)

        allergenes_valeurs = {}
        for nom in ALLERGENES:
            col_source = colonnes_allergenes.get(nom)
            valeur = ws_source.cell(row=r, column=col_source).value if col_source else None
            allergenes_valeurs[nom] = "o" if _normalise(valeur) == "o" else None

        lignes.append({
            "marche": None,
            "lot": lot,
            "designation": designation,
            "unite_ref": valeur_unite_ref,
            "poids": poids_kg,
            "commentaire_unite": commentaire_unite,
            "fournisseur": fournisseur,
            "cond": cond_a_ecrire,
            "desig_fourn": designation,
            "ref": ws_source.cell(row=r, column=colonnes["ref"]).value,
            "prix_ref": ws_source.cell(row=r, column=colonnes["prix_ref"]).value,
            "unite_cmd": ws_source.cell(row=r, column=colonnes["unite_cmd"]).value,
            "prix_cmd": ws_source.cell(row=r, column=colonnes["prix_cmd"]).value,
            "local": local_x,
            "bio": bio_x,
            "label": label,
            "origine": valeur_origine,
            "commentaire_origine": commentaire_origine,
            "allergenes": allergenes_valeurs,
            "tva": ws_source.cell(row=r, column=colonnes["tva"]).value,
        })

    lignes.sort(key=lambda l: (
        _normalise(l["fournisseur"]), _normalise(l["marche"]), _normalise(l["lot"]),
    ))

    return ecrit_classeur_canonique(lignes, chemin_sortie, "REGAL")


def ecrit_classeur_canonique(lignes, chemin_sortie, nom_feuille):
    """Écrit un classeur au format canonique (même aspect que raw_data/) à
    partir d'une liste de « lignes » (dicts avec les mêmes clés que celles
    construites par transforme_catalogue_regal) — réutilisable par n'importe
    quel script de prétraitement, pas seulement celui du REGAL."""
    wb_dst = openpyxl.Workbook()
    ws_dst = wb_dst.active
    ws_dst.title = nom_feuille
    _ecrit_ligne1_info(ws_dst)
    _ecrit_entetes(ws_dst)
    _applique_largeurs_colonnes(ws_dst)
    ws_dst.row_dimensions[1].height = 43.5
    ws_dst.row_dimensions[2].height = 57.65

    for i, ligne in enumerate(lignes):
        r = i + 3
        _ecrit(ws_dst, r, COL_MARCHE, ligne["marche"])
        _ecrit(ws_dst, r, COL_LOT, ligne["lot"])
        _ecrit(ws_dst, r, COL_DESIGNATION, ligne["designation"])
        cel_unite_ref = _ecrit(ws_dst, r, COL_UNITE_REF, ligne["unite_ref"])
        if ligne["commentaire_unite"]:
            cel_unite_ref.comment = Comment(ligne["commentaire_unite"], "audit_geral")
        if ligne["poids"] is not None:
            _ecrit(ws_dst, r, COL_POIDS_REF, ligne["poids"])
        _ecrit(ws_dst, r, COL_FOURNISSEUR, ligne["fournisseur"])
        _ecrit(ws_dst, r, COL_COND, ligne["cond"])
        _ecrit(ws_dst, r, COL_DESIG_FOURN, ligne["desig_fourn"])
        _ecrit(ws_dst, r, COL_REF, ligne["ref"])
        _ecrit(ws_dst, r, COL_PRIX_REF, ligne["prix_ref"])
        _ecrit(ws_dst, r, COL_UNITE_CMD, ligne["unite_cmd"])
        _ecrit(ws_dst, r, COL_PRIX_CMD, ligne["prix_cmd"])
        _ecrit(ws_dst, r, COL_LOCAL, ligne["local"])
        _ecrit(ws_dst, r, COL_BIO, ligne["bio"])
        _ecrit(ws_dst, r, COL_LABEL, ligne["label"])

        cel_origine = _ecrit(ws_dst, r, COL_ORIGINE, ligne["origine"])
        if ligne["commentaire_origine"]:
            cel_origine.fill = COULEUR_ORIGINE_INCERTAINE
            cel_origine.comment = Comment(ligne["commentaire_origine"], "audit_geral")

        for j, nom in enumerate(ALLERGENES):
            _ecrit(ws_dst, r, COL_GLUTEN + j, ligne["allergenes"][nom])

        _ecrit(ws_dst, r, COL_TVA, ligne["tva"])

    if lignes:
        _applique_grille_complete(ws_dst, 3, 2 + len(lignes))

    wb_dst.save(chemin_sortie)
    return len(lignes)
