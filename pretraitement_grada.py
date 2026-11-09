"""
Prétraitement — famille "GRADA" (consultations BPU avec une ligne de tags
internes _REFERENCE/_DESIGNATION/_PRIX/_UG/...) : AGRICOURT, VDR, Lot 8
Boucherie-GRADA-PPF. Les trois partagent la même structure (prix par unité
de gestion + texte de conditionnement minimum de vente + lignes de section
qui servent de Lot), mais avec des noms de tags et des colonnes qui varient
d'un fichier à l'autre — toutes les colonnes sont donc repérées dynamiquement
(par tag ou par texte d'en-tête), jamais par position fixe.

Réutilise l'écriture canonique de pretraitement.py (REGAL), sans le modifier.
"""

import re

import openpyxl
from openpyxl.comments import Comment

import pretraitement as pt
from causes_format import CauseFormatInvalide

COULEUR_A_VERIFIER = pt.COULEUR_ORIGINE_INCERTAINE

RE_POIDS_SEUL = re.compile(r"^\s*(\d+[.,]?\d*)\s*(kg|g|l|ml)\b", re.I)
RE_PRIX_PAR_PIECE = re.compile(r"(\d+[.,]\d+)\s*€?\s*/\s*piece", re.I)
RE_SIGNE_BIO = re.compile(r"\bab\b", re.I)
# Le mot "bio" lui-même (pas seulement la sigle "AB") doit aussi déclencher Bio=X et sortir du
# Label — vu le 2026-10-30 sur BPUR FRESCA ("BIO, V - LABEL" restait entièrement en Label, Bio
# jamais coché, car "bio" ne contient pas "ab" comme mot séparé). Même logique que
# pretraitement_a_formater._bio_et_label_ardeche, appliquée ici à la fonction partagée par
# AGRICOURT/VDR/BOUCHERIE GRADA et BPUR FRESCA.
RE_SIGNE_BIO_MOT = re.compile(r"\bbio(?:logique)?\b", re.I)
RE_LIGNE_LOT = re.compile(r"^\s*lot\s*\d+.*$", re.I | re.M)


def _combine_lot(titre_lot, section):
    """Combine le titre du lot (niveau fichier, ex: 'Lot 8 Boucherie') et la
    section repérée ligne à ligne (ex: 'Boeuf') en une seule valeur de Lot."""
    if titre_lot and section:
        return f"{titre_lot} - {section}"
    return titre_lot or section


def extrait_titre_lot(titre_brut):
    """Repère, dans le titre multi-lignes du document (ex: 'AOOE fourniture...
    \\nLOT 13 Fruits et légumes et 4ème et 5ème gamme BIO'), la ligne qui
    donne le vrai nom du lot (celle qui commence par « Lot N »), et la
    renvoie telle quelle — ou None si rien ne correspond."""
    if not titre_brut:
        return None
    m = RE_LIGNE_LOT.search(str(titre_brut))
    return m.group(0).strip() if m else None


def poids_total_conditionnement(cond):
    """Lit soit « N x {poids}{unité} » (ex: Boucherie '1 x 2kg') soit un
    « {poids}{unité} » nu (ex: AGRICOURT '18,5Kg'), et renvoie le poids TOTAL
    en kg du conditionnement (N × poids si N>1), ou None si rien ne correspond."""
    if not cond:
        return None
    n, valeur, unite = pt._n_cond_et_quantite(cond)
    if n is not None:
        if unite not in ("kg", "g"):
            return None
        poids_unitaire = valeur / 1000 if unite == "g" else valeur
        return n * poids_unitaire
    m = RE_POIDS_SEUL.match(str(cond).strip())
    if not m:
        return None
    valeur = float(m.group(1).replace(",", "."))
    unite = m.group(2).lower()
    if unite == "l":
        return None  # volume, pas un poids — pas de sens ici
    return valeur / 1000 if unite == "g" else valeur


def prix_depuis_commentaire(commentaire):
    """Lit un texte du genre '5,884€/PIECE' et renvoie 5.884, ou None."""
    if not commentaire:
        return None
    m = RE_PRIX_PAR_PIECE.search(str(commentaire))
    return float(m.group(1).replace(",", ".")) if m else None


def _nombre(valeur):
    """Convertit un prix qui peut être un nombre ou un texte à virgule
    française (ex: '18,344') en float, ou None si rien d'exploitable."""
    if valeur is None or valeur == "":
        return None
    if isinstance(valeur, (int, float)):
        return float(valeur)
    try:
        return float(str(valeur).replace(",", "."))
    except ValueError:
        return None


def _bio_et_label(signe):
    """« AB » ou le mot « bio » (ex: « AB, équitable », « BIO, V - LABEL ») -> Bio=X, et le
    reste du texte (hors la mention bio elle-même) va dans Label s'il reste quelque chose (ex:
    'équitable', 'V - LABEL'). Sans mention bio, tout le texte va dans Label (ex: 'LABEL
    ROUGE')."""
    if not signe:
        return None, None
    norm = pt._normalise(signe)
    if not (RE_SIGNE_BIO.search(norm) or RE_SIGNE_BIO_MOT.search(norm)):
        return None, signe
    reste = RE_SIGNE_BIO.sub("", signe)
    reste = RE_SIGNE_BIO_MOT.sub("", reste)
    reste = reste.strip(" ,;").strip()
    return "X", (reste or None)


def _trouve_ligne_tags(ws, tag_recherche="_REFERENCE", max_lignes=40):
    derniere_ligne_cherchee = min(max_lignes, ws.max_row)
    for r in range(1, derniere_ligne_cherchee + 1):
        for c in range(1, ws.max_column + 1):
            if str(ws.cell(row=r, column=c).value or "").strip() == tag_recherche:
                return r
    raise CauseFormatInvalide(
        "SECTION_INTROUVABLE",
        f"Ligne de tags introuvable dans « {ws.title} » (cherché {tag_recherche!r})",
        onglet=ws.title, cellule=f"lignes 1 à {derniere_ligne_cherchee}",
    )


def _tags_de_la_ligne(ws, ligne_tags):
    """Renvoie {tag: [colonnes]} — une liste, car certains tags (ex: _PRIX,
    _UG) apparaissent deux fois (cotation initiale puis révisée)."""
    tags = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=ligne_tags, column=c).value
        if v:
            tags.setdefault(str(v).strip(), []).append(c)
    return tags


def _trouve_colonne_texte(ws, ligne_max, texte_recherche):
    cible = pt._normalise(texte_recherche)
    for r in range(1, ligne_max):
        for c in range(1, ws.max_column + 1):
            if cible in pt._normalise(ws.cell(row=r, column=c).value):
                return c
    return None


def _lots_par_section(ws, ligne_debut, ligne_fin, col_libelle, col_reference):
    """Une ligne de section (Libellé rempli, Référence vide) sert d'étiquette
    de Lot pour toutes les lignes de produit qui suivent, jusqu'à la
    prochaine section. Renvoie {ligne: lot_ou_none}, seulement pour les
    lignes de produit (Référence remplie)."""
    lots = {}
    lot_courant = None
    for r in range(ligne_debut, ligne_fin + 1):
        libelle = ws.cell(row=r, column=col_libelle).value
        reference = ws.cell(row=r, column=col_reference).value
        if libelle and not reference:
            lot_courant = str(libelle).strip()
            continue
        if reference:
            lots[r] = lot_courant
    return lots


def _derniere_ligne_avec_reference(ws, ligne_debut, col_reference, limite=500):
    derniere = ligne_debut
    for r in range(ligne_debut, min(ws.max_row, ligne_debut + limite) + 1):
        if ws.cell(row=r, column=col_reference).value:
            derniere = r
    return derniere


def _ligne_vide_de_base(fournisseur, lot, designation, unite_ref, poids, commentaire_unite,
                         cond, desig_fourn, ref, prix_ref, unite_cmd, prix_cmd,
                         local, bio, label, origine, commentaire_origine, marche=None):
    return {
        "marche": marche, "lot": lot, "designation": designation, "unite_ref": unite_ref,
        "poids": poids, "commentaire_unite": commentaire_unite, "fournisseur": fournisseur,
        "cond": cond, "desig_fourn": desig_fourn, "ref": ref, "prix_ref": prix_ref,
        "unite_cmd": unite_cmd, "prix_cmd": prix_cmd, "local": local, "bio": bio,
        "label": label, "origine": origine, "commentaire_origine": commentaire_origine,
        "allergenes": {nom: None for nom in pt.ALLERGENES}, "tva": None,
    }


def _traite_prix_kg(prix, ug_norm, cond):
    """Cas normal (UG = kg) : Unité de référence = KG, Unité de commande =
    KG aussi, prix de commande = prix × poids total du conditionnement.
    Renvoie (unite_ref, unite_cmd, prix_cmd, commentaire_ou_none)."""
    poids = poids_total_conditionnement(cond)
    if poids is None:
        return "KG", "KG", prix, (
            f"Aucun poids trouvé dans le conditionnement ('{cond}') ; prix de commande "
            "laissé égal au prix de référence (pas de conversion possible)."
        )
    return "KG", "KG", round(prix * poids, 3), None


def _traite_prix_exception(prix, ug_brut, cond):
    """UG différent de kg (ex: Pièce/Botte/Colis) sans prix explicite ailleurs
    (cas AGRICOURT) : pas de multiplication risquée — on garde PC des deux
    côtés avec le même prix, et on flague pour vérification humaine."""
    commentaire = (
        f"Unité de gestion d'origine : « {ug_brut} » (pas en kilo) — prix laissé tel quel, "
        f"sans conversion ; vérifier si le conditionnement ('{cond}') implique de multiplier "
        "le prix par la quantité qu'il indique."
    )
    return "PC", "PC", prix, commentaire
