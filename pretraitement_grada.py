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
    """« AB » (ou « AB, équitable ») -> Bio=X, et le reste du texte (hors
    « AB » lui-même) va dans Label s'il reste quelque chose (ex: 'équitable').
    Sans « AB », tout le texte va dans Label (ex: 'LABEL ROUGE')."""
    if not signe:
        return None, None
    if not RE_SIGNE_BIO.search(pt._normalise(signe)):
        return None, signe
    reste = RE_SIGNE_BIO.sub("", signe).strip(" ,;").strip()
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


def transforme_agricourt(chemin_source, chemin_sortie):
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    ws = wb["Consultation"]

    marche = extrait_titre_lot(ws.cell(row=2, column=2).value)

    ligne_tags = _trouve_ligne_tags(ws)
    tags = _tags_de_la_ligne(ws, ligne_tags)

    col_libelle = _trouve_colonne_texte(ws, ligne_tags, "libellé")
    col_desig_fourn = tags["_NOM_PUBLIC"][-1]
    col_reference = tags["_REFERENCE"][-1]
    col_prix = tags["_PRIX"][-1]
    col_ug = tags["_UG"][-1]
    col_cond = _trouve_colonne_texte(ws, ligne_tags, "conditionnement minimum de vente")
    col_signe = tags.get("_SIGNES", [None])[-1]
    col_origine = _trouve_colonne_texte(ws, ligne_tags, "origine matière première")
    col_local_oui_non = _trouve_colonne_texte(ws, ligne_tags, "oui ou non")

    ligne_debut = ligne_tags + 1
    ligne_fin = _derniere_ligne_avec_reference(ws, ligne_debut, col_reference)
    lots = _lots_par_section(ws, ligne_debut, ligne_fin, col_libelle, col_reference)

    lignes = []
    for r, section in lots.items():
        libelle = ws.cell(row=r, column=col_libelle).value
        desig_fourn = ws.cell(row=r, column=col_desig_fourn).value
        reference = ws.cell(row=r, column=col_reference).value
        prix = _nombre(ws.cell(row=r, column=col_prix).value)
        ug = ws.cell(row=r, column=col_ug).value
        cond = ws.cell(row=r, column=col_cond).value if col_cond else None
        signe = ws.cell(row=r, column=col_signe).value if col_signe else None
        origine_brute = ws.cell(row=r, column=col_origine).value if col_origine else None
        oui_non = ws.cell(row=r, column=col_local_oui_non).value if col_local_oui_non else None

        if not prix:
            continue

        ug_norm = pt._normalise(ug)
        if ug_norm in ("eur/kg", "kg", "€/kg"):
            unite_ref, unite_cmd, prix_cmd, commentaire_unite = _traite_prix_kg(prix, ug_norm, cond)
        else:
            unite_ref, unite_cmd, prix_cmd, commentaire_unite = _traite_prix_exception(prix, ug, cond)

        local_x = "X" if pt._normalise(oui_non) == "oui" else None
        bio_x, label = _bio_et_label(signe)

        valeur_origine, _, commentaire_origine = pt.traite_origine(origine_brute)

        lignes.append(_ligne_vide_de_base(
            fournisseur="AGRICOURT", lot=_combine_lot(marche, section), designation=libelle, unite_ref=unite_ref,
            poids=None, commentaire_unite=commentaire_unite, cond=cond, desig_fourn=desig_fourn,
            ref=reference, prix_ref=prix, unite_cmd=unite_cmd, prix_cmd=prix_cmd,
            local=local_x, bio=bio_x, label=label, origine=valeur_origine,
            commentaire_origine=commentaire_origine, marche=None,
        ))

    lignes.sort(key=lambda l: (pt._normalise(l["lot"]), pt._normalise(l["designation"])))
    return pt.ecrit_classeur_canonique(lignes, chemin_sortie, "AGRICOURT")


def transforme_vdr(chemin_source, chemin_sortie, nom_feuille=None):
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    ws = wb[nom_feuille] if nom_feuille else wb[wb.sheetnames[-1]]

    marche = extrait_titre_lot(ws.cell(row=2, column=1).value)

    ligne_tags = _trouve_ligne_tags(ws)
    tags = _tags_de_la_ligne(ws, ligne_tags)

    col_libelle = _trouve_colonne_texte(ws, ligne_tags, "libellé")
    col_desig_fourn = tags["_DESIGNATION"][-1]
    col_reference = tags["_REFERENCE"][-1]
    col_prix = tags["_PRIX"][-1]
    col_ug = tags["_UG"][-1]
    col_cond = _trouve_colonne_texte(ws, ligne_tags, "conditionnement minimum de vente")
    col_origine = _trouve_colonne_texte(ws, ligne_tags, "origine")
    col_commentaires = _trouve_colonne_texte(ws, ligne_tags, "commentaires")

    ligne_debut = ligne_tags + 2  # +1 = ligne de section ("Poulet"...), +1 = 1re donnée
    ligne_fin = _derniere_ligne_avec_reference(ws, ligne_debut, col_reference)
    lots = _lots_par_section(ws, ligne_tags + 1, ligne_fin, col_libelle, col_reference)

    lignes = []
    for r, section in lots.items():
        libelle = ws.cell(row=r, column=col_libelle).value
        desig_fourn = ws.cell(row=r, column=col_desig_fourn).value
        reference = ws.cell(row=r, column=col_reference).value
        prix = _nombre(ws.cell(row=r, column=col_prix).value)
        ug = ws.cell(row=r, column=col_ug).value
        cond = ws.cell(row=r, column=col_cond).value if col_cond else None
        origine_brute = ws.cell(row=r, column=col_origine).value if col_origine else None
        commentaires = ws.cell(row=r, column=col_commentaires).value if col_commentaires else None

        if not prix:
            continue

        ug_norm = pt._normalise(ug)
        if ug_norm in ("piece", "pièce"):
            prix_piece = prix_depuis_commentaire(commentaires)
            if prix_piece is not None:
                unite_ref, unite_cmd, prix_cmd = "KG", "PIECE", prix_piece
                commentaire_unite = None
            else:
                unite_ref, unite_cmd, prix_cmd, commentaire_unite = _traite_prix_exception(
                    prix, ug, cond
                )
        else:
            unite_ref, unite_cmd, prix_cmd = "KG", "KG", prix
            commentaire_unite = None

        valeur_origine, _, commentaire_origine = pt.traite_origine(origine_brute)

        lignes.append(_ligne_vide_de_base(
            fournisseur="VDR", lot=_combine_lot(marche, section), designation=libelle, unite_ref=unite_ref,
            poids=None, commentaire_unite=commentaire_unite, cond=cond, desig_fourn=desig_fourn,
            ref=reference, prix_ref=prix, unite_cmd=unite_cmd, prix_cmd=prix_cmd,
            local=None, bio=None, label=None, origine=valeur_origine,
            commentaire_origine=commentaire_origine, marche=None,
        ))

    lignes.sort(key=lambda l: (pt._normalise(l["lot"]), pt._normalise(l["designation"])))
    return pt.ecrit_classeur_canonique(lignes, chemin_sortie, "VDR")


def transforme_boucherie(chemin_source, chemin_sortie):
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    ws = wb["MajTarifs"]

    marche = extrait_titre_lot(ws.cell(row=2, column=1).value)

    ligne_tags = _trouve_ligne_tags(ws)
    tags = _tags_de_la_ligne(ws, ligne_tags)

    col_libelle = tags.get("_DESIGNATION_CONSULTATION", [None])[-1] or \
        _trouve_colonne_texte(ws, ligne_tags, "libellé")
    col_desig_fourn = tags["_DESIGNATION"][-1]
    col_reference = tags["_REFERENCE"][-1]
    col_prix = tags["_PRIX"][-1]
    col_ug = tags["_UG"][-1]
    col_cond = _trouve_colonne_texte(ws, ligne_tags, "conditionnement minimum de vente")
    col_signe = tags.get("_SIGNE_QUALITE", [None])[-1]

    ligne_debut = ligne_tags + 1
    ligne_fin = _derniere_ligne_avec_reference(ws, ligne_debut, col_reference)
    lots = _lots_par_section(ws, ligne_debut, ligne_fin, col_libelle, col_reference)

    lignes = []
    for r, section in lots.items():
        libelle = ws.cell(row=r, column=col_libelle).value
        desig_fourn = ws.cell(row=r, column=col_desig_fourn).value
        reference = ws.cell(row=r, column=col_reference).value
        prix = _nombre(ws.cell(row=r, column=col_prix).value)
        cond = ws.cell(row=r, column=col_cond).value if col_cond else None
        signe = ws.cell(row=r, column=col_signe).value if col_signe else None

        if not prix:
            continue

        unite_ref, unite_cmd, prix_cmd, commentaire_unite = _traite_prix_kg(prix, "kg", cond)
        bio_x, label = _bio_et_label(signe)

        lignes.append(_ligne_vide_de_base(
            fournisseur="BOUCHERIE GRADA", lot=_combine_lot(marche, section), designation=libelle, unite_ref=unite_ref,
            poids=None, commentaire_unite=commentaire_unite, cond=cond, desig_fourn=desig_fourn,
            ref=reference, prix_ref=prix, unite_cmd=unite_cmd, prix_cmd=prix_cmd,
            local=None, bio=bio_x, label=label, origine=None, commentaire_origine=None,
            marche=None,
        ))

    lignes.sort(key=lambda l: (pt._normalise(l["lot"]), pt._normalise(l["designation"])))
    return pt.ecrit_classeur_canonique(lignes, chemin_sortie, "BOUCHERIE")
