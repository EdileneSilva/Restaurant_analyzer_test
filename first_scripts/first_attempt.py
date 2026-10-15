"""
Transformation et Vérification — Marché alimentaire (lots frais)
==================================================================

Principe général (validé avec Edilene sur des cas réels du fichier) :
LE PRIX DE COMMANDE (M) EST TOUJOURS CONSIDÉRÉ COMME LE PRIX RÉEL VOULU PAR
LA PERSONNE. Le script ne le recalcule JAMAIS à l'aveugle. La question n'est
jamais "quel est le bon M", mais "qu'est-ce que K et les unités devraient
dire pour être cohérents avec ce M".

Ce que fait ce script, à partir du fichier MDC_IMPORT_MARCHE original :

  1. Supprime les lignes où Désignation de l'article ET Référence produit
     sont TOUTES LES DEUX vides, et les lignes strictement dupliquées.

  2. Convertit toutes les formules en valeurs fixes AVANT toute suppression
     de ligne (openpyxl ne réajuste pas les références de formule quand on
     supprime des lignes).

  3. Pour chaque ligne, dans cet ordre de priorité :

     a) Si le Conditionnement décrit un multiplicateur ("N x Qté" ou
        "N x M x Qté") : compare le ratio RÉEL (Prix commande ÷ Prix
        référence) aux valeurs candidates (N, ou N, M, N×M pour le format
        triple). Si un candidat correspond, c'est la confirmation que K est
        déjà le prix À LA PIÈCE et M le prix DU LOT ENTIER — dans ce cas
        AUCUN PRIX N'EST MODIFIÉ, seuls les libellés changent :
        Unité de référence -> "PC", Unité de commande -> "LOT DE {N}".
        Si aucun candidat ne correspond, rien n'est changé mais la ligne
        est signalée avec les candidats et le ratio observé, pour revue
        manuelle.

     b) Sinon, si un poids/volume de référence fiable existe (colonne E,
        confirmé si possible par la Désignation ou le Conditionnement) :
        l'unité de référence est convertie en KG/L, et le prix de référence
        (K) est DIVISÉ par ce poids — car E indique que K était en fait le
        prix pour CETTE quantité, pas pour 1kg/1L. M n'est jamais modifié.
        La colonne E est vidée après usage (elle n'a plus lieu d'être une
        fois l'unité en KG/L, conformément à la note de Julien L.).

     c) Sinon, l'unité "unité"/"pièce" est uniformisée en "PC" sans toucher
        aux prix ; les autres cas restent inchangés et sont signalés en
        bleu pour vérification manuelle.

  4. Vérifie la cohérence entre "Unité de commande" et le Conditionnement/
     Désignation (signalement seul) — sauf sur les lignes déjà traitées
     comme "LOT DE N", où l'unité de commande vient d'être définie par ce
     traitement même.

  5. Toute valeur modifiée est surlignée ET commentée. Toujours une COPIE
     du fichier d'origine, jamais l'original. Aucune colonne ajoutée.

Comment l'utiliser :
    pip install openpyxl
    python transformer_mdc.py
"""

import re
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

# ============================================================================
# CONFIGURATION
# ============================================================================

FICHIER_ENTREE = "MDC_IMPORT_MARCHE - Lots frais 2026 exel.xlsx"
FICHIER_SORTIE = "MDC_IMPORT_MARCHE_-_Lots_frais_2026_transforme.xlsx"
NOM_FEUILLE = "Feuille 1"
LIGNE_DEBUT_DONNEES = 3

COL_LOT = 2
COL_DESIGNATION = 3
COL_UNITE_REF = 4
COL_POIDS_REF = 5
COL_CONDITIONNEMENT = 8
COL_REF_PRODUIT = 10
COL_PRIX_REF = 11
COL_UNITE_COMMANDE = 12
COL_PRIX_COMMANDE = 13

DERNIERE_COLONNE_CLE_DOUBLON = 13

TOLERANCE_RATIO = 0.05   # 5% pour valider un candidat de LOT
TOLERANCE_SOURCES = 0.10  # 10% pour comparer E / Désignation / Conditionnement entre eux

COULEUR_LOT_CONFIRME = "FFD9EAD3"        # vert — LOT confirmé par le ratio M/K, rien recalculé
COULEUR_UNITE_TRANSFORMEE = "FFCFE2F3"   # bleu clair plus soutenu — K divisé par E, M inchangé
COULEUR_A_VERIFIER = "FFBDD7EE"          # bleu clair — pas de conversion fiable possible
COULEUR_RATIO_SUSPECT = "FFFCE5CD"       # orange — multiplicateur détecté mais ratio ne correspond à rien
COULEUR_COMMANDE_SUSPECTE = "FFFCE5CD"   # orange — unité de commande incohérente avec conditionnement/désignation
COULEUR_COUSTERON = "FFFFFF00"        # jaune — ligne "type cousteron" (conditionnement + référence vides)

# ============================================================================
# EXTRACTION DU TEXTE "CONDITIONNEMENT" / "DÉSIGNATION"
# ============================================================================

FACTEUR_VERS_KG_OU_L = {
    "kg": 1, "kilo": 1, "l": 1, "litre": 1,
    "g": 0.001, "gramme": 0.001,
    "mg": 0.000001, "cl": 0.01, "ml": 0.001,
}

RE_TRIPLE = re.compile(
    r"^\s*(\d+)\s*x\s*(\d+)\s*x\s*([\d,.]+)\s*(kg|kilo|g|gramme|l|litre|ml|cl|mg)\s*$", re.I)
RE_DOUBLE = re.compile(
    r"^\s*(\d+)\s*x\s*([\d,.]+)\s*(kg|kilo|g|gramme|l|litre|ml|cl|mg)\s*$", re.I)
RE_TYPE = re.compile(
    r"^\s*(sac|pi[eè]ce|seau|brique|barquette|bidon|outre|colis|pot|sachet|carton|poche)\s+(?:de\s+)?"
    r"([\d,.]+)\s*(kg|kilo|g|gramme|l|litre|ml|cl|mg)\s*$", re.I)
RE_LIBRE = re.compile(r"([\d,.]+)\s*(kg|kilo|g|gramme|l|litre|ml|cl|mg)\b", re.I)

RE_PIECES_DOUBLE = re.compile(
    r"^\s*(sac|seau|brique|barquette|bidon|outre|colis|pot|sachet|carton|poche)\s+de\s+"
    r"(\d+)\s*x\s*(\d+)\s*pi[eè]ces\s*$", re.I)
RE_PIECES_SIMPLE = re.compile(
    r"^\s*(sac|seau|brique|barquette|bidon|outre|colis|pot|sachet|carton|poche)\s+de\s+"
    r"(\d+)\s*pi[eè]ces\s*$", re.I)


def extraire_nombre_pieces(texte):
    """Nombre de pièces décrites dans le conditionnement (ex: 'Seau de 75 pièces' -> 75,
    'Colis de 4 x 10 pièces' -> 40)."""
    if not texte:
        return None
    texte = str(texte).strip()
    m = RE_PIECES_DOUBLE.match(texte)
    if m:
        return int(m.group(2)) * int(m.group(3))
    m = RE_PIECES_SIMPLE.match(texte)
    if m:
        return int(m.group(2))
    return None

PACKAGE_TYPES = ("barquette", "seau", "colis", "carton", "poche", "sac",
                  "bidon", "brique", "outre", "pot")
APPROXIMATIF = re.compile(r"\benv\b|\benviron\b", re.I)


def _nombre(s):
    return float(str(s).replace(",", "."))


def base_depuis_unite(unite):
    return "kg" if unite.lower() in ("kg", "kilo", "g", "gramme", "mg") else "l"


def normalise_texte(u):
    if u is None:
        return ""
    texto = str(u).strip().lower()
    substituicoes = str.maketrans("èéêëàâäîïôöûüç", "eeeeaaaiioouuc")
    return texto.translate(substituicoes)


def candidats_lot(texte):
    """
    Si le Conditionnement décrit un multiplicateur, renvoie la liste des
    valeurs candidates pour la taille du lot :
      "100 x 10g"        -> [100]
      "12 x 4 x 100g"     -> [12, 4, 48]
      "6 x 1L" / "6 x 1Kg" -> [6]
    Renvoie [] si aucun multiplicateur n'est détecté.
    """
    if not texte:
        return []
    texte = str(texte).strip()

    m = RE_TRIPLE.match(texte)
    if m:
        n1, n2 = int(m.group(1)), int(m.group(2))
        return [n1, n2, n1 * n2]

    m = RE_DOUBLE.match(texte)
    if m:
        n, qty, unite = m.groups()
        return [int(n)]

    return []


def formata_poids_kg(kg):
    """Formate un poids en kg avec 3 décimales (valeur numérique, ex: 0.3, 0.053)."""
    return round(kg, 3)


def peso_por_piece_kg(texte):
    """Poids par pièce, en kg, déduit d'un conditionnement à multiplicateur."""
    if not texte:
        return None
    texte = str(texte).strip()
    m = RE_TRIPLE.match(texte)
    if m:
        _, _, qty, unite = m.groups()
        return _nombre(qty) * FACTEUR_VERS_KG_OU_L[unite.lower()]
    m = RE_DOUBLE.match(texte)
    if m:
        _, qty, unite = m.groups()
        return _nombre(qty) * FACTEUR_VERS_KG_OU_L[unite.lower()]
    return None


def detecte_piece_de_poids(texte):
    """'Pièce de Xkg'/'Pièce de Xg' (uniquement le mot 'pièce', pas les autres types
    d'emballage) -> X en kg, sinon None."""
    if not texte:
        return None
    m = RE_TYPE.match(str(texte).strip())
    if m:
        mot, qty, unite = m.groups()
        if normalise_texte(mot) == "piece":
            return _nombre(qty) * FACTEUR_VERS_KG_OU_L[unite.lower()]
    return None


def extrai_numero_conditionnement(texte):
    """
    Cas d'un conditionnement qui est juste un nombre, avec ou sans unité collée :
      "10 KG"  -> (10.0, "kg")
      "5"      -> (5.0, "sans_unite")
    Renvoie (None, None) si rien de tel n'est trouvé.
    """
    if not texte:
        return None, None
    texte = str(texte).strip()

    m = re.match(r"^\s*([\d,.]+)\s*(kg|kilo|g|gramme|l|litre|ml|cl|mg)\s*$", texte, re.I)
    if m:
        qty, unite = m.groups()
        return _nombre(qty) * FACTEUR_VERS_KG_OU_L[unite.lower()], base_depuis_unite(unite)

    m = re.match(r"^\s*(\d+(?:[.,]\d+)?)\s*$", texte)
    if m:
        return _nombre(m.group(1)), "sans_unite"

    return None, None


def cherche_unite_designation(designation):
    """Cherche un mot-type de contenant dans la désignation (sac, sachet, pièce, barquette...).
    Cherche les mots les plus longs/spécifiques en premier (ex: 'sachet' avant 'sac', car
    'sac' est une sous-chaîne de 'sachet')."""
    if not designation:
        return None
    texte = normalise_texte(designation)
    mots = sorted(PACKAGE_TYPES + ("piece", "bouteille", "sachet"), key=len, reverse=True)
    for mot in mots:
        if mot in texte:
            return mot.upper()
    return None


def detecte_unite_directe_kg_l(texte):
    """Détecte le cas d'exception : 'N x 1kg' ou 'N x 1L' -> (N, 'KG' ou 'L')."""
    if not texte:
        return None, None
    m = RE_DOUBLE.match(str(texte).strip())
    if m:
        n, qty, unite = m.groups()
        if _nombre(qty) == 1 and unite.lower() in ("kg", "l", "litre"):
            return int(n), ("KG" if unite.lower() == "kg" else "L")
    return None, None


def extraire_quantite(texte):
    """(quantité_totale_en_kg_ou_l, unité_base, fiable) pour un conditionnement SANS
    multiplicateur (paquet unique) : 'Sac de 2,5Kg', 'Barquette de 300g'..."""
    if not texte:
        return None, None, False
    texte = str(texte).strip()

    m = RE_TYPE.match(texte)
    if m:
        _, qty, unite = m.groups()
        total = _nombre(qty)
        return total * FACTEUR_VERS_KG_OU_L[unite.lower()], base_depuis_unite(unite), True

    m = RE_LIBRE.search(texte)
    if m:
        qty, unite = m.groups()
        return _nombre(qty) * FACTEUR_VERS_KG_OU_L[unite.lower()], base_depuis_unite(unite), False

    return None, None, False


def unite_directe_si_quantite_1(texte):
    """Renvoie 'KG'/'L' si, dans le motif du conditionnement ('N x Qté(unité)' ou
    'N x M x Qté(unité)'), le nombre juste à côté de l'unité vaut 1 ET que cette unité est
    directement kg/L (pas g/ml à convertir) — ex: '6 x 1L', '6 x 1Kg'. Dans ce cas l'unité de
    référence doit rester kg/L (pas devenir PC), même si l'unité de commande devient LOT DE N.
    Renvoie None sinon (ex: '4 x 125g', '100 x 10g' -> le nombre à côté de l'unité n'est pas 1)."""
    if not texte:
        return None
    texte = str(texte).strip()
    m = RE_TRIPLE.match(texte)
    if m:
        qty, unite = m.group(3), m.group(4)
    else:
        m = RE_DOUBLE.match(texte)
        if m:
            qty, unite = m.group(2), m.group(3)
        else:
            return None
    if abs(_nombre(qty) - 1) < 0.01 and unite.lower() in ("kg", "l", "litre"):
        return "KG" if unite.lower() == "kg" else "L"
    return None


def mot_type_conditionnement(texte):
    """Le mot-type d'emballage capturé par RE_TYPE dans le conditionnement lui-même
    (ex: 'Sachet 500g' -> 'SACHET'), s'il y en a un."""
    if not texte:
        return None
    m = RE_TYPE.match(str(texte).strip())
    return m.group(1).upper() if m else None


def extraire_quantite_designation(designation):
    if not designation:
        return None, None
    m = RE_LIBRE.search(str(designation))
    if m:
        qty, unite = m.groups()
        return _nombre(qty) * FACTEUR_VERS_KG_OU_L[unite.lower()], base_depuis_unite(unite)
    return None, None


def resout_quantite_reference(poids_ref_num, qte_cond, fiable_cond, qte_desig):
    """Confronte jusqu'à 3 sources (E, Conditionnement à paquet unique, Désignation) et
    renvoie (quantite_retenue, avertissement)."""
    candidats = []
    if poids_ref_num:
        candidats.append(("E (poids de référence)", poids_ref_num))
    if fiable_cond and qte_cond:
        candidats.append(("Conditionnement", qte_cond))
    if qte_desig:
        candidats.append(("Désignation", qte_desig))

    if not candidats:
        return None, ""
    if len(candidats) == 1:
        return candidats[0][1], ""

    base = candidats[0][1]
    concordent = [c for c in candidats if abs(c[1] - base) <= TOLERANCE_SOURCES * base]
    divergent = [c for c in candidats if c not in concordent]

    if len(concordent) >= 2:
        valeur = sum(v for _, v in concordent) / len(concordent)
        aviso = ""
        if divergent:
            aviso = (" ATTENTION : " +
                      ", ".join(f"{n}={round(v, 3)}" for n, v in divergent) +
                      " ne correspond PAS aux autres sources qui concordent (" +
                      ", ".join(f"{n}={round(v, 3)}" for n, v in concordent) +
                      ") — probablement une erreur isolée sur cette source-là.")
        return valeur, aviso

    aviso = (" ATTENTION : les sources ne concordent pas du tout (" +
              ", ".join(f"{n}={round(v, 3)}" for n, v in candidats) +
              ") — à vérifier manuellement.")
    return (poids_ref_num or candidats[0][1]), aviso


# ============================================================================
# TRAITEMENT D'UNE LIGNE
# ============================================================================

class Resultat:
    def __init__(self):
        self.nouveau_d = None
        self.nouveau_e = "GARDER"   # "GARDER", "VIDER", ou une valeur
        self.nouveau_k = None
        self.nouveau_l = None
        self.commentaire_d = None
        self.commentaire_k = None
        self.commentaire_l = None
        self.commentaire_m = None
        self.couleur_d = None
        self.couleur_k = None
        self.couleur_l = None
        self.couleur_m = None


COULEUR_PRIX_OUTLIER = "FFEA9999"     # rouge clair — prix par kg très différent des produits comparables (pas utilisée pour l'instant)
COULEUR_PRIX_INEXPLIQUE = "FFFFFF00"  # jaune — même unité des deux côtés, ratio ≠ 1, rien dans le conditionnement ne l'explique


def extrai_premier_mot(designation):
    """Premier mot alphabétique de la désignation (ex: 'pâté de campagne' -> 'pate')."""
    if not designation:
        return None
    texte = normalise_texte(designation)
    m = re.match(r"[a-z]+", texte)
    if m and len(m.group(0)) >= 3:
        return m.group(0)
    return None


def analyse_coherence_prix_categorie(ws):
    """
    Compare le prix au kg de chaque produit à celui de produits similaires (même premier
    mot de désignation, même lot). S'il n'y a pas assez de produits similaires (< 3), utilise
    la moyenne de tout le lot comme repère à la place. Signale (sans corriger) les écarts
    de plus de 50% par rapport à la médiane du groupe utilisé.
    """
    import statistics

    donnees = []
    for linha in range(LIGNE_DEBUT_DONNEES, ws.max_row + 1):
        d = ws.cell(row=linha, column=COL_UNITE_REF).value
        k = ws.cell(row=linha, column=COL_PRIX_REF).value
        lot = ws.cell(row=linha, column=COL_LOT).value
        desig = ws.cell(row=linha, column=COL_DESIGNATION).value
        if d == "KG" and k and lot and desig:
            donnees.append((linha, lot, extrai_premier_mot(desig), k))

    grupos_especificos = {}
    for _, lot, palavra, k in donnees:
        if palavra:
            grupos_especificos.setdefault((lot, palavra), []).append(k)

    total_flags = 0
    for linha, lot, palavra, k in donnees:
        valores = grupos_especificos.get((lot, palavra)) if palavra else None
        if not valores or len(valores) < 3:
            continue  # pas assez de produits similaires pour comparer -> on ne signale rien
        grupo_usado = f"produits similaires ('{palavra}...', {len(valores)} produits)"

        mediane = statistics.median(valores)
        if mediane > 0 and abs(k - mediane) / mediane > 0.5:
                cel = ws.cell(row=linha, column=COL_PRIX_REF)
                cel.fill = PatternFill(start_color=COULEUR_PRIX_OUTLIER, end_color=COULEUR_PRIX_OUTLIER, fill_type="solid")
                cel.comment = Comment(
                    f"Prix au kg ({k} €) très différent de la médiane parmi les {grupo_usado} "
                    f"(médiane = {round(mediane, 2)} €). Écart de {round(abs(k - mediane) / mediane * 100)}%. "
                    "À vérifier — peut être justifié (qualité, origine, bio...), mais l'écart "
                    "est important.",
                    "Script de transformation")
                total_flags += 1
    return total_flags


def traite_ligne(unite_ref, poids_ref, designation, conditionnement, unite_cmd, prix_ref, prix_cmd):
    r = Resultat()
    ur = normalise_texte(unite_ref)
    uc = normalise_texte(unite_cmd)

    # --- Priorité 0 : D et L physiquement incompatibles (l'un en kg, l'autre en L) ---
    if (ur == "kg" and uc in ("l", "litre")) or (ur in ("l", "litre") and uc == "kg"):
        qte_cond, base_cond, fiable_cond = extraire_quantite(conditionnement)
        if not fiable_cond:
            valor_cond, base_cond2 = extrai_numero_conditionnement(conditionnement)
            if base_cond2 not in (None, "sans_unite"):
                qte_cond, base_cond, fiable_cond = valor_cond, base_cond2, True
        if not fiable_cond:
            n_direct, base_directe = detecte_unite_directe_kg_l(conditionnement)
            if n_direct:
                qte_cond, base_cond, fiable_cond = float(n_direct), base_directe.lower(), True

        if fiable_cond and qte_cond and prix_ref and prix_cmd:
            ratio = prix_cmd / prix_ref
            base_vraie = base_cond.upper()
            if abs(ratio - 1) <= TOLERANCE_RATIO or abs(ratio - qte_cond) <= TOLERANCE_RATIO * qte_cond:
                base_directe = unite_directe_si_quantite_1(conditionnement)
                if base_directe and qte_cond > 1 and abs(ratio - qte_cond) <= TOLERANCE_RATIO * qte_cond:
                    # "N x 1kg/1L" : le nombre à côté de l'unité vaut 1 -> l'unité de
                    # référence reste kg/L, seule l'unité de commande devient LOT DE N.
                    r.nouveau_d = base_directe
                    r.nouveau_l = f"LOT DE {int(qte_cond)}"
                    r.couleur_d = COULEUR_LOT_CONFIRME
                    r.couleur_l = COULEUR_LOT_CONFIRME
                    r.commentaire_d = (
                        f"Incohérence corrigée : unité de référence ('{unite_ref}') et unité "
                        f"de commande ('{unite_cmd}') étaient physiquement incompatibles "
                        f"(kg vs L). Le conditionnement ('{conditionnement}') confirme un lot "
                        f"de {int(qte_cond)} unités de 1{base_directe.lower()} — l'unité de "
                        f"référence reste '{base_directe}', l'unité de commande devient "
                        f"LOT DE {int(qte_cond)}. Prix inchangés.")
                    r.commentaire_l = r.commentaire_d
                    return r
                if unite_ref != base_vraie:
                    r.nouveau_d = base_vraie
                if unite_cmd != base_vraie:
                    r.nouveau_l = base_vraie
                r.couleur_d = COULEUR_UNITE_TRANSFORMEE
                r.couleur_l = COULEUR_UNITE_TRANSFORMEE
                r.commentaire_d = (
                    f"Incohérence corrigée : unité de référence ('{unite_ref}') et unité de "
                    f"commande ('{unite_cmd}') étaient physiquement incompatibles (kg vs L). "
                    f"Le conditionnement ('{conditionnement}') confirme la vraie base : "
                    f"'{base_vraie}'. Les deux unités sont uniformisées, prix inchangés.")
                r.commentaire_l = r.commentaire_d
                return r

        r.couleur_l = COULEUR_RATIO_SUSPECT
        r.commentaire_l = (
            f"Incohérence : unité de référence ('{unite_ref}') et unité de commande "
            f"('{unite_cmd}') sont physiquement incompatibles (kg vs L), et le conditionnement "
            f"('{conditionnement}') ne permet pas de confirmer laquelle est correcte. "
            "À vérifier manuellement.")
        return r

    # --- Priorité A : Conditionnement avec multiplicateur -> vérifie le ratio M/K ---
    candidats = candidats_lot(conditionnement)
    ratio_deja_coherent = False
    if candidats and prix_ref and prix_cmd:
        ratio = prix_cmd / prix_ref
        candidat_ok = next((c for c in candidats if c != 1 and abs(ratio - c) <= TOLERANCE_RATIO * c), None)

        if candidat_ok:
            base_directe = unite_directe_si_quantite_1(conditionnement)
            if base_directe:
                r.nouveau_d = base_directe
                r.nouveau_e = "VIDER" if poids_ref else "GARDER"
            else:
                r.nouveau_d = "PC"
                peso_kg = peso_por_piece_kg(conditionnement)
                r.nouveau_e = formata_poids_kg(peso_kg) if peso_kg else "VIDER"
            r.nouveau_l = f"LOT DE {candidat_ok}"
            r.couleur_d = COULEUR_LOT_CONFIRME
            r.couleur_l = COULEUR_LOT_CONFIRME
            r.commentaire_d = (
                f"Confirmé : le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} "
                f"= {round(ratio, 3)}) correspond au multiplicateur {candidat_ok} déduit du "
                f"conditionnement ('{conditionnement}'). "
                + (f"Le nombre à côté de l'unité vaut 1, donc l'unité de référence reste en "
                   f"'{r.nouveau_d}' — seule l'unité de commande précise le lot. " if base_directe else
                   "Le prix de référence est donc bien à la pièce/au pot. ")
                + "Aucun prix modifié.")
            r.commentaire_l = r.commentaire_d
            return r
        elif abs(ratio - 1) <= TOLERANCE_RATIO:
            # K et M sont déjà cohérents entre eux (ratio ~1) — rien d'anormal, on continue
            # simplement vers la normalisation habituelle (pas de signalement nécessaire).
            ratio_deja_coherent = True
        else:
            r.couleur_d = COULEUR_RATIO_SUSPECT
            r.commentaire_d = (
                f"Le conditionnement ('{conditionnement}') suggère un multiplicateur parmi "
                f"{candidats}, mais le ratio réel Prix commande ÷ Prix référence "
                f"({prix_cmd} ÷ {prix_ref} = {round(ratio, 3) if prix_ref else '?'}) ne "
                f"correspond à aucun d'eux. À vérifier manuellement — aucun prix modifié.")
            return r

    # --- Priorité B : unité déjà KG/L, ou convertible via E ---
    designation_approximative = bool(APPROXIMATIF.search(str(designation or "")))

    if ur == "kg" or ur in ("l", "litre"):
        base_actuelle = "KG" if ur == "kg" else "L"
        if unite_ref != base_actuelle:
            r.nouveau_d = base_actuelle  # normalisation de casse seule

        k_final = prix_ref

        if not ratio_deja_coherent and poids_ref and not designation_approximative:
            try:
                poids = float(str(poids_ref).replace(",", "."))
                if poids > 0 and prix_ref:
                    nouveau_k = round(prix_ref / poids, 4)
                    r.nouveau_d = base_actuelle
                    r.nouveau_k = nouveau_k
                    r.nouveau_e = "VIDER"
                    r.couleur_d = COULEUR_UNITE_TRANSFORMEE
                    r.couleur_k = COULEUR_UNITE_TRANSFORMEE
                    r.commentaire_k = (
                        f"Le poids de référence (E={poids}) indique que le prix de référence "
                        f"({prix_ref}) était en fait donné pour {poids} {base_actuelle.lower()}, "
                        f"pas pour 1. Prix normalisé = {prix_ref} / {poids} = {nouveau_k} "
                        f"€/{base_actuelle.lower()}. Le prix de commande ({prix_cmd}) n'est "
                        "PAS modifié — il reste le prix réel de commande.")
                    k_final = nouveau_k
            except (TypeError, ValueError):
                pass
        elif ratio_deja_coherent and poids_ref:
            r.nouveau_e = "VIDER"

        # Règle : le ratio observé décide tout. Si ratio ≈ 1, tout va bien (même base) — sauf
        # si l'unité de commande n'était pas déjà kg/l, auquel cas on la corrige. Si ratio ≠ 1,
        # ce n'est PAS un prix "par kg/L" : on cherche ce que le conditionnement/désignation dit
        # sur l'emballage réel et on relabelle l'unité de commande en conséquence (peu importe
        # que le nombre trouvé ait une unité collée ou non — "10 KG" ou juste "10", même logique).
        if r.nouveau_k is None and not ratio_deja_coherent and prix_ref and prix_cmd:
            ratio_actuel = prix_cmd / prix_ref

            if abs(ratio_actuel - 1) <= TOLERANCE_RATIO:
                if normalise_texte(unite_cmd) not in ("kg", "l", "litre"):
                    r.nouveau_l = base_actuelle
                    r.couleur_l = COULEUR_UNITE_TRANSFORMEE
                    r.commentaire_l = (
                        f"Unité de commande corrigée de '{unite_cmd}' vers '{base_actuelle}' : "
                        f"le ratio prix commande ÷ prix référence est de 1 (prix identiques : "
                        f"{prix_cmd} = {prix_ref}), donc même base que la référence. Le libellé "
                        f"d'origine semblait erroné. Prix inchangés.")
                # sinon : ratio déjà cohérent et unité déjà correcte -> rien à signaler.
            else:
                valor_cond, base_cond2 = extrai_numero_conditionnement(conditionnement)
                if valor_cond is None:
                    qte_type, base_type, fiable_type = extraire_quantite(conditionnement)
                    if fiable_type:
                        valor_cond, base_cond2 = qte_type, base_type

                if valor_cond is not None and abs(ratio_actuel - valor_cond) <= TOLERANCE_RATIO * valor_cond:
                    unite_trouvee = mot_type_conditionnement(conditionnement) or cherche_unite_designation(designation)
                    r.nouveau_l = unite_trouvee or "SAC"
                    r.couleur_l = COULEUR_UNITE_TRANSFORMEE
                    r.commentaire_l = (
                        f"Unité de commande précisée en '{r.nouveau_l}' : le ratio prix "
                        f"commande/prix référence ({round(ratio_actuel, 3)}) correspond au "
                        f"conditionnement ('{conditionnement}'), donc le prix de commande "
                        f"correspond à cet emballage, pas au kilo/litre."
                        + (f" Mot trouvé dans la désignation." if unite_trouvee else
                           " Aucun mot de contenant trouvé dans la désignation — 'SAC' utilisé "
                           "par défaut, à confirmer."))
                elif valor_cond is not None:
                    r.couleur_l = COULEUR_RATIO_SUSPECT
                    r.commentaire_l = (
                        f"Doute sur l'unité de commande : le ratio observé entre les prix "
                        f"({prix_cmd} ÷ {prix_ref} = {round(ratio_actuel, 3)}) ne correspond "
                        f"PAS au nombre trouvé dans le conditionnement ('{conditionnement}' -> "
                        f"{valor_cond}). À vérifier manuellement.")
                else:
                    # Même unité des deux côtés, ratio ≠ 1, et rien dans le conditionnement (ni
                    # nombre, ni poids reconnaissable) n'explique ce ratio — cas typique des
                    # lots viande/poisson (calibres, fourchettes de poids en texte libre...).
                    r.couleur_m = COULEUR_PRIX_INEXPLIQUE
                    r.commentaire_m = (
                        f"À quoi correspond ce prix ? Le prix de référence ({prix_ref}) et le "
                        f"prix de commande ({prix_cmd}) sont dans la même unité "
                        f"('{base_actuelle}') mais ne sont pas égaux (ratio = "
                        f"{round(ratio_actuel, 3)}), et rien dans le conditionnement "
                        f"('{conditionnement}') ne permet d'expliquer cet écart (probablement "
                        "un prix par calibre/taille, à confirmer manuellement).")

        # Règle : "Pièce de Xkg" -> si le ratio (prix commande / prix référence final)
        # correspond au poids de la pièce, l'unité de commande devient "PIECE"
        x_piece = detecte_piece_de_poids(conditionnement)
        if x_piece and k_final and prix_cmd:
            ratio_piece = prix_cmd / k_final
            if abs(ratio_piece - x_piece) <= TOLERANCE_RATIO * x_piece:
                r.nouveau_l = "PIECE"
                r.couleur_l = COULEUR_UNITE_TRANSFORMEE
                r.commentaire_l = (
                    f"Unité de commande uniformisée en 'PIECE' : le ratio prix commande ÷ prix "
                    f"référence ({prix_cmd} ÷ {k_final} = {round(ratio_piece, 3)}) correspond au "
                    f"poids de la pièce décrit dans le conditionnement ('{conditionnement}' = "
                    f"{x_piece} kg).")

        return r

    if ur in ("unite", "piece"):
        poids_ref_num = None
        if poids_ref:
            try:
                poids_ref_num = float(str(poids_ref).replace(",", "."))
            except (TypeError, ValueError):
                poids_ref_num = None

        qte_cond, base_cond, fiable_cond = extraire_quantite(conditionnement)
        qte_desig, base_desig = extraire_quantite_designation(designation)

        # Exception : si le prix de commande est déjà égal au prix de référence (comme pour
        # une baguette : K=M=0,85), la personne compare déjà "à la pièce" volontairement —
        # on ne force pas la conversion en KG, on uniformise juste en PC avec le poids en
        # kg (3 décimales), sans toucher aux prix.
        if prix_ref and prix_cmd and abs(prix_ref - prix_cmd) <= TOLERANCE_RATIO * max(prix_ref, prix_cmd):
            poids_kg = None
            if poids_ref_num:
                poids_kg = poids_ref_num
            elif qte_desig:
                poids_kg = qte_desig
            elif fiable_cond and qte_cond:
                poids_kg = qte_cond
            r.nouveau_d = "PC"
            if poids_kg:
                r.nouveau_e = formata_poids_kg(poids_kg)
            r.couleur_d = COULEUR_UNITE_TRANSFORMEE
            r.commentaire_d = (
                f"Unité uniformisée de '{unite_ref}' vers 'PC' : le prix de commande "
                f"({prix_cmd}) est déjà égal au prix de référence ({prix_ref}), signe que la "
                "comparaison se fait volontairement à la pièce (ex : une baguette). Prix "
                "inchangés.")
            return r

        if not designation_approximative:
            poids, aviso = resout_quantite_reference(poids_ref_num, qte_cond, fiable_cond, qte_desig)
            if poids and poids > 0 and prix_ref:
                nouveau_k = round(prix_ref / poids, 4)
                r.nouveau_d = "KG"
                r.nouveau_k = nouveau_k
                r.nouveau_e = "VIDER"
                r.couleur_d = COULEUR_UNITE_TRANSFORMEE
                r.couleur_k = COULEUR_UNITE_TRANSFORMEE
                r.commentaire_k = (
                    f"Unité convertie de '{unite_ref}' vers 'KG' : poids de référence retenu "
                    f"= {round(poids, 4)} kg. Prix normalisé = {prix_ref} / {round(poids, 4)} "
                    f"= {nouveau_k} €/kg.{aviso} Le prix de commande n'est PAS modifié.")
                return r

        r.nouveau_d = "PC"
        r.couleur_d = COULEUR_A_VERIFIER
        r.commentaire_d = (
            f"Unité uniformisée de '{unite_ref}' vers 'PC' — aucun poids fiable trouvé pour "
            "convertir en KG (poids de référence manquant, ou désignation approximative). "
            "Aucun prix modifié.")
        # Règle : toute unité PC doit afficher un poids en kg (3 décimales) s'il y en a un connu
        if poids_ref_num:
            r.nouveau_e = formata_poids_kg(poids_ref_num)
        return r

    if ur in PACKAGE_TYPES:
        poids_ref_num = None
        if poids_ref:
            try:
                poids_ref_num = float(str(poids_ref).replace(",", "."))
            except (TypeError, ValueError):
                poids_ref_num = None

        nb_pieces = extraire_nombre_pieces(conditionnement)

        # Cas "œufs" : le conditionnement décrit un NOMBRE DE PIÈCES (pas un poids total).
        # Dans ce cas, la comparaison de marché se fait à la pièce, pas au kg : le prix
        # affiché (K=M) est le prix du contenant entier -> devient le prix de commande,
        # et le prix de référence (par pièce) = prix de commande ÷ nombre de pièces.
        if nb_pieces and prix_cmd:
            nouveau_k = round(prix_cmd / nb_pieces, 4)
            r.nouveau_d = "PC"
            r.nouveau_k = nouveau_k
            r.nouveau_l = unite_ref  # garde le type de contenant d'origine (Seau, Colis...)
            if poids_ref_num:
                r.nouveau_e = formata_poids_kg(poids_ref_num)
            r.couleur_d = COULEUR_UNITE_TRANSFORMEE
            r.couleur_k = COULEUR_UNITE_TRANSFORMEE
            r.couleur_l = COULEUR_UNITE_TRANSFORMEE
            r.commentaire_k = (
                f"Le conditionnement ('{conditionnement}') décrit {nb_pieces} pièces, pas un "
                f"poids total — comparaison au marché à la pièce. Prix de référence par pièce = "
                f"prix de commande ÷ nombre de pièces = {prix_cmd} / {nb_pieces} = {nouveau_k} "
                f"€/pièce. Le prix de commande n'est PAS modifié ; l'unité de commande reprend "
                f"le contenant d'origine ('{unite_ref}').")
            r.commentaire_l = r.commentaire_k
            return r

        qte_cond, base_cond, fiable_cond = extraire_quantite(conditionnement)
        qte_desig, base_desig = extraire_quantite_designation(designation)
        poids, aviso = resout_quantite_reference(poids_ref_num, qte_cond, fiable_cond, qte_desig)
        base = base_cond or base_desig or "kg"

        if poids and poids > 0 and prix_ref:
            nouveau_k = round(prix_ref / poids, 4)
            r.nouveau_d = base.upper()
            r.nouveau_k = nouveau_k
            r.nouveau_e = "VIDER"
            r.couleur_d = COULEUR_UNITE_TRANSFORMEE
            r.couleur_k = COULEUR_UNITE_TRANSFORMEE
            r.commentaire_k = (
                f"Unité convertie de '{unite_ref}' vers '{base.upper()}' : quantité de "
                f"référence retenue = {round(poids, 4)} {base}. Prix normalisé = {prix_ref} / "
                f"{round(poids, 4)} = {nouveau_k} €/{base}.{aviso} "
                "Le prix de commande n'est PAS modifié.")
        else:
            r.couleur_d = COULEUR_A_VERIFIER
            r.commentaire_d = (
                f"Reste en '{unite_ref}' : pas de quantité fiable trouvée pour convertir en "
                "KG/L. Aucun prix modifié.")
        return r

    r.couleur_d = COULEUR_A_VERIFIER
    r.commentaire_d = f"Unité de référence inhabituelle ('{unite_ref}') — non convertie automatiquement."
    return r


# ============================================================================
# ÉTAPE — COHÉRENCE UNITÉ DE COMMANDE <-> CONDITIONNEMENT / DÉSIGNATION
# ============================================================================

CATEGORIES_COMMANDE = {
    "kg": "kg", "pu": "pc", "pot": "pc", "l": "l", "piece": "pc", "pièce": "pc",
    "bqt": "barquette", "sac": "sac", "bt.": "bouteille", "bt": "bouteille",
}


def categorie_commande(unite_cmd):
    return CATEGORIES_COMMANDE.get(normalise_texte(unite_cmd), normalise_texte(unite_cmd))


def verifie_unite_commande(unite_cmd, conditionnement, designation):
    if not unite_cmd:
        return False, "Unité de commande vide."

    cat = categorie_commande(unite_cmd)
    cond = normalise_texte(conditionnement)
    desig = normalise_texte(designation)
    texte_combine = f"{cond} {desig}"

    if not conditionnement and not designation:
        return True, None
    if cat in ("kg", "l"):
        return True, None

    mots_types = ["sac", "pièce", "piece", "seau", "brique", "barquette", "bidon",
                  "outre", "colis", "pot", "sachet", "carton", "poche", "bouteille"]
    trouve = [m for m in mots_types if m in texte_combine]
    if not trouve:
        return True, None

    if cat == "pc" and any(m in ("pièce", "piece", "pot", "sachet") for m in trouve):
        return True, None
    if cat == "barquette" and "barquette" in trouve:
        return True, None
    if cat == "sac" and "sac" in trouve:
        return True, None
    if cat == "bouteille" and any(m in ("bouteille", "brique") for m in trouve):
        return True, None

    source = "Conditionnement" if any(m in cond for m in trouve) else "Désignation de l'article"
    texte_source = conditionnement if source == "Conditionnement" else designation
    return False, (f"L'unité de commande ('{unite_cmd}') ne semble pas correspondre au "
                    f"type de contenant décrit dans « {source} » ('{texte_source}').")


# ============================================================================
# NETTOYAGE PRÉALABLE : LIGNES VIDES ET DOUBLONS
# ============================================================================

def normalise_pour_doublon(valeur):
    """Normalise une valeur pour comparer deux lignes comme doublons : texte en minuscules,
    espaces superflus retirés, et ligatures comme 'œ'/'æ' éclatées en 'oe'/'ae' (sinon
    'blancs d'œufs' et 'blancs d'oeufs' ne sont pas reconnus comme identiques)."""
    if isinstance(valeur, str):
        texte = valeur.strip().lower()
        texte = texte.replace("œ", "oe").replace("æ", "ae")
        texte = re.sub(r"\s+", " ", texte)
        return texte
    return valeur


RE_NUMERO_LOT_SEUL = re.compile(r"\b(\d)\b")


def normalise_numero_lot(valeur):
    """Ajoute un zéro devant chaque numéro de lot à un seul chiffre (1-9 -> 01-09),
    pour que le tri alphabétique de la colonne corresponde à l'ordre numérique."""
    if not isinstance(valeur, str):
        return valeur
    return RE_NUMERO_LOT_SEUL.sub(lambda m: f"0{m.group(1)}", valeur)


def trouve_lignes_a_supprimer(ws, derniere_ligne):
    a_supprimer = set()
    lignes_vues = {}

    for r in range(LIGNE_DEBUT_DONNEES, derniere_ligne + 1):
        designation = ws.cell(row=r, column=COL_DESIGNATION).value
        reference = ws.cell(row=r, column=COL_REF_PRODUIT).value

        if not designation and not reference:
            a_supprimer.add(r)
            continue

        # Aussi : pas de désignation ET aucun des deux prix -> ligne inutilisable
        prix_ref_v = ws.cell(row=r, column=COL_PRIX_REF).value
        prix_cmd_v = ws.cell(row=r, column=COL_PRIX_COMMANDE).value
        if not designation and not prix_ref_v and not prix_cmd_v:
            a_supprimer.add(r)
            continue

        cle = tuple(
            normalise_pour_doublon(ws.cell(row=r, column=c).value)
            for c in range(1, DERNIERE_COLONNE_CLE_DOUBLON + 1)
        )
        if any(v is not None for v in cle):
            if cle in lignes_vues:
                a_supprimer.add(r)
            else:
                lignes_vues[cle] = r

    return a_supprimer


# ============================================================================
# PROGRAMME PRINCIPAL
# ============================================================================

COULEUR_COPIE_POSSIBLE = "FFD5A6BD"   # rose — même Référence produit + même prix de référence (possible copie)


def signale_copies_possibles(ws):
    """Surligne les groupes de lignes qui partagent la même Référence produit ET le même
    prix de référence (K) — signe probable que l'une est une copie/variante de l'autre
    (ex: faute de frappe dans la désignation), même si le reste diffère."""
    groupes = {}
    for linha in range(LIGNE_DEBUT_DONNEES, ws.max_row + 1):
        ref = ws.cell(row=linha, column=COL_REF_PRODUIT).value
        k = ws.cell(row=linha, column=COL_PRIX_REF).value
        if ref and k:
            groupes.setdefault((ref, k), []).append(linha)

    total = 0
    for (ref, k), lignes in groupes.items():
        if len(lignes) < 2:
            continue
        for linha in lignes:
            autres = [str(l) for l in lignes if l != linha]
            cel = ws.cell(row=linha, column=COL_DESIGNATION)
            cel.fill = PatternFill(start_color=COULEUR_COPIE_POSSIBLE, end_color=COULEUR_COPIE_POSSIBLE, fill_type="solid")
            cel.comment = Comment(
                f"Copie possible : même Référence produit ({ref}) et même prix de référence "
                f"({k}) que la ligne {' et '.join(autres)}. À comparer pour voir laquelle est "
                "correcte (peut être une simple faute de frappe dans la désignation).",
                "Script de transformation")
            total += 1
    return total


def main():
    wb_valeurs = openpyxl.load_workbook(FICHIER_ENTREE, data_only=True)
    ws_valeurs = wb_valeurs[NOM_FEUILLE]
    wb = openpyxl.load_workbook(FICHIER_ENTREE, data_only=False)
    ws = wb[NOM_FEUILLE]

    # Gèle toutes les formules en valeurs AVANT toute suppression de ligne
    for linha in range(LIGNE_DEBUT_DONNEES, ws.max_row + 1):
        for coluna in range(1, ws.max_column + 1):
            cel = ws.cell(row=linha, column=coluna)
            if isinstance(cel.value, str) and cel.value.startswith("="):
                cel.value = ws_valeurs.cell(row=linha, column=coluna).value

    derniere_ligne = LIGNE_DEBUT_DONNEES
    r = LIGNE_DEBUT_DONNEES
    while r <= ws_valeurs.max_row:
        if ws_valeurs.cell(row=r, column=COL_DESIGNATION).value is not None:
            derniere_ligne = r
        elif all(ws_valeurs.cell(row=r + i, column=COL_DESIGNATION).value is None for i in range(20)):
            break
        r += 1

    a_supprimer = trouve_lignes_a_supprimer(ws_valeurs, ws_valeurs.max_row)
    for ligne_no in sorted(a_supprimer, reverse=True):
        ws.delete_rows(ligne_no, 1)
        ws_valeurs.delete_rows(ligne_no, 1)

    # Remplit les Lot/catégorie manquants (colonne B) quand la ligne du dessus et celle du
    # dessous ont la même valeur — la ligne "sandwich" reçoit cette même valeur.
    for linha in range(LIGNE_DEBUT_DONNEES + 1, ws.max_row):
        if ws.cell(row=linha, column=COL_LOT).value is None:
            valeur_dessus = ws.cell(row=linha - 1, column=COL_LOT).value
            valeur_dessous = ws.cell(row=linha + 1, column=COL_LOT).value
            if valeur_dessus is not None and valeur_dessus == valeur_dessous:
                ws.cell(row=linha, column=COL_LOT).value = valeur_dessus
                ws_valeurs.cell(row=linha, column=COL_LOT).value = valeur_dessus

    # Numéro de lot à un seul chiffre (1-9) -> zéro devant (01, 02...), pour que le tri
    # alphabétique corresponde à l'ordre numérique. Gère aussi les cellules avec plusieurs
    # numéros (ex: "9 & 10 - viandes fraîches" -> "09 & 10 - viandes fraîches").
    for linha in range(LIGNE_DEBUT_DONNEES, ws.max_row + 1):
        valeur = ws.cell(row=linha, column=COL_LOT).value
        nouvelle_valeur = normalise_numero_lot(valeur)
        if nouvelle_valeur != valeur:
            ws.cell(row=linha, column=COL_LOT).value = nouvelle_valeur

    stats = {"lignes": 0, "lignes_supprimees": len(a_supprimer),
              "lot_confirme": 0, "ratio_suspect": 0, "unite_transformee": 0,
              "a_verifier": 0, "commande_suspecte": 0, "prix_inexplique": 0}

    for ligne in range(LIGNE_DEBUT_DONNEES, ws_valeurs.max_row + 1):
        designation = ws_valeurs.cell(row=ligne, column=COL_DESIGNATION).value
        unite_ref = ws_valeurs.cell(row=ligne, column=COL_UNITE_REF).value
        if not designation and not unite_ref:
            continue

        stats["lignes"] += 1

        poids_ref = ws_valeurs.cell(row=ligne, column=COL_POIDS_REF).value
        conditionnement = ws_valeurs.cell(row=ligne, column=COL_CONDITIONNEMENT).value
        reference = ws_valeurs.cell(row=ligne, column=COL_REF_PRODUIT).value
        unite_cmd = ws_valeurs.cell(row=ligne, column=COL_UNITE_COMMANDE).value
        prix_ref = ws_valeurs.cell(row=ligne, column=COL_PRIX_REF).value
        prix_cmd = ws_valeurs.cell(row=ligne, column=COL_PRIX_COMMANDE).value

        if not conditionnement and not reference:
            for coluna in range(1, ws.max_column + 1):
                ws.cell(row=ligne, column=coluna).fill = PatternFill(
                    start_color=COULEUR_COUSTERON, end_color=COULEUR_COUSTERON, fill_type="solid")
            ws.cell(row=ligne, column=COL_DESIGNATION).comment = Comment(
                "Conditionnement ET Référence produit vides en même temps — ligne "
                "incomplète, à vérifier/compléter manuellement.",
                "Script de transformation")
            stats["a_verifier"] += 1

        if prix_ref is None:
            cel = ws.cell(row=ligne, column=COL_PRIX_REF)
            cel.fill = PatternFill(start_color=COULEUR_A_VERIFIER, end_color=COULEUR_A_VERIFIER, fill_type="solid")
            cel.comment = Comment("Prix de référence manquant.", "Script de transformation")
            stats["a_verifier"] += 1
        if prix_cmd is None:
            cel = ws.cell(row=ligne, column=COL_PRIX_COMMANDE)
            cel.fill = PatternFill(start_color=COULEUR_A_VERIFIER, end_color=COULEUR_A_VERIFIER, fill_type="solid")
            cel.comment = Comment("Prix de commande manquant.", "Script de transformation")
            stats["a_verifier"] += 1

        resultat = None
        if unite_ref:
            resultat = traite_ligne(unite_ref, poids_ref, designation, conditionnement,
                                     unite_cmd, prix_ref, prix_cmd)
        else:
            cel_d = ws.cell(row=ligne, column=COL_UNITE_REF)
            cel_d.fill = PatternFill(start_color=COULEUR_A_VERIFIER, end_color=COULEUR_A_VERIFIER, fill_type="solid")
            cel_d.comment = Comment(
                "Unité de référence vide — à compléter manuellement.",
                "Script de transformation")
            stats["a_verifier"] += 1

        if resultat:
            if resultat.nouveau_d is not None:
                ws.cell(row=ligne, column=COL_UNITE_REF).value = resultat.nouveau_d
            if resultat.couleur_d or resultat.commentaire_d:
                cel_d = ws.cell(row=ligne, column=COL_UNITE_REF)
                if resultat.couleur_d:
                    cel_d.fill = PatternFill(start_color=resultat.couleur_d, end_color=resultat.couleur_d, fill_type="solid")
                if resultat.commentaire_d:
                    cel_d.comment = Comment(resultat.commentaire_d, "Script de transformation")

            if resultat.nouveau_k is not None:
                ws.cell(row=ligne, column=COL_PRIX_REF).value = resultat.nouveau_k
            if resultat.couleur_k or resultat.commentaire_k:
                cel_k = ws.cell(row=ligne, column=COL_PRIX_REF)
                if resultat.couleur_k:
                    cel_k.fill = PatternFill(start_color=resultat.couleur_k, end_color=resultat.couleur_k, fill_type="solid")
                if resultat.commentaire_k:
                    cel_k.comment = Comment(resultat.commentaire_k, "Script de transformation")

            if resultat.nouveau_l is not None:
                ws.cell(row=ligne, column=COL_UNITE_COMMANDE).value = resultat.nouveau_l
            if resultat.couleur_l or resultat.commentaire_l:
                cel_l = ws.cell(row=ligne, column=COL_UNITE_COMMANDE)
                if resultat.couleur_l:
                    cel_l.fill = PatternFill(start_color=resultat.couleur_l, end_color=resultat.couleur_l, fill_type="solid")
                if resultat.commentaire_l:
                    cel_l.comment = Comment(resultat.commentaire_l, "Script de transformation")

            if resultat.couleur_m or resultat.commentaire_m:
                cel_m = ws.cell(row=ligne, column=COL_PRIX_COMMANDE)
                if resultat.couleur_m:
                    cel_m.fill = PatternFill(start_color=resultat.couleur_m, end_color=resultat.couleur_m, fill_type="solid")
                    if resultat.couleur_m == COULEUR_PRIX_INEXPLIQUE:
                        stats["prix_inexplique"] += 1
                if resultat.commentaire_m:
                    cel_m.comment = Comment(resultat.commentaire_m, "Script de transformation")

            if resultat.nouveau_e == "VIDER":
                ws.cell(row=ligne, column=COL_POIDS_REF).value = None
            elif resultat.nouveau_e != "GARDER":
                ws.cell(row=ligne, column=COL_POIDS_REF).value = resultat.nouveau_e

            if resultat.couleur_d == COULEUR_LOT_CONFIRME:
                stats["lot_confirme"] += 1
            elif resultat.couleur_d == COULEUR_RATIO_SUSPECT:
                stats["ratio_suspect"] += 1
            elif resultat.couleur_d == COULEUR_UNITE_TRANSFORMEE:
                stats["unite_transformee"] += 1
            elif resultat.couleur_d == COULEUR_A_VERIFIER:
                stats["a_verifier"] += 1

        # --- Cohérence unité de commande <-> conditionnement/désignation ---
        # (sautée si la ligne vient d'être traitée en "LOT DE N", où L a été redéfini exprès)
        if unite_cmd and not (resultat and resultat.nouveau_l is not None):
            ok, message = verifie_unite_commande(unite_cmd, conditionnement, designation)
            if not ok:
                stats["commande_suspecte"] += 1
                cel_l = ws.cell(row=ligne, column=COL_UNITE_COMMANDE)
                cel_l.fill = PatternFill(start_color=COULEUR_COMMANDE_SUSPECTE, end_color=COULEUR_COMMANDE_SUSPECTE, fill_type="solid")
                cel_l.comment = Comment(message, "Script de transformation")

    # Analyse de cohérence des prix par catégorie : désactivée pour l'instant, en attendant
    # que les données soient bien nettoyées d'abord (à reprendre plus tard).
    # stats["prix_outlier"] = analyse_coherence_prix_categorie(ws)

    stats["copies_possibles"] = signale_copies_possibles(ws)

    # --- Sécurité : toute ligne totalement vide (aucun contenu dans aucune colonne) perd
    # toute sa mise en forme (couleur, bordure, format) — elle ne doit garder aucune trace
    # visuelle si elle n'a plus rien d'écrit.
    for linha in range(LIGNE_DEBUT_DONNEES, ws.max_row + 1):
        if all(ws.cell(row=linha, column=c).value is None for c in range(1, ws.max_column + 1)):
            for c in range(1, ws.max_column + 1):
                cel = ws.cell(row=linha, column=c)
                cel.fill = PatternFill(fill_type=None)
                cel.number_format = "General"
                cel.comment = None

    # --- Mise en forme : Calibri 11 noir + alignement + format numérique, à partir ligne 3 ---
    from openpyxl.styles import Font, Alignment, Border, Side
    fonte_padrao = Font(name="Calibri", size=11, color="FF000000")
    alinh_esquerda = Alignment(horizontal="left")
    alinh_centro = Alignment(horizontal="center")
    alinh_direita = Alignment(horizontal="right")
    borda_fina = Border(
        left=Side(style="thin", color="FF000000"), right=Side(style="thin", color="FF000000"),
        top=Side(style="thin", color="FF000000"), bottom=Side(style="thin", color="FF000000"))

    colunas_centro = (COL_UNITE_REF, COL_POIDS_REF, COL_LOT_FOURNISSEUR := 7, COL_UNITE_COMMANDE)
    colunas_preco = (COL_PRIX_REF, COL_PRIX_COMMANDE)
    cores_destacadas = {
        COULEUR_LOT_CONFIRME, COULEUR_UNITE_TRANSFORMEE, COULEUR_A_VERIFIER,
        COULEUR_RATIO_SUSPECT, COULEUR_COMMANDE_SUSPECTE, COULEUR_COUSTERON, COULEUR_PRIX_OUTLIER,
        COULEUR_PRIX_INEXPLIQUE, COULEUR_COPIE_POSSIBLE,
    }
    sem_preenchimento = PatternFill(fill_type=None)

    for linha in range(LIGNE_DEBUT_DONNEES, ws.max_row + 1):
        linha_tem_conteudo = any(
            ws.cell(row=linha, column=c).value is not None for c in range(1, ws.max_column + 1))
        for coluna in range(1, ws.max_column + 1):
            cel = ws.cell(row=linha, column=coluna)
            cel.font = fonte_padrao
            if linha_tem_conteudo:
                cel.border = borda_fina

            # Ne garde que les couleurs que le script lui-même a posées ; toute autre
            # couleur (résidu du fichier d'origine, formatage antérieur...) est retirée.
            if cel.fill and cel.fill.patternType:
                try:
                    cor_atual = cel.fill.fgColor.rgb
                except Exception:
                    cor_atual = None
                if cor_atual not in cores_destacadas:
                    cel.fill = sem_preenchimento

            # Unité de commande : corrections d'abréviation + tout en majuscules
            if coluna == COL_UNITE_COMMANDE and isinstance(cel.value, str):
                valeur = cel.value.strip()
                if valeur.lower() == "bt.":
                    valeur = "BQT"
                cel.value = valeur.upper()

            # Unité de référence : tout en majuscules (couvre kg->KG, l->L, etc. directement)
            if coluna == COL_UNITE_REF and isinstance(cel.value, str):
                cel.value = cel.value.upper()

            if coluna in colunas_centro:
                cel.alignment = alinh_centro
                if coluna == COL_POIDS_REF and isinstance(cel.value, (int, float)):
                    cel.number_format = "0.000"
            elif coluna in colunas_preco:
                cel.alignment = alinh_direita
                cel.number_format = '#,##0.00" €"'
            elif coluna == COL_REF_PRODUIT:
                cel.alignment = alinh_esquerda
                if isinstance(cel.value, (int, float)):
                    cel.number_format = "0"
            else:
                cel.alignment = alinh_esquerda

    # Nettoie les entrées de style "orphelines" (lignes sans aucune cellule, mais avec un
    # style posé au niveau de la ligne elle-même) qui peuvent subsister après delete_rows()
    # au-delà de la dernière ligne réellement utilisée.
    for numero_ligne in list(ws.row_dimensions.keys()):
        if numero_ligne > ws.max_row:
            del ws.row_dimensions[numero_ligne]

    wb.save(FICHIER_SORTIE)

    print("=== Résumé ===")
    print(f"Lignes supprimées (vides ou dupliquées) : {stats['lignes_supprimees']}")
    print(f"Lignes analysées ensuite : {stats['lignes']}")
    print(f"LOT confirmé, rien recalculé (vert) : {stats['lot_confirme']}")
    print(f"Ratio suspect, multiplicateur ne correspond pas (orange) : {stats['ratio_suspect']}")
    print(f"Unité transformée, K divisé par E, M inchangé (bleu soutenu) : {stats['unite_transformee']}")
    print(f"À vérifier manuellement (bleu clair) : {stats['a_verifier']}")
    print(f"Unité de commande suspecte (orange) : {stats['commande_suspecte']}")
    print(f"Prix inexpliqué, 'à quoi correspond ce prix ?' (jaune) : {stats['prix_inexplique']}")
    print(f"Copies possibles (rose) : {stats['copies_possibles']}")
    print(f"Fichier enregistré : {FICHIER_SORTIE}")
    print("Le fichier original n'a PAS été modifié.")


if __name__ == "__main__":
    main()