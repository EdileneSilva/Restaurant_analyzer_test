"""
Audit de cohérence unités/prix — GÉNÉRALISÉ à tous les fichiers "Base article"
================================================================================

Ce script traite automatiquement TOUS les fichiers .xlsx du dossier raw_data/, quel que
soit leur format exact (nombre de colonnes, présence ou non d'une colonne "Circuit court",
1, 2 ou 3 blocs Fournisseur remplis...) — les colonnes sont repérées par leur EN-TÊTE
(texte), pas par leur numéro, ce qui permet de traiter aussi bien les fichiers
"Base article groupement N" que le fichier "MDC_IMPORT_MARCHE" sans réglage manuel.

Ce que fait ce script, pour chaque produit et chaque fournisseur rempli :
  - vérifie que le ratio Prix de commande ÷ Prix de référence (BPU) correspond à ce que
    décrit le Conditionnement (multiplicateur, récipient nommé — sachet/carton/barquette/
    cageot/bouteille/pièce...) ;
  - normalise la lecture des unités (kg/kilo/kilogramme, l/litre, g/gr/gramme, ml/cl...) et
    des récipients, quelle que soit la façon dont la personne les a écrits ;
  - repère les lignes strictement dupliquées (surlignées en ROUGE, 2e occurrence et
    suivantes) ;
  - repère les références produit en double avec le même prix (surlignées en ROSE) ;
  - repère les valeurs manquantes (prix, unité de référence) ;
  - ne fait JAMAIS de correction automatique du prix ou de l'unité — il se contente de
    surligner la cellule et d'écrire un commentaire expliquant l'incohérence. Une personne
    doit ensuite valider ; ce script est une pré-analyse pour accélérer ce travail, pas une
    décision finale.

Les deux SEULES valeurs que ce script écrit réellement (jamais un prix) :
  1. La colonne Lot : un zéro est ajouté devant les numéros à un seul chiffre (1-9 -> 01-09).
  2. Le Poids à l'unité de référence, UNIQUEMENT quand elle est vide, que l'unité de
     référence est à la pièce ('PC'/'pièce'/'unité'), que le conditionnement identifie sans
     ambiguïté ce qu'est une pièce et son poids/volume, et que le ratio des prix confirme le
     nombre total de pièces.

Usage :
    python audit_geral.py
    (lit tous les .xlsx de raw_data/, écrit les fichiers audités + rapports dans
    transformed_files/)
"""

import os
import re
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import first_scripts.first_attempt_v2_prototype as motor  # noqa: E402  (règles de cohérence déjà validées)

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill, Font, Alignment

# ============================================================================
# DOSSIERS
# ============================================================================

DOSSIER_PROJET = os.path.dirname(os.path.abspath(__file__))
DOSSIER_ENTREE = os.path.join(DOSSIER_PROJET, "raw_data")
DOSSIER_SORTIE = os.path.join(DOSSIER_PROJET, "transformed_files")

# ============================================================================
# EXTENSION DE VOCABULAIRE — récipients, unités et fautes de frappe fréquentes,
# accumulées au fil des fichiers déjà traités (fruits/légumes/viande ET laitier/traiteur).
# Ces objets surchargent ceux importés de first_attempt_v2_prototype.py, sans modifier ce
# fichier source partagé.
# ============================================================================

MOTS_CONTENANT = ["sac", "seau", "brique", "barquette", "barq", "bidon", "outre", "colis",
                  "pot", "sachet", "carton", "poche", "cageot", "boite", "filet", "bouteille",
                  "bouteile",  # faute de frappe fréquente ('BOUTEILES', 1 seul L)
                  "caisse", "plaque", "plateau", "clayette", "pack", "gourde", "botte"]
_MOTS_REGEX = "pi[eè]ce|" + "|".join(MOTS_CONTENANT)

# "GR" (abréviation très fréquente de gramme) est absent du vocabulaire d'origine, qui ne
# connaît que "g"/"gramme" — ce qui faisait échouer le parsing de la plupart des
# conditionnements ('COLIS DE 10 X 250 GR'...). "gr" doit être listé AVANT "g" dans
# l'alternance (sinon "g" seul matche et laisse le "r" de côté, ce qui fait échouer
# l'ancrage de fin de motif).
_UNITES_REGEX = "kg|kilo|gr|g|gramme|l|litre|ml|cl|mg"
motor.FACTEUR_VERS_KG_OU_L["gr"] = 0.001
motor.base_depuis_unite = lambda unite: (
    "kg" if unite.lower() in ("kg", "kilo", "g", "gr", "gramme", "mg") else "l")

motor.PACKAGE_TYPES = tuple(MOTS_CONTENANT)
motor.RE_TYPE = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de\s+)?([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
motor.RE_CONTENANT_MULTIPLIE = re.compile(
    rf"^\s*(\d+)\s*({_MOTS_REGEX})s?\s+(?:de\s+)?([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
motor.RE_TRIPLE = re.compile(
    rf"^\s*(\d+)\s*x\s*(\d+)\s*x\s*([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
motor.RE_DOUBLE = re.compile(
    rf"^\s*(\d+)\s*x\s*([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
motor.RE_LIBRE = re.compile(rf"([\d,.]+)\s*({_UNITES_REGEX})\b", re.I)
# Comme RE_LIBRE, mais tolère un multiplicateur 'x' collé juste après l'unité, sans espace
# ('1,2KGx6') — fréquent dans la Désignation fournisseur (texte libre), contrairement au
# Conditionnement, plus structuré, où RE_LIBRE (avec \b) suffit.
RE_QUANTITE_DESIGNATION = re.compile(rf"([\d,.]+)\s*({_UNITES_REGEX})(?![a-wyzà-öø-ÿ])", re.I)
# Le fichier MDC d'origine abrège barquette/pot en "BQT" ; les fichiers "groupement" écrivent
# le mot en entier. On garde le mot en entier par défaut (le cas le plus fréquent observé) —
# une éventuelle abréviation spécifique à un fichier reste à ajuster si besoin plus tard.
motor.CONTENANT_ABREGE = {
    "sac": "SAC", "piece": "PIECE", "seau": "SEAU", "brique": "BRIQUE",
    "barquette": "BARQUETTE", "barq": "BARQUETTE",
    "bidon": "BIDON", "outre": "OUTRE", "colis": "COLIS", "pot": "POT", "sachet": "SACHET",
    "carton": "CARTON", "poche": "POCHE", "bouteille": "BOUTEILLE", "bouteile": "BOUTEILLE",
    "cageot": "CAGEOT", "boite": "BOITE", "filet": "FILET",
    "caisse": "CAISSE", "plaque": "PLAQUE", "plateau": "PLATEAU", "clayette": "CLAYETTE",
}

# "SV" / "S/V" (sous-vide) apparaît comme mot inséré entre le récipient et sa quantité
# ('PIECE SV DE 7 KG') — on le retire avant toute analyse du Conditionnement, ce qui
# simplifie tous les motifs ci-dessous.
RE_SV = re.compile(r"\bS\s*/?\s*V\b", re.I)

# Qualificatifs de matière/texture parfois insérés entre le récipient et sa quantité
# ('CARTON DE 10 SACHETS KRAFT DE 1KG') et qui n'apportent aucune information de quantité —
# on les retire aussi avant toute analyse, comme pour "S/V".
RE_QUALIFICATIF_CONTENANT = re.compile(r"\bKRAFT\b", re.I)

# Référence de plage de cellules Excel collée par erreur au milieu du texte (ex: une personne
# a sélectionné "H33:H66" et l'a collée dans la cellule par inadvertance), ce qui casse
# complètement la lecture du Conditionnement — ex: 'CARTON DE 4 POCHES H33:H66DE 5 KG' au lieu
# de 'CARTON DE 4 POCHES DE 5 KG'. Pas de \b de fin : la plage colle parfois directement au mot
# suivant, sans espace.
RE_PLAGE_CELLULE_COLLEE = re.compile(r"\b[A-Z]{1,3}\d{1,7}\s*:\s*[A-Z]{1,3}\d{1,7}")


def nettoie_conditionnement(texte):
    if not isinstance(texte, str):
        return texte
    propre = RE_SV.sub("", texte)
    propre = RE_QUALIFICATIF_CONTENANT.sub("", propre)
    propre = re.sub(r"\s{2,}", " ", propre).strip(" ,")
    return propre


def retire_plage_cellule_collee(conditionnement):
    """Repère une plage de cellules Excel ('H33:H66') collée par erreur dans le
    Conditionnement et la retire. Renvoie (texte_nettoyé, True) si une correction a été
    faite, sinon (texte_inchangé, False)."""
    if not isinstance(conditionnement, str) or not RE_PLAGE_CELLULE_COLLEE.search(conditionnement):
        return conditionnement, False
    nettoye = RE_PLAGE_CELLULE_COLLEE.sub(" ", conditionnement)
    nettoye = re.sub(r"\s{2,}", " ", nettoye).strip(" ,")
    return nettoye, True


# Formats absents du vocabulaire d'origine, fréquents dans les fichiers fruits/légumes/viande :
#   "CARTON DE 2 POCHES DE 4 KG" (récipient contenant N sous-récipients d'une quantité donnée)
#   "CARTON DE 3 X 2 KG"         (récipient puis multiplicateur "N x Qté")
RE_CONTENANT_IMBRIQUE = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de\s+)?(\d+)\s*({_MOTS_REGEX})s?\s+(?:de\s+)?"
    rf"([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
RE_CONTENANT_PUIS_MULT = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de\s+)?(\d+)\s*x\s*([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)

# Triple imbrication ("CARTON DE 8 PACKS DE 4 GOURDES DE 90 GR" : un carton de 8 packs,
# chaque pack contenant 4 gourdes de 90g) — la "pièce" vendue est le niveau INTERMÉDIAIRE (le
# pack), pas le carton entier ni la gourde individuelle ; voir piece_carton_candidats().
RE_CONTENANT_TRIPLE_IMBRIQUE = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de\s+)?(\d+)\s*({_MOTS_REGEX})s?\s+(?:de\s+)?(\d+)\s*"
    rf"({_MOTS_REGEX})s?\s+(?:de\s+)?([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)

# "2 PIECES DE 2 KG" (= 4 kg au total) : aucun récipient n'est nommé, la ligne commence
# directement par "N PIECES DE Qté" — quand l'unité de référence est au kg/L, le nombre de
# pièces sert de multiplicateur du poids (contrairement à piece_carton_candidats(), qui
# s'applique quand l'unité de référence est déjà à la pièce).
RE_N_PIECES_DE_QTY = re.compile(
    rf"^\s*(\d+)\s*pi[eè]ces?\s+(?:de\s+)?([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)

# Conditionnement qui est JUSTE une quantité, sans récipient ni multiplicateur ('1,425KG',
# '1,5KG') — fréquent en charcuterie/traiteur. Sans ce motif, ces lignes ne passent par AUCUNE
# des vérifications ci-dessus (ni piece_carton_candidats, ni determine_base_et_quantite), donc
# un prix de référence resté identique au prix de commande (copié tel quel au lieu d'être
# divisé par le poids réel du conditionnement) n'était jamais recalculé ni signalé.
RE_QUANTITE_SEULE = re.compile(rf"^\s*([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)

_extraire_quantite_original = motor.extraire_quantite


def extraire_quantite_etendue(texte):
    texte = nettoie_conditionnement(texte)
    if texte:
        m = RE_CONTENANT_IMBRIQUE.match(texte)
        if m:
            _mot_ext, n, _mot_int, qty, unite = m.groups()
            total = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()] * int(n)
            return total, motor.base_depuis_unite(unite), True
        m = RE_CONTENANT_PUIS_MULT.match(texte)
        if m:
            _mot_ext, n, qty, unite = m.groups()
            total = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()] * int(n)
            return total, motor.base_depuis_unite(unite), True
        m = RE_N_PIECES_DE_QTY.match(texte)
        if m:
            n, qty, unite = m.groups()
            total = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()] * int(n)
            return total, motor.base_depuis_unite(unite), True
    return _extraire_quantite_original(texte)


motor.extraire_quantite = extraire_quantite_etendue


def extraire_quantite_designation_etendue(designation):
    """Comme motor.extraire_quantite_designation, mais tolère un multiplicateur collé juste
    après l'unité sans espace ('750GX8', '100Gx28'), fréquent dans la Désignation de
    l'article/fournisseur — motor.extraire_quantite_designation (avec RE_LIBRE, qui exige un
    \\b) ne voit alors rien du tout."""
    if not designation:
        return None, None
    m = RE_QUANTITE_DESIGNATION.search(str(designation))
    if m:
        qty, unite = m.groups()
        return motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], motor.base_depuis_unite(unite)
    return None, None


motor.extraire_quantite_designation = extraire_quantite_designation_etendue

_mot_type_conditionnement_original = motor.mot_type_conditionnement


def mot_type_conditionnement_etendu(texte):
    """Comme mot_type_conditionnement, mais reconnaît aussi le récipient EXTÉRIEUR des
    formats imbriqué/multiplié ('CARTON DE 2 POCHES DE 4 KG' ou 'CARTON DE 3 X 2 KG' ->
    'CARTON'), et le format sans récipient 'N PIECES DE Qté' -> 'N PIECES'."""
    texte = nettoie_conditionnement(texte)
    if texte:
        for regex in (RE_CONTENANT_IMBRIQUE, RE_CONTENANT_PUIS_MULT):
            m = regex.match(texte)
            if m:
                mot_ext = m.group(1)
                return motor.CONTENANT_ABREGE.get(motor.normalise_texte(mot_ext), mot_ext.upper())
        m = RE_N_PIECES_DE_QTY.match(texte)
        if m:
            return f"{m.group(1)} PIECES"
    return _mot_type_conditionnement_original(texte)


motor.mot_type_conditionnement = mot_type_conditionnement_etendu

_recipient_conditionnement_original = motor.recipient_conditionnement


def recipient_conditionnement_etendu(texte):
    """Comme motor.recipient_conditionnement, mais ne force PAS 'LOT DE N' quand un nom de
    récipient est présent dans le conditionnement (ex: '4 BARQUETTES DE 500G' -> 'BARQUETTE',
    pas 'LOT DE 4') — l'original le faisait pour TOUT ce qui correspond à
    'N contenant DE qté unité', même quand ce "contenant" est un vrai nom (barquette, pot...),
    perdant l'information. 'LOT DE N' reste utilisé quand le mot est le générique 'pièce'
    (qui ne nomme rien de spécifique), ou quand aucun nom n'est trouvé du tout."""
    if not texte:
        return None
    texte_propre = nettoie_conditionnement(str(texte).strip())
    m = motor.RE_CONTENANT_MULTIPLIE.match(texte_propre)
    if m:
        n, mot, _qty, _unite = m.groups()
        if motor.normalise_texte(mot) == "piece":
            return f"LOT DE {int(n)}"
        return motor.CONTENANT_ABREGE.get(motor.normalise_texte(mot), mot.upper())
    return _recipient_conditionnement_original(texte_propre)


motor.recipient_conditionnement = recipient_conditionnement_etendu

# Alias d'unité de référence : certains fichiers utilisent "PC" là où le vocabulaire
# d'origine attend "unité"/"pièce". On substitue avant l'appel puis on restitue le texte
# d'origine dans les commentaires, pour ne pas modifier le fichier source partagé.
ALIAS_UNITE_REF = {"pc": "unité"}

RE_CONTENANT_NX_M_PIECES = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de\s+)?(\d+)\s*x\s*(\d+)\s*pi[eè]ces?\s*$", re.I)

# Plusieurs références dans une même cellule (une ligne par référence séparée par une
# virgule ou un retour à la ligne, ex: '43582, 4353, 43521, 43619' ou '813072 - Mexicaine\n
# 813064 -Provençale') — à séparer/vérifier par une personne, jamais deviné automatiquement.
RE_REFERENCES_MULTIPLES = re.compile(r"[,\n]")


def piece_carton_candidats(unite_ref, conditionnement):
    """Si l'unité de référence est déjà à la pièce ('PC'/'pièce'/'unité'), renvoie
    (candidats, mot_piece, poids_piece, base_unite) :
      'CARTON DE 6 PIECES DE 420 GR' -> ([6], 'piece', 0.42, 'kg')
      'CARTON 14 SACHETS DE 350 GR'  -> ([14], 'sachet', 0.35, 'kg') (le "sachet" EST la pièce)
      'CARTON 6 BOUTEILLES DE 75 CL' -> ([6], 'bouteille', 0.75, 'l')
      'CARTON DE 10 X 250 GR'        -> ([10], None, 0.25, 'kg')
      'CARTON DE 3 X 18 PIECES'      -> ([3, 18, 54], None, None, None) (pas de poids donné)
    'candidats' est la liste des nombres qui devraient correspondre au ratio Prix commande ÷
    Prix référence — le DERNIER est toujours le nombre total de pièces. 'mot_piece',
    'poids_piece' et 'base_unite' identifient ce qu'EST une pièce et son poids/volume, pour
    renseigner la colonne Poids à l'unité de référence quand c'est fiable. Renvoie None si
    aucun motif ne correspond."""
    if motor.normalise_texte(unite_ref) not in ("pc", "piece", "pièce", "unite", "unité"):
        return None
    if not conditionnement:
        return None
    m = RE_CONTENANT_TRIPLE_IMBRIQUE.match(conditionnement)
    if m:
        # "CARTON DE 8 PACKS DE 4 GOURDES DE 90 GR" : la "pièce" vendue est le pack
        # (niveau intermédiaire), pas la gourde individuelle ni le carton entier — le nombre
        # total de packs (N) est ce qui doit correspondre au ratio, et son poids est celui
        # de ses M gourdes de qty chacune.
        _mot_ext, n, mot_piece, m_inner, _mot_inner, qty, unite = m.groups()
        poids_piece = round(int(m_inner) * motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        return [int(n)], motor.normalise_texte(mot_piece), poids_piece, motor.base_depuis_unite(unite)
    m = RE_CONTENANT_IMBRIQUE.match(conditionnement)
    if m:
        _mot_ext, n, mot_piece, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        return [int(n)], motor.normalise_texte(mot_piece), poids_piece, motor.base_depuis_unite(unite)
    m = RE_CONTENANT_PUIS_MULT.match(conditionnement)
    if m:
        mot_ext, n, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        # Le récipient nommé ('CARTON DE 8 X 250 GR') garde son nom en unité de commande —
        # "LOT DE N" est réservé aux formats SANS aucun nom de récipient ('100 x 10g').
        return [int(n)], motor.normalise_texte(mot_ext), poids_piece, motor.base_depuis_unite(unite)
    m = motor.RE_CONTENANT_MULTIPLIE.match(conditionnement)
    if m:
        n, mot_piece, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        return [int(n)], motor.normalise_texte(mot_piece), poids_piece, motor.base_depuis_unite(unite)
    m = RE_CONTENANT_NX_M_PIECES.match(conditionnement)
    if m:
        mot_ext, n1, n2 = m.group(1), int(m.group(2)), int(m.group(3))
        return [n1, n2, n1 * n2], motor.normalise_texte(mot_ext), None, None
    # Récipient nommé contenant SEULEMENT un nombre de pièces, sans poids ('CAGEOT DE 30
    # PIECES') : le poids réel est inconnu (souvent variable), mais le nombre de pièces suffit
    # à vérifier le ratio — pas besoin de deviner un poids via la désignation pour expliquer le
    # prix (ce qui donnerait un résultat en €/kg non pertinent ici).
    m = RE_CONTENANT_PIECES_SEULEMENT.match(conditionnement)
    if m:
        mot_ext, n = m.group(1), int(m.group(2))
        # Deux lectures possibles, selon ce que confirme le ratio : soit N exemplaires séparés
        # (ratio = N, ex: 'CAGEOT DE 30 PIECES' d'œufs vendus à l'unité), soit UN SEUL
        # contenant divisé en N parts (ratio = 1, ex: 'PLAQUE DE 30 PIECES' = un gâteau coupé
        # en 30 parts, vendu entier) — le 1 est ajouté comme candidat valable lui aussi.
        return [1, n], motor.normalise_texte(mot_ext), None, None
    # Multiplicateur SANS récipient nommé ('100 x 10g', '4 x 125g') : quand l'unité de
    # référence est déjà à la pièce, c'est N qui doit correspondre au ratio (comparaison par
    # pièce), PAS un total en kg/L à recalculer — contrairement à determine_base_et_quantite(),
    # qui suppose au contraire que l'unité de référence devrait être en kg/L.
    m = motor.RE_TRIPLE.match(conditionnement)
    if m:
        n1, n2, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        n1, n2 = int(n1), int(n2)
        return [n1, n2, n1 * n2], None, poids_piece, motor.base_depuis_unite(unite)
    m = motor.RE_DOUBLE.match(conditionnement)
    if m:
        n, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        return [int(n)], None, poids_piece, motor.base_depuis_unite(unite)
    return None


def determine_base_et_quantite(conditionnement):
    """Cherche, dans le conditionnement, un récipient de quantité TOTALE connue, via un motif
    PROPREMENT reconnu (pas une recherche floue en texte libre) :
      'Seau de 5L'                 -> ('l', 5.0, 'seau')
      'Sachet 500g'                -> ('kg', 0.5, 'sachet')
      '6 x 1L'                     -> ('l', 6.0, None)          (pas de récipient nommé)
      'CARTON DE 2 POCHES DE 4 KG' -> ('kg', 8.0, 'carton')
      'CARTON DE 3 X 2 KG'         -> ('kg', 6.0, 'carton')
    Renvoie None si aucun motif propre ne correspond (ex: 'SAC KRAFT DE 5 KG', où le mot en
    trop 'KRAFT' casse tous ces motifs — ce cas reste géré par la recherche floue, moins
    fiable, dans reconsidere_prix_inexplique)."""
    if not conditionnement:
        return None
    # "pièce" n'est PAS un récipient : "2 PIECES DE 2 KG" ne veut pas dire "un récipient
    # nommé PIECE contenant 2 kg" (ce qui donnerait, à tort, une unité de commande 'PIECE' au
    # singulier) mais bien "2 pièces, 2 kg au total" — ce format reste géré plus loin par
    # motor.traite_ligne (via mot_type_conditionnement_etendu, qui renvoie '2 PIECES').
    m = RE_CONTENANT_IMBRIQUE.match(conditionnement)
    if m and motor.normalise_texte(m.group(1)) != "piece":
        mot_ext, n, _mot_int, qty, unite = m.groups()
        quantite = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()] * int(n)
        return motor.base_depuis_unite(unite), quantite, motor.normalise_texte(mot_ext)
    m = RE_CONTENANT_PUIS_MULT.match(conditionnement)
    if m and motor.normalise_texte(m.group(1)) != "piece":
        mot_ext, n, qty, unite = m.groups()
        quantite = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()] * int(n)
        return motor.base_depuis_unite(unite), quantite, motor.normalise_texte(mot_ext)
    m = motor.RE_CONTENANT_MULTIPLIE.match(conditionnement)
    if m and motor.normalise_texte(m.group(2)) != "piece":
        n, mot_ext, qty, unite = m.groups()
        quantite = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()] * int(n)
        return motor.base_depuis_unite(unite), quantite, motor.normalise_texte(mot_ext)
    m = motor.RE_TYPE.match(conditionnement)
    if m:
        mot_ext, qty, unite = m.groups()
        quantite = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()]
        return motor.base_depuis_unite(unite), quantite, motor.normalise_texte(mot_ext)
    m = motor.RE_DOUBLE.match(conditionnement)
    if m:
        n, qty, unite = m.groups()
        quantite = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()] * int(n)
        return motor.base_depuis_unite(unite), quantite, None
    m = RE_QUANTITE_SEULE.match(conditionnement)
    if m:
        qty, unite = m.groups()
        quantite = motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()]
        return motor.base_depuis_unite(unite), quantite, None
    return None


_MOTS_CONTENANT_CANONIQUES = [
    "sac", "seau", "brique", "barquette", "bidon", "outre", "colis", "pot", "sachet",
    "carton", "poche", "cageot", "boite", "filet", "bouteille", "caisse", "plaque",
    "plateau", "clayette", "piece",
]
_MOTS_CONTENANT_CONNUS = set(MOTS_CONTENANT) | {"piece", "pièce"}


def _distance_un(a, b):
    """Vrai si a et b sont à une lettre près l'un de l'autre (insertion, suppression ou
    substitution d'une seule lettre) — suffisant pour repérer une faute de frappe simple."""
    if a == b:
        return True
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(1 for x, y in zip(a, b) if x != y) <= 1
    long, court = (a, b) if len(a) > len(b) else (b, a)
    i = j = diff = 0
    while i < len(long) and j < len(court):
        if long[i] == court[j]:
            i += 1
            j += 1
        else:
            i += 1
            diff += 1
            if diff > 1:
                return False
    return True


def corrige_faute_frappe_contenant(conditionnement):
    """Repère un mot du Conditionnement qui ne correspond à AUCUN mot de contenant connu mais
    qui est à une seule lettre près d'un mot connu (lettre manquante/en trop/substituée) — ex:
    'BARQUETE' (un seul T) au lieu de 'BARQUETTE'. Renvoie (texte_corrigé, (mot_original,
    mot_corrigé)) si une correction a été trouvée, sinon (texte, None). N'est appelé qu'en
    dernier recours, quand l'analyse normale (motifs propres) a déjà échoué."""
    if not conditionnement:
        return conditionnement, None
    for mot in re.findall(r"[A-ZÀ-Ÿ]{4,}", conditionnement.upper()):
        mot_normalise = motor.normalise_texte(mot)
        # Pluriel simple ('PIECES' -> 'piece') : ce n'est pas une faute de frappe, donc on
        # vérifie le mot SANS son éventuel 's' final avant de chercher une faute de frappe.
        mot_singulier = mot_normalise[:-1] if (mot_normalise.endswith("s")
                                                and len(mot_normalise) > 4) else mot_normalise
        if mot_normalise in _MOTS_CONTENANT_CONNUS or mot_singulier in _MOTS_CONTENANT_CONNUS:
            continue
        for canonique in _MOTS_CONTENANT_CANONIQUES:
            if _distance_un(mot_singulier, canonique):
                corrige = re.sub(rf"\b{re.escape(mot)}\b", canonique.upper(), conditionnement,
                                  count=1, flags=re.I)
                return corrige, (mot, canonique.upper())
    return conditionnement, None


def traite_ligne_etendue(unite_ref, poids_ref, designation, conditionnement, unite_cmd,
                          prix_ref, prix_cmd, desig_fourn=None):
    """Délègue à _traite_ligne_etendue_nucleo. La correction d'une éventuelle faute de frappe
    sur un mot de contenant est faite en amont, directement sur la cellule Conditionnement
    (voir audite_fournisseur) — ce que cette fonction reçoit est donc déjà la version
    corrigée."""
    resultat = _traite_ligne_etendue_nucleo(unite_ref, poids_ref, designation, conditionnement,
                                             unite_cmd, prix_ref, prix_cmd, desig_fourn)

    # Garde-fou : une couleur posée sans commentaire (oubli dans motor.traite_ligne, ex: la
    # conversion PC -> KG par poids de référence ne renseigne que commentaire_k) rendrait la
    # cellule changée SANS AUCUN signe visible (marque() a besoin des deux). On synthétise un
    # commentaire minimal plutôt que de laisser passer un changement invisible.
    for champ, valeur_ancienne in (("d", unite_ref), ("l", unite_cmd), ("k", prix_ref), ("m", prix_cmd)):
        couleur = getattr(resultat, f"couleur_{champ}")
        if couleur and not getattr(resultat, f"commentaire_{champ}"):
            nouvelle_valeur = getattr(resultat, f"nouveau_{champ}", None)
            setattr(resultat, f"commentaire_{champ}",
                    f"Valeur changée de '{valeur_ancienne}' vers '{nouvelle_valeur}'." if nouvelle_valeur is not None
                    else "Incohérence détectée — à vérifier manuellement.")
    return resultat


def _traite_ligne_etendue_nucleo(unite_ref, poids_ref, designation, conditionnement, unite_cmd,
                                  prix_ref, prix_cmd, desig_fourn=None):
    conditionnement = nettoie_conditionnement(conditionnement)

    resultat_pieces = piece_carton_candidats(unite_ref, conditionnement)
    if resultat_pieces and prix_ref and prix_cmd:
        candidats_pieces, mot_piece, poids_piece, base_unite = resultat_pieces
        ratio = prix_cmd / prix_ref
        r = motor.Resultat()
        total_pieces = candidats_pieces[-1]
        candidat_ok = next((c for c in candidats_pieces if abs(ratio - c) <= motor.TOLERANCE_RATIO * c), None)

        # L'unité de référence est bien à la pièce ici (pas un poids à recalculer en kg/L) —
        # toujours normaliser en 'PC', silencieusement si c'était déjà 'pc'/'pièce'/'unité'.
        r.nouveau_d = "PC"
        if motor.normalise_texte(unite_ref) != "pc":
            r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
            r.commentaire_d = f"Unité de référence normalisée de '{unite_ref}' vers 'PC'."

        if (candidat_ok is not None and candidat_ok == total_pieces and not mot_piece
                and motor.normalise_texte(unite_cmd) != f"lot de {total_pieces}"):
            # Aucun récipient nommé dans le conditionnement ('100 x 10g', '4 x 125g') — le
            # ratio confirme bien un LOT de N pièces, pas un poids total à recalculer.
            r.nouveau_l = f"LOT DE {total_pieces}"
            r.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
            r.commentaire_l = (
                f"Unité de commande précisée en 'LOT DE {total_pieces}' : le ratio Prix "
                f"commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = {round(ratio, 3)}) "
                f"correspond au conditionnement ('{conditionnement}'), sans récipient nommé.")

        if poids_piece and not poids_ref:
            # Le poids d'UNE pièce est écrit explicitement dans le conditionnement (ex: 'DE
            # 500 GR'), indépendamment du nombre total de pièces — ce n'est pas une déduction
            # à partir du ratio des prix, donc on la renseigne même quand le ratio ne confirme
            # PAS le nombre total de pièces (ça reste une incohérence de prix/quantité à
            # signaler séparément, mais le poids d'une pièce, lui, n'est pas en cause).
            r.nouveau_e = poids_piece
            confirmation = (
                f"Confirmé par le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ "
                f"{prix_ref} = {round(ratio, 3)}) qui correspond bien à {total_pieces} pièces."
                if candidat_ok == total_pieces else
                "Le nombre total de pièces n'est en revanche pas confirmé par le ratio des "
                "prix — voir le signalement correspondant.")
            r.explication_e = (
                f"Poids à l'unité de référence renseigné à partir du conditionnement "
                f"('{conditionnement}') : {poids_piece} {base_unite.upper()} par pièce "
                f"(sur {total_pieces} au total dans ce {mot_piece or 'conditionnement'}). "
                f"{confirmation}")

        if candidat_ok is None:
            r.couleur_m = motor.COULEUR_RATIO_SUSPECT
            r.commentaire_m = (
                f"Le conditionnement ('{conditionnement}') indique {candidats_pieces} pièce(s) "
                f"possibles, et l'unité de référence est déjà à la pièce ('{unite_ref}') — le "
                f"ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                f"{round(ratio, 3)}) ne correspond à aucun d'eux. À vérifier manuellement.")
        elif candidat_ok == 1 and candidat_ok != total_pieces:
            # Le ratio confirme que c'est le CONTENANT ENTIER (pas chacune de ses parts) qui
            # est vendu (ex: 'PLAQUE DE 30 PIECES' = un gâteau entier coupé en 30 parts) — on
            # cherche alors le poids de ce contenant dans la Désignation fournisseur ou la
            # Désignation de l'article pour compléter E.
            poids_trouve, base_trouvee = motor.extraire_quantite_designation(desig_fourn or "")
            if not (poids_trouve and base_trouvee == "kg"):
                poids_trouve, base_trouvee = motor.extraire_quantite_designation(designation or "")
            if poids_trouve and base_trouvee == "kg" and not poids_ref:
                r.nouveau_e = round(poids_trouve, 3)
                r.explication_e = (
                    f"Poids à l'unité de référence renseigné à {round(poids_trouve, 3)} kg : le "
                    f"conditionnement ('{conditionnement}') décrit {total_pieces} parts, mais "
                    f"le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                    f"{round(ratio, 3)}) confirme que le prix est pour le contenant ENTIER, pas "
                    "pour une part individuelle — poids trouvé dans la désignation.")
        elif candidat_ok != total_pieces:
            r.couleur_d = motor.COULEUR_RATIO_SUSPECT
            r.commentaire_d = (
                f"L'unité de référence est déclarée à la pièce ('{unite_ref}'), mais le ratio "
                f"observé ({prix_cmd} ÷ {prix_ref} = {round(ratio, 3)}) correspond à "
                f"{candidat_ok}, un SOUS-groupe du conditionnement ('{conditionnement}'), pas "
                f"au nombre total de pièces ({total_pieces}). Le prix de référence semble être "
                f"au prix du sous-lot de {candidat_ok}, pas à la pièce individuelle — à "
                "vérifier.")
        return r  # cohérent (ou signalé ci-dessus) — le poids par pièce n'est pas utilisé

    # Unité de référence en kg/L (ou nommant elle-même un récipient, ex: 'seau', 'barquette')
    # ET le conditionnement identifie PROPREMENT (motif reconnu, pas une recherche floue) un
    # récipient de quantité connue -> le prix de commande (M, toujours considéré comme réel)
    # sert à RECALCULER le prix de référence (K = M / quantité), l'unité de référence devient
    # KG/L, et l'unité de commande devient le nom du récipient. Un format flou (ex: 'SAC KRAFT
    # DE 5 KG', où un mot en trop casse le motif propre) ne passe PAS par ici — il reste
    # seulement signalé (voir reconsidere_prix_inexplique), jamais recalculé automatiquement.
    #
    # Exception : si l'unité de référence et l'unité de commande sont physiquement
    # INCOMPATIBLES (l'une en kg, l'autre en L — ex: D='KG' mais L='L'), ce bloc ne sait pas
    # traiter ce cas correctement (il changerait D silencieusement d'après le conditionnement,
    # sans jamais regarder si le prix de commande confirme vraiment la bonne base). On laisse
    # directement la main à motor.traite_ligne, qui a une règle dédiée à ce cas précis
    # (Priorité 0) : KG/L si K = M, sinon cherche une confirmation dans le conditionnement,
    # sinon signale l'incohérence sans rien changer.
    base_ref_norm_precoce = motor.normalise_texte(unite_ref)
    base_cmd_norm_precoce = motor.normalise_texte(unite_cmd) if unite_cmd else None
    kg_l_incompatibles = (
        (base_ref_norm_precoce == "kg" and base_cmd_norm_precoce in ("l", "litre"))
        or (base_ref_norm_precoce in ("l", "litre") and base_cmd_norm_precoce == "kg"))

    # Remarque : quand l'unité de référence est déjà à la pièce ('PC'/'pièce'/'unité'), on
    # laisse quand même ce bloc s'en charger si le Conditionnement décrit un récipient UNIQUE
    # de quantité connue (ex: 'CARTON DE 2,5 KG') — piece_carton_candidats() a la priorité et a
    # déjà renvoyé None (donc ce n'est pas un format de comptage de pièces ambigu) ; ici, le
    # prix de commande (M, toujours fiable) donne un prix par kg/L cohérent avec le prix par
    # pièce d'origine. Ne PAS le faire mélangerait un prix "par pièce" avec un poids total de
    # récipient sans rapport (ex: diviser un prix-par-pièce par le poids TOTAL d'un carton de
    # plusieurs pièces), ce qui donnerait un prix au kilo n'importe quoi.
    if prix_ref and prix_cmd and not kg_l_incompatibles:
        info_contenant = determine_base_et_quantite(conditionnement)
        if info_contenant:
            base, quantite, mot = info_contenant
            if quantite and quantite > 0:
                r = motor.Resultat()

                # Garde-fou : une quantité par récipient inférieure à 1 gramme/1 mL est
                # physiquement improbable pour un produit alimentaire — signe quasi certain
                # d'une confusion d'unité dans le Conditionnement (GR écrit au lieu de KG, ou
                # ML au lieu de L). On ne recalcule rien automatiquement dans ce cas, on se
                # contente de signaler le Conditionnement lui-même.
                if quantite < 0.001:
                    r.flag_conditionnement = (
                        motor.COULEUR_RATIO_SUSPECT,
                        f"Le conditionnement ('{conditionnement}') indique une quantité de "
                        f"{quantite} {base} par récipient, ce qui est physiquement improbable "
                        "pour un produit alimentaire — probable confusion d'unité (GR écrit au "
                        "lieu de KG, ou ML au lieu de L). À vérifier avant de valider ; rien n'a "
                        "été recalculé automatiquement.")
                    return r

                ratio_prix = prix_cmd / prix_ref
                tolerance_qte = max(0.05, 0.05 * quantite)

                # La "Désignation fournisseur" confirme parfois une AUTRE quantité que celle
                # lue dans le Conditionnement (ex: Conditionnement 'PIECE DE 2,1 KG' saisi avec
                # les chiffres inversés, alors que la Désignation fournisseur dit '...1,2KGx6...'
                # et que le ratio Prix commande ÷ Prix référence vaut justement 1,2) : dans ce
                # cas, on fait confiance au ratio (confirmé par la désignation), pas au
                # Conditionnement, et on ne recalcule rien automatiquement.
                quantite_desig = None
                for qty_str, unite_str in RE_QUANTITE_DESIGNATION.findall(desig_fourn or ""):
                    if motor.base_depuis_unite(unite_str) != base:
                        continue
                    try:
                        quantite_desig = motor._nombre(qty_str) * motor.FACTEUR_VERS_KG_OU_L[unite_str.lower()]
                    except (KeyError, ValueError):
                        continue
                    break

                if (quantite_desig and abs(quantite_desig - quantite) > tolerance_qte
                        and abs(ratio_prix - quantite_desig) <= max(0.05, 0.05 * quantite_desig)):
                    r.flag_conditionnement = (
                        motor.COULEUR_RATIO_SUSPECT,
                        f"Le conditionnement ('{conditionnement}') indique {round(quantite, 3)} "
                        f"{base}, mais le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ "
                        f"{prix_ref} = {round(ratio_prix, 3)}) correspond en fait à "
                        f"{round(quantite_desig, 3)} {base}, ce que confirme la désignation "
                        f"fournisseur ('{desig_fourn}'). Le conditionnement semble contenir une "
                        "erreur de saisie — à vérifier. Prix de référence non recalculé "
                        "automatiquement dans ce cas.")
                    return r

                # Si le ratio observé est un nombre ENTIER (différent de 1) mais ne correspond
                # pas à la quantité lue dans le Conditionnement, le ratio (basé sur M, toujours
                # considéré comme réel) est plus digne de confiance qu'un Conditionnement qui
                # semble incohérent avec lui — on ne divise pas automatiquement, on laisse un
                # commentaire sur le prix de commande pour qu'une personne vérifie le
                # Conditionnement. Le cas ratio = 1 est volontairement EXCLU de cette règle : il
                # correspond le plus souvent à une confusion kilo/récipient réel (ex: prix
                # donné par 'seau' recopié tel quel au lieu d'être divisé) — ce cas-là doit être
                # recalculé normalement plus bas (K = M ÷ quantité), pas juste signalé.
                ratio_entier = round(ratio_prix)
                if (ratio_entier != 0 and ratio_entier != 1
                        and abs(ratio_prix - ratio_entier) < 0.02
                        and abs(ratio_prix - quantite) > tolerance_qte):
                    r.couleur_m = motor.COULEUR_RATIO_SUSPECT
                    r.commentaire_m = (
                        f"Le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                        f"{round(ratio_prix, 3)}) est un nombre entier, ce qui suggère une "
                        f"quantité réelle de {ratio_entier} {base} par récipient, alors que le "
                        f"conditionnement ('{conditionnement}') indique {round(quantite, 3)} "
                        f"{base}. Vérifiez si le conditionnement est correct avant de valider — "
                        "prix de référence non recalculé automatiquement dans ce cas.")
                    return r

                base_ref_norm = motor.normalise_texte(unite_ref)
                d_novo = base.upper()
                d_change = base_ref_norm not in ("kg", "l", "litre")
                nouveau_k = round(prix_cmd / quantite, 4)
                diff_relative = abs(nouveau_k - prix_ref) / prix_ref if prix_ref else abs(nouveau_k - prix_ref)
                # En dessous de 5% d'écart, le prix de référence reste quasi identique
                # (arrondis, centimes) — ne rien changer évite de déranger pour rien.
                k_change = diff_relative > 0.05
                l_novo = motor.CONTENANT_ABREGE.get(mot, mot.upper()) if mot else None
                # Récipient d'exactement 1 kg/L, avec une unité de commande déjà 'KG'/'L' :
                # les deux conventions (garder KG/L, ou préciser le récipient) sont aussi
                # valables l'une que l'autre — ne pas forcer le changement.
                if (l_novo and abs(quantite - 1) < 0.01
                        and motor.normalise_texte(unite_cmd) in ("kg", "l", "litre")):
                    l_novo = None
                l_change = l_novo and motor.normalise_texte(l_novo) != motor.normalise_texte(unite_cmd)

                r.nouveau_d = d_novo
                if d_change:
                    r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
                    r.commentaire_d = (
                        f"Unité de référence changée de '{unite_ref}' vers '{d_novo}' : le "
                        f"conditionnement ('{conditionnement}') décrit un récipient, pas "
                        "directement une unité de mesure.")

                if k_change:
                    r.nouveau_k = nouveau_k
                    r.couleur_k = motor.COULEUR_UNITE_TRANSFORMEE
                    qte_affichee = round(quantite, 6) if quantite < 1 else round(quantite, 3)
                    r.commentaire_k = (
                        f"Prix de référence recalculé : {prix_cmd} ÷ {qte_affichee} "
                        f"{base} = {nouveau_k} (ancien prix de référence : {prix_ref}). Le "
                        f"conditionnement ('{conditionnement}') indique {qte_affichee} {base} "
                        f"pour ce récipient ; le prix de commande ({prix_cmd}) reste le prix "
                        "réel, inchangé.")

                if l_change:
                    r.nouveau_l = l_novo
                    r.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
                    r.commentaire_l = (
                        f"Unité de commande changée en '{l_novo}' d'après le conditionnement "
                        f"('{conditionnement}').")

                return r  # coherent ou signalé ci-dessus — D/K/L déjà tranchés ici

    alias = ALIAS_UNITE_REF.get(motor.normalise_texte(unite_ref))
    resultat = motor.traite_ligne(alias or unite_ref, poids_ref, designation, conditionnement,
                                   unite_cmd, prix_ref, prix_cmd)
    if alias:
        for attribut in ("commentaire_d", "commentaire_k", "commentaire_l", "commentaire_m"):
            texte = getattr(resultat, attribut)
            if texte:
                setattr(resultat, attribut, texte.replace(f"'{alias}'", f"'{unite_ref}'"))
    reconsidere_prix_inexplique(resultat, unite_ref, conditionnement, unite_cmd, prix_ref, prix_cmd)
    reconsidere_recipient_de_1_unite(resultat, unite_ref, unite_cmd, conditionnement)
    reconsidere_ratio_un_contre_contenant(resultat, unite_ref, unite_cmd, conditionnement, prix_ref, prix_cmd)
    return resultat


def reconsidere_ratio_un_contre_contenant(resultat, unite_ref, unite_cmd, conditionnement, prix_ref, prix_cmd):
    """Cas 'SAC KRAFT DE 5 KG' avec K = M (ratio = 1) : le motif propre échoue (mot en trop),
    donc determine_base_et_quantite() ne l'a pas recalculé, et le ratio ≈ 1 fait que le
    correcteur d'origine se contente de relabelliser l'unité de commande vers KG/L, sans
    remarquer que le conditionnement suggère une quantité (5 kg) totalement différente du
    ratio observé (1). On cherche cette quantité en texte libre, et si elle contredit
    nettement le ratio, on remplace le relabel par un signalement clair du prix incohérent —
    SAUF quand l'unité de commande d'origine était dans une base physiquement DIFFÉRENTE de
    l'unité de référence (ex: D='KG', L='L') : dans ce cas, le relabel (vers le récipient,
    ex: 'SACHET') corrige une vraie incohérence physique et doit être gardé, même si le prix
    est EN PLUS signalé comme suspect."""
    base_cible = motor.normalise_texte(unite_ref)
    if base_cible not in ("kg", "l", "litre") or not conditionnement or not prix_ref or not prix_cmd:
        return
    ratio_observe = prix_cmd / prix_ref
    if abs(ratio_observe - 1) > motor.TOLERANCE_RATIO:
        return  # ce cas ne concerne que K = M

    base_cmd = motor.normalise_texte(unite_cmd) if unite_cmd else None
    base_cmd = "kg" if base_cmd == "kg" else ("l" if base_cmd in ("l", "litre") else None)
    # Ne garde le relabel que si l'unité de commande d'ORIGINE était dans l'AUTRE base kg/L
    # (vraie incompatibilité physique, ex: D='KG' et L='L' d'origine). Si elle n'était ni kg
    # ni L (ex: déjà 'POCHE', pas reconnue proprement par recipient_conditionnement à cause
    # d'un format inhabituel), on garde l'ancien comportement (annule le relabel, signale
    # juste le prix) plutôt que de la remplacer par 'KG'/'L' au passage.
    annule_relabel = base_cmd is None or base_cmd == base_cible

    for qty_str, unite_str in motor.RE_LIBRE.findall(conditionnement):
        if motor.base_depuis_unite(unite_str) != ("kg" if base_cible == "kg" else "l"):
            continue
        try:
            quantite = motor._nombre(qty_str) * motor.FACTEUR_VERS_KG_OU_L[unite_str.lower()]
        except (KeyError, ValueError):
            continue
        if quantite > 0 and abs(quantite - 1) > motor.TOLERANCE_RATIO:
            if annule_relabel:
                resultat.nouveau_l = None
                resultat.couleur_l = None
                resultat.commentaire_l = None
            resultat.couleur_m = motor.COULEUR_RATIO_SUSPECT
            resultat.commentaire_m = (
                f"Le conditionnement ('{conditionnement}') suggère {round(quantite, 3)} "
                f"{base_cible}, ce qui voudrait dire que le prix de commande ({prix_cmd}) "
                f"devrait valoir environ {round(prix_ref * quantite, 4)} (soit {round(quantite, 3)} "
                f"fois le prix de référence {prix_ref}) — au lieu de ça il est identique au prix "
                "de référence (ratio = 1). Un des deux prix semble incorrect — à vérifier "
                "manuellement.")
            return


def recipient_est_exactement_1_unite(conditionnement, base_cible):
    """Vrai si le Conditionnement décrit un récipient contenant EXACTEMENT 1 unité de base
    (ex: 'SACHET DE 1 KG', 'BARQUETTE DE 1 L')."""
    if not conditionnement:
        return False
    m = motor.RE_TYPE.match(conditionnement)
    if not m:
        return False
    _mot, qty, unite = m.groups()
    if motor.base_depuis_unite(unite) != base_cible:
        return False
    return abs(motor._nombre(qty) - 1) < 0.01


def reconsidere_recipient_de_1_unite(resultat, unite_ref, unite_cmd, conditionnement):
    """Quand le récipient contient EXACTEMENT 1 kg/L et que l'unité de commande était DÉJÀ
    dans la MÊME base que l'unité de référence (ex: les deux en kg), les deux conventions
    (garder KG/L, ou préciser le récipient) sont aussi valables l'une que l'autre — ce n'est
    pas une incohérence à signaler. Si l'unité de commande était dans une base DIFFÉRENTE
    (kg vs L), en revanche, ce n'était pas qu'une question de convention mais une vraie
    incohérence physique que le récipient vient de corriger — on ne l'annule pas."""
    if resultat.couleur_l is None or not resultat.nouveau_l:
        return
    base_cible = motor.normalise_texte(unite_ref)
    base_cible = "kg" if base_cible == "kg" else ("l" if base_cible in ("l", "litre") else None)
    base_cmd = motor.normalise_texte(unite_cmd)
    base_cmd = "kg" if base_cmd == "kg" else ("l" if base_cmd in ("l", "litre") else None)
    if not base_cible or base_cmd != base_cible:
        return
    if recipient_est_exactement_1_unite(conditionnement, base_cible):
        resultat.nouveau_l = None
        resultat.couleur_l = None
        resultat.commentaire_l = None


RE_CONTENANT_PIECES_SEULEMENT = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de\s+)?(\d+)\s*pi[eè]ces?\s*$", re.I)


def reconsidere_prix_inexplique(resultat, unite_ref, conditionnement, unite_cmd, prix_ref, prix_cmd):
    """Quand le BPU est au kg/L et que rien dans le conditionnement n'a pu expliquer le
    ratio Prix commande ÷ BPU via les motifs reconnus, deux cas différents :
    1. Le conditionnement ne donne qu'un NOMBRE DE PIÈCES sans poids ('CAGEOT DE 6 PIECES')
       et la commande se fait dans ce même récipient : le poids réel est inconnu (souvent
       variable) — impossible de vérifier, donc pas d'incohérence à signaler.
    2. Un mot ou code non reconnu casse le motif ('SAC KRAFT DE 5 KG') — on cherche quand
       même, en texte libre, un nombre suivi d'une unité n'importe où dans le
       Conditionnement, et si ce nombre correspond au ratio observé, ce n'est pas non plus
       une incohérence, juste un format non standard."""
    if resultat.couleur_m != motor.COULEUR_PRIX_INEXPLIQUE:
        return
    base_cible = motor.normalise_texte(unite_ref)
    if base_cible not in ("kg", "l", "litre") or not conditionnement or not prix_ref:
        return

    m = RE_CONTENANT_PIECES_SEULEMENT.match(conditionnement)
    if m and unite_cmd and motor.normalise_texte(unite_cmd) == motor.normalise_texte(m.group(1)):
        resultat.couleur_m = None
        resultat.commentaire_m = None
        return

    ratio_observe = prix_cmd / prix_ref
    candidats_trouves = []
    for qty_str, unite_str in motor.RE_LIBRE.findall(conditionnement):
        if motor.base_depuis_unite(unite_str) != ("kg" if base_cible == "kg" else "l"):
            continue
        try:
            valeur_kg = motor._nombre(qty_str) * motor.FACTEUR_VERS_KG_OU_L[unite_str.lower()]
        except (KeyError, ValueError):
            continue
        if valeur_kg <= 0:
            continue
        candidats_trouves.append(valeur_kg)
        if abs(ratio_observe - valeur_kg) <= motor.TOLERANCE_RATIO * valeur_kg:
            resultat.couleur_m = None
            resultat.commentaire_m = None
            return

    if candidats_trouves:
        # Le conditionnement suggère bien une quantité (ex: 'SAC KRAFT DE 5 KG' -> 5), mais le
        # ratio observé ne correspond à AUCUNE d'elles — trop incertain pour recalculer
        # automatiquement (le motif n'est pas propre), mais le message pointe directement le
        # problème : un des deux prix semble faux, pas juste "pas d'explication".
        quantite = candidats_trouves[0]
        resultat.commentaire_m = (
            f"Le conditionnement ('{conditionnement}') suggère {round(quantite, 3)} {base_cible} "
            f"— dans ce cas le prix de commande devrait valoir {round(quantite, 3)} fois le prix "
            f"de référence, soit {round(prix_ref * quantite, 4)}, mais il vaut {prix_cmd} (ratio "
            f"observé = {round(ratio_observe, 3)}). Un des deux prix semble incorrect — à "
            "vérifier manuellement (pas recalculé automatiquement, le format n'est pas standard).")


def verifie_unite_commande_etendue(unite_cmd, conditionnement, designation):
    """Copie de motor.verifie_unite_commande, avec des corrections :
    1. Vocabulaire étendu (cageot/boite/filet/caisse/plaque/plateau/clayette...) ;
    2. Un mot de récipient qui s'auto-désigne est reconnu comme cohérent pour N'IMPORTE
       QUEL mot de la liste (pas seulement barquette/sac/bouteille comme à l'origine) ;
    3. "N PIECES" (comme dans '2 PIECES DE 2 KG') est équivalent à "pièce"."""
    if not unite_cmd:
        return False, "Unité de commande vide."

    conditionnement = nettoie_conditionnement(conditionnement)
    cat = motor.categorie_commande(unite_cmd)
    cond = motor.normalise_texte(conditionnement)
    desig = motor.normalise_texte(designation)
    texte_combine = f"{cond} {desig}"

    if not conditionnement and not designation:
        return True, None
    if cat in ("kg", "l"):
        return True, None
    if re.match(r"^\d+\s*pi[eè]ces?$", cat):
        return True, None

    mots_types = ["sac", "pièce", "piece", "seau", "brique", "barquette", "barq", "bidon",
                  "outre", "colis", "pot", "sachet", "carton", "poche", "bouteille", "bouteile",
                  "cageot", "boite", "filet", "caisse", "plaque", "plateau", "clayette"]
    trouve = [m for m in mots_types if m in texte_combine]
    if not trouve:
        return True, None

    cat_singulier = cat[:-1] if cat.endswith("s") and cat[:-1] in mots_types else cat
    if cat_singulier in trouve:
        return True, None
    if cat == "pc" and any(m in ("pièce", "piece", "pot", "sachet") for m in trouve):
        return True, None
    if cat == "barquette" and "barquette" in trouve:
        return True, None
    if cat == "sac" and "sac" in trouve:
        return True, None
    if cat == "bouteille" and any(m in ("bouteille", "brique") for m in trouve):
        return True, None

    # Avant de conclure à une incohérence, vérifie si un mot du Conditionnement est juste une
    # faute de frappe d'un contenant connu (ex: 'BARQUETE') — dans ce cas, la faute est déjà
    # signalée séparément sur le Conditionnement ; pas la peine de signaler AUSSI l'unité de
    # commande comme incohérente avec elle.
    conditionnement_corrige, typo = corrige_faute_frappe_contenant(conditionnement)
    if typo:
        cond_corrige = motor.normalise_texte(conditionnement_corrige)
        trouve_corrige = [m for m in mots_types if m in f"{cond_corrige} {desig}"]
        if cat_singulier in trouve_corrige or (cat == "pc" and any(
                m in ("pièce", "piece", "pot", "sachet") for m in trouve_corrige)) or (
                cat == "barquette" and "barquette" in trouve_corrige) or (
                cat == "sac" and "sac" in trouve_corrige) or (cat == "bouteille" and any(
                m in ("bouteille", "brique") for m in trouve_corrige)):
            return True, None

    source = "Conditionnement" if any(m in cond for m in trouve) else "Désignation de l'article"
    texte_source = conditionnement if source == "Conditionnement" else designation
    return False, (f"L'unité de commande ('{unite_cmd}') ne semble pas correspondre au "
                    f"type de contenant décrit dans « {source} » ('{texte_source}').")


# ============================================================================
# DÉTECTION AUTOMATIQUE DES COLONNES — par en-tête, pas par numéro. C'est ce qui permet à
# ce même script de traiter des fichiers dont la mise en page diffère légèrement (colonne
# "Circuit court" présente ou non, 1 à 3 blocs Fournisseur...).
# ============================================================================

def _normalise_entete(valeur):
    if valeur is None:
        return ""
    return re.sub(r"\s+", " ", str(valeur)).strip().lower()


def localise_feuille_et_colonnes(ws):
    """Cherche, dans les 5 premières lignes de la feuille, la ligne d'en-têtes contenant
    'désignation de l'article' et 'unité de référence'. Si trouvée, renvoie un dict avec la
    ligne de début des données et les colonnes repérées (Lot, Désignation, Unité de
    référence, Poids à l'unité de référence, et la liste des blocs Fournisseur remplis avec
    leurs 6 colonnes : conditionnement, désignation fournisseur, référence, prix de
    référence, unité de commande, prix de commande). Renvoie None si cette feuille ne
    correspond pas au format attendu (ex: la feuille 'lot 3 Produits élaborés')."""
    ligne_entete = None
    for r in range(1, 6):
        textes = [_normalise_entete(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        if any("désignation de l'article" in t for t in textes) and \
           any("unité de référence" in t for t in textes):
            ligne_entete = r
            break
    if ligne_entete is None:
        return None

    col_marche = col_lot = col_designation = col_unite_ref = col_poids_ref = None
    fournisseurs_col = []  # [(numero, colonne_du_label_fournisseur)]
    for c in range(1, ws.max_column + 1):
        texte = _normalise_entete(ws.cell(row=ligne_entete, column=c).value)
        if texte == "marché":
            col_marche = c
        elif texte == "lot":
            col_lot = c
        elif "désignation de l'article" in texte:
            col_designation = c
        elif "poids à l'unité de référence" in texte:
            col_poids_ref = c
        elif "unité de référence" in texte:
            col_unite_ref = c
        else:
            m = re.match(r"fournisseur\s*(\d+)", texte)
            if m:
                fournisseurs_col.append((int(m.group(1)), c))

    if col_designation is None or col_unite_ref is None or not fournisseurs_col:
        return None

    fournisseurs = []
    for numero, col in sorted(fournisseurs_col):
        fournisseurs.append({
            "nom": f"Fournisseur {numero}",
            "cond": col + 1, "desig_fourn": col + 2, "ref": col + 3,
            "prix_ref": col + 4, "unite_cmd": col + 5, "prix_cmd": col + 6,
        })

    derniere_colonne_cle = max(f["prix_cmd"] for f in fournisseurs) + 6  # marge (Local/Bio/.../Indice I0)

    return {
        "ligne_debut_donnees": ligne_entete + 1,
        "col_marche": col_marche,
        "col_lot": col_lot,
        "col_designation": col_designation,
        "col_unite_ref": col_unite_ref,
        "col_poids_ref": col_poids_ref,
        "fournisseurs": fournisseurs,
        "derniere_colonne_cle_doublon": derniere_colonne_cle,
    }


def _base_comparable(unite):
    """Regroupe les unités comparables entre elles pour la comparaison de prix : KG, L, et
    PC/pièce/unité (toutes regroupées sous 'pc', car le prix y est déjà par pièce, comme en
    kg/L). Renvoie None si l'unité n'est pas de celles-là (pas comparable directement)."""
    base = motor.normalise_texte(unite)
    if base == "kg":
        return "kg"
    if base in ("l", "litre"):
        return "l"
    if base in ("pc", "piece", "pièce", "unite", "unité"):
        return "pc"
    return None


def reconsidere_recalcul_k_vs_indice_prix(resultat, indice_prix, designation, lot_valor,
                                           prix_ref, prix_cmd):
    """Avant d'appliquer un recalcul de K proposé par le bloc 'conditionnement' (K = prix de
    commande ÷ quantité), vérifie que ce nouveau K est bien PLUS plausible que l'ancien, selon
    l'indice de prix comparables — pas l'inverse. Cas typique : le prix de référence (K) est
    déjà cohérent avec les produits comparables, mais le prix de commande (M) contient une
    erreur de saisie (ex: deux zéros oubliés) ; le recalcul, qui suppose M toujours fiable,
    donne alors un K ABSURDE alors que l'ancien était bon. Dans ce cas, on annule le recalcul
    et on signale plutôt le prix de commande, en expliquant à quoi il devrait ressembler."""
    if not indice_prix or resultat.nouveau_k is None or not prix_ref:
        return
    base_final = _base_comparable(resultat.nouveau_d) if resultat.nouveau_d else None
    if base_final not in ("kg", "l"):
        return

    stats = indice_prix["produit"].get((base_final, motor.normalise_pour_doublon(designation)))
    portee = f"des lignes comparables pour '{designation}'"
    if not stats and lot_valor:
        stats = indice_prix["categorie"].get((base_final, motor.normalise_pour_doublon(lot_valor)))
        portee = f"la catégorie '{lot_valor}'"
    if not stats:
        return
    mediane, n_observations = stats

    ancien_plausible = (mediane / 3) <= prix_ref <= (mediane * 3)
    nouveau_implausible = not ((mediane / 5) <= resultat.nouveau_k <= (mediane * 5))
    if not (ancien_plausible and nouveau_implausible):
        return

    quantite = prix_cmd / resultat.nouveau_k
    prix_cmd_attendu = round(prix_ref * quantite, 2)
    resultat.nouveau_k = None
    resultat.couleur_k = None
    resultat.commentaire_k = None
    resultat.couleur_m = motor.COULEUR_RATIO_SUSPECT
    resultat.commentaire_m = (
        f"Le prix de référence actuel ({prix_ref}) est cohérent avec {portee} "
        f"({n_observations} lignes, médiane {round(mediane, 4)} / {base_final}), mais le prix "
        f"de commande ({prix_cmd}) ne correspond pas à la quantité du conditionnement "
        f"({round(quantite, 3)} {base_final}) : attendu environ {prix_cmd_attendu}. "
        "Vérifiez le prix de commande (ex: zéro(s) oublié(s) lors de la saisie) — le prix de "
        "référence n'a PAS été recalculé, il semble déjà correct.")


def construit_indice_prix_depuis_classeurs(classeurs):
    """Cœur de construit_indice_prix(), séparé pour pouvoir être appelé aussi bien depuis des
    fichiers sur disque (CLI) que depuis des classeurs déjà en mémoire (ex: fichiers déposés
    dans l'application Streamlit). 'classeurs' est une liste de classeurs openpyxl ouverts
    avec data_only=True. Voir construit_indice_prix() pour le détail du résultat."""
    brut_produit = {}
    brut_categorie = {}
    for wb in classeurs:
        for nom_feuille in wb.sheetnames:
            ws = wb[nom_feuille]
            info = localise_feuille_et_colonnes(ws)
            if info is None:
                continue
            col_designation = info["col_designation"]
            col_unite_ref = info["col_unite_ref"]
            col_lot = info["col_lot"]
            for ligne in range(info["ligne_debut_donnees"], ws.max_row + 1):
                designation = ws.cell(row=ligne, column=col_designation).value
                unite_ref = ws.cell(row=ligne, column=col_unite_ref).value
                base = _base_comparable(unite_ref) if unite_ref else None
                if not designation or not base:
                    continue
                cle_produit = (base, motor.normalise_pour_doublon(designation))
                lot_valor = ws.cell(row=ligne, column=col_lot).value if col_lot else None
                cle_categorie = (base, motor.normalise_pour_doublon(lot_valor)) if lot_valor else None
                for fournisseur in info["fournisseurs"]:
                    prix_ref = ws.cell(row=ligne, column=fournisseur["prix_ref"]).value
                    if isinstance(prix_ref, (int, float)) and prix_ref > 0:
                        brut_produit.setdefault(cle_produit, []).append(float(prix_ref))
                        if cle_categorie:
                            brut_categorie.setdefault(cle_categorie, []).append(float(prix_ref))

    # Réduit directement aux statistiques utiles (médiane, nombre d'observations) : inutile
    # de garder toutes les valeurs brutes une fois la médiane calculée.
    def resume(brut, minimum):
        return {cle: (statistics.median(prix), len(prix))
                for cle, prix in brut.items() if len(prix) >= minimum}

    return {"produit": resume(brut_produit, 3), "categorie": resume(brut_categorie, 8)}


def construit_indice_prix(dossier_entree):
    """Parcourt, EN LECTURE SEULE, tous les fichiers de raw_data/ pour relever tous les prix
    de référence (BPU) observés, toutes lignes et tous fournisseurs confondus, dans tous les
    fichiers, regroupés de deux façons :
      - 'produit' : (unité de base, désignation de l'article normalisée) — comparaison la
        plus précise (même produit exact), mais parfois trop peu d'observations ;
      - 'categorie' : (unité de base, Lot normalisé, ex: '13 - charcuterie') — bien plus
        d'observations, utile quand le produit exact est unique dans les fichiers (donc
        jamais comparé à rien), au prix d'une comparaison moins précise (produits différents
        dans un même Lot).
    Sert à repérer un prix par kg/L/pièce très éloigné de celui des produits comparables
    (ex : un facteur ×10 ou ×100, souvent le signe d'une erreur de saisie plutôt que d'un
    vrai écart de prix)."""
    if not os.path.isdir(dossier_entree):
        return {"produit": {}, "categorie": {}}
    classeurs = []
    for nom_fichier in sorted(os.listdir(dossier_entree)):
        if not nom_fichier.lower().endswith(".xlsx") or nom_fichier.startswith("~$"):
            continue
        try:
            classeurs.append(openpyxl.load_workbook(os.path.join(dossier_entree, nom_fichier), data_only=True))
        except Exception:
            continue
    resultat = construit_indice_prix_depuis_classeurs(classeurs)
    for wb in classeurs:
        wb.close()
    return resultat


# ============================================================================
# MARQUAGE — couleurs et commentaires (jamais de prix modifié, sauf les 2 exceptions notées
# en tête de fichier).
# ============================================================================

AUTEUR_COMMENTAIRE = "Audit de cohérence (lecture seule)"
COULEUR_DUPLICATA = "FFFF0000"  # rouge plein — toute la ligne, pas juste une cellule

COULEURS_INCOHERENCE = {
    motor.COULEUR_UNITE_TRANSFORMEE,
    motor.COULEUR_A_VERIFIER,
    motor.COULEUR_RATIO_SUSPECT,
    motor.COULEUR_COMMANDE_SUSPECTE,
    motor.COULEUR_PRIX_INEXPLIQUE,
    motor.COULEUR_COUSTERON,
}


def majuscule_si_inchangee(valeur):
    """Renvoie la version MAJUSCULE d'une unité (D ou L) quand elle n'est pas déjà toute en
    majuscules, sinon None. Utilisé pour forcer la casse (ex: 'kg' -> 'KG', 'pièce' ->
    'PIÈCE') même quand aucune autre règle n'a de raison de toucher la cellule — JAMAIS
    marqué (pas de couleur ni de commentaire), puisque ce n'est qu'une question de casse."""
    if not isinstance(valeur, str):
        return None
    maj = valeur.upper()
    return maj if maj != valeur else None


def marque(cel, couleur, fournisseur_nom, texte):
    cel.fill = PatternFill(start_color=couleur, end_color=couleur, fill_type="solid")
    prefixe = f"[{fournisseur_nom}] " if fournisseur_nom else ""
    ligne = f"{prefixe}{texte}"
    if cel.comment:
        cel.comment.text = cel.comment.text + "\n\n" + ligne
    else:
        cel.comment = Comment(ligne, AUTEUR_COMMENTAIRE)


RE_CATEGORIE_LEGUME = re.compile(r"l[ée]gume|verdure|fruit", re.I)


def applique_defaut_legume(resultat, lot_valor, designation, conditionnement, unite_cmd,
                            prix_ref=None, prix_cmd=None):
    """Pour un produit de catégorie légume/verdure/fruit : si rien (ni le conditionnement, ni
    E) n'a permis d'expliquer le ratio des prix (couleur_m = prix inexpliqué), OU si l'unité
    de commande reste un simple 'KG'/'L' alors que le ratio Prix commande ÷ Prix référence est
    clairement différent de 1 (un conditionnement bare comme '12 KG' donne un calcul cohérent,
    mais 'KG' comme unité de commande n'a pas de sens pour un prix qui est en réalité celui
    d'un lot/sac de 12 kg, pas d'un seul kilo) — et qu'aucun mot de récipient n'est
    identifiable dans le conditionnement, on regarde la désignation ; si elle non plus ne
    donne rien, on suppose 'SAC' par défaut (l'emballage le plus courant pour ces produits) —
    à confirmer, mais ça évite de laisser la case avec une unité qui ne veut rien dire."""
    prix_inexplique = resultat.couleur_m == motor.COULEUR_PRIX_INEXPLIQUE
    label_kg_l_malgre_ratio = (
        resultat.nouveau_l is None
        and motor.normalise_texte(unite_cmd) in ("kg", "l", "litre")
        and prix_ref and prix_cmd
        and abs(prix_cmd / prix_ref - 1) > motor.TOLERANCE_RATIO)
    if not (prix_inexplique or label_kg_l_malgre_ratio):
        return
    if not lot_valor or not RE_CATEGORIE_LEGUME.search(str(lot_valor)):
        return
    if motor.mot_type_conditionnement(conditionnement):
        return  # le conditionnement donne déjà un mot de récipient, pas besoin de deviner
    mot_designation = motor.cherche_unite_designation(designation)
    nouveau_l = mot_designation or "SAC"
    if motor.normalise_texte(nouveau_l) == motor.normalise_texte(unite_cmd):
        return
    resultat.couleur_m = None
    resultat.nouveau_l = nouveau_l
    resultat.couleur_l = motor.COULEUR_A_VERIFIER
    origine = "la désignation" if mot_designation else "aucune information disponible (par défaut)"
    raison = (f"prix non expliqué par le conditionnement ('{conditionnement}')" if prix_inexplique
              else f"le conditionnement ('{conditionnement}') ne nomme aucun récipient, et le "
                   "ratio Prix commande ÷ Prix référence est différent de 1 — 'KG' seul n'a pas "
                   "de sens comme unité de commande dans ce cas")
    resultat.commentaire_l = (
        f"Catégorie légume/fruit, {raison} — unité de commande proposée d'après {origine} : "
        f"'{nouveau_l}'. À confirmer.")


def audite_fournisseur(ws, ws_valeurs, ligne, designation, unite_ref, poids_ref, fournisseur,
                        col_poids_ref, lot_valor=None, indice_prix=None):
    conditionnement = ws_valeurs.cell(row=ligne, column=fournisseur["cond"]).value
    desig_fourn = ws_valeurs.cell(row=ligne, column=fournisseur["desig_fourn"]).value
    reference = ws_valeurs.cell(row=ligne, column=fournisseur["ref"]).value
    unite_cmd = ws_valeurs.cell(row=ligne, column=fournisseur["unite_cmd"]).value
    prix_ref = ws_valeurs.cell(row=ligne, column=fournisseur["prix_ref"]).value
    prix_cmd = ws_valeurs.cell(row=ligne, column=fournisseur["prix_cmd"]).value

    if reference is None and conditionnement is None and prix_ref is None and prix_cmd is None:
        return 0  # ce fournisseur n'est simplement pas renseigné pour cette ligne

    total = 0

    # Référence de plage de cellules Excel ('H33:H66') collée par erreur au milieu du
    # Conditionnement : aucune ambiguïté possible (ce n'est jamais du texte produit légitime),
    # donc on corrige directement la cellule, contrairement au reste du Conditionnement qui
    # n'est lui jamais modifié.
    if conditionnement:
        conditionnement_sans_plage, corrige = retire_plage_cellule_collee(conditionnement)
        if corrige:
            cel_cond = ws.cell(row=ligne, column=fournisseur["cond"])
            cel_cond.value = conditionnement_sans_plage
            cel_cond.fill = PatternFill(start_color=motor.COULEUR_UNITE_TRANSFORMEE,
                                         end_color=motor.COULEUR_UNITE_TRANSFORMEE, fill_type="solid")
            texte = (
                f"Conditionnement corrigé : une référence de plage de cellules Excel "
                f"('{conditionnement}' contenait une plage du type 'H33:H66') semble avoir été "
                "collée par erreur au milieu du texte. Valeur nettoyée automatiquement — "
                "vérifiez qu'elle est bien correcte.")
            if cel_cond.comment:
                cel_cond.comment.text = cel_cond.comment.text + "\n\n" + texte
            else:
                cel_cond.comment = Comment(texte, AUTEUR_COMMENTAIRE)
            total += 1
            conditionnement = conditionnement_sans_plage

    # Faute de frappe sur un mot de contenant (ex: 'BARQUETE' au lieu de 'BARQUETTE') :
    # seulement quand l'analyse normale échouerait sinon (pas de faux positif sur un mot qui
    # se trouve juste être proche d'un contenant connu par coïncidence). Correction purement
    # orthographique, sans ambiguïté sur le fond — appliquée directement et silencieusement,
    # sans couleur ni commentaire.
    if conditionnement and unite_ref:
        conditionnement_propre = nettoie_conditionnement(conditionnement)
        if (piece_carton_candidats(unite_ref, conditionnement_propre) is None
                and determine_base_et_quantite(conditionnement_propre) is None):
            conditionnement_corrige, typo = corrige_faute_frappe_contenant(conditionnement_propre)
            if typo:
                ws.cell(row=ligne, column=fournisseur["cond"]).value = conditionnement_corrige
                conditionnement = conditionnement_corrige

    if prix_ref is None:
        marque(ws.cell(row=ligne, column=fournisseur["prix_ref"]), motor.COULEUR_A_VERIFIER,
               fournisseur["nom"], "Prix de référence manquant.")
        total += 1
    if prix_cmd is None:
        marque(ws.cell(row=ligne, column=fournisseur["prix_cmd"]), motor.COULEUR_A_VERIFIER,
               fournisseur["nom"], "Prix de commande manquant.")
        total += 1

    if unite_ref and prix_ref is not None and prix_cmd is not None:
        resultat = traite_ligne_etendue(unite_ref, poids_ref, designation, conditionnement,
                                         unite_cmd, prix_ref, prix_cmd, desig_fourn)

        applique_defaut_legume(resultat, lot_valor, designation, conditionnement, unite_cmd,
                                prix_ref, prix_cmd)

        flag_cond = getattr(resultat, "flag_conditionnement", None)
        if flag_cond:
            couleur_cond, texte_cond = flag_cond
            marque(ws.cell(row=ligne, column=fournisseur["cond"]), couleur_cond,
                   fournisseur["nom"], texte_cond)
            total += 1

        reconsidere_recalcul_k_vs_indice_prix(resultat, indice_prix, designation, lot_valor,
                                               prix_ref, prix_cmd)

        d_final = resultat.nouveau_d if resultat.nouveau_d is not None else unite_ref
        base_final_d = motor.normalise_texte(d_final) if d_final else None

        if col_poids_ref and base_final_d in ("kg", "l", "litre"):
            # L'unité de référence est en kg/L (qu'elle l'ait toujours été, ou qu'elle vienne
            # de le devenir) : la colonne Poids à l'unité de référence ne sert QUE pour les
            # unités à la pièce (poids d'UNE pièce) — si elle contient encore une valeur, elle
            # n'a plus de sens, donc on la vide. Pas de couleur ni de commentaire : conséquence
            # mécanique du passage en kg/L, pas une incohérence en soi.
            cel_e = ws.cell(row=ligne, column=col_poids_ref)
            if cel_e.value is not None:
                cel_e.value = None
        else:
            explication_e = getattr(resultat, "explication_e", None)
            if col_poids_ref and resultat.nouveau_e not in ("GARDER", "VIDER", None):
                if not explication_e:
                    # motor.traite_ligne renseigne parfois nouveau_e directement (ex:
                    # l'exception K = M, "comparaison volontaire à la pièce") sans poser
                    # explication_e (propre à audit_geral) — on en synthétise une à partir de
                    # ce qu'on sait plutôt que de remplir la case sans aucune explication
                    # visible.
                    explication_e = (
                        f"Poids à l'unité de référence renseigné à {resultat.nouveau_e} kg, "
                        f"déduit du Conditionnement ('{conditionnement}'), de la Désignation "
                        f"fournisseur ('{desig_fourn}') ou de la Désignation de l'article "
                        f"('{designation}'). " + (resultat.commentaire_d or ""))
                cel_e = ws.cell(row=ligne, column=col_poids_ref)
                if cel_e.value is None:
                    cel_e.value = resultat.nouveau_e
                    cel_e.fill = PatternFill(start_color=motor.COULEUR_UNITE_TRANSFORMEE,
                                              end_color=motor.COULEUR_UNITE_TRANSFORMEE, fill_type="solid")
                    cel_e.comment = Comment(f"[{fournisseur['nom']}] {explication_e}", AUTEUR_COMMENTAIRE)
                    total += 1

        d_differe = (resultat.nouveau_d is not None
                     and motor.normalise_texte(resultat.nouveau_d) != motor.normalise_texte(unite_ref))
        l_differe = (resultat.nouveau_l is not None
                     and motor.normalise_texte(resultat.nouveau_l) != motor.normalise_texte(unite_cmd))
        k_differe = (resultat.nouveau_k is not None
                     and abs(resultat.nouveau_k - prix_ref) > 0.0005)

        # Écrit toujours la valeur normalisée (D en KG/L/PC, L quand il est fiable, K quand
        # recalculé) — même une simple normalisation silencieuse (kg -> KG) est appliquée.
        # Seule la mise en évidence (couleur + commentaire) dépend d'un changement réel. Et si
        # aucune règle n'a eu de raison de toucher D/L, on force quand même la casse en
        # MAJUSCULES (jamais marqué, ce n'est qu'une question de casse).
        if resultat.nouveau_d is not None:
            valeur_d = resultat.nouveau_d.upper() if isinstance(resultat.nouveau_d, str) else resultat.nouveau_d
            ws.cell(row=ligne, column=fournisseur["_col_unite_ref"]).value = valeur_d
        else:
            maj = majuscule_si_inchangee(unite_ref)
            if maj:
                ws.cell(row=ligne, column=fournisseur["_col_unite_ref"]).value = maj
        if resultat.nouveau_l is not None:
            valeur_l = resultat.nouveau_l.upper() if isinstance(resultat.nouveau_l, str) else resultat.nouveau_l
            ws.cell(row=ligne, column=fournisseur["unite_cmd"]).value = valeur_l
        else:
            maj = majuscule_si_inchangee(unite_cmd)
            if maj:
                ws.cell(row=ligne, column=fournisseur["unite_cmd"]).value = maj
        if resultat.nouveau_k is not None:
            ws.cell(row=ligne, column=fournisseur["prix_ref"]).value = resultat.nouveau_k

        if resultat.couleur_d in COULEURS_INCOHERENCE and resultat.commentaire_d and (
                d_differe or resultat.nouveau_d is None):
            marque(ws.cell(row=ligne, column=fournisseur["_col_unite_ref"]),
                   resultat.couleur_d, fournisseur["nom"], resultat.commentaire_d)
            total += 1
        if resultat.couleur_k in COULEURS_INCOHERENCE and resultat.commentaire_k and (
                k_differe or resultat.nouveau_k is None):
            marque(ws.cell(row=ligne, column=fournisseur["prix_ref"]), resultat.couleur_k,
                   fournisseur["nom"], resultat.commentaire_k)
            total += 1
        if resultat.couleur_l in COULEURS_INCOHERENCE and resultat.commentaire_l and (
                l_differe or resultat.nouveau_l is None):
            marque(ws.cell(row=ligne, column=fournisseur["unite_cmd"]), resultat.couleur_l,
                   fournisseur["nom"], resultat.commentaire_l)
            total += 1
        if resultat.couleur_m in COULEURS_INCOHERENCE and resultat.commentaire_m:
            marque(ws.cell(row=ligne, column=fournisseur["prix_cmd"]), resultat.couleur_m,
                   fournisseur["nom"], resultat.commentaire_m)
            total += 1

        if indice_prix:
            base_final = _base_comparable(resultat.nouveau_d or unite_ref)
            k_final = resultat.nouveau_k if resultat.nouveau_k is not None else prix_ref
            # 'PC' exclu : le prix à la pièce ne dit rien du prix au kilo tant que le poids de
            # la pièce n'est pas connu avec fiabilité (sinon D serait déjà devenu 'KG') — le
            # comparer à d'autres prix à la pièce, de tailles de pièce différentes, n'a pas de
            # sens et ne fait que produire de faux signalements.
            if base_final in ("kg", "l") and k_final and k_final > 0:
                def signale_prix_outlier(stats_groupe, seuil, portee, origine):
                    mediane, n_observations = stats_groupe
                    ratio = k_final / mediane
                    if ratio <= seuil and ratio >= (1 / seuil):
                        return 0
                    comparaison = (f"×{round(ratio, 1)}" if ratio >= 1
                                    else f"÷{round(1 / ratio, 1)}")
                    marque(ws.cell(row=ligne, column=fournisseur["prix_ref"]),
                           motor.COULEUR_PRIX_OUTLIER, fournisseur["nom"],
                           f"Prix de référence ({k_final} / {base_final}) très éloigné de "
                           f"{portee} ({n_observations} lignes comparables{origine}) : médiane "
                           f"{round(mediane, 4)} / {base_final}, soit un facteur {comparaison}. "
                           "Peut être une erreur de saisie (zéro en trop/en moins, confusion "
                           "d'unité...) ou un vrai écart de prix — à vérifier.")
                    return 1

                stats_produit = indice_prix["produit"].get(
                    (base_final, motor.normalise_pour_doublon(designation)))
                if stats_produit:
                    total += signale_prix_outlier(
                        stats_produit, 3, f"celui des autres lignes de '{designation}'", "")

                if lot_valor:
                    stats_categorie = indice_prix["categorie"].get(
                        (base_final, motor.normalise_pour_doublon(lot_valor)))
                    if stats_categorie:
                        total += signale_prix_outlier(
                            stats_categorie, 5, f"celui de la catégorie '{lot_valor}'",
                            f", catégorie '{lot_valor}'")
    else:
        # Prix manquant(s) : aucune des règles ci-dessus n'a pu s'exécuter, mais la casse de
        # D/L doit quand même être forcée en MAJUSCULES (jamais marqué).
        maj_d = majuscule_si_inchangee(unite_ref)
        if maj_d:
            ws.cell(row=ligne, column=fournisseur["_col_unite_ref"]).value = maj_d
        maj_l = majuscule_si_inchangee(unite_cmd)
        if maj_l:
            ws.cell(row=ligne, column=fournisseur["unite_cmd"]).value = maj_l

    if unite_cmd:
        ok, message = verifie_unite_commande_etendue(unite_cmd, conditionnement, designation)
        if not ok:
            marque(ws.cell(row=ligne, column=fournisseur["unite_cmd"]), motor.COULEUR_COMMANDE_SUSPECTE,
                   fournisseur["nom"], message)
            total += 1

    if reference and RE_REFERENCES_MULTIPLES.search(str(reference)):
        marque(ws.cell(row=ligne, column=fournisseur["ref"]), motor.COULEUR_A_VERIFIER,
               fournisseur["nom"],
               f"Plusieurs références semblent présentes dans cette cellule ('{reference}'). "
               "Vérifiez laquelle est la bonne, ou s'il s'agit de plusieurs produits "
               "différents, séparez-les en une ligne par article/référence.")
        total += 1

    return total


def detecte_lignes_dupliquees(ws_valeurs, ligne_debut, col_designation, derniere_colonne):
    """Repère les lignes dupliquées : toutes les colonnes comparées identiques, SAUF
    éventuellement la Désignation elle-même (espace en trop, suffixe 'CE2'...). Exige un
    minimum de valeurs réellement remplies pour éviter que deux lignes très peu renseignées
    (juste Marché/Lot/Fournisseur) ne soient prises pour des doublons."""
    MINIMUM_CHAMPS_REMPLIS = 8

    colonnes_cle = [c for c in range(1, derniere_colonne + 1) if c != col_designation]
    lignes_vues = {}
    designation_normalisee = {}
    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        designation = ws_valeurs.cell(row=ligne, column=col_designation).value
        if not designation:
            continue
        designation_normalisee[ligne] = motor.normalise_pour_doublon(designation)
        valeurs = [ws_valeurs.cell(row=ligne, column=c).value for c in colonnes_cle]
        if sum(1 for v in valeurs if v not in (None, "")) < MINIMUM_CHAMPS_REMPLIS:
            continue
        cle = tuple(motor.normalise_pour_doublon(v) for v in valeurs)
        lignes_vues.setdefault(cle, []).append(ligne)

    doublons = {}
    for lignes in lignes_vues.values():
        if len(lignes) < 2:
            continue
        designations_identiques = len({designation_normalisee[l] for l in lignes}) == 1
        for ligne in lignes:
            doublons[ligne] = ([l for l in lignes if l != ligne], designations_identiques)
    return doublons


def marque_lignes_dupliquees(ws, doublons, col_designation):
    total = 0
    for ligne, (autres, designations_identiques) in doublons.items():
        premiere_occurrence = min([ligne] + autres)
        if ligne == premiere_occurrence:
            continue

        fill_rouge = PatternFill(start_color=COULEUR_DUPLICATA, end_color=COULEUR_DUPLICATA,
                                  fill_type="solid")
        for colonne in range(1, ws.max_column + 1):
            ws.cell(row=ligne, column=colonne).fill = fill_rouge

        cel = ws.cell(row=ligne, column=col_designation)
        if designations_identiques:
            texte = (f"Ligne strictement dupliquée (contenu identique) de la ligne "
                     f"{premiere_occurrence}. Toute la ligne est surlignée en rouge — ce "
                     "script ne supprime rien, à vous de choisir laquelle garder.")
        else:
            texte = (f"Ligne dupliquée de la ligne {premiere_occurrence} : toutes les colonnes "
                     "sont identiques (référence, prix, fournisseur, conditionnement...), SAUF "
                     "la Désignation, qui diffère légèrement (espace, suffixe 'CE2'...). Toute "
                     "la ligne est surlignée en rouge — ce script ne supprime rien, à vous de "
                     "choisir laquelle garder.")
        if cel.comment:
            cel.comment.text = cel.comment.text + "\n\n" + texte
        else:
            cel.comment = Comment(texte, AUTEUR_COMMENTAIRE)
        total += 1
    return total


def signale_copies_possibles_fournisseur(ws, ws_valeurs, ligne_debut, fournisseur):
    groupes = {}
    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        ref = ws_valeurs.cell(row=ligne, column=fournisseur["ref"]).value
        k = ws_valeurs.cell(row=ligne, column=fournisseur["prix_ref"]).value
        if ref and k:
            groupes.setdefault((ref, k), []).append(ligne)

    total = 0
    for (ref, k), lignes in groupes.items():
        if len(lignes) < 2:
            continue
        for ligne in lignes:
            autres = [str(l) for l in lignes if l != ligne]
            marque(ws.cell(row=ligne, column=fournisseur["ref"]), motor.COULEUR_COPIE_POSSIBLE,
                   fournisseur["nom"],
                   f"Copie possible : même Référence produit ({ref}) et même prix de "
                   f"référence ({k}) que la ligne {' et '.join(autres)}. À comparer — peut "
                   "être une simple variante/faute de frappe de désignation.")
            total += 1
    return total


def comble_marche_lot(ws, ws_valeurs, ligne_debut, col_designation, col_marche, col_lot):
    """Quand une ligne a tout de rempli SAUF Marché et Lot : si la ligne juste au-dessus et
    celle juste en-dessous ont le MÊME Marché/Lot, on les recopie (ligne "sandwich"). Sinon,
    la cellule est surlignée pour demander de préciser le Marché/Lot correspondant."""
    if not col_marche and not col_lot:
        return 0, 0
    total_rempli = total_signale = 0
    for ligne in range(max(ligne_debut + 1, ligne_debut), ws_valeurs.max_row):
        if not ws_valeurs.cell(row=ligne, column=col_designation).value:
            continue
        marche_vide = col_marche and not ws_valeurs.cell(row=ligne, column=col_marche).value
        lot_vide = col_lot and not ws_valeurs.cell(row=ligne, column=col_lot).value
        if not (marche_vide or lot_vide):
            continue

        marche_dessus = ws_valeurs.cell(row=ligne - 1, column=col_marche).value if col_marche else None
        marche_dessous = ws_valeurs.cell(row=ligne + 1, column=col_marche).value if col_marche else None
        lot_dessus = ws_valeurs.cell(row=ligne - 1, column=col_lot).value if col_lot else None
        lot_dessous = ws_valeurs.cell(row=ligne + 1, column=col_lot).value if col_lot else None

        marche_ok = (not col_marche) or (marche_dessus is not None and marche_dessus == marche_dessous)
        lot_ok = (not col_lot) or (lot_dessus is not None and lot_dessus == lot_dessous)

        if marche_ok and lot_ok:
            if col_marche and marche_vide:
                ws.cell(row=ligne, column=col_marche).value = marche_dessus
            if col_lot and lot_vide:
                ws.cell(row=ligne, column=col_lot).value = lot_dessus
            total_rempli += 1
        else:
            if col_marche and marche_vide:
                marque(ws.cell(row=ligne, column=col_marche), motor.COULEUR_A_VERIFIER, None,
                       "Marché vide, et les lignes voisines ne s'accordent pas — préciser le "
                       "Marché correspondant à cette ligne.")
            if col_lot and lot_vide:
                marque(ws.cell(row=ligne, column=col_lot), motor.COULEUR_A_VERIFIER, None,
                       "Lot vide, et les lignes voisines ne s'accordent pas — préciser le Lot "
                       "correspondant à cette ligne.")
            total_signale += 1
    return total_rempli, total_signale


def applique_formatage_uniforme(ws, ligne_debut):
    """Même police (Calibri 11) et même alignement (aligné à gauche, horizontal et vertical) pour
    toutes les colonnes, à partir de la ligne de données — calqué sur le style déjà utilisé
    dans les fichiers 'groupement'. N'affecte ni les couleurs de fond ni les commentaires.
    Garde le retour à la ligne automatique (wrap_text) de chaque cellule tel qu'il était : une
    cellule avec un texte long (ex: 'Lot' avec la liste des établissements) peut avoir une
    hauteur de ligne déjà fixée en fonction de ce retour à la ligne — l'écraser en imposant
    wrap_text=False laisserait le texte sur une seule ligne dans une ligne restée très haute,
    donnant l'impression d'un espacement énorme entre les lignes."""
    fonte = Font(name="Calibri", size=11)
    for ligne in range(ligne_debut, ws.max_row + 1):
        for colonne in range(1, ws.max_column + 1):
            cel = ws.cell(row=ligne, column=colonne)
            cel.font = fonte
            wrap_text = cel.alignment.wrap_text if cel.alignment else None
            cel.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap_text)


def normalise_lots(ws, ws_valeurs, ligne_debut, col_lot):
    if not col_lot:
        return 0
    total = 0
    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        valeur = ws_valeurs.cell(row=ligne, column=col_lot).value
        nouvelle_valeur = motor.normalise_numero_lot(valeur)
        if nouvelle_valeur != valeur:
            ws.cell(row=ligne, column=col_lot).value = nouvelle_valeur
            total += 1
    return total


TITRES_CATEGORIES_RAPPORT = {
    "doublons": "Lignes strictement dupliquées",
    "references_doublon": "Références en double (même référence + même prix)",
    "unite_ref_vide": "Unité de référence vide",
    "unite_ref": "Unité de référence incohérente",
    "prix_ref": "Prix de référence (BPU) incohérent",
    "unite_cmd": "Unité de commande incohérente",
    "prix_cmd": "Prix de commande incohérent / inexpliqué",
    "poids_rempli": "Poids à l'unité de référence — rempli automatiquement",
    "valeurs_manquantes": "Valeurs manquantes (prix de référence ou de commande)",
}


def calcule_categories_rapport(ws, ws_valeurs, ligne_debut, col_designation, col_unite_ref,
                                col_poids_ref, fournisseurs):
    """Relit les commentaires posés sur la feuille et les classe par catégorie — séparé de
    genere_rapport() pour que l'application Streamlit puisse aussi construire sa propre
    présentation (titres + compteurs, détails repliés) à partir des mêmes données."""
    categories = {cle: [] for cle in TITRES_CATEGORIES_RAPPORT}

    colonnes_ref = {f["ref"] for f in fournisseurs}
    colonnes_prix_ref = {f["prix_ref"] for f in fournisseurs}
    colonnes_unite_cmd = {f["unite_cmd"] for f in fournisseurs}
    colonnes_prix_cmd = {f["prix_cmd"] for f in fournisseurs}

    for ligne in range(ligne_debut, ws.max_row + 1):
        designation = ws_valeurs.cell(row=ligne, column=col_designation).value
        for colonne in range(1, ws.max_column + 1):
            cel = ws.cell(row=ligne, column=colonne)
            if not cel.comment:
                continue
            texte = cel.comment.text

            if colonne == col_designation and "dupliquée" in texte:
                categories["doublons"].append((ligne, designation, texte))
            elif colonne in colonnes_ref and "Copie possible" in texte:
                categories["references_doublon"].append((ligne, designation, texte))
            elif colonne == col_unite_ref:
                if "vide" in texte:
                    categories["unite_ref_vide"].append((ligne, designation, texte))
                else:
                    categories["unite_ref"].append((ligne, designation, texte))
            elif col_poids_ref and colonne == col_poids_ref:
                categories["poids_rempli"].append((ligne, designation, texte))
            elif colonne in colonnes_prix_ref:
                if "manquant" in texte:
                    categories["valeurs_manquantes"].append((ligne, designation, texte))
                else:
                    categories["prix_ref"].append((ligne, designation, texte))
            elif colonne in colonnes_unite_cmd:
                categories["unite_cmd"].append((ligne, designation, texte))
            elif colonne in colonnes_prix_cmd:
                if "manquant" in texte:
                    categories["valeurs_manquantes"].append((ligne, designation, texte))
                else:
                    categories["prix_cmd"].append((ligne, designation, texte))

    return categories


def genere_rapport(ws, ws_valeurs, nom_feuille, ligne_debut, col_designation, col_unite_ref,
                    col_poids_ref, fournisseurs, categories=None):
    if categories is None:
        categories = calcule_categories_rapport(ws, ws_valeurs, ligne_debut, col_designation,
                                                 col_unite_ref, col_poids_ref, fournisseurs)

    total_general = sum(len(v) for v in categories.values())
    lignes_sortie = [
        f"# Rapport d'audit — {nom_feuille}",
        "",
        f"Total d'éléments signalés : **{total_general}**",
        "",
    ]
    for cle, titre in TITRES_CATEGORIES_RAPPORT.items():
        elements = categories[cle]
        lignes_sortie.append(f"## {titre} ({len(elements)})")
        lignes_sortie.append("")
        if not elements:
            lignes_sortie.append("_(aucun)_")
        for ligne, designation, texte in elements:
            desig_courte = (designation or "").strip()
            lignes_sortie.append(f"- **Ligne {ligne}** — {desig_courte}")
            lignes_sortie.append(f"  {texte}")
        lignes_sortie.append("")

    return "\n".join(lignes_sortie)


RE_FORMULE_EXTERNE = re.compile(r"^=.?\[\d+\]")


def gele_formules_externes(ws, ws_valeurs):
    """Certaines cellules contiennent une formule pointant vers un classeur EXTERNE
    (ex: "='[1]BPU '!B4"), dont le lien n'est plus disponible. openpyxl ne sait pas
    réenregistrer ce type de lien correctement (il perd une des relations internes du
    fichier), ce qui transforme ces cellules en #NAME? à l'ouverture — alors qu'on n'a
    jamais touché à ces cellules. On fige donc leur dernière valeur calculée (déjà
    présente dans ws_valeurs, lu avec data_only=True) à la place de la formule, pour que
    la valeur reste visible et correcte."""
    total = 0
    for row in ws.iter_rows():
        for cel in row:
            if cel.data_type == "f" and isinstance(cel.value, str) and RE_FORMULE_EXTERNE.match(cel.value):
                valeur = ws_valeurs.cell(row=cel.row, column=cel.column).value
                ws.cell(row=cel.row, column=cel.column).value = valeur
                total += 1
    return total


def traite_feuille(ws, ws_valeurs, info, indice_prix=None):
    gele_formules_externes(ws, ws_valeurs)
    ligne_debut = info["ligne_debut_donnees"]
    col_marche = info["col_marche"]
    col_lot = info["col_lot"]
    col_designation = info["col_designation"]
    col_unite_ref = info["col_unite_ref"]
    col_poids_ref = info["col_poids_ref"]
    fournisseurs = info["fournisseurs"]
    for f in fournisseurs:
        f["_col_unite_ref"] = col_unite_ref  # utilisé par audite_fournisseur

    stats = {"lignes_avec_incoherence": 0, "marques": 0, "unite_ref_vide": 0,
              "lignes_dupliquees": 0, "copies_possibles": 0, "lots_normalises": 0,
              "marche_lot_rempli": 0, "marche_lot_signale": 0, "lignes_vides_a_supprimer": 0}

    stats["marche_lot_rempli"], stats["marche_lot_signale"] = comble_marche_lot(
        ws, ws_valeurs, ligne_debut, col_designation, col_marche, col_lot)

    stats["lots_normalises"] = normalise_lots(ws, ws_valeurs, ligne_debut, col_lot)

    for fournisseur in fournisseurs:
        stats["copies_possibles"] += signale_copies_possibles_fournisseur(
            ws, ws_valeurs, ligne_debut, fournisseur)

    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        designation = ws_valeurs.cell(row=ligne, column=col_designation).value
        if not designation:
            # Pas de désignation : si en plus AUCUN fournisseur n'a les DEUX prix (de référence
            # ET de commande) remplis, cette ligne ne peut de toute façon pas être auditée (ni
            # servir à rien en l'état) — qu'elle ait par ailleurs un conditionnement/une
            # référence renseignés ou non, elle est candidate à la suppression. On ne la saute
            # donc pas silencieusement, on la signale.
            aucun_prix_complet = all(
                ws_valeurs.cell(row=ligne, column=f["prix_ref"]).value is None
                or ws_valeurs.cell(row=ligne, column=f["prix_cmd"]).value is None
                for f in fournisseurs)
            if aucun_prix_complet and any(
                    ws_valeurs.cell(row=ligne, column=c).value is not None
                    for c in range(1, info["derniere_colonne_cle_doublon"] + 1)):
                # Toute la ligne en rouge, comme pour les doublons — candidate à la
                # suppression, pas seulement une cellule signalée.
                fill_rouge = PatternFill(start_color=COULEUR_DUPLICATA, end_color=COULEUR_DUPLICATA,
                                          fill_type="solid")
                for colonne in range(1, ws.max_column + 1):
                    ws.cell(row=ligne, column=colonne).fill = fill_rouge
                marque(ws.cell(row=ligne, column=col_designation), COULEUR_DUPLICATA, None,
                       "Ligne sans désignation, et sans prix de référence ET de commande "
                       "réunis pour aucun fournisseur — probablement à supprimer (ligne "
                       "résiduelle/incomplète).")
                stats["marques"] += 1
                stats["lignes_avec_incoherence"] += 1
                stats["lignes_vides_a_supprimer"] += 1
            continue

        unite_ref = ws_valeurs.cell(row=ligne, column=col_unite_ref).value
        poids_ref = ws_valeurs.cell(row=ligne, column=col_poids_ref).value if col_poids_ref else None
        lot_valor = ws_valeurs.cell(row=ligne, column=col_lot).value if col_lot else None

        if not unite_ref:
            marque(ws.cell(row=ligne, column=col_unite_ref), motor.COULEUR_A_VERIFIER, None,
                   "Unité de référence vide — à compléter manuellement.")
            stats["unite_ref_vide"] += 1
            stats["marques"] += 1
            continue

        # Force la casse en MAJUSCULES dès maintenant (jamais marqué) : couvre aussi les
        # lignes où AUCUN fournisseur n'a de prix (donc audite_fournisseur() ne s'exécute
        # jamais assez pour s'en charger lui-même).
        maj_d = majuscule_si_inchangee(unite_ref)
        if maj_d:
            ws.cell(row=ligne, column=col_unite_ref).value = maj_d

        ligne_a_un_probleme = False
        for fournisseur in fournisseurs:
            n = audite_fournisseur(ws, ws_valeurs, ligne, designation, unite_ref, poids_ref,
                                    fournisseur, col_poids_ref, lot_valor, indice_prix)
            if n:
                ligne_a_un_probleme = True
                stats["marques"] += n
        if ligne_a_un_probleme:
            stats["lignes_avec_incoherence"] += 1

    # Appliqué en dernier pour que le rouge des lignes dupliquées ne soit jamais recouvert par
    # une autre couleur posée cellule par cellule plus haut.
    doublons = detecte_lignes_dupliquees(ws_valeurs, ligne_debut, col_designation,
                                         info["derniere_colonne_cle_doublon"])
    doublons_ref = detecte_doublons_par_reference(ws_valeurs, ligne_debut, fournisseurs)
    for ligne, (autres, _) in doublons_ref.items():
        if ligne not in doublons:
            doublons[ligne] = (autres, False)
    stats["lignes_dupliquees"] = marque_lignes_dupliquees(ws, doublons, col_designation)

    applique_formatage_uniforme(ws, ligne_debut)

    return stats


def detecte_doublons_par_reference(ws_valeurs, ligne_debut, fournisseurs):
    """Même Référence produit + même Conditionnement + même Prix de référence, pour un même
    fournisseur : signe très fort qu'il s'agit de la MÊME ligne de produit, même si la
    désignation diffère un peu (ex: 'œufs' vs 'œuf') et même si le poids/prix de commande
    (souvent recalculés ou mal copiés) diffèrent. Traité comme une ligne dupliquée (rouge),
    pas seulement 'référence en double' (rose)."""
    lignes_vues = {}
    for fournisseur in fournisseurs:
        for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
            ref = ws_valeurs.cell(row=ligne, column=fournisseur["ref"]).value
            cond = ws_valeurs.cell(row=ligne, column=fournisseur["cond"]).value
            k = ws_valeurs.cell(row=ligne, column=fournisseur["prix_ref"]).value
            if ref is None or cond is None or k is None:
                continue
            cle = (fournisseur["nom"], motor.normalise_pour_doublon(ref),
                   motor.normalise_pour_doublon(cond), k)
            lignes_vues.setdefault(cle, []).append(ligne)

    doublons = {}
    for lignes in lignes_vues.values():
        if len(lignes) < 2:
            continue
        for ligne in lignes:
            autres = [l for l in lignes if l != ligne]
            if ligne in doublons:
                doublons[ligne] = (sorted(set(doublons[ligne][0]) | set(autres)), False)
            else:
                doublons[ligne] = (autres, False)
    return doublons


def audite_classeur(wb_valeurs, wb, indice_prix=None):
    """Audite TOUTES les feuilles reconnues d'un classeur déjà ouvert (data_only=True pour
    wb_valeurs, data_only=False pour wb, le classeur réellement modifié) — cœur partagé par
    main() (fichiers sur disque) et par l'application Streamlit (fichiers déposés en mémoire).
    Modifie wb sur place (couleurs, commentaires, Lot, Poids à l'unité de référence) et
    renvoie la liste des feuilles traitées : [{"nom", "fournisseurs", "stats", "rapport",
    "categories"}, ...] ("categories" : mêmes données que le rapport, mais sous forme de
    dict {cle: [(ligne, designation, texte), ...]}, utile pour un affichage interactif comme
    celui de l'application Streamlit). Liste vide si aucune feuille de ce classeur n'a le
    format attendu."""
    resultats = []
    for nom_feuille in wb.sheetnames:
        ws_valeurs = wb_valeurs[nom_feuille]
        ws = wb[nom_feuille]
        info = localise_feuille_et_colonnes(ws_valeurs)
        if info is None:
            continue  # feuille dans un format différent (ex: 'lot 3 Produits élaborés')

        stats = traite_feuille(ws, ws_valeurs, info, indice_prix)
        categories = calcule_categories_rapport(ws, ws_valeurs, info["ligne_debut_donnees"],
                                                  info["col_designation"], info["col_unite_ref"],
                                                  info["col_poids_ref"], info["fournisseurs"])
        rapport = genere_rapport(ws, ws_valeurs, nom_feuille, info["ligne_debut_donnees"],
                                  info["col_designation"], info["col_unite_ref"],
                                  info["col_poids_ref"], info["fournisseurs"], categories)
        resultats.append({
            "nom": nom_feuille,
            "fournisseurs": [f["nom"] for f in info["fournisseurs"]],
            "stats": stats,
            "rapport": rapport,
            "categories": categories,
        })
    return resultats


def main():
    if not os.path.isdir(DOSSIER_ENTREE):
        print(f"Dossier introuvable : {DOSSIER_ENTREE}")
        return
    os.makedirs(DOSSIER_SORTIE, exist_ok=True)

    fichiers = sorted(f for f in os.listdir(DOSSIER_ENTREE) if f.lower().endswith(".xlsx")
                       and not f.startswith("~$"))
    if not fichiers:
        print(f"Aucun fichier .xlsx trouvé dans {DOSSIER_ENTREE}")
        return

    print("Construction de l'indice de prix comparables (lecture de tous les fichiers)...")
    indice_prix = construit_indice_prix(DOSSIER_ENTREE)
    print(f"  {len(indice_prix['produit'])} produits comparables (≥ 3 lignes) et "
          f"{len(indice_prix['categorie'])} catégories comparables (≥ 8 lignes) repérés.")

    for nom_fichier in fichiers:
        chemin_entree = os.path.join(DOSSIER_ENTREE, nom_fichier)
        print(f"\n=== {nom_fichier} ===")
        try:
            wb_valeurs = openpyxl.load_workbook(chemin_entree, data_only=True)
            wb = openpyxl.load_workbook(chemin_entree, data_only=False)
        except Exception as exc:
            print(f"  Impossible d'ouvrir ce fichier : {exc}")
            continue

        nom_base = os.path.splitext(nom_fichier)[0]
        resultats = audite_classeur(wb_valeurs, wb, indice_prix)

        if not resultats:
            print("  Aucune feuille reconnue dans ce fichier (format inattendu) — ignoré.")
            continue

        for resultat in resultats:
            nom_feuille = resultat["nom"]
            stats = resultat["stats"]
            print(f"  Feuille '{nom_feuille}' : colonnes repérées ({', '.join(resultat['fournisseurs'])})")

            total_marques = stats["marques"] + stats["lignes_dupliquees"] + stats["copies_possibles"]
            print(f"    Marché/Lot recopiés (lignes voisines identiques) : {stats['marche_lot_rempli']}")
            print(f"    Marché/Lot signalés (à préciser manuellement) : {stats['marche_lot_signale']}")
            print(f"    Lignes vides (ni désignation ni prix) peintes en rouge : {stats['lignes_vides_a_supprimer']}")
            print(f"    Numéros de lot normalisés : {stats['lots_normalises']}")
            print(f"    Unité de référence vide : {stats['unite_ref_vide']}")
            print(f"    Lignes dupliquées peintes en rouge : {stats['lignes_dupliquees']}")
            print(f"    Références en double : {stats['copies_possibles']}")
            print(f"    Lignes avec incohérence signalée : {stats['lignes_avec_incoherence']}")
            print(f"    Total de cellules marquées : {total_marques}")

            nom_rapport = f"{nom_base} - rapport.md" if len(wb.sheetnames) <= 2 else \
                f"{nom_base} - {nom_feuille} - rapport.md"
            with open(os.path.join(DOSSIER_SORTIE, nom_rapport), "w", encoding="utf-8") as f:
                f.write(resultat["rapport"])

        chemin_sortie = os.path.join(DOSSIER_SORTIE, f"{nom_base} - audit.xlsx")
        wb.save(chemin_sortie)
        print(f"  Fichier enregistré : {chemin_sortie}")

    print("\nTerminé. Aucun prix ni unité n'a été modifié dans les fichiers sources — seules "
          "la colonne Lot et, quand elle était vide et confirmée de façon fiable, la colonne "
          "Poids à l'unité de référence ont été renseignées. Tout le reste n'est que "
          "surlignage + commentaire, à valider par une personne.")


if __name__ == "__main__":
    main()
