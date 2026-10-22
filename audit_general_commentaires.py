"""
Audit de cohérence unités/prix — VERSION "COMMENTAIRES SEULEMENT"
================================================================================

Variante de `audit_general.py` demandée pour une version du fichier où AUCUNE case n'est
modifiée, à deux exceptions près : le "X" de Bio et le nom du label (ex: "LABEL ROUGE")
quand ils sont trouvés dans le texte. Tout le reste — même ce qui, dans `audit_general.py`,
était rempli automatiquement avec confiance (Unité de référence, Unité de commande, Poids à
l'unité de référence, Nombre de pièces, Marché/Lot recopiés, numéro de Lot avec le zéro
devant...) — n'est ici que surligné avec un commentaire expliquant simplement ce qu'on pense
que la valeur devrait être. Les commentaires sont volontairement écrits en langage courant,
sans vocabulaire technique (pas de "ratio", de "÷", de formule) : compréhensibles par
quelqu'un qui ne travaille pas avec ce fichier au quotidien.

Lit les mêmes fichiers que `audit_general.py` (dossier `pret_pour_audit/`), mais écrit dans un
dossier séparé (`commented_files/`) pour ne jamais mélanger les deux versions.

Usage :
    python audit_general_commentaires.py
    (lit tous les .xlsx de pret_pour_audit/, écrit les fichiers commentés + rapports dans
    commented_files/)
"""

import os
import re
import statistics
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import first_scripts.first_attempt_v2_prototype as motor  # noqa: E402  (règles de cohérence déjà validées)

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill, Font, Alignment

# ============================================================================
# DOSSIERS
# ============================================================================

DOSSIER_PROJET = os.path.dirname(os.path.abspath(__file__))
DOSSIER_ENTREE = os.path.join(DOSSIER_PROJET, "pret_pour_audit")
DOSSIER_SORTIE = os.path.join(DOSSIER_PROJET, "commented_files")

# ============================================================================
# EXTENSION DE VOCABULAIRE — récipients, unités et fautes de frappe fréquentes,
# accumulées au fil des fichiers déjà traités (fruits/légumes/viande ET laitier/traiteur).
# Ces objets surchargent ceux importés de first_attempt_v2_prototype.py, sans modifier ce
# fichier source partagé.
# ============================================================================

MOTS_CONTENANT = ["sac", "seau", "brique", "barquette", "barq", "bidon", "outre", "colis",
                  "pot", "sachet", "carton", "poche", "cageot", "boite", "boîte", "filet",
                  "bouteille",
                  "bouteile",  # faute de frappe fréquente ('BOUTEILES', 1 seul L)
                  "caisse", "plaque", "plateau", "clayette", "pack", "gourde", "botte",
                  "bac", "coffret", "plaquette", "paquet", "etui", "étui", "bombe",
                  "bocal", "flacon", "tube"]
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

# Tolérance utilisée partout pour valider un ratio Prix commande ÷ Prix référence contre une
# quantité attendue (5 % jugé trop permissif, 2 % trop strict — des cas réels ne passaient pas
# à très peu près — stabilisé à 3 %).
motor.TOLERANCE_RATIO = 0.03

# Palette de couleurs simplifiée (demandée pour remplacer les 8 couleurs d'origine, dont
# certaines se recouvraient sémantiquement) :
# - Mudança direta / changement confirmé (D/K/L recalculés OU LOT confirmé par le ratio,
#   mêmes conventions, aucune vraie différence) : une seule couleur.
# - Signalements "incohérence" / "sans explication" : chacun sa couleur, mais fusionnés entre
#   les variantes qui ne faisaient de toute façon que des doublons historiques.
# - Cellule vide / aucune conversion fiable : gris neutre.
# - Copie possible (même Référence+prix, ou même Désignation de l'article) : une seule couleur.
# - Ligne dupliquée entière : reste en rouge plein (inchangé).
motor.COULEUR_LOT_CONFIRME = motor.COULEUR_UNITE_TRANSFORMEE = "FF93CDDC"
motor.COULEUR_A_VERIFIER = "FFBFBFBF"
motor.COULEUR_RATIO_SUSPECT = motor.COULEUR_COMMANDE_SUSPECTE = "FFD5A6BD"
motor.COULEUR_COUSTERON = motor.COULEUR_PRIX_INEXPLIQUE = "FF92D050"
motor.COULEUR_COPIE_POSSIBLE = "FFFF6969"
# Champ obligatoire vide (bloque l'intégration en base) — distinct des 6 couleurs ci-dessus,
# qui parlent toutes de cohérence/changement, jamais de complétude. Les champs optionnels vides
# réutilisent le gris existant (COULEUR_A_VERIFIER), déjà "cellule vide" partout ailleurs.
COULEUR_CHAMP_OBLIGATOIRE_VIDE = "FFFABF8F"

motor.PACKAGE_TYPES = tuple(MOTS_CONTENANT)
# Un ou plusieurs mots qualificatifs (matière, ou...) insérés entre le récipient et sa
# quantité ('CAISSE BOIS 10 KG', 'CAGEOT BOIS OU CARTON 5KG', 'COLIS NUE 2 KG'...) — observé
# dans des fichiers réels, mots variables (bois/carton/nue/poly/verre/ou...), jamais une liste
# fermée à maintenir : on tolère n'importe quelle suite de mots ALPHABÉTIQUES à cet endroit
# (jamais "de" lui-même, pour ne pas interférer avec l'alternative "(?:de\s+)?" déjà
# existante). Le récipient retenu reste le premier mot (groupe 1, ancré en ^) : les mots
# suivants ne sont que des qualificatifs de matière, jamais un second récipient. Un marqueur
# de comptage collé 'x<N>' ('COLIS X960 5 KG' : un colis de 960 pièces, 5 kg au total) est
# toléré ici de la même façon — il décrit le nombre de pièces, une information à part du poids
# total qui suit, pas un second récipient non plus.
_MOT_QUALIFICATIF = r"(?:(?!de\b)(?:[A-Za-zÀ-ÿ]+|x\d+)\s+)*"
motor.RE_TYPE = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+{_MOT_QUALIFICATIF}(?:de\s+)?([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
motor.RE_CONTENANT_MULTIPLIE = re.compile(
    rf"^\s*(\d+)\s*({_MOTS_REGEX})s?\s+{_MOT_QUALIFICATIF}(?:de\s+)?([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
motor.RE_TRIPLE = re.compile(
    rf"^\s*(\d+)\s*x\s*(\d+)\s*x\s*([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
motor.RE_DOUBLE = re.compile(
    rf"^\s*(\d+)\s*x\s*([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)
motor.RE_LIBRE = re.compile(rf"([\d,.]+)\s*({_UNITES_REGEX})\b", re.I)
# Comme RE_LIBRE, mais tolère un multiplicateur 'x' collé juste après l'unité, sans espace
# ('1,2KGx6') — fréquent dans la Désignation fournisseur (texte libre), contrairement au
# Conditionnement, plus structuré, où RE_LIBRE (avec \b) suffit.
RE_QUANTITE_DESIGNATION = re.compile(rf"([\d,.]+)\s*({_UNITES_REGEX})(?![a-wyzà-öø-ÿ])", re.I)
# Mot de récipient trouvé dans le Conditionnement -> mot écrit dans l'Unité de commande. Pour
# un mot déjà dans la liste d'unités acceptées par la base (voir la feuille 'Unités' du
# fichier de référence métier), on garde ce mot tel quel (juste la bonne orthographe/casse).
# Pour un mot identifié comme NON reconnu par la base (catalogue
# unites_commande_non_reconnues.xlsx, colonne "Unité de remplacement", remplie par le métier),
# on écrit directement le remplacement officiel au lieu du mot d'origine — ex: 'COLIS' (non
# reconnu) devient 'CARTON', 'PLAQUE' devient 'PLATEAU', 'PACK' devient 'LOT'. Ne concerne QUE
# l'Unité de commande (seul endroit où ce dictionnaire est utilisé dans ce fichier) — l'Unité
# de référence n'y touche jamais (voir plus haut dans ce fichier : son vocabulaire reste
# volontairement simple, PC/KG/L, jamais un nom de contenant).
motor.CONTENANT_ABREGE = {
    # Déjà acceptés par la base : mot gardé tel quel.
    "sac": "SAC", "piece": "PIECE", "seau": "SEAU",
    "barquette": "BARQUETTE", "barq": "BARQUETTE",
    "bidon": "BIDON", "pot": "POT", "sachet": "SACHET",
    "carton": "CARTON", "poche": "POCHE", "bouteille": "BOUTEILLE", "bouteile": "BOUTEILLE",
    "boite": "BOITE", "boîte": "BOITE",
    "plateau": "PLATEAU", "paquet": "PAQUET", "botte": "BOTTE",
    # Non reconnus par la base : remplacés par l'équivalent officiel
    # (unites_commande_non_reconnues.xlsx).
    "colis": "CARTON", "cageot": "PLATEAU", "caisse": "CARTON", "filet": "SAC",
    "etui": "PAQUET", "étui": "PAQUET", "plaque": "PLATEAU", "pack": "LOT",
    "brique": "UNITÉ", "coffret": "BOITE", "clayette": "PLATEAU", "gourde": "SACHET",
    "bocal": "POT", "outre": "BIDON", "flacon": "BOUTEILLE", "plaquette": "PLATEAU",
    "tube": "UNITÉ", "bombe": "UNITÉ", "bac": "BOITE",
}

# Abréviations déjà utilisées telles quelles dans certains fichiers sources (ex: le fichier
# MDC abrège systématiquement "barquette" en "BQT") — on uniformise vers le mot en entier
# (standardisation voulue), mais SANS marquer la cellule : ce n'est pas une incohérence,
# juste une autre façon d'écrire la même chose.
ABREVIATIONS_CONTENANT_RECONNUES = {"bqt": "barquette"}

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
    # Parenthèse finale purement informative (ex: '(15 TR)' pour 'tranches'), sans aucune
    # quantité reconnue dedans — on la retire : elle casse sinon la détection du motif propre,
    # qui doit se terminer exactement après la quantité (ex: 'POCHE SV DE 1KG (15 TR)'). Si la
    # parenthèse contient elle-même une quantité reconnue (ex: '(2,2L)'), on la laisse — elle
    # porte potentiellement une information utile ailleurs.
    m_parenthese = re.search(r"\s*\(([^()]*)\)\s*$", propre)
    if m_parenthese and not motor.RE_LIBRE.search(m_parenthese.group(1)):
        propre = propre[:m_parenthese.start()]
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

# Même chose que RE_CONTENANT_PUIS_MULT, mais avec 'x' (pas 'de') AVANT le premier nombre, et
# un mot 'ST'/'SACHET(S)' collé juste après ce nombre ('CARTON X 5 ST X 1 KG' : un carton de 5
# sachets de 1 kg chacun, 5 kg au total) — 'st' n'est PAS ajouté au vocabulaire général des
# récipients (trop court, risque de faux positifs ailleurs) : il n'est reconnu qu'à cette
# position précise, entre le nombre de sous-unités et le second 'x'.
RE_CONTENANT_PUIS_MULT_SACHET = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de|x)\s+(\d+)\s*(?:st|sachets?)\.?\s*x\s*([\d,.]+)\s*({_UNITES_REGEX})\s*$",
    re.I)

# Même chose que RE_CONTENANT_PUIS_MULT, mais dans l'ordre inverse — le poids D'UNE pièce
# d'abord, puis le nombre de pièces collé en 'x<N>' ('COLIS 30 G X144' : 144 pièces de 30 g
# chacune, pas 30 g au total) — observé dans les mêmes fichiers que le motif 'CONTENANT x<N>
# POIDS' ci-dessus (voir _MOT_QUALIFICATIF), juste avec le poids et le compte intervertis.
RE_CONTENANT_POIDS_PUIS_X = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de\s+)?([\d,.]+)\s*({_UNITES_REGEX})\s*x\s*(\d+)\s*$", re.I)

# Triple imbrication ("CARTON DE 8 PACKS DE 4 GOURDES DE 90 GR" : un carton de 8 packs,
# chaque pack contenant 4 gourdes de 90g) — la "pièce" vendue est le niveau INTERMÉDIAIRE (le
# pack), pas le carton entier ni la gourde individuelle ; voir piece_carton_candidats().
RE_CONTENANT_TRIPLE_IMBRIQUE = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+(?:de\s+)?(\d+)\s*({_MOTS_REGEX})s?\s+(?:de\s+)?(\d+)\s*"
    rf"({_MOTS_REGEX})s?\s+(?:de\s+)?([\d,.]+)\s*({_UNITES_REGEX})\s*$", re.I)

# Unité de référence elle-même écrite comme "CONTENANT DE N" (ex: 'plateau de 60', 'caisse de
# 180') — observé dans Base IANOR AGAPE 16 (œufs), où ce texte est à la fois l'unité de
# référence ET le conditionnement, sans aucun poids. Contrairement à tous les motifs ci-dessus
# (qui s'appliquent au Conditionnement quand l'unité de référence est DÉJÀ 'PC'), celui-ci
# s'applique à l'unité de référence elle-même : le nombre total (N) doit correspondre au
# ratio Prix commande ÷ Prix référence pour que la conversion en 'PC' (prix par pièce
# individuelle, comparable entre fournisseurs quel que soit le conditionnement) soit confirmée.
RE_CONTENANT_SEUL_SANS_POIDS = re.compile(rf"^\s*({_MOTS_REGEX})s?\s+de\s+(\d+)\s*$", re.I)

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


def extraire_dernier_poids_designation(designation):
    """Comme extraire_quantite_designation_etendue, mais renvoie le DERNIER nombre+unité
    trouvé, pas le premier — utile quand la désignation détaille d'abord un poids PAR PIÈCE
    avant de donner le poids TOTAL du contenant entier (ex: 'GALETTES (100X15GR) 1.5 KG' :
    15GR est le poids d'une galette, 1.5KG est le poids total des 100 — c'est ce dernier qu'il
    faut retenir quand le ratio confirme que le prix est pour le contenant entier, pas pour une
    pièce)."""
    if not designation:
        return None, None
    trouvailles = RE_QUANTITE_DESIGNATION.findall(str(designation))
    if not trouvailles:
        return None, None
    qty, unite = trouvailles[-1]
    return motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], motor.base_depuis_unite(unite)

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

# Mot de comptage final ('pièces', 'tranches', 'portions', 'biscuits'...) : jamais une liste
# fermée à maintenir (le mot varie selon le produit : tranches de pain, biscuits, galettes...),
# donc on accepte N'IMPORTE QUEL mot à cet endroit — SAUF une unité de poids/volume reconnue
# ('kg', 'g'...), qui doit rester traitée par les motifs dédiés au poids (RE_TYPE,
# RE_CONTENANT_MULTIPLIE...), jamais confondue avec un simple compte de pièces. Le 's?' couvre
# le pluriel ('litres', 'kilos'...) : sans lui, 'litres' passait le négatif (le \b échoue entre
# 'litre' et le 's' collé) et 'SEAU DE 5 LITRES' était pris à tort pour '5 pièces'.
_MOT_COMPTAGE_FINAL = rf"(?!(?:{_UNITES_REGEX})s?\b)[A-Za-zÀ-ÿ]+"

RE_CONTENANT_NX_M_PIECES = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+{_MOT_QUALIFICATIF}(?:de\s+)?(\d+)\s*x\s*(\d+)\s*{_MOT_COMPTAGE_FINAL}\s*$", re.I)

# "COLIS DE 12 BARQUETTES", "COLIS DE 12 BOTTES" : récipient extérieur contenant N récipients
# nommés (pas littéralement "pièces"), sans aucun poids donné. Le récipient intérieur EST la
# pièce vendue (comme 'sachet' dans 'CARTON 14 SACHETS DE 350 GR'), le récipient extérieur est
# ce qui se commande.
RE_CONTENANT_PUIS_CONTENANT_SEUL = re.compile(
    rf"^\s*({_MOTS_REGEX})s?\s+{_MOT_QUALIFICATIF}(?:de\s+)?(\d+)\s*({_MOTS_REGEX})s?\s*$", re.I)

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
        # de ses M gourdes de qty chacune. mot_ext (carton) et m_inner (nombre de gourdes par
        # pack) sont renvoyés en plus pour permettre, UNIQUEMENT dans ce cas précis confirmé
        # par le ratio, de préciser l'unité de référence/commande au lieu de 'PC' générique
        # (voir useful_stuff/Exemples unités et prix 6.xlsx, "CAS 3") — jamais pour les autres
        # motifs ci-dessous, où 'PC' générique reste correct (un seul niveau de contenant).
        mot_ext, n, mot_piece, m_inner, _mot_inner, qty, unite = m.groups()
        poids_piece = round(int(m_inner) * motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        return ([int(n)], motor.normalise_texte(mot_piece), poids_piece, motor.base_depuis_unite(unite),
                motor.normalise_texte(mot_ext), int(m_inner))
    m = RE_CONTENANT_IMBRIQUE.match(conditionnement)
    if m:
        mot_ext, n, mot_piece, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        # '1' est ajouté comme candidat valable (contenant entier vendu en un seul morceau,
        # ex: ratio = 1 malgré un conditionnement qui détaille N pièces de tel poids) — même
        # logique que RE_CONTENANT_PIECES_SEULEMENT plus bas. Le récipient EXTÉRIEUR (mot_ext,
        # ex: 'carton') n'est renvoyé en 5e position (comme pour RE_CONTENANT_TRIPLE_IMBRIQUE)
        # QUE si le récipient intérieur est RÉEL (ex: 'bouteille' dans 'CARTON 6 BOUTEILLES DE
        # 75 CL') : c'est alors lui l'unité de référence et le récipient extérieur l'unité de
        # commande (même logique CAS-3 que le triple niveau). Quand l'intérieur n'est que
        # 'piece'/'pièce' (compte générique, ex: 'CARTON 12 PIECES DE 500 GR'), mot_ext_imbrique
        # reste None comme avant : sinon la détection "vendu au poids total, pas à la pièce"
        # plus haut (qui exige justement 'not mot_ext_imbrique') ne s'appliquerait plus.
        mot_piece_norm = motor.normalise_texte(mot_piece)
        mot_ext_imbrique = motor.normalise_texte(mot_ext) if mot_piece_norm not in ("piece", "pièce") else None
        return ([1, int(n)], mot_piece_norm, poids_piece, motor.base_depuis_unite(unite),
                mot_ext_imbrique, None)
    m = RE_CONTENANT_PUIS_MULT.match(conditionnement)
    if m:
        mot_ext, n, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        # Le récipient nommé ('CARTON DE 8 X 250 GR') garde son nom en unité de commande —
        # "LOT DE N" est réservé aux formats SANS aucun nom de récipient ('100 x 10g').
        return [1, int(n)], motor.normalise_texte(mot_ext), poids_piece, motor.base_depuis_unite(unite), None, None
    m = RE_CONTENANT_POIDS_PUIS_X.match(conditionnement)
    if m:
        mot_ext, qty, unite, n = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        return [1, int(n)], motor.normalise_texte(mot_ext), poids_piece, motor.base_depuis_unite(unite), None, None
    m = motor.RE_CONTENANT_MULTIPLIE.match(conditionnement)
    if m:
        n, mot_piece, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        return [int(n)], motor.normalise_texte(mot_piece), poids_piece, motor.base_depuis_unite(unite), None, None
    m = RE_CONTENANT_NX_M_PIECES.match(conditionnement)
    if m:
        mot_ext, n1, n2 = m.group(1), int(m.group(2)), int(m.group(3))
        return [n1, n2, n1 * n2], motor.normalise_texte(mot_ext), None, None, None, None
    # Récipient extérieur contenant N récipients intérieurs NOMMÉS ('CARTON DE 12 PLAQUES',
    # 'COLIS DE 12 BARQUETTES') : vérifié AVANT RE_CONTENANT_PIECES_SEULEMENT ci-dessous,
    # devenu généraliste (n'importe quel mot de comptage final) — sinon ce motif plus
    # spécifique (récipient intérieur réellement nommé) ne serait jamais atteint.
    m = RE_CONTENANT_PUIS_CONTENANT_SEUL.match(conditionnement)
    if (m and motor.normalise_texte(m.group(1)) != motor.normalise_texte(m.group(3))
            and motor.normalise_texte(m.group(3)) not in ("piece", "pièce")):
        mot_ext, n, mot_int = m.group(1), int(m.group(2)), m.group(3)
        # Même logique que RE_CONTENANT_PIECES_SEULEMENT (1 = contenant entier vendu en un
        # seul morceau, n = pièces séparées), mais avec un récipient intérieur nommé (pas
        # juste "pièces") : il devient l'unité de référence si confirmé, le récipient
        # extérieur reste l'unité de commande (imbrication à UN SEUL niveau, pas de poids
        # connu). 'n' (pièces intérieures par récipient extérieur, ex: 12 plaques par carton)
        # N'EST PAS un nombre de sous-pièces internes au récipient intérieur lui-même (ici,
        # une plaque n'est composée d'aucune sous-pièce) — le dernier élément du tuple reste
        # donc None, pas 'n' : le réutiliser ferait croire à tort à un troisième niveau
        # d'imbrication (nb_pieces_internes), ce qui ferait choisir 'PAQUET' au lieu de 'PC'
        # pour l'unité de référence, et remplirait la colonne F avec une valeur qui ne décrit
        # pas ce qu'elle est censée décrire (pièces par UNITÉ DE RÉFÉRENCE, pas par récipient
        # extérieur).
        return ([1, n], motor.normalise_texte(mot_int), None, None,
                motor.normalise_texte(mot_ext), None)
    # Récipient nommé contenant SEULEMENT un nombre de pièces, sans poids ('CAGEOT DE 30
    # PIECES', 'SAC DE 100 BISCUITS') : le poids réel est inconnu (souvent variable), mais le
    # nombre de pièces suffit à vérifier le ratio — pas besoin de deviner un poids via la
    # désignation pour expliquer le prix (ce qui donnerait un résultat en €/kg non pertinent
    # ici).
    m = RE_CONTENANT_PIECES_SEULEMENT.match(conditionnement)
    if m:
        mot_ext, n = m.group(1), int(m.group(2))
        # Deux lectures possibles, selon ce que confirme le ratio : soit N exemplaires séparés
        # (ratio = N, ex: 'CAGEOT DE 30 PIECES' d'œufs vendus à l'unité), soit UN SEUL
        # contenant divisé en N parts (ratio = 1, ex: 'PLAQUE DE 30 PIECES' = un gâteau coupé
        # en 30 parts, vendu entier) — le 1 est ajouté comme candidat valable lui aussi.
        return [1, n], motor.normalise_texte(mot_ext), None, None, None, None
    m = RE_LOT_PIECES_SEULEMENT.match(conditionnement)
    if m:
        # "LOT DE 4 PIECES" : aucun récipient nommé, juste un comptage — le marqueur 'lot'
        # n'est PAS un contenant (volontairement absent de CONTENANT_ABREGE) : les chemins qui
        # décident de l'unité de commande doivent le traiter comme équivalent à "pas de mot_piece
        # du tout", pour que 'LOT DE N' soit proposé normalement.
        return [1, int(m.group(1))], "lot", None, None, None, None
    # Multiplicateur SANS récipient nommé ('100 x 10g', '4 x 125g') : quand l'unité de
    # référence est déjà à la pièce, c'est N qui doit correspondre au ratio (comparaison par
    # pièce), PAS un total en kg/L à recalculer — contrairement à determine_base_et_quantite(),
    # qui suppose au contraire que l'unité de référence devrait être en kg/L.
    m = motor.RE_TRIPLE.match(conditionnement)
    if m:
        n1, n2, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        n1, n2 = int(n1), int(n2)
        return [n1, n2, n1 * n2], None, poids_piece, motor.base_depuis_unite(unite), None, None
    m = motor.RE_DOUBLE.match(conditionnement)
    if m:
        n, qty, unite = m.groups()
        poids_piece = round(motor._nombre(qty) * motor.FACTEUR_VERS_KG_OU_L[unite.lower()], 3)
        return [int(n)], None, poids_piece, motor.base_depuis_unite(unite), None, None
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
    m = RE_CONTENANT_PUIS_MULT_SACHET.match(conditionnement)
    if m:
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
    "carton", "poche", "cageot", "boite", "boîte", "filet", "bouteille", "caisse", "plaque",
    "plateau", "clayette", "bac", "coffret", "plaquette", "paquet", "etui", "étui", "bombe",
    "bocal", "flacon", "tube",
    "piece",
]
_MOTS_CONTENANT_CONNUS = set(MOTS_CONTENANT) | {"piece", "pièce"}


def mot_contenant_dans_conditionnement(conditionnement):
    """Cherche, par limite de mot (pas une sous-chaîne), le premier mot de récipient connu
    explicitement nommé dans le Conditionnement — ex: 'COLIS DE 60 PIECES' -> 'colis' (le mot
    le plus à gauche dans le texte l'emporte quand plusieurs sont présents, le reste n'étant
    souvent qu'un qualificatif de matière). 'piece'/'pièce' est volontairement exclu : dans '60
    PIECES', ce n'est pas un récipient mais un simple compte. Renvoie None si aucun mot net ne
    ressort."""
    if not conditionnement:
        return None
    texte = motor.normalise_texte(conditionnement)
    meilleure_position, meilleur_mot = None, None
    for mot in _MOTS_CONTENANT_CANONIQUES:
        if mot == "piece":
            continue
        m = re.search(rf"\b{re.escape(mot)}s?\b", texte)
        if m and (meilleure_position is None or m.start() < meilleure_position):
            meilleure_position, meilleur_mot = m.start(), mot
    return meilleur_mot


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


RE_FORMAT_BOITE = re.compile(r"\bboi?tes?\s*(\d+)\s*/\s*(\d+)\b", re.I)


def reconsidere_format_boite(resultat, conditionnement, unite_cmd):
    """Les formats de boîte de conserve normalisés ('BOITE 5/1', 'BOITE 3/1', 'BOITE 4/4'...)
    sont des unités à part entière reconnues par la base (feuille 'Unités' du fichier de
    référence métier), distinctes d'une 'BOITE' générique. Quand le conditionnement précise ce
    format mais que l'unité de commande (actuelle, ou déjà proposée par les règles
    précédentes) ne retient que 'BOITE' nue, on récupère le format."""
    if not conditionnement:
        return
    m = RE_FORMAT_BOITE.search(conditionnement)
    if not m:
        return
    valeur_actuelle = resultat.nouveau_l if resultat.nouveau_l is not None else unite_cmd
    if not valeur_actuelle or motor.normalise_texte(valeur_actuelle) != "boite":
        return  # déjà précisé, ou pas une 'BOITE' du tout : rien à faire
    format_boite = f"{m.group(1)}/{m.group(2)}"
    resultat.nouveau_l = f"BOITE {format_boite}"
    resultat.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
    resultat.commentaire_l = (
        f"Unité de commande précisée en 'BOITE {format_boite}' : le conditionnement "
        f"('{conditionnement}') indique ce format normalisé de boîte, une unité à part "
        "entière distincte de 'BOITE' générique.")


def reconsidere_poids_piece_existant(resultat, poids_ref, unite_ref, conditionnement,
                                      designation, desig_fourn):
    """'Poids à l'unité de référence' (E) n'est rempli AUTOMATIQUEMENT que quand il est vide —
    s'il contient déjà une valeur, le moteur partagé part du principe qu'elle est bonne et ne
    la recalcule ni ne la vérifie JAMAIS (motor.Resultat.nouveau_e = 'GARDER' dans ce cas). Si
    le conditionnement (ou, à défaut, la désignation) permet de déduire indépendamment un poids
    par pièce différent, on le signale — jamais modifié automatiquement, comme pour tout
    prix/poids : une personne valide toujours elle-même."""
    if not poids_ref:
        return
    d_final = resultat.nouveau_d if resultat.nouveau_d is not None else unite_ref
    base_d = motor.normalise_texte(d_final) if d_final else None
    if not base_d or base_d in ("kg", "l", "litre"):
        return  # ce poids ne concerne que les unités comptées (PC ou un récipient nommé),
        # jamais une unité déjà au poids/volume
    try:
        poids_ref_num = float(str(poids_ref).replace(",", "."))
    except (TypeError, ValueError):
        return
    if poids_ref_num <= 0:
        return

    poids_detecte, source = None, None
    if base_d == "pc":
        resultat_pieces = piece_carton_candidats(unite_ref, conditionnement)
        if resultat_pieces and resultat_pieces[2] and resultat_pieces[3] == "kg":
            poids_detecte = resultat_pieces[2]
            source = f"le conditionnement ('{conditionnement}')"
    if poids_detecte is None:
        # Format simple 'PIECE DE 2KG' ou 'BARQUETTE DE 300G' (un seul niveau, pas de
        # décompte) : passe par determine_base_et_quantite(), pas piece_carton_candidats()
        # (qui ne reconnaît pas ce motif sans compte, et exige de toute façon que l'unité de
        # référence soit déjà 'PC'). Le mot retenu doit correspondre à l'unité de référence
        # elle-même — 'piece'/'pièce' quand D='PC', ou le même récipient que D (ex:
        # 'barquette' quand D='Barquette') — sinon la quantité trouvée est le poids du
        # contenant ENTIER (ex: un autre récipient imbriqué), pas celui d'UNE unité de
        # référence, et ce serait une mauvaise comparaison.
        info = determine_base_et_quantite(nettoie_conditionnement(conditionnement))
        if info and info[0] == "kg" and info[2]:
            mot_norm = motor.normalise_texte(info[2])
            # Si D='PC' vient d'une unité de référence d'origine qui nommait elle-même un
            # récipient (ex: 'Barquette' collapsé en 'PC' par le ratio≈1 du moteur partagé —
            # voir first_attempt_v2_prototype.py), le mot d'origine reste valide pour la
            # comparaison, pas seulement 'piece'/'pièce' générique.
            mots_valides = ({"piece", "pièce", motor.normalise_texte(unite_ref)} if base_d == "pc"
                             else {base_d})
            if mot_norm in mots_valides:
                poids_detecte = info[1]
                source = f"le conditionnement ('{conditionnement}')"
    if poids_detecte is None and base_d == "pc":
        for texte, nom_source in ((designation, "la désignation de l'article"),
                                   (desig_fourn, "la désignation fournisseur")):
            if not texte or RE_QUANTITE_APPROXIMATIVE.search(texte) or re.search(r"calibr", texte, re.I):
                # Poids calibré/variable (ex: 'calibrés 53/63g') : plusieurs valeurs valables
                # selon le calibre réellement livré, pas UN poids précis à comparer — le
                # premier nombre trouvé par le regex ('63g') ne serait qu'une des bornes,
                # pas forcément "la bonne", une fausse alerte de précision.
                continue
            poids, base = extraire_quantite_designation_etendue(texte)
            if poids and base == "kg":
                poids_detecte = poids
                source = f"{nom_source} ('{texte}')"
                break
    if poids_detecte is None or poids_detecte <= 0:
        return
    if abs(poids_ref_num - poids_detecte) <= motor.TOLERANCE_RATIO * poids_detecte:
        return  # cohérent, rien à signaler

    resultat.couleur_e = motor.COULEUR_RATIO_SUSPECT
    resultat.commentaire_e = (
        f"Il y a déjà un poids écrit ici ({poids_ref_num} kg), mais {source} donne plutôt "
        f"à penser que le poids d'une pièce est de {round(poids_detecte, 4)} kg — à vérifier, "
        "rien n'est changé automatiquement.")


RE_N_PIECES_UNITE_CMD = re.compile(r"^(\d+)\s*pi[eè]ces?$", re.I)


def reconsidere_lot_pieces_brutes(resultat):
    """Quand l'unité de commande finalement proposée n'est qu'un décompte brut recopié du
    conditionnement (ex: '2 PIECES', sans aucun récipient nommé — cf. '2 PIECES DE 2 KG') : ce
    n'est pas une vraie unité, juste le texte du conditionnement répété. On la fait passer par
    le même traitement que 'LOT DE N' (voir RE_LOT_DE_N_UNITE_CMD dans audite_fournisseur) en
    la réécrivant dans ce format, pour que l'unité de commande finale devienne 'LOT' et que le
    nombre reste visible dans le Conditionnement."""
    valeur = resultat.nouveau_l
    if not isinstance(valeur, str):
        return
    m = RE_N_PIECES_UNITE_CMD.match(valeur.strip())
    if not m:
        return
    resultat.nouveau_l = f"LOT DE {m.group(1)}"
    resultat.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
    resultat.commentaire_l = (
        f"Unité de commande '{valeur}' n'est qu'un décompte recopié du conditionnement, pas "
        "une vraie unité — regroupée sous 'LOT' (le nombre reste visible dans le "
        "Conditionnement).")


def traite_ligne_etendue(unite_ref, poids_ref, designation, conditionnement, unite_cmd,
                          prix_ref, prix_cmd, desig_fourn=None):
    """Délègue à _traite_ligne_etendue_nucleo. La correction d'une éventuelle faute de frappe
    sur un mot de contenant est faite en amont, directement sur la cellule Conditionnement
    (voir audite_fournisseur) — ce que cette fonction reçoit est donc déjà la version
    corrigée."""
    resultat = _traite_ligne_etendue_nucleo(unite_ref, poids_ref, designation, conditionnement,
                                             unite_cmd, prix_ref, prix_cmd, desig_fourn)
    reconsidere_format_boite(resultat, conditionnement, unite_cmd)
    reconsidere_lot_pieces_brutes(resultat)
    reconsidere_poids_piece_existant(resultat, poids_ref, unite_ref, conditionnement,
                                      designation, desig_fourn)

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

    # Unité de référence écrite directement comme "CONTENANT DE N" (ex: 'plateau de 60'),
    # sans aucun poids — seul cas où l'unité de référence elle-même (pas le conditionnement)
    # porte ce motif. Converti en 'PC' avec un prix de référence recalculé (par pièce
    # individuelle), UNIQUEMENT si le prix de référence n'est pas déjà cohérent avec ce calcul
    # — pour permettre la comparaison du prix à l'unité entre fournisseurs quel que soit le
    # conditionnement (plateau de 60, caisse de 180...). Court-circuite le reste de la
    # fonction : ce motif est trop spécifique pour interagir avec piece_carton_candidats, qui
    # suppose que l'unité de référence est déjà 'PC'.
    m_unite_seule = RE_CONTENANT_SEUL_SANS_POIDS.match(unite_ref) if unite_ref else None
    if m_unite_seule and prix_ref and prix_cmd:
        mot_contenant, n_texte = m_unite_seule.groups()
        n = int(n_texte)
        mot_contenant = motor.normalise_texte(mot_contenant)
        prix_par_piece = round(prix_cmd / n, 4)
        r = motor.Resultat()
        r.nouveau_d = "PC"
        r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
        r.commentaire_d = (
            f"Unité de référence précisée en 'PC' (prix par pièce individuelle, comparable "
            f"entre fournisseurs quel que soit le conditionnement) : l'unité de référence "
            f"d'origine ('{unite_ref}') décrivait en fait le contenant de commande, pas une "
            "unité de mesure.")
        r.nouveau_l = mot_contenant.upper()
        r.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
        r.commentaire_l = (
            f"Unité de commande précisée en '{mot_contenant.upper()}' (contenant de commande, "
            "sans le décompte qui est maintenant implicite dans le prix par pièce).")
        r.nouveau_f = n
        r.commentaire_f = f"Renseigné à {n} : nombre de pièces dans ce '{mot_contenant}'."
        r.contenant_imbrique_confirme = True  # pour valider l'unité de commande CORRIGÉE, pas l'ancienne
        if round(prix_ref, 4) != prix_par_piece:
            r.couleur_k = motor.COULEUR_RATIO_SUSPECT
            r.nouveau_k = prix_par_piece
            r.commentaire_k = (
                f"Prix de référence par pièce suggéré : {prix_cmd} ÷ {n} = {prix_par_piece} "
                f"(prix de référence actuel : {prix_ref}, qui ne semble pas divisé par le "
                "nombre de pièces) — à vérifier, non modifié automatiquement.")
        return r

    resultat_pieces = piece_carton_candidats(unite_ref, conditionnement)
    if resultat_pieces and prix_ref and prix_cmd:
        candidats_pieces, mot_piece, poids_piece, base_unite, mot_ext_imbrique, nb_pieces_internes = resultat_pieces
        ratio = prix_cmd / prix_ref
        r = motor.Resultat()
        total_pieces = candidats_pieces[-1]
        candidat_ok = next((c for c in candidats_pieces if abs(ratio - c) <= motor.TOLERANCE_RATIO * c), None)

        # Le ratio ne confirme aucun nombre de pièces, mais le poids d'une pièce est connu —
        # teste si le ratio correspond plutôt au POIDS TOTAL (total_pieces × poids d'une
        # pièce) : le produit est alors vendu/comparé au poids, pas à la pièce (ex: miel en
        # carton de 12 pots de 500g, prix déjà au kilo) — le prix de référence n'est PAS
        # recalculé (même nombre, juste réinterprété en kg/L), seule l'unité change.
        if candidat_ok is None and poids_piece and not mot_ext_imbrique:
            poids_total = round(total_pieces * poids_piece, 6)
            if poids_total > 0 and abs(ratio - poids_total) <= motor.TOLERANCE_RATIO * poids_total:
                r2 = motor.Resultat()
                r2.nouveau_d = base_unite.upper()
                r2.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
                r2.commentaire_d = (
                    f"Unité de référence changée de '{unite_ref}' vers '{base_unite.upper()}' : "
                    f"le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                    f"{round(ratio, 3)}) ne correspond à aucun nombre de pièces, mais correspond "
                    f"exactement au poids total du conditionnement ('{conditionnement}') : "
                    f"{total_pieces} × {poids_piece} = {poids_total} {base_unite}. Le prix de "
                    f"référence ({prix_ref}) n'est pas recalculé, juste réinterprété au "
                    f"{base_unite}.")
                return r2

        if (mot_ext_imbrique and mot_piece not in ("piece", "pièce")
                and candidat_ok is not None and candidat_ok == total_pieces):
            # Contenant à double imbrication confirmé par le ratio des prix (ex: "CARTON DE 8
            # PACKS DE 4 GOURDES DE 90 GR", 8 packs/carton confirmé ; ou "CARTON 6 BOUTEILLES DE
            # 75 CL", 6 bouteilles confirmé) — règle métier documentée dans
            # useful_stuff/Exemples unités et prix 6.xlsx ("CAS 3"). Exclu quand le récipient
            # intérieur n'est que 'piece'/'pièce' (compte générique, pas un vrai récipient, ex:
            # 'CARTON 6 PIECES DE 300 GR') : D doit rester 'PC' générique, seul le récipient
            # extérieur (mot_ext_imbrique) reste pertinent pour L, géré dans le 'else' ci-dessous.
            #
            # Le nom précis du contenant intermédiaire (pack, bouteille...) n'est PAS repris tel
            # quel en D : le métier recommande de garder l'unité de référence simple (PC/KG/L —
            # voir la note dans la feuille 'Unités' du fichier de référence). Deux cas :
            #   - le contenant intermédiaire n'a AUCUNE subdivision interne (nb_pieces_internes
            #     vide, ex: une bouteille) : il EST la pièce atomique -> 'PC' suffit.
            #   - le contenant intermédiaire regroupe lui-même N pièces (nb_pieces_internes
            #     rempli, ex: un pack de 4 gourdes) : 'PC' serait trompeur (une seule pièce qui
            #     en contient plusieurs) -> 'PAQUET' (terme générique, accepté par la base)
            #     décrit mieux ce niveau, le nombre de pièces restant précisé à part.
            mot_ext_novo = motor.CONTENANT_ABREGE.get(mot_ext_imbrique, mot_ext_imbrique.upper())
            if nb_pieces_internes:
                r.nouveau_d = "PAQUET"
                r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
                r.commentaire_d = (
                    f"Unité de référence précisée en 'PAQUET' : ce niveau ('{mot_piece}') n'est "
                    f"pas une pièce atomique, il regroupe lui-même {nb_pieces_internes} pièces "
                    f"individuelles, d'après le conditionnement ('{conditionnement}') — le ratio "
                    f"Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                    f"{round(ratio, 3)}) confirme {total_pieces} {mot_piece}s par "
                    f"{mot_ext_imbrique} au total.")
                r.nouveau_f = nb_pieces_internes
                r.commentaire_f = (
                    f"Renseigné à {nb_pieces_internes} : nombre d'éléments individuels par "
                    f"'{mot_piece}', d'après le conditionnement ('{conditionnement}').")
            else:
                r.nouveau_d = "PC"
                if motor.normalise_texte(unite_ref) not in ("pc", "piece", "pièce", "unite", "unité"):
                    r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
                    r.commentaire_d = f"Unité de référence normalisée de '{unite_ref}' vers 'PC'."
            r.nouveau_l = mot_ext_novo
            r.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
            r.commentaire_l = (
                f"Unité de commande précisée en '{mot_ext_novo}' pour correspondre "
                f"au contenant extérieur décrit dans le conditionnement ('{conditionnement}').")
            r.contenant_imbrique_confirme = True
        else:
            # L'unité de référence est bien à la pièce ici (pas un poids à recalculer en
            # kg/L) — toujours normaliser en 'PC', silencieusement si c'était déjà une variante
            # de 'pc'/'piece'/'pièce'/'unite'/'unité' (simple nomenclature, pas un changement
            # d'unité réel).
            r.nouveau_d = "PC"
            if motor.normalise_texte(unite_ref) not in ("pc", "piece", "pièce", "unite", "unité"):
                r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
                r.commentaire_d = f"Unité de référence normalisée de '{unite_ref}' vers 'PC'."
            if mot_ext_imbrique:
                # Le ratio ne confirme pas le nombre de pièces intérieures, mais le récipient
                # EXTÉRIEUR décrit dans le conditionnement reste fiable (texte explicite) —
                # on le reprend quand même en unité de commande, que candidat_ok soit 1 (tout
                # le contenant vendu en un seul morceau) ou ne corresponde à rien.
                l_novo = motor.CONTENANT_ABREGE.get(mot_ext_imbrique, mot_ext_imbrique.upper())
                if motor.normalise_texte(l_novo) != motor.normalise_texte(unite_cmd):
                    r.nouveau_l = l_novo
                    r.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
                    r.commentaire_l = (
                        f"Unité de commande changée en '{l_novo}' d'après le conditionnement "
                        f"('{conditionnement}').")

        if (candidat_ok is not None and candidat_ok == total_pieces
                and mot_piece in (None, "lot")
                and motor.normalise_texte(unite_cmd) != f"lot de {total_pieces}"):
            # Aucun récipient PHYSIQUE nommé dans le conditionnement — soit aucun mot du tout
            # ('100 x 10g', '4 x 125g'), soit le marqueur 'lot' ('LOT DE 4 PIECES', qui ne
            # désigne pas un récipient mais directement un lot) — le ratio confirme bien un
            # LOT de N pièces, pas un poids total à recalculer.
            r.nouveau_l = f"LOT DE {total_pieces}"
            r.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
            r.commentaire_l = (
                f"Unité de commande précisée en 'LOT DE {total_pieces}' : le ratio Prix "
                f"commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = {round(ratio, 3)}) "
                f"correspond au conditionnement ('{conditionnement}'), sans récipient nommé.")
        elif (mot_piece and mot_piece not in ("lot", "piece", "pièce") and not mot_ext_imbrique):
            # Récipient nommé dans le conditionnement ('COLIS DE 15 PIECES' -> 'colis') : on le
            # reprend dans l'unité de commande, QUE LE RATIO confirme ou non le nombre de
            # pièces — le nom du récipient vient directement du texte (fiable), indépendamment
            # de la cohérence du prix (qui reste signalée séparément si besoin). 'piece'/'pièce'
            # est exclu : ce n'est pas un récipient réel (juste un compte générique, ex: 'CARTON
            # 6 PIECES DE 300 GR' où RE_CONTENANT_IMBRIQUE renvoie 'piece' comme mot_piece faute
            # de récipient intérieur nommé) — l'imposer en L écraserait à tort un vrai récipient
            # déjà correct ('CARTON').
            l_novo = motor.CONTENANT_ABREGE.get(mot_piece, mot_piece.upper())
            if motor.normalise_texte(l_novo) != motor.normalise_texte(unite_cmd):
                r.nouveau_l = l_novo
                r.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
                r.commentaire_l = (
                    f"Unité de commande changée en '{l_novo}' d'après le conditionnement "
                    f"('{conditionnement}').")

        if poids_piece and not poids_ref:
            # Le poids d'UNE pièce est écrit explicitement dans le conditionnement (ex: 'DE
            # 500 GR'), indépendamment du nombre total de pièces — ce n'est pas une déduction
            # à partir du ratio des prix, donc on la renseigne même quand le ratio ne confirme
            # PAS le nombre total de pièces (ça reste une incohérence de prix/quantité à
            # signaler séparément, mais le poids d'une pièce, lui, n'est pas en cause).
            r.nouveau_e = poids_piece
            unite_confirmation = f"{mot_piece}s" if mot_ext_imbrique else "pièces"
            confirmation = (
                f"Confirmé par le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ "
                f"{prix_ref} = {round(ratio, 3)}) qui correspond bien à {total_pieces} "
                f"{unite_confirmation}."
                if candidat_ok == total_pieces else
                "Le nombre total de pièces n'est en revanche pas confirmé par le ratio des "
                "prix — voir le signalement correspondant.")
            if mot_ext_imbrique:
                unite_designee = mot_piece
                precision_quantite = f"(sur {total_pieces} {mot_piece}s par {mot_ext_imbrique})"
            else:
                unite_designee = "pièce"
                precision_quantite = f"(sur {total_pieces} au total dans ce {mot_piece or 'conditionnement'})"
            r.explication_e = (
                f"Poids à l'unité de référence renseigné à partir du conditionnement "
                f"('{conditionnement}') : {poids_piece} {base_unite.upper()} par {unite_designee} "
                f"{precision_quantite}. {confirmation}")
        elif (mot_piece and candidat_ok == total_pieces and not poids_ref):
            # Récipient nommé ('COLIS DE 32 PIECES', 'COFFRET DE 100 PIECES'...) ou marqueur
            # 'lot' ('LOT DE N PIECES') ne donnant aucun poids — mais une fois le nombre de
            # pièces confirmé par le ratio (donc la pièce elle-même bien identifiée), le poids
            # unitaire se lit souvent dans la Désignation (fournisseur ou article), ex:
            # '(125Gx4)x12', '(100Gx4x8)' -> 0.125/0.100 kg par pièce (seul le PREMIER poids
            # trouvé compte : c'est celui d'une pièce, pas le poids total du contenant).
            poids_trouve, base_trouvee = motor.extraire_quantite_designation(desig_fourn or "")
            if not poids_trouve:
                poids_trouve, base_trouvee = motor.extraire_quantite_designation(designation or "")
            if poids_trouve and base_trouvee == "kg":
                r.nouveau_e = round(poids_trouve, 3)
                r.explication_e = (
                    f"Poids à l'unité de référence renseigné à {round(poids_trouve, 3)} kg "
                    f"d'après la désignation, confirmé par le ratio Prix commande ÷ Prix "
                    f"référence ({prix_cmd} ÷ {prix_ref} = {round(ratio, 3)}) qui correspond "
                    f"bien à {total_pieces} pièces.")

        if candidat_ok is None:
            r.couleur_m = motor.COULEUR_RATIO_SUSPECT
            r.commentaire_m = (
                f"Le conditionnement ('{conditionnement}') indique {candidats_pieces} pièce(s) "
                f"possibles, et l'unité de référence est déjà à la pièce ('{unite_ref}') — le "
                f"ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                f"{round(ratio, 3)}) ne correspond à aucun d'eux. À vérifier manuellement.")
        elif candidat_ok == 1 and candidat_ok != total_pieces:
            # Le ratio confirme que c'est le CONTENANT ENTIER (pas chacune de ses parts) qui
            # est vendu (ex: 'PLAQUE DE 30 PIECES' = un gâteau entier coupé en 30 parts, 'COLIS
            # 30 G X100' = un colis de 100 pièces de 30 g). Si le poids d'UNE pièce est déjà
            # connu depuis le conditionnement lui-même (poids_piece, ex: les '30 G' de 'COLIS 30
            # G X100'), le poids total du contenant entier se déduit directement (poids_piece ×
            # total_pieces) — bien plus fiable qu'une recherche en texte libre, qui confondrait
            # ce poids PAR pièce avec le poids total recherché ici. On ne cherche dans la
            # désignation qu'à défaut (ex: 'PLAQUE DE 30 PIECES', qui ne donne aucun poids).
            if poids_piece:
                poids_trouve, base_trouvee = round(poids_piece * total_pieces, 3), base_unite
            else:
                # On prend le DERNIER nombre trouvé, pas le premier : la désignation détaille
                # parfois d'abord un poids PAR PIÈCE avant de donner le poids total (ex:
                # '(100X15GR) 1.5 KG' — 15GR est le poids d'une galette, 1.5KG le poids des 100,
                # celui qu'il faut ici puisque le prix est pour le tout).
                poids_trouve, base_trouvee = extraire_dernier_poids_designation(desig_fourn or "")
                if not (poids_trouve and base_trouvee == "kg"):
                    poids_trouve, base_trouvee = extraire_dernier_poids_designation(designation or "")
            if poids_trouve and base_trouvee == "kg" and not poids_ref:
                r.nouveau_e = round(poids_trouve, 3)
                r.explication_e = (
                    f"Poids à l'unité de référence renseigné à {round(poids_trouve, 3)} kg : le "
                    f"conditionnement ('{conditionnement}') décrit {total_pieces} parts, mais "
                    f"le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                    f"{round(ratio, 3)}) confirme que le prix est pour le contenant ENTIER, pas "
                    "pour une part individuelle — poids trouvé dans la désignation.")
        elif candidat_ok != total_pieces:
            # Le ratio confirme un SOUS-groupe intermédiaire (ex: 'CARTON DE 12 X 36
            # TRANCHES' : ratio = 12, le carton contient 12 paquets de 36 tranches chacun — le
            # prix est pour UN paquet, pas pour une tranche ni pour le carton entier) : on
            # cherche le poids de CE sous-groupe dans la désignation (premier nombre trouvé,
            # comme pour le cas 'pièce individuelle' ci-dessus — pas de poids total en jeu ici)
            # avant de se contenter de signaler l'incohérence.
            poids_trouve, base_trouvee = motor.extraire_quantite_designation(desig_fourn or "")
            if not (poids_trouve and base_trouvee == "kg"):
                poids_trouve, base_trouvee = motor.extraire_quantite_designation(designation or "")
            if poids_trouve and base_trouvee == "kg" and not poids_ref:
                r.nouveau_e = round(poids_trouve, 3)
                r.explication_e = (
                    f"Poids à l'unité de référence renseigné à {round(poids_trouve, 3)} kg : le "
                    f"ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                    f"{round(ratio, 3)}) correspond à {candidat_ok}, un sous-groupe du "
                    f"conditionnement ('{conditionnement}') — poids trouvé dans la désignation "
                    "pour ce sous-groupe.")
            else:
                r.couleur_d = motor.COULEUR_RATIO_SUSPECT
                r.commentaire_d = (
                    f"L'unité de référence est déclarée à la pièce ('{unite_ref}'), mais le "
                    f"ratio observé ({prix_cmd} ÷ {prix_ref} = {round(ratio, 3)}) correspond à "
                    f"{candidat_ok}, un SOUS-groupe du conditionnement ('{conditionnement}'), "
                    f"pas au nombre total de pièces ({total_pieces}). Le prix de référence "
                    f"semble être au prix du sous-lot de {candidat_ok}, pas à la pièce "
                    "individuelle — à vérifier.")
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
                # Alignée sur motor.TOLERANCE_RATIO (2%, la tolérance standard du projet) —
                # c'était à tort 5% ici (0.05 * quantite), ce qui laissait passer en silence des
                # écarts entre 2% et 5% (ex: ratio=5 vs quantité réelle=4.8, 4% d'écart). Le
                # plancher absolu (0.05 kg/L) reste séparé : il évite de signaler un simple bruit
                # d'arrondi sur de très petites quantités, quel que soit le pourcentage retenu.
                tolerance_qte = max(0.05, motor.TOLERANCE_RATIO * quantite)

                # Ratio ≈ 1 (prix de commande = prix de référence) avec une unité de référence
                # à la pièce : confirmé à plusieurs reprises sur des fichiers réels (ex: 'PIECE
                # DE 2KG', 'BARQUETTE DE 300G' dont le prix de référence et le prix de commande
                # sont identiques) — le produit est vendu et payé à la pièce/au récipient entier,
                # le conditionnement ne fait que PRÉCISER son poids, ce n'est pas un prix au
                # poids à recalculer. On renseigne juste le poids (E) si absent, sans toucher
                # K/L. Inclut aussi le cas où l'unité de référence d'ORIGINE est elle-même un
                # mot de récipient (ex: D='Barquette', pas une valeur BPU valide) : même
                # logique, sauf que D doit alors être corrigé vers 'PC' (couleur, car ce n'est
                # plus une simple nomenclature équivalente).
                if ((base_ref_norm_precoce in ("pc", "pièce", "piece", "unite", "unité")
                     or base_ref_norm_precoce in _MOTS_CONTENANT_CONNUS)
                        and abs(ratio_prix - 1) <= motor.TOLERANCE_RATIO):
                    r = motor.Resultat()
                    r.nouveau_d = "PC"
                    if base_ref_norm_precoce not in ("pc", "pièce", "piece", "unite", "unité"):
                        r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
                        r.commentaire_d = (
                            f"Unité de référence changée de '{unite_ref}' vers 'PC' : le prix "
                            f"de commande ({prix_cmd}) est égal au prix de référence "
                            f"({prix_ref}), le produit est vendu à la pièce/au récipient entier "
                            f"('{unite_ref}' n'est de toute façon pas une unité BPU valide).")
                    # Normalise L vers le récipient nommé dans le conditionnement, si différent
                    # — sinon les deux restent incohérents entre eux sans raison (ex: 'UNITE'/
                    # 'PU' au lieu de 'PC'/'PIECE').
                    if mot:
                        l_novo = motor.CONTENANT_ABREGE.get(mot, mot.upper())
                        if motor.normalise_texte(l_novo) != motor.normalise_texte(unite_cmd):
                            r.nouveau_l = l_novo
                            # Abréviation déjà reconnue du bon récipient (ex: 'BQT' pour
                            # 'barquette') : uniformise en silence, pas une incohérence.
                            abrege_reconnu = ABREVIATIONS_CONTENANT_RECONNUES.get(
                                motor.normalise_texte(unite_cmd))
                            if abrege_reconnu != mot:
                                r.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
                                r.commentaire_l = (
                                    f"Unité de commande changée en '{l_novo}' d'après le "
                                    f"conditionnement ('{conditionnement}').")
                    if not poids_ref:
                        r.nouveau_e = round(quantite, 3)
                        r.explication_e = (
                            f"Poids à l'unité de référence renseigné à {round(quantite, 3)} "
                            f"{base} d'après le conditionnement ('{conditionnement}') : le prix "
                            f"de commande ({prix_cmd}) est égal au prix de référence ({prix_ref}), "
                            "le produit est donc vendu à la pièce/au récipient entier, pas au "
                            f"{base}.")
                    return r

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
                    # Avant de se contenter de signaler : si le conditionnement est justement
                    # écrit 'CONTENANT DE N X QTÉ UNITÉ' (ex: 'CARTON DE 9 X 500 GR') et que ce
                    # N littéral correspond exactement au ratio entier observé, le prix n'est
                    # pas incohérent — il est simplement donné par SOUS-paquet (ici 500 g), pas
                    # au kilo : l'unité de référence devient 'PC', le poids d'un sous-paquet
                    # (quantité totale ÷ N) renseigne E, et les prix restent intouchés (ils
                    # étaient déjà corrects pour cette lecture-là).
                    m_mult = (RE_CONTENANT_PUIS_MULT.match(conditionnement)
                              or motor.RE_CONTENANT_MULTIPLIE.match(conditionnement))
                    n_texte = None
                    if m_mult:
                        n_texte = int(m_mult.group(2) if m_mult.re is RE_CONTENANT_PUIS_MULT
                                      else m_mult.group(1))
                    if n_texte == ratio_entier:
                        poids_sous_paquet = round(quantite / ratio_entier, 3)
                        r.nouveau_d = "PC"
                        r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
                        r.commentaire_d = (
                            f"Unité de référence changée de '{unite_ref}' vers 'PC' : le "
                            f"conditionnement ('{conditionnement}') décrit {ratio_entier} "
                            f"sous-paquets de {poids_sous_paquet} {base} chacun, et le ratio "
                            f"Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                            f"{round(ratio_prix, 3)}) confirme que le prix de référence est déjà "
                            "celui d'UN sous-paquet, pas du kilo. Prix inchangés.")
                        if not poids_ref:
                            r.nouveau_e = poids_sous_paquet
                        return r
                    r.couleur_m = motor.COULEUR_RATIO_SUSPECT
                    r.commentaire_m = (
                        f"Le ratio Prix commande ÷ Prix référence ({prix_cmd} ÷ {prix_ref} = "
                        f"{round(ratio_prix, 3)}) suggère une "
                        f"quantité réelle de {ratio_entier} {base} par récipient, alors que le "
                        f"conditionnement ('{conditionnement}') indique {round(quantite, 3)} "
                        f"{base}. Vérifiez si le conditionnement est correct avant de valider — "
                        "prix de référence non recalculé automatiquement dans ce cas.")
                    return r

                base_ref_norm = motor.normalise_texte(unite_ref)
                d_novo = base.upper()
                d_change = base_ref_norm not in ("kg", "l", "litre")

                # Avant de confirmer le passage en kg/L : si l'unité de référence d'origine
                # était déjà à la pièce, et que la Désignation (article ou fournisseur)
                # mentionne un poids PAR PIÈCE (ex: 'CHOUQUETTE CRÈME 18,5 GR') qui, divisé dans
                # le poids total du conditionnement ('CARTON DE 2,5 KG', qui ne donne qu'un
                # poids total, aucun compte), retombe sur le ratio Prix commande ÷ Prix
                # référence observé : l'unité de référence à la pièce était la bonne lecture, le
                # conditionnement ne révèle juste pas directement le nombre de pièces. Repris
                # pour l'instant seulement de la Désignation (pas du Conditionnement, qui a déjà
                # été examiné en vain par piece_carton_candidats()/determine_base_et_quantite()).
                if d_change:
                    for texte_desig in (designation, desig_fourn):
                        poids_piece_desig, base_poids_desig = extraire_quantite_designation_etendue(texte_desig)
                        if not poids_piece_desig or base_poids_desig != base:
                            continue
                        candidat_pieces = quantite / poids_piece_desig
                        if abs(ratio_prix - candidat_pieces) <= motor.TOLERANCE_RATIO * candidat_pieces:
                            r2 = motor.Resultat()
                            r2.nouveau_d = "PC"
                            if not poids_ref:
                                r2.nouveau_e = round(poids_piece_desig, 4)
                                r2.explication_e = (
                                    f"Le poids d'une pièce semble être {round(poids_piece_desig, 4)} "
                                    f"kg, d'après la désignation ('{texte_desig}').")
                            return r2

                nouveau_k = round(prix_cmd / quantite, 4)
                diff_relative = abs(nouveau_k - prix_ref) / prix_ref if prix_ref else abs(nouveau_k - prix_ref)
                # Alignée sur motor.TOLERANCE_RATIO (2%, la tolérance standard du projet) —
                # c'était à tort 5% ici, ce qui laissait passer en silence des écarts entre 2%
                # et 5% (ex: 'COLIS DE 2 X 2,2 KG' avec un ratio de 4.6 au lieu de 4.4, 4.5%
                # d'écart). En dessous de ce seuil, le prix de référence reste quasi identique
                # (arrondis, centimes) — ne rien changer évite de déranger pour rien.
                k_change = diff_relative > motor.TOLERANCE_RATIO
                l_novo = motor.CONTENANT_ABREGE.get(mot, mot.upper()) if mot else None
                # Aucun récipient nommé ('6 x 1L'...), mais chaque unité fait exactement 1
                # kg/L : comme pour les produits à la pièce, on précise "LOT DE N" plutôt que de
                # laisser l'unité de commande telle quelle — confirmé sur fichiers réels (ex:
                # lait vendu en 'LOT DE 6' bouteilles de 1L chacune).
                if l_novo is None:
                    m_double = motor.RE_DOUBLE.match(conditionnement)
                    if m_double:
                        n_unites, qte_unite, _unite_texte = m_double.groups()
                        if abs(motor._nombre(qte_unite) - 1) < 0.01:
                            l_novo = f"LOT DE {int(n_unites)}"
                l_change = l_novo and motor.normalise_texte(l_novo) != motor.normalise_texte(unite_cmd)
                l_change_silencieux = False
                if l_change and mot:
                    # L'unité de commande actuelle est déjà une abréviation reconnue du BON
                    # récipient (ex: 'BQT' pour 'barquette') — pas une incohérence, juste une
                    # autre façon d'écrire la même chose : on uniformise vers le mot en entier,
                    # mais silencieusement (pas de couleur/commentaire pour une simple réécriture).
                    abrege_reconnu = ABREVIATIONS_CONTENANT_RECONNUES.get(
                        motor.normalise_texte(unite_cmd))
                    if abrege_reconnu == mot:
                        l_change_silencieux = True

                r.nouveau_d = d_novo
                if d_change:
                    r.couleur_d = motor.COULEUR_UNITE_TRANSFORMEE
                    r.commentaire_d = (
                        f"Unité de référence changée de '{unite_ref}' vers '{d_novo}' : le "
                        f"conditionnement ('{conditionnement}') décrit un récipient, pas "
                        "directement une unité de mesure.")

                if k_change:
                    r.nouveau_k = nouveau_k
                    r.couleur_k = motor.COULEUR_RATIO_SUSPECT
                    qte_affichee = round(quantite, 6) if quantite < 1 else round(quantite, 3)
                    r.commentaire_k = (
                        f"Prix de référence suggéré : {prix_cmd} ÷ {qte_affichee} "
                        f"{base} = {nouveau_k} (prix de référence actuel : {prix_ref}). Le "
                        f"conditionnement ('{conditionnement}') indique {qte_affichee} {base} "
                        f"pour ce récipient — à vérifier, non modifié automatiquement.")
                    if mot in ("piece", "pièce"):
                        # 'piece' générique (pas un vrai récipient nommé) : l'incohérence peut
                        # aussi bien venir d'une Unité de référence mal choisie (l'article est
                        # peut-être vendu à la pièce, pas au kg/L) que d'un prix de référence
                        # erroné — les deux pistes sont à vérifier, pas seulement le prix.
                        r.commentaire_k += (
                            f" Autre piste à considérer : l'article est peut-être vendu à la "
                            f"pièce, pas au {base} — dans ce cas, ce serait l'Unité de "
                            f"référence ('{unite_ref}') qu'il faudrait changer en 'PC', plutôt "
                            "que le prix.")

                if l_change:
                    r.nouveau_l = l_novo
                    if not l_change_silencieux:
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
    reconsidere_prix_inexplique(resultat, unite_ref, conditionnement, unite_cmd, prix_ref, prix_cmd,
                                 designation, desig_fourn)
    reconsidere_recipient_de_1_unite(resultat, unite_ref, unite_cmd, conditionnement)
    reconsidere_ratio_un_contre_contenant(resultat, unite_ref, unite_cmd, conditionnement, prix_ref, prix_cmd)
    reconsidere_contenant_nombre_pieces(resultat, unite_ref, conditionnement)
    return resultat


def reconsidere_contenant_nombre_pieces(resultat, unite_ref, conditionnement):
    """Cas 'nombre de pièces' du moteur de base (ex: œufs, 'Colis de 60 pièces') : l'unité de
    commande reprend telle quelle l'ancienne unité de référence ('Poche', 'Carton'...) au lieu
    du récipient réellement nommé dans le conditionnement ('Colis'). Si le conditionnement
    nomme clairement un récipient reconnu, on le préfère à ce texte recopié."""
    if resultat.nouveau_d != "PC" or resultat.nouveau_l != unite_ref:
        return
    mot = mot_contenant_dans_conditionnement(conditionnement)
    if not mot:
        return
    l_novo = motor.CONTENANT_ABREGE.get(mot, mot.upper())
    if motor.normalise_texte(l_novo) == motor.normalise_texte(unite_ref):
        return
    resultat.nouveau_l = l_novo
    resultat.commentaire_l = (
        f"Unité de commande précisée en '{l_novo}' d'après le conditionnement "
        f"('{conditionnement}'), plutôt que l'ancienne unité de référence ('{unite_ref}').")


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

    # Calibre/poids approximatif ou variable d'une pièce à l'autre (ex: '110g environ',
    # 'calibres 90 et 110g') : ne décrit pas une quantité totale, et quand l'unité de commande
    # est déjà dans la même base que l'unité de référence, ratio ≈ 1 est justement le résultat
    # ATTENDU (achat au poids) — rien à signaler, ce nombre ne doit pas être traité comme une
    # quantité à multiplier.
    if base_cmd == base_cible and RE_QUANTITE_APPROXIMATIVE.search(conditionnement):
        return

    # D'abord un multiplicateur explicite devant un mot de récipient en texte libre (ex: '4
    # barquettes de 400g'), sinon un simple nombre + unité isolé.
    candidats_libres = []
    for n_str, qty_str, unite_str in RE_MULTIPLICATEUR_CONTENANT_LIBRE.findall(conditionnement):
        try:
            candidats_libres.append((
                motor._nombre(qty_str) * motor.FACTEUR_VERS_KG_OU_L[unite_str.lower()] * int(n_str),
                unite_str))
        except (KeyError, ValueError):
            continue
    if not candidats_libres:
        for qty_str, unite_str in motor.RE_LIBRE.findall(conditionnement):
            try:
                candidats_libres.append((
                    motor._nombre(qty_str) * motor.FACTEUR_VERS_KG_OU_L[unite_str.lower()],
                    unite_str))
            except (KeyError, ValueError):
                continue

    candidats_pertinents = [
        quantite for quantite, unite_str in candidats_libres
        if motor.base_depuis_unite(unite_str) == ("kg" if base_cible == "kg" else "l") and quantite > 0]

    # Si au moins un nombre du texte CONFIRME déjà le ratio (ex: 'SACHET DE 1 KG, PIECE DE 120
    # GR' — le '1 KG' confirme, le '120 GR' n'est qu'un détail par pièce qu'il ne faut pas
    # traiter comme une quantité totale contredisant le premier) : rien à signaler. On en
    # profite pour préciser l'unité de commande avec le récipient nommé dans le texte (même
    # recherche que reconsidere_contenant_nombre_pieces, insensible aux virgules), si elle
    # n'était encore qu'un simple rappel de la base kg/L.
    if any(abs(quantite - 1) <= motor.TOLERANCE_RATIO for quantite in candidats_pertinents):
        if base_cmd == base_cible:
            mot = mot_contenant_dans_conditionnement(conditionnement)
            if mot:
                l_novo = motor.CONTENANT_ABREGE.get(mot, mot.upper())
                if motor.normalise_texte(l_novo) != motor.normalise_texte(unite_cmd):
                    resultat.nouveau_l = l_novo
                    resultat.couleur_l = motor.COULEUR_UNITE_TRANSFORMEE
                    resultat.commentaire_l = (
                        f"Unité de commande précisée en '{l_novo}' d'après le conditionnement "
                        f"('{conditionnement}').")
        return

    # Aucun nombre ne confirme le ratio : on signale avec le PREMIER trouvé (comportement
    # d'origine, inchangé).
    for quantite in candidats_pertinents:
        if abs(quantite - 1) > motor.TOLERANCE_RATIO:
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
    rf"^\s*({_MOTS_REGEX})s?\s+{_MOT_QUALIFICATIF}(?:de\s+)?(\d+)\s*{_MOT_COMPTAGE_FINAL}\s*$", re.I)

# "LOT DE N PIECES" (ex: yaourts vendus en lot de 2/4/6/8) : même logique que ci-dessus, mais
# "lot" n'est volontairement PAS un récipient physique (pas dans _MOTS_REGEX) — c'est un
# simple marqueur de comptage, jamais un nom à reprendre tel quel dans l'unité de commande
# (voir le marqueur spécial 'lot' dans piece_carton_candidats()).
RE_LOT_PIECES_SEULEMENT = re.compile(r"^\s*lots?\s+(?:de\s+)?(\d+)\s*pi[eè]ces?\s*$", re.I)

# Détecte 'LOT DE N' écrit tel quel en unité de commande (toutes origines confondues — piece_
# carton_candidats, le relabel kg/L 'N x 1kg/1L'...) : l'unité de commande elle-même ne garde
# que 'LOT' (le nombre n'est pas vraiment une unité), et le nombre est reporté au début du
# Conditionnement à la place (voir le point d'écriture de L dans audite_fournisseur).
RE_LOT_DE_N_UNITE_CMD = re.compile(r"^lot\s+de\s+(\d+)$", re.I)

# Multiplicateur explicite devant un mot de récipient, n'importe où dans un texte libre non
# structuré (ex: 'sous-atmosphère - 4 barquettes de 400g') — contrairement aux motifs stricts
# ci-dessus (ancrés en ^...$), celui-ci sert à la recherche de dernier recours dans
# reconsidere_prix_inexplique(), sur du texte qui ne respecte aucun format propre.
RE_MULTIPLICATEUR_CONTENANT_LIBRE = re.compile(
    rf"(\d+)\s*(?:{_MOTS_REGEX})s?\s+de\s+([\d,.]+)\s*({_UNITES_REGEX})", re.I)

# Fourchette de poids/volume (ex: 'PIECE DE 1,5 A 2,5 KG', '1-3KG') — poids variable d'une
# pièce à l'autre, précisé par une plage plutôt qu'un nombre unique.
RE_PLAGE_LIBRE = re.compile(rf"([\d,.]+)\s*(?:[aà]|-)\s*([\d,.]+)\s*({_UNITES_REGEX})\b", re.I)

# Calibre/poids approximatif ou variable d'une pièce à l'autre (ex: '110g environ', '100g et
# plus', 'calibres 90 et 110g') — décrit la taille indicative d'UNE pièce, pas une quantité
# totale à multiplier (voir reconsidere_prix_inexplique).
RE_QUANTITE_APPROXIMATIVE = re.compile(r"\benviron\b|\bet\s+plus\b|^\s*calibres?\b", re.I)


def _cherche_candidats_quantite_libre(texte, base_cible):
    """Cherche, dans un texte libre (Conditionnement OU Désignation), toute quantité
    exploitable dans la base visée (kg/L) : d'abord une fourchette ('1,5 A 2,5 KG'), puis un
    multiplicateur devant un mot de récipient ('4 barquettes de 400g'), puis un simple nombre
    + unité isolé. Renvoie (candidats, plage) où 'plage' est (basse, haute) si une fourchette a
    été trouvée (sinon None) et 'candidats' la liste des quantités uniques trouvées."""
    if not texte:
        return [], None
    m_plage = RE_PLAGE_LIBRE.search(texte)
    if m_plage:
        borne_basse, borne_haute, unite_str = m_plage.groups()
        if motor.base_depuis_unite(unite_str) == ("kg" if base_cible == "kg" else "l"):
            try:
                facteur = motor.FACTEUR_VERS_KG_OU_L[unite_str.lower()]
                basse = motor._nombre(borne_basse) * facteur
                haute = motor._nombre(borne_haute) * facteur
                return [], (min(basse, haute), max(basse, haute))
            except (KeyError, ValueError):
                pass

    candidats = []
    for n_str, qty_str, unite_str in RE_MULTIPLICATEUR_CONTENANT_LIBRE.findall(texte):
        if motor.base_depuis_unite(unite_str) != ("kg" if base_cible == "kg" else "l"):
            continue
        try:
            candidats.append(motor._nombre(qty_str) * motor.FACTEUR_VERS_KG_OU_L[unite_str.lower()]
                              * int(n_str))
        except (KeyError, ValueError):
            continue
    if not candidats:
        for qty_str, unite_str in motor.RE_LIBRE.findall(texte):
            if motor.base_depuis_unite(unite_str) != ("kg" if base_cible == "kg" else "l"):
                continue
            try:
                candidats.append(motor._nombre(qty_str) * motor.FACTEUR_VERS_KG_OU_L[unite_str.lower()])
            except (KeyError, ValueError):
                continue
    return [c for c in candidats if c > 0], None


def reconsidere_prix_inexplique(resultat, unite_ref, conditionnement, unite_cmd, prix_ref,
                                 prix_cmd, designation=None, desig_fourn=None):
    """Quand le BPU est au kg/L et que rien dans le conditionnement n'a pu expliquer le
    ratio Prix commande ÷ BPU via les motifs reconnus, plusieurs cas différents :
    1. Le conditionnement ne donne qu'un NOMBRE DE PIÈCES sans poids ('CAGEOT DE 6 PIECES')
       et la commande se fait dans ce même récipient : le poids réel est inconnu (souvent
       variable) — impossible de vérifier, donc pas d'incohérence à signaler.
    2. Un mot ou code non reconnu casse le motif ('SAC KRAFT DE 5 KG'), ou la quantité n'est
       écrite que dans la Désignation (de l'article ou du fournisseur), pas le Conditionnement
       ('SEAU CONSIGNÉ' + désignation '... SEAU 5L') — on cherche quand même, en texte libre,
       une quantité qui explique le ratio observé.
    3. Rien nulle part (ni Conditionnement ni Désignation) ne permet de vérifier quoi que ce
       soit — impossible de confirmer OU d'infirmer, donc pas d'incohérence à signaler non
       plus (juste un signalement vague ne ferait qu'ajouter du bruit sans aide réelle)."""
    if resultat.couleur_m != motor.COULEUR_PRIX_INEXPLIQUE:
        return
    base_cible = motor.normalise_texte(unite_ref)
    if base_cible not in ("kg", "l", "litre") or not prix_ref:
        return

    m = RE_CONTENANT_PIECES_SEULEMENT.match(conditionnement or "")
    if m and unite_cmd and motor.normalise_texte(unite_cmd) == motor.normalise_texte(m.group(1)):
        resultat.couleur_m = None
        resultat.commentaire_m = None
        nb_pieces = int(m.group(2))
        if prix_cmd and abs(prix_cmd / prix_ref - nb_pieces) <= motor.TOLERANCE_RATIO * nb_pieces:
            # Aucun poids n'est donné nulle part (impossible de vérifier un recalcul en kg/L),
            # mais le ratio Prix commande ÷ Prix référence colle exactement au nombre de pièces
            # annoncé par le conditionnement — signe que le prix de référence pourrait en
            # réalité être un prix à la pièce, étiqueté à tort en kg/L. Simple suggestion à
            # vérifier par une personne, jamais changé seul (le poids réel d'une pièce reste
            # inconnu, donc pas de recalcul automatique).
            resultat.couleur_d = motor.COULEUR_RATIO_SUSPECT
            resultat.commentaire_d = (
                f"Le conditionnement ('{conditionnement}') ne donne que {nb_pieces} pièces, "
                "sans aucun poids, et le prix semble correspondre exactement à ce nombre de "
                f"pièces. L'unité de référence devrait peut-être être 'PC' plutôt que "
                f"'{unite_ref}' — à vérifier.")
        return

    ratio_observe = prix_cmd / prix_ref

    # Calibre/poids APPROXIMATIF ou variable d'une pièce à l'autre (ex: '110g environ',
    # '120-130g environ', '100g et plus', 'calibres 90 et 110g') : ce nombre décrit la taille
    # indicative d'UNE pièce/filet (poisson...), pas une quantité totale à multiplier — la
    # recherche libre ci-dessous s'y tromperait (elle traiterait '110g environ' comme si
    # c'était la quantité totale du conditionnement). Quand l'unité de commande est déjà dans
    # la MÊME base que l'unité de référence (les deux en kg, ou les deux en L), le prix devrait
    # être identique des deux côtés (achat au poids, peu importe la taille exacte de la pièce) —
    # ratio attendu = 1, sinon incohérence claire de prix/unité, sans rapport avec ce calibre.
    if conditionnement and RE_QUANTITE_APPROXIMATIVE.search(conditionnement):
        base_cmd = motor.normalise_texte(unite_cmd) if unite_cmd else None
        base_cmd = "kg" if base_cmd == "kg" else ("l" if base_cmd in ("l", "litre") else None)
        if base_cmd == base_cible:
            if abs(ratio_observe - 1) <= motor.TOLERANCE_RATIO:
                resultat.couleur_m = None
                resultat.commentaire_m = None
            else:
                resultat.commentaire_m = (
                    f"Le conditionnement ('{conditionnement}') ne donne qu'un calibre/poids "
                    f"approximatif (pas une quantité à multiplier) : l'unité de référence et "
                    f"l'unité de commande sont toutes les deux en {base_cible}, donc le prix "
                    f"devrait être identique des deux côtés, mais le prix de référence "
                    f"({prix_ref}) et le prix de commande ({prix_cmd}) diffèrent (ratio = "
                    f"{round(ratio_observe, 3)}). Incohérence de prix/unité à vérifier "
                    "manuellement.")
            return

    texte_designation = " ".join(t for t in (designation, desig_fourn) if t)

    for source, texte in (("Conditionnement", conditionnement), ("Désignation", texte_designation)):
        candidats, plage = _cherche_candidats_quantite_libre(texte, base_cible)
        if plage:
            if plage[0] <= ratio_observe <= plage[1]:
                resultat.couleur_m = None
                resultat.commentaire_m = None
                base_cmd = motor.normalise_texte(unite_cmd) if unite_cmd else None
                if base_cmd not in ("kg", "l", "litre"):
                    # Le poids de la pièce varie dans un intervalle (pas une valeur fixe) :
                    # payer un prix fixe par pièce suppose un poids moyen, alors que le poids
                    # réellement livré peut être n'importe où dans l'intervalle — une Unité de
                    # commande au kg/L (même prix que l'Unité de référence, facturé sur le
                    # poids réellement pesé/livré) serait plus fidèle.
                    resultat.couleur_l = motor.COULEUR_A_VERIFIER
                    resultat.commentaire_l = (
                        f"Le poids d'une pièce varie dans une fourchette "
                        f"({plage[0]}-{plage[1]} {base_cible}), pas un poids fixe — le poids "
                        "réellement livré peut varier. À considérer : passer cette colonne en "
                        f"'{base_cible.upper()}' (avec le même prix que l'Unité de référence) "
                        "serait peut-être plus juste qu'un prix fixe par pièce.")
                return
            continue
        for valeur_kg in candidats:
            if abs(ratio_observe - valeur_kg) <= motor.TOLERANCE_RATIO * valeur_kg:
                resultat.couleur_m = None
                resultat.commentaire_m = None
                return
        if candidats:
            # Une quantité est bien trouvée (dans le Conditionnement ou la Désignation), mais
            # le ratio observé ne correspond à aucune d'elles — trop incertain pour recalculer
            # automatiquement, mais le message pointe directement le problème : un des deux
            # prix semble faux, pas juste "pas d'explication".
            quantite = candidats[0]
            resultat.commentaire_m = (
                f"Le {source.lower()} ('{texte}') suggère {round(quantite, 3)} {base_cible} "
                f"— dans ce cas le prix de commande devrait valoir {round(quantite, 3)} fois le "
                f"prix de référence, soit {round(prix_ref * quantite, 4)}, mais il vaut "
                f"{prix_cmd} (ratio observé = {round(ratio_observe, 3)}). Un des deux prix "
                "semble incorrect — à vérifier manuellement (pas recalculé automatiquement, le "
                "format n'est pas standard).")
            return

    # Rien trouvé nulle part pour vérifier quoi que ce soit — un signalement vague n'aiderait
    # pas, on reste silencieux plutôt que d'ajouter du bruit.
    resultat.couleur_m = None
    resultat.commentaire_m = None


# ============================================================================
# ORIGINE — contrairement à pretraitement.py (qui corrige automatiquement une valeur non
# reconnue en 'France'), l'audit ne modifie JAMAIS la valeur : seulement un signalement, à
# charge d'une personne de corriger. Liste de pays volontairement plus large que celle de
# pretraitement.py (observée directement dans les fichiers de référence).
# ============================================================================
ORIGINES_LARGES = {"france", "ue", "union europeenne", "hors ue", "hors union europeenne", "monde"}
PAYS_ORIGINE_CONNUS = {
    "espagne", "italie", "belgique", "chine", "royaume-uni", "royaume uni", "chili", "perou",
    "irlande", "egypte", "asie", "allemagne", "portugal", "pays-bas", "maroc", "pologne",
    "canada", "costa rica", "inde", "malaisie", "norvege", "bangladesh", "bresil", "usa",
    "etats-unis", "nouvelle zelande", "islande", "grece",
}


# BIO/Label détectés depuis la Désignation quand mentionnés explicitement (voir
# audite_fournisseur) — vocabulaire volontairement restreint à des mentions non ambiguës.
# Négation ('NON BIO', 'SANS BIO') explicitement exclue pour ne pas marquer l'inverse.
RE_BIO_DETECTE = re.compile(r"\bbio(?:logique)?\b", re.I)
RE_BIO_NEGATION = re.compile(r"\b(?:non|sans)[\s-]+bio(?:logique)?\b", re.I)
RE_LABEL_DETECTE = re.compile(r"\blabel\s+rouge\b|\baop\b|\baoc\b|\bigp\b|\bstg\b|\bce2\b", re.I)


def _normalise_origine(valeur):
    texte = re.sub(r"\(.*?\)", "", str(valeur)).strip().lower()  # retire qualificatif entre parenthèses, ex: "France (79)"
    texte = "".join(c for c in unicodedata.normalize("NFD", texte) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texte).strip()


def verifie_origine(valeur):
    """Vérifie qu'Origine est soit une des 4 grandes catégories (France/UE/Hors UE/Monde),
    soit un seul pays reconnu. Renvoie (True, None) si reconnu/vide, (False, message) sinon —
    ne modifie jamais la valeur, juste un signalement."""
    if not valeur:
        return True, None  # vide : optionnelle (cf. format_cible.md), pas le rôle de cette règle
    norm = _normalise_origine(valeur)
    if norm in ORIGINES_LARGES or norm in PAYS_ORIGINE_CONNUS:
        return True, None
    if any(sep in norm for sep in (",", "/")):
        return False, (
            f"Origine ('{valeur}') semble combiner plusieurs valeurs — préciser laquelle "
            "s'applique réellement à ce produit (une seule valeur attendue).")
    return False, (
        f"Origine ('{valeur}') non reconnue : ni une des 4 catégories (France, UE, Hors UE, "
        "Monde), ni un pays de la liste connue — vérifier l'orthographe, ou que ce pays existe "
        "bien (liste à élargir si besoin).")


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
    if RE_FORMAT_BOITE.fullmatch(cat):
        # 'BOITE 5/1', 'BOITE 3/1'... : format de boîte normalisé (voir
        # reconsidere_format_boite) — reste une 'boite' pour la comparaison ci-dessous avec
        # le mot trouvé dans le conditionnement/la désignation.
        cat = "boite"
    cond = motor.normalise_texte(conditionnement)
    desig = motor.normalise_texte(designation)
    texte_combine = f"{cond} {desig}"

    if not conditionnement and not designation:
        return True, None
    if cat in ("kg", "l"):
        return True, None
    if re.match(r"^\d+\s*pi[eè]ces?$", cat):
        return True, None
    m_lot = re.match(r"^lot de (\d+)$", cat)
    if m_lot:
        n = int(m_lot.group(1))
        # Le nombre N n'est pas toujours écrit tel quel (ex: '12 x 4 x 100g' pour un lot de 48 :
        # 48 = 12 × 4, jamais écrit littéralement) — on accepte aussi le produit de deux
        # nombres présents dans le conditionnement.
        nombres = [int(x) for x in re.findall(r"\d+", cond)]
        if n in nombres or any(a * b == n for i, a in enumerate(nombres) for b in nombres[i + 1:]):
            return True, None

    mots_types = ["sac", "pièce", "piece", "seau", "brique", "barquette", "barq", "bidon",
                  "outre", "colis", "pot", "sachet", "carton", "poche", "bouteille", "bouteile",
                  "cageot", "boite", "boîte", "filet", "caisse", "plaque", "plateau", "clayette",
                  "bac", "coffret", "plaquette", "paquet", "etui", "étui", "bombe",
                  "bocal", "flacon", "tube"]
    trouve = [m for m in mots_types if m in texte_combine]
    if not trouve:
        return True, None

    cat_singulier = cat[:-1] if cat.endswith("s") and cat[:-1] in mots_types else cat
    if cat_singulier in trouve:
        return True, None
    # Le mot trouvé est peut-être NON reconnu par la base (CONTENANT_ABREGE porte alors son
    # remplacement officiel, pas le mot lui-même) — l'unité de commande actuelle est ce
    # remplacement, pas une incohérence (ex: conditionnement 'ETUI 1,25 KG', unité de commande
    # déjà 'PAQUET', le remplacement officiel d'ETUI). Comparé au texte normalisé de
    # l'unité de commande elle-même, pas à 'cat' : categorie_commande() regroupe certains mots
    # dans une catégorie plus large (ex: 'POT' -> 'pc'), ce qui ferait rater la comparaison
    # exacte dont on a besoin ici.
    remplacements_trouves = {motor.normalise_texte(motor.CONTENANT_ABREGE.get(m, m)) for m in trouve}
    if motor.normalise_texte(unite_cmd) in remplacements_trouves:
        return True, None
    if cat == "pc" and any(m in ("pièce", "piece", "pot", "sachet") for m in trouve):
        return True, None
    if cat == "barquette" and any(m in ("barquette", "barq") for m in trouve):
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
                cat == "barquette" and any(m in ("barquette", "barq") for m in trouve_corrige)) or (
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

    col_marche = col_lot = col_designation = col_unite_ref = col_poids_ref = col_nb_pieces = None
    fournisseurs_col = []  # [(numero, colonne_du_label_fournisseur)]
    for c in range(1, ws.max_column + 1):
        texte = _normalise_entete(ws.cell(row=ligne_entete, column=c).value)
        if texte == "marché":
            col_marche = c
        elif texte == "lot":
            col_lot = c
        elif "désignation de l'article" in texte:
            col_designation = c
        elif texte.startswith("poids"):
            # Match par préfixe (pas juste "poids à l'unité de référence" en substring) :
            # certains fichiers (ex: Base IANOR AGAPE 16) utilisent un texte différent ("Poids
            # de la PC (si Unité de référence = PC)") qui contient lui-même la phrase "unité de
            # référence" entre parenthèses — une recherche en substring sur "unité de
            # référence" matchait alors À TORT cette colonne Poids comme étant la colonne
            # Unité de référence (jamais la bonne colonne 4, toujours vide ensuite).
            col_poids_ref = c
        elif "nombre de" in texte and "pièces" in texte:
            col_nb_pieces = c
        elif texte.startswith("unité") and "référence" in texte:
            # Préfixe "unité" (pas substring "unité de référence") pour couvrir aussi le texte
            # "Unité Référence (BPU)..." (sans "de") observé dans Base IANOR AGAPE 16, sans
            # jamais matcher la colonne Poids ci-dessus (qui commence par "poids", pas "unité").
            col_unite_ref = c
        else:
            m = re.match(r"fournisseur\s*(\d+)", texte)
            if m:
                fournisseurs_col.append((int(m.group(1)), c))

    if col_designation is None or col_unite_ref is None or not fournisseurs_col:
        return None

    fournisseurs = []
    cols_fournisseur_triees = sorted(c for _, c in fournisseurs_col)
    for numero, col in sorted(fournisseurs_col):
        # Origine n'est pas à un décalage fixe par rapport au bloc : le bloc Fournisseur 3 n'a
        # pas de colonne "Circuit court", ce qui décale toutes les colonnes suivantes d'une
        # position par rapport aux blocs 1/2 — on la retrouve donc par son texte d'en-tête,
        # cherchée seulement dans la plage de colonnes de CE bloc (jusqu'au bloc suivant).
        prochaine_col = next((c2 for c2 in cols_fournisseur_triees if c2 > col), ws.max_column + 1)
        col_origine = col_bio = col_label = None
        for c2 in range(col + 1, prochaine_col):
            texte_c2 = _normalise_entete(ws.cell(row=ligne_entete, column=c2).value)
            if texte_c2 == "origine":
                col_origine = c2
            elif texte_c2.startswith("bio"):
                col_bio = c2
            elif texte_c2.startswith("label"):
                col_label = c2
        fournisseurs.append({
            "nom": f"Fournisseur {numero}",
            "cond": col + 1, "desig_fourn": col + 2, "ref": col + 3,
            "prix_ref": col + 4, "unite_cmd": col + 5, "prix_cmd": col + 6,
            "origine": col_origine, "bio": col_bio, "label": col_label,
        })

    derniere_colonne_cle = max(f["prix_cmd"] for f in fournisseurs) + 6  # marge (Local/Bio/.../Indice I0)

    return {
        "ligne_debut_donnees": ligne_entete + 1,
        "col_marche": col_marche,
        "col_lot": col_lot,
        "col_designation": col_designation,
        "col_unite_ref": col_unite_ref,
        "col_poids_ref": col_poids_ref,
        "col_nb_pieces": col_nb_pieces,
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
    """Parcourt, EN LECTURE SEULE, tous les fichiers de pret_pour_audit/ pour relever tous les prix
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
                        col_poids_ref, lot_valor=None, indice_prix=None, col_nb_pieces=None):
    conditionnement = ws_valeurs.cell(row=ligne, column=fournisseur["cond"]).value
    desig_fourn = ws_valeurs.cell(row=ligne, column=fournisseur["desig_fourn"]).value
    reference = ws_valeurs.cell(row=ligne, column=fournisseur["ref"]).value
    unite_cmd = ws_valeurs.cell(row=ligne, column=fournisseur["unite_cmd"]).value
    prix_ref = ws_valeurs.cell(row=ligne, column=fournisseur["prix_ref"]).value
    prix_cmd = ws_valeurs.cell(row=ligne, column=fournisseur["prix_cmd"]).value

    # Un prix écrit en texte (ex: '104.5' saisi comme chaîne, pas comme nombre) fait planter
    # tous les calculs (prix_cmd / prix_ref) plus loin — on tente de le convertir, et si ce
    # n'est vraiment pas un nombre, on le signale au lieu de planter.
    total_prix_texte = 0
    for champ, valeur in (("prix_ref", prix_ref), ("prix_cmd", prix_cmd)):
        if isinstance(valeur, str):
            try:
                convertie = motor._nombre(valeur)
            except ValueError:
                marque(ws.cell(row=ligne, column=fournisseur[champ]), motor.COULEUR_A_VERIFIER,
                       fournisseur["nom"],
                       f"Valeur ('{valeur}') écrite en texte et pas reconnue comme un nombre.")
                total_prix_texte += 1
                convertie = None
            if champ == "prix_ref":
                prix_ref = convertie
            else:
                prix_cmd = convertie

    if reference is None and conditionnement is None and prix_ref is None and prix_cmd is None:
        return total_prix_texte  # ce fournisseur n'est simplement pas renseigné pour cette ligne

    total = total_prix_texte

    # Produit vendu au poids sans conditionnement fixe (CAS 2 documenté par le métier dans
    # useful_stuff/Exemples unités et prix 6.xlsx — viandes/poissons/fromages/fruits et
    # légumes) : Conditionnement vide, Unité de référence = Unité de commande (KG ou L), et
    # prix de référence = prix de commande (aucune conversion nécessaire) — ce n'est pas une
    # donnée manquante, juste l'absence de récipient fixe. Rempli à 'VRAC' pour le distinguer
    # d'une vraie case vide (qui, elle, reste un problème potentiel à vérifier manuellement).
    if (not conditionnement and prix_ref is not None and prix_cmd is not None
            and unite_ref and motor.normalise_texte(unite_ref) in ("kg", "l", "litre")
            and unite_cmd and motor.normalise_texte(unite_cmd) == motor.normalise_texte(unite_ref)
            and prix_ref == prix_cmd):
        # Rien n'est écrit dans cette version : juste un commentaire simple expliquant ce qu'on
        # aurait mis.
        marque(ws.cell(row=ligne, column=fournisseur["cond"]), motor.COULEUR_UNITE_TRANSFORMEE,
               fournisseur["nom"],
               "Le conditionnement est vide, mais tout indique que ce produit est vendu au "
               "poids, sans emballage fixe (le prix d'achat et le prix de référence sont "
               "identiques) : on pourrait écrire 'VRAC' ici.")
        total += 1
        conditionnement = "VRAC"

    # Référence de plage de cellules Excel ('H33:H66') collée par erreur au milieu du
    # Conditionnement : repéré, mais pas corrigé dans cette version (rien n'est écrit).
    if conditionnement:
        conditionnement_sans_plage, corrige = retire_plage_cellule_collee(conditionnement)
        if corrige:
            marque(ws.cell(row=ligne, column=fournisseur["cond"]), motor.COULEUR_A_VERIFIER,
                   fournisseur["nom"],
                   "Il semble y avoir un morceau de cellules Excel collé par erreur au milieu "
                   f"du texte (quelque chose comme 'H33:H66'). Sans ça, le texte serait : "
                   f"'{conditionnement_sans_plage}'.")
            total += 1
            conditionnement = conditionnement_sans_plage

    # Faute de frappe sur un mot de contenant (ex: 'BARQUETE' au lieu de 'BARQUETTE') :
    # seulement quand l'analyse normale échouerait sinon (pas de faux positif sur un mot qui
    # se trouve juste être proche d'un contenant connu par coïncidence). Repéré, mais rien
    # n'est écrit dans cette version.
    if conditionnement and unite_ref:
        conditionnement_propre = nettoie_conditionnement(conditionnement)
        if (piece_carton_candidats(unite_ref, conditionnement_propre) is None
                and determine_base_et_quantite(conditionnement_propre) is None):
            conditionnement_corrige, typo = corrige_faute_frappe_contenant(conditionnement_propre)
            if typo:
                mot_original, mot_corrige = typo
                marque(ws.cell(row=ligne, column=fournisseur["cond"]), motor.COULEUR_A_VERIFIER,
                       fournisseur["nom"],
                       f"Le mot '{mot_original}' ressemble à une faute de frappe pour "
                       f"'{mot_corrige}' — à vérifier.")
                total += 1
                conditionnement = conditionnement_corrige

    # Cellule vide : la couleur seule suffit, pas de commentaire (allège le fichier).
    if prix_ref is None:
        cel = ws.cell(row=ligne, column=fournisseur["prix_ref"])
        cel.fill = PatternFill(start_color=motor.COULEUR_A_VERIFIER,
                                end_color=motor.COULEUR_A_VERIFIER, fill_type="solid")
        total += 1
    if prix_cmd is None:
        cel = ws.cell(row=ligne, column=fournisseur["prix_cmd"])
        cel.fill = PatternFill(start_color=motor.COULEUR_A_VERIFIER,
                                end_color=motor.COULEUR_A_VERIFIER, fill_type="solid")
        total += 1

    resultat = None
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

        # Comparaison du prix à la médiane des autres lignes/catégories désactivée à la
        # demande — seule la cohérence interne de la ligne (unité/prix/conditionnement) reste
        # vérifiée.
        # reconsidere_recalcul_k_vs_indice_prix(resultat, indice_prix, designation, lot_valor,
        #                                        prix_ref, prix_cmd)

        d_final = resultat.nouveau_d if resultat.nouveau_d is not None else unite_ref
        base_final_d = motor.normalise_texte(d_final) if d_final else None

        if col_poids_ref and base_final_d in ("kg", "l", "litre"):
            # L'unité de référence semble être en kg/L — la colonne Poids à l'unité de
            # référence ne sert QUE pour les unités à la pièce. Rien n'est vidé dans cette
            # version (on ne touche à aucune cellule) : juste un repère si une valeur y
            # traîne déjà et ne semble plus avoir de sens.
            cel_e = ws.cell(row=ligne, column=col_poids_ref)
            if cel_e.value is not None:
                marque(cel_e, motor.COULEUR_A_VERIFIER, fournisseur["nom"],
                       "Le produit semble être vendu au kilo ou au litre — la valeur ici ne "
                       "sert normalement que pour un produit vendu à la pièce.")
                total += 1
        else:
            if col_poids_ref and resultat.nouveau_e not in ("GARDER", "VIDER", None):
                cel_e = ws.cell(row=ligne, column=col_poids_ref)
                if cel_e.value is None:
                    marque(cel_e, motor.COULEUR_UNITE_TRANSFORMEE, fournisseur["nom"],
                           f"Le poids d'une pièce semble être {resultat.nouveau_e} kg, d'après "
                           "l'emballage (conditionnement) ou la désignation du fournisseur/de "
                           "l'article.")
                    total += 1
            # Poids à l'unité de référence DÉJÀ rempli (pas vide) : jamais recalculé ni vérifié
            # par le moteur partagé (il part du principe que la valeur existante est bonne) —
            # reconsidere_poids_piece_existant() comble ce manque en le confrontant à ce que le
            # conditionnement/la désignation suggèrent indépendamment.
            couleur_e = getattr(resultat, "couleur_e", None)
            commentaire_e = getattr(resultat, "commentaire_e", None)
            if col_poids_ref and couleur_e and commentaire_e:
                marque(ws.cell(row=ligne, column=col_poids_ref), couleur_e,
                       fournisseur["nom"], commentaire_e)
                total += 1

        nouveau_f = getattr(resultat, "nouveau_f", None)
        if col_nb_pieces and nouveau_f is not None:
            cel_f = ws.cell(row=ligne, column=col_nb_pieces)
            if cel_f.value is None:
                marque(cel_f, motor.COULEUR_UNITE_TRANSFORMEE, fournisseur["nom"],
                       f"Le nombre de pièces semble être {nouveau_f}, d'après l'emballage "
                       "(conditionnement).")
                total += 1

        d_differe = (resultat.nouveau_d is not None
                     and motor.normalise_texte(resultat.nouveau_d) != motor.normalise_texte(unite_ref))
        l_differe = (resultat.nouveau_l is not None
                     and motor.normalise_texte(resultat.nouveau_l) != motor.normalise_texte(unite_cmd))
        k_differe = (resultat.nouveau_k is not None
                     and abs(resultat.nouveau_k - prix_ref) > 0.0005)

        # VERSION "COMMENTAIRES SEULEMENT" : aucune cellule D/K/L n'est modifiée, même quand
        # le script est confiant. On se contente d'un commentaire simple, en langage courant,
        # expliquant ce qu'on pense que la valeur devrait être.
        if d_differe:
            marque(ws.cell(row=ligne, column=fournisseur["_col_unite_ref"]),
                   motor.COULEUR_UNITE_TRANSFORMEE, fournisseur["nom"],
                   f"D'après la comparaison des prix et l'emballage, l'unité de référence "
                   f"semble plutôt être '{resultat.nouveau_d}' (ici : '{unite_ref}').")
            total += 1
        if l_differe:
            # Cas particulier 'LOT DE N' : expliqué directement dans le Conditionnement
            # (jamais écrit, juste commenté), pas la peine de le répéter sur l'unité de
            # commande avec le même niveau de détail.
            valeur_l_affichee = resultat.nouveau_l.upper() if isinstance(resultat.nouveau_l, str) else resultat.nouveau_l
            m_lot = RE_LOT_DE_N_UNITE_CMD.match(valeur_l_affichee) if isinstance(valeur_l_affichee, str) else None
            if m_lot:
                n_lot = m_lot.group(1)
                texte_cond_actuel = ws.cell(row=ligne, column=fournisseur["cond"]).value
                if not (texte_cond_actuel and re.search(r"\blot\b", str(texte_cond_actuel), re.I)):
                    marque(ws.cell(row=ligne, column=fournisseur["cond"]),
                           motor.COULEUR_UNITE_TRANSFORMEE, fournisseur["nom"],
                           f"On dirait que ce produit se commande par lot de {n_lot} pièces.")
                    total += 1
                marque(ws.cell(row=ligne, column=fournisseur["unite_cmd"]),
                       motor.COULEUR_UNITE_TRANSFORMEE, fournisseur["nom"],
                       f"L'unité de commande semble être un lot de {n_lot} pièces (ici : "
                       f"'{unite_cmd}').")
            else:
                marque(ws.cell(row=ligne, column=fournisseur["unite_cmd"]),
                       motor.COULEUR_UNITE_TRANSFORMEE, fournisseur["nom"],
                       f"D'après l'emballage, l'unité de commande semble plutôt être "
                       f"'{resultat.nouveau_l}' (ici : '{unite_cmd}').")
            total += 1

        if resultat.couleur_d in COULEURS_INCOHERENCE and resultat.commentaire_d and (
                not d_differe) and resultat.nouveau_d is None:
            marque(ws.cell(row=ligne, column=fournisseur["_col_unite_ref"]),
                   motor.COULEUR_A_VERIFIER, fournisseur["nom"],
                   f"L'unité de référence ('{unite_ref}') ne semble pas correspondre à "
                   "l'emballage ou au prix — à vérifier à la main.")
            total += 1
        if k_differe:
            marque(ws.cell(row=ligne, column=fournisseur["prix_ref"]),
                   motor.COULEUR_RATIO_SUSPECT, fournisseur["nom"],
                   f"D'après la comparaison des prix, le prix de référence semble plutôt "
                   f"devoir être {round(resultat.nouveau_k, 4)}€ (ici : {prix_ref}€) — à "
                   "vérifier avant de changer quoi que ce soit.")
            total += 1
        elif resultat.couleur_k in COULEURS_INCOHERENCE and resultat.commentaire_k:
            marque(ws.cell(row=ligne, column=fournisseur["prix_ref"]),
                   motor.COULEUR_A_VERIFIER, fournisseur["nom"],
                   f"Le prix de référence ('{prix_ref}') ne semble pas correspondre à "
                   "l'emballage ou au prix de commande — à vérifier à la main.")
            total += 1
        if (not l_differe) and resultat.couleur_l in COULEURS_INCOHERENCE and resultat.commentaire_l:
            marque(ws.cell(row=ligne, column=fournisseur["unite_cmd"]),
                   motor.COULEUR_A_VERIFIER, fournisseur["nom"],
                   f"L'unité de commande ('{unite_cmd}') ne semble pas correspondre à "
                   "l'emballage — à vérifier à la main.")
            total += 1
        if resultat.couleur_m in COULEURS_INCOHERENCE and resultat.commentaire_m:
            marque(ws.cell(row=ligne, column=fournisseur["prix_cmd"]),
                   motor.COULEUR_A_VERIFIER, fournisseur["nom"],
                   f"Le prix de commande ('{prix_cmd}') ne semble pas correspondre à "
                   "l'emballage ou au prix de référence — à vérifier à la main.")
            total += 1

        # Signalement des prix "outliers" (comparés à la médiane des autres lignes/catégories)
        # désactivé à la demande — seule la cohérence interne de la ligne reste vérifiée.
        # if indice_prix:
        #     base_final = _base_comparable(resultat.nouveau_d or unite_ref)
        #     k_final = resultat.nouveau_k if resultat.nouveau_k is not None else prix_ref
        #     if base_final in ("kg", "l") and k_final and k_final > 0:
        #         def signale_prix_outlier(stats_groupe, seuil, portee, origine):
        #             mediane, n_observations = stats_groupe
        #             ratio = k_final / mediane
        #             if ratio <= seuil and ratio >= (1 / seuil):
        #                 return 0
        #             comparaison = (f"×{round(ratio, 1)}" if ratio >= 1
        #                             else f"÷{round(1 / ratio, 1)}")
        #             marque(ws.cell(row=ligne, column=fournisseur["prix_ref"]),
        #                    motor.COULEUR_PRIX_OUTLIER, fournisseur["nom"],
        #                    f"Prix de référence ({k_final} / {base_final}) très éloigné de "
        #                    f"{portee} ({n_observations} lignes comparables{origine}) : médiane "
        #                    f"{round(mediane, 4)} / {base_final}, soit un facteur {comparaison}. "
        #                    "Peut être une erreur de saisie (zéro en trop/en moins, confusion "
        #                    "d'unité...) ou un vrai écart de prix — à vérifier.")
        #             return 1
        #
        #         stats_produit = indice_prix["produit"].get(
        #             (base_final, motor.normalise_pour_doublon(designation)))
        #         if stats_produit:
        #             total += signale_prix_outlier(
        #                 stats_produit, 3, f"celui des autres lignes de '{designation}'", "")
        #
        #         if lot_valor:
        #             stats_categorie = indice_prix["categorie"].get(
        #                 (base_final, motor.normalise_pour_doublon(lot_valor)))
        #             if stats_categorie:
        #                 total += signale_prix_outlier(
        #                     stats_categorie, 5, f"celui de la catégorie '{lot_valor}'",
        #                     f", catégorie '{lot_valor}'")
    # else : prix manquant(s), aucune analyse possible — rien à faire ni à signaler ici.

    # Si l'unité de commande vient d'être corrigée automatiquement (mot repris du
    # conditionnement, contenant à double imbrication, 'LOT DE N'...), valider contre la
    # NOUVELLE valeur, pas l'ancienne — sinon on re-signale à tort ce qui vient justement
    # d'être corrigé.
    unite_cmd_a_verifier = (resultat.nouveau_l
                             if resultat is not None and resultat.nouveau_l is not None
                             else unite_cmd)
    if unite_cmd_a_verifier:
        ok, message = verifie_unite_commande_etendue(unite_cmd_a_verifier, conditionnement,
                                                       designation)
        if not ok:
            # Le récipient nommé dans le Conditionnement reste fiable même quand la quantité
            # ne peut pas être confirmée (ex: 'CARTON 10 POCHE', aucun poids donné) — on le
            # reprend directement, sauf quand l'unité de commande actuelle est déjà KG/L
            # (commander au poids reste un choix valable, voir plus haut dans cette fonction).
            base_cmd_actuelle = motor.normalise_texte(unite_cmd_a_verifier)
            mot_trouve = (None if base_cmd_actuelle in ("kg", "l", "litre")
                          else mot_contenant_dans_conditionnement(conditionnement))
            abrege = motor.CONTENANT_ABREGE.get(mot_trouve, mot_trouve.upper()) if mot_trouve else None
            cel_l = ws.cell(row=ligne, column=fournisseur["unite_cmd"])
            if abrege and motor.normalise_texte(abrege) != base_cmd_actuelle:
                marque(cel_l, motor.COULEUR_UNITE_TRANSFORMEE, fournisseur["nom"],
                       f"D'après l'emballage, l'unité de commande semble plutôt être "
                       f"'{abrege}'.")
            else:
                marque(cel_l, motor.COULEUR_A_VERIFIER, fournisseur["nom"],
                       "L'unité de commande ne semble pas correspondre à l'emballage — à "
                       "vérifier à la main.")
            total += 1

    col_origine = fournisseur.get("origine")
    if col_origine:
        valeur_origine = ws_valeurs.cell(row=ligne, column=col_origine).value
        ok, message = verifie_origine(valeur_origine)
        if not ok:
            marque(ws.cell(row=ligne, column=col_origine), motor.COULEUR_A_VERIFIER,
                   fournisseur["nom"], message)
            total += 1

    # Bio/Label détectés depuis la Désignation (article ou fournisseur) quand mentionnés
    # explicitement — ne remplit que des cellules vides, ne déduit jamais rien. Une cellule
    # vide ne veut donc dire QUE "non détecté dans le texte", pas "confirmé non bio/non
    # labellisé" — colorée (comme les autres remplissages automatiques) mais sans commentaire.
    def _case_vide(valeur):
        # Une case vidée au clavier garde parfois une chaîne vide ('') plutôt que None — à
        # traiter comme vide elle aussi, sinon le remplissage automatique ne se déclenche
        # jamais sur ces cases-là (ex: 'CE2' jamais rempli dans Label alors que la case
        # affichait vide, parce qu'elle contenait '' et pas None).
        return valeur is None or (isinstance(valeur, str) and not valeur.strip())

    texte_bio_label = f"{designation or ''} {desig_fourn or ''}"
    col_bio = fournisseur.get("bio")
    if col_bio:
        cel_bio = ws.cell(row=ligne, column=col_bio)
        if (_case_vide(cel_bio.value) and RE_BIO_DETECTE.search(texte_bio_label)
                and not RE_BIO_NEGATION.search(texte_bio_label)):
            cel_bio.value = "X"
            cel_bio.fill = PatternFill(start_color=motor.COULEUR_UNITE_TRANSFORMEE,
                                        end_color=motor.COULEUR_UNITE_TRANSFORMEE, fill_type="solid")
    col_label = fournisseur.get("label")
    if col_label:
        cel_label = ws.cell(row=ligne, column=col_label)
        if _case_vide(cel_label.value):
            m_label = RE_LABEL_DETECTE.search(texte_bio_label)
            if m_label:
                cel_label.value = m_label.group().upper()
                cel_label.fill = PatternFill(start_color=motor.COULEUR_UNITE_TRANSFORMEE,
                                              end_color=motor.COULEUR_UNITE_TRANSFORMEE, fill_type="solid")

    if reference and RE_REFERENCES_MULTIPLES.search(str(reference)):
        marque(ws.cell(row=ligne, column=fournisseur["ref"]), motor.COULEUR_A_VERIFIER,
               fournisseur["nom"],
               f"Plusieurs références semblent présentes dans cette cellule ('{reference}'). "
               "Vérifiez laquelle est la bonne, ou s'il s'agit de plusieurs produits "
               "différents, séparez-les en une ligne par article/référence.")
        total += 1

    return total


def detecte_lignes_dupliquees(ws_valeurs, ligne_debut, col_designation, derniere_colonne):
    """Repère les lignes STRICTEMENT dupliquées : toutes les colonnes identiques, Désignation
    de l'article comprise (tolère juste un espace en trop/la casse, via normalise_pour_doublon).
    Une même Référence + même prix mais une Désignation réellement différente (ex: 'BIO' vs
    'LR') n'est PAS une duplicata complète — ce cas reste couvert par
    signale_copies_possibles_fournisseur()/signale_designation_dupliquee(), qui ne se
    trompent pas sur un simple écart de référence/prix entre deux produits distincts. Exige un
    minimum de valeurs réellement remplies pour éviter que deux lignes très peu renseignées
    (juste Marché/Lot/Fournisseur) ne soient prises pour des doublons."""
    MINIMUM_CHAMPS_REMPLIS = 8

    colonnes_cle = list(range(1, derniere_colonne + 1))
    lignes_vues = {}
    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        designation = ws_valeurs.cell(row=ligne, column=col_designation).value
        if not designation:
            continue
        valeurs = [ws_valeurs.cell(row=ligne, column=c).value for c in colonnes_cle]
        if sum(1 for v in valeurs if v not in (None, "")) < MINIMUM_CHAMPS_REMPLIS:
            continue
        cle = tuple(motor.normalise_pour_doublon(v) for v in valeurs)
        lignes_vues.setdefault(cle, []).append(ligne)

    doublons = {}
    for lignes in lignes_vues.values():
        if len(lignes) < 2:
            continue
        for ligne in lignes:
            doublons[ligne] = ([l for l in lignes if l != ligne], True)
    return doublons


def marque_lignes_dupliquees(ws, doublons, col_designation):
    total = 0
    for ligne, (autres, _designations_identiques) in doublons.items():
        premiere_occurrence = min([ligne] + autres)
        if ligne == premiere_occurrence:
            continue

        fill_rouge = PatternFill(start_color=COULEUR_DUPLICATA, end_color=COULEUR_DUPLICATA,
                                  fill_type="solid")
        # 'ws.max_column' scanne TOUTES les cellules déjà touchées de la feuille à chaque
        # accès (propriété non mise en cache d'openpyxl) — le ressortir du 'for ligne'
        # transformerait un simple O(lignes dupliquées) en O(lignes × cellules de la feuille),
        # déjà vu provoquer un ralentissement de plusieurs dizaines de secondes par fichier.
        derniere_colonne = ws.max_column
        for colonne in range(1, derniere_colonne + 1):
            ws.cell(row=ligne, column=colonne).fill = fill_rouge

        cel = ws.cell(row=ligne, column=col_designation)
        texte = (f"Ligne strictement dupliquée (contenu identique) de la ligne "
                 f"{premiere_occurrence}. Toute la ligne est surlignée en rouge — ce "
                 "script ne supprime rien, à vous de choisir laquelle garder.")
        if cel.comment:
            cel.comment.text = cel.comment.text + "\n\n" + texte
        else:
            cel.comment = Comment(texte, AUTEUR_COMMENTAIRE)
        total += 1
    return total


def signale_designation_dupliquee(ws, ws_valeurs, ligne_debut, col_designation, fournisseurs,
                                   lignes_a_exclure, col_lot=None):
    """Signale les lignes qui partagent exactement la même Désignation de l'article qu'une
    autre ligne, sur deux plans indépendants (une ligne peut recevoir les deux signalements) :
      1. Au moins un même fournisseur (même produit décrit deux fois chez le MÊME fournisseur —
         pas besoin que le reste de la ligne diffère aussi, juste que ce ne soit pas déjà un
         doublon complet). Même désignation mais fournisseurs totalement différents n'est PAS
         signalé pour ce premier cas : avoir le même produit chez différents fournisseurs est
         normal, pas une entrée en double.
      2. Un Lot DIFFÉRENT (quand renseigné des deux côtés) : deux lots distincts ne devraient
         pas décrire le même produit avec un texte rigoureusement identique — il manque
         forcément une information distinctive (variété, poids, origine, label...) dans au
         moins l'une des deux désignations, sinon rien ne permet de les différencier.
    'lignes_a_exclure' : lignes qui font partie d'un doublon COMPLET (toute la ligne identique,
    pas seulement la désignation) — déjà repeintes en rouge par ailleurs, elles ne comptent
    pas ici."""
    groupes = {}
    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        if ligne in lignes_a_exclure:
            continue
        designation = ws_valeurs.cell(row=ligne, column=col_designation).value
        if designation:
            groupes.setdefault(motor.normalise_pour_doublon(designation), []).append(ligne)

    def noms_fournisseurs(ligne):
        noms = set()
        for f in fournisseurs:
            valeur = ws_valeurs.cell(row=ligne, column=f["cond"] - 1).value
            if valeur:
                noms.add(motor.normalise_pour_doublon(valeur))
        return noms

    def lot_de(ligne):
        if not col_lot:
            return None
        valeur = ws_valeurs.cell(row=ligne, column=col_lot).value
        return str(valeur).strip() if valeur else None

    total = 0
    for lignes in groupes.values():
        if len(lignes) < 2:
            continue
        noms_par_ligne = {l: noms_fournisseurs(l) for l in lignes}
        lots_par_ligne = {l: lot_de(l) for l in lignes}
        for ligne in lignes:
            marque_ici = False

            autres_fournisseur = [l for l in lignes if l != ligne
                                   and (noms_par_ligne[l] & noms_par_ligne[ligne])]
            if autres_fournisseur:
                marque(ws.cell(row=ligne, column=col_designation), motor.COULEUR_COPIE_POSSIBLE,
                       None,
                       f"Même Désignation de l'article que la ligne "
                       f"{' et '.join(str(l) for l in autres_fournisseur)} (même fournisseur) — "
                       "à comparer, peut être une vraie variante ou une entrée en double.")
                marque_ici = True

            if lots_par_ligne[ligne]:
                autres_lot = [(l, lots_par_ligne[l]) for l in lignes
                              if l != ligne and lots_par_ligne[l]
                              and lots_par_ligne[l] != lots_par_ligne[ligne]]
                if autres_lot:
                    marque(ws.cell(row=ligne, column=col_designation), motor.COULEUR_RATIO_SUSPECT,
                           None,
                           "Même Désignation de l'article que la ligne "
                           f"{' et '.join(str(l) for l, _ in autres_lot)}, mais un Lot différent "
                           f"({' , '.join(sorted({lo for _, lo in autres_lot}))} au lieu de "
                           f"'{lots_par_ligne[ligne]}') — deux lots distincts ne peuvent pas "
                           "partager une désignation rigoureusement identique : il manque une "
                           "information distinctive (variété, poids, origine, label...) dans au "
                           "moins l'une des deux lignes.")
                    marque_ici = True

            if marque_ici:
                total += 1
    return total


def signale_champs_obligatoires_optionnels(ws, ligne_debut, col_marche, col_lot,
                                            col_designation, col_unite_ref, fournisseurs,
                                            lignes_a_supprimer):
    """Pour chaque ligne qui contient AU MOINS une donnée (une ligne totalement vide est déjà
    gérée ailleurs comme candidate à la suppression, et une ligne entièrement repeinte en rouge
    — doublon complet ou résiduelle sans désignation ni prix — est ignorée ici via
    'lignes_a_supprimer', pour ne pas recouvrir partiellement ce rouge) :
      - signale en orange les champs OBLIGATOIRES vides (Marché, Lot, Désignation de
        l'article, Unité de référence, et par bloc Fournisseur : le nom du fournisseur, l'Unité
        de commande, le Prix de référence, le Prix de commande) — sans eux, l'intégration dans
        la base ne sera pas possible ;
      - signale en gris les champs OPTIONNELS vides, par bloc Fournisseur : Conditionnement,
        Désignation fournisseur, Référence produit — conseillés, pas bloquants.
    Un bloc Fournisseur N entièrement vide n'est jamais signalé : il n'a simplement pas
    participé à cette ligne (le bloc Fournisseur 2/3 n'existe que pour comparaison avec le 1,
    pas un champ à remplir systématiquement). Lit l'état FINAL des cellules (ws, pas ws_valeurs)
    : exécuté en tout dernier, après tout remplissage automatique (Marché/Lot recopiés, Unité de
    référence déduite...), pour ne signaler que ce qui reste réellement vide."""
    def vide(valeur):
        return valeur is None or (isinstance(valeur, str) and valeur.strip() == "")

    total = 0

    def marque_champ(ligne, colonne, obligatoire):
        nonlocal total
        if obligatoire:
            marque(ws.cell(row=ligne, column=colonne), COULEUR_CHAMP_OBLIGATOIRE_VIDE, None,
                   "Champ obligatoire vide : sans cette information, l'intégration ne sera pas "
                   "possible.")
        else:
            marque(ws.cell(row=ligne, column=colonne), motor.COULEUR_A_VERIFIER, None,
                   "Champ optionnel vide : il est conseillé de le renseigner pour une "
                   "meilleure intégration de vos données.")
        total += 1

    for ligne in range(ligne_debut, ws.max_row + 1):
        if ligne in lignes_a_supprimer:
            continue

        marche_val = ws.cell(row=ligne, column=col_marche).value if col_marche else None
        lot_val = ws.cell(row=ligne, column=col_lot).value if col_lot else None
        designation_val = ws.cell(row=ligne, column=col_designation).value
        unite_ref_val = ws.cell(row=ligne, column=col_unite_ref).value

        blocs = []
        for f in fournisseurs:
            valeurs = {
                "nom": ws.cell(row=ligne, column=f["cond"] - 1).value,
                "cond": ws.cell(row=ligne, column=f["cond"]).value,
                "desig_fourn": ws.cell(row=ligne, column=f["desig_fourn"]).value,
                "ref": ws.cell(row=ligne, column=f["ref"]).value,
                "prix_ref": ws.cell(row=ligne, column=f["prix_ref"]).value,
                "unite_cmd": ws.cell(row=ligne, column=f["unite_cmd"]).value,
                "prix_cmd": ws.cell(row=ligne, column=f["prix_cmd"]).value,
            }
            blocs.append((f, valeurs))

        ligne_a_des_donnees = (
            not vide(marche_val) or not vide(lot_val) or not vide(designation_val)
            or not vide(unite_ref_val)
            or any(any(not vide(v) for v in valeurs.values()) for _, valeurs in blocs))
        if not ligne_a_des_donnees:
            continue  # ligne totalement vide : gérée ailleurs (candidate à la suppression)

        if col_marche and vide(marche_val):
            marque_champ(ligne, col_marche, True)
        if col_lot and vide(lot_val):
            marque_champ(ligne, col_lot, True)
        if vide(designation_val):
            marque_champ(ligne, col_designation, True)
        if vide(unite_ref_val):
            marque_champ(ligne, col_unite_ref, True)

        for f, valeurs in blocs:
            if not any(not vide(v) for v in valeurs.values()):
                continue  # bloc Fournisseur entièrement vide : n'a pas participé, on l'ignore
            if vide(valeurs["nom"]):
                marque_champ(ligne, f["cond"] - 1, True)
            if vide(valeurs["unite_cmd"]):
                marque_champ(ligne, f["unite_cmd"], True)
            if vide(valeurs["prix_ref"]):
                marque_champ(ligne, f["prix_ref"], True)
            if vide(valeurs["prix_cmd"]):
                marque_champ(ligne, f["prix_cmd"], True)
            if vide(valeurs["cond"]):
                marque_champ(ligne, f["cond"], False)
            if vide(valeurs["desig_fourn"]):
                marque_champ(ligne, f["desig_fourn"], False)
            if vide(valeurs["ref"]):
                marque_champ(ligne, f["ref"], False)
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
                marque(ws.cell(row=ligne, column=col_marche), motor.COULEUR_UNITE_TRANSFORMEE,
                       None,
                       f"Les lignes juste au-dessus et en-dessous ont le même Marché "
                       f"('{marche_dessus}') — celui-ci devrait sans doute être le même.")
            if col_lot and lot_vide:
                marque(ws.cell(row=ligne, column=col_lot), motor.COULEUR_UNITE_TRANSFORMEE,
                       None,
                       f"Les lignes juste au-dessus et en-dessous ont le même Lot "
                       f"('{lot_dessus}') — celui-ci devrait sans doute être le même.")
            total_rempli += 1
        else:
            if col_marche and marche_vide:
                marque(ws.cell(row=ligne, column=col_marche), motor.COULEUR_A_VERIFIER, None,
                       "Le Marché est vide, et les lignes voisines ne permettent pas de "
                       "deviner lequel mettre — à préciser à la main.")
            if col_lot and lot_vide:
                marque(ws.cell(row=ligne, column=col_lot), motor.COULEUR_A_VERIFIER, None,
                       "Le Lot est vide, et les lignes voisines ne permettent pas de deviner "
                       "lequel mettre — à préciser à la main.")
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
    # 'ws.max_column' n'est pas mis en cache par openpyxl : le ressortir du 'for ligne'
    # recalculerait la borne à chaque ligne en reparcourant toutes les cellules déjà touchées
    # de la feuille (O(lignes × cellules) au lieu de O(lignes × colonnes) — c'était la cause
    # principale d'un ralentissement de plusieurs dizaines de secondes par fichier).
    derniere_ligne = ws.max_row
    derniere_colonne = ws.max_column
    for ligne in range(ligne_debut, derniere_ligne + 1):
        for colonne in range(1, derniere_colonne + 1):
            cel = ws.cell(row=ligne, column=colonne)
            cel.font = fonte
            wrap_text = cel.alignment.wrap_text if cel.alignment else None
            cel.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap_text)


def desocculte_lignes_et_colonnes(ws):
    """Réaffiche toutes les lignes ET colonnes masquées — résidu fréquent d'un filtre
    automatique (lignes) ou d'un masquage manuel de colonne (ex: une colonne 'Désignation
    fournisseur' repliée pour gagner de la place à l'écran), tous deux laissés actifs à
    l'enregistrement. Une ligne/colonne masquée contient souvent une donnée bien réelle (pas
    vide), invisible à la relecture si elle n'est pas réaffichée. Purement cosmétique : ne
    touche à aucune donnée, et ne modifie pas les boutons de filtre eux-mêmes (une personne peut
    toujours filtrer/masquer à nouveau si besoin)."""
    total = 0
    for dim in ws.row_dimensions.values():
        if dim.hidden:
            dim.hidden = False
            total += 1
    for dim in ws.column_dimensions.values():
        if dim.hidden:
            dim.hidden = False
            total += 1
    return total


def normalise_lots(ws, ws_valeurs, ligne_debut, col_lot):
    if not col_lot:
        return 0
    total = 0
    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        valeur = ws_valeurs.cell(row=ligne, column=col_lot).value
        nouvelle_valeur = motor.normalise_numero_lot(valeur)
        if nouvelle_valeur != valeur:
            marque(ws.cell(row=ligne, column=col_lot), motor.COULEUR_UNITE_TRANSFORMEE, None,
                   f"Le numéro de lot s'écrit normalement avec un zéro devant quand il n'a "
                   f"qu'un chiffre : '{nouvelle_valeur}' au lieu de '{valeur}'.")
            total += 1
    return total


TITRES_CATEGORIES_RAPPORT = {
    "doublons": "Lignes strictement dupliquées (ou candidates à la suppression)",
    "references_doublon": "Références en double (même référence + même prix)",
    "designation_dupliquee": "Même désignation qu'une autre ligne (même fournisseur, ou Lot différent)",
    "champs_obligatoires_vides": "Champs obligatoires vides (bloquent l'intégration)",
    "champs_optionnels_vides": "Champs optionnels vides (conseillés)",
    "unite_ref": "Unité de référence incohérente",
    "prix_ref": "Prix de référence (BPU) incohérent / suggéré",
    "unite_cmd": "Unité de commande incohérente",
    "prix_cmd": "Prix de commande incohérent / inexpliqué",
    "conditionnement": "Conditionnement — signalement ou complément",
    "poids_rempli": "Poids à l'unité de référence — suggestion (rien n'est rempli)",
    "origine": "Origine non reconnue / ambiguë",
}


def calcule_categories_rapport(ws, ws_valeurs, ligne_debut, col_designation, col_unite_ref,
                                col_poids_ref, fournisseurs):
    """Relit les commentaires posés sur la feuille et les classe par catégorie — séparé de
    genere_rapport() pour que l'application Streamlit puisse aussi construire sa propre
    présentation (titres + compteurs, détails repliés) à partir des mêmes données.
    Les champs obligatoires/optionnels vides (signale_champs_obligatoires_optionnels) sont
    repérés par leur TEXTE en premier, avant tout classement par colonne : ce signalement peut
    tomber sur n'importe quelle colonne concernée (Marché, Lot, nom du Fournisseur, Désignation
    fournisseur...), dont certaines n'ont par ailleurs aucune autre catégorie dédiée."""
    categories = {cle: [] for cle in TITRES_CATEGORIES_RAPPORT}

    colonnes_ref = {f["ref"] for f in fournisseurs}
    colonnes_prix_ref = {f["prix_ref"] for f in fournisseurs}
    colonnes_unite_cmd = {f["unite_cmd"] for f in fournisseurs}
    colonnes_prix_cmd = {f["prix_cmd"] for f in fournisseurs}
    colonnes_cond = {f["cond"] for f in fournisseurs}
    colonnes_origine = {f["origine"] for f in fournisseurs if f.get("origine")}

    # Mis en cache hors de la boucle : voir la note dans applique_formatage_uniforme (même
    # piège de performance avec 'ws.max_column'/'ws.max_row' non mis en cache par openpyxl).
    derniere_ligne = ws.max_row
    derniere_colonne = ws.max_column
    for ligne in range(ligne_debut, derniere_ligne + 1):
        designation = ws_valeurs.cell(row=ligne, column=col_designation).value
        for colonne in range(1, derniere_colonne + 1):
            cel = ws.cell(row=ligne, column=colonne)
            if not cel.comment:
                continue
            texte = cel.comment.text

            if "Champ obligatoire vide" in texte:
                categories["champs_obligatoires_vides"].append((ligne, designation, texte))
            elif "Champ optionnel vide" in texte:
                categories["champs_optionnels_vides"].append((ligne, designation, texte))
            elif colonne == col_designation and (
                    "dupliquée" in texte or "probablement à supprimer" in texte):
                categories["doublons"].append((ligne, designation, texte))
            elif colonne == col_designation and (
                    "(même fournisseur)" in texte or "Lot différent" in texte):
                categories["designation_dupliquee"].append((ligne, designation, texte))
            elif colonne in colonnes_ref and "Copie possible" in texte:
                categories["references_doublon"].append((ligne, designation, texte))
            elif colonne == col_unite_ref:
                categories["unite_ref"].append((ligne, designation, texte))
            elif col_poids_ref and colonne == col_poids_ref:
                categories["poids_rempli"].append((ligne, designation, texte))
            elif colonne in colonnes_prix_ref:
                categories["prix_ref"].append((ligne, designation, texte))
            elif colonne in colonnes_unite_cmd:
                categories["unite_cmd"].append((ligne, designation, texte))
            elif colonne in colonnes_prix_cmd:
                categories["prix_cmd"].append((ligne, designation, texte))
            elif colonne in colonnes_cond:
                categories["conditionnement"].append((ligne, designation, texte))
            elif colonne in colonnes_origine:
                categories["origine"].append((ligne, designation, texte))

    return categories


def calcule_valeurs_uniques(ws_valeurs, ligne_debut, colonne):
    """Compte les valeurs EXACTEMENT distinctes d'une colonne (espaces en début/fin compris —
    aucune normalisation) : deux cellules qui semblent identiques à l'œil mais diffèrent par un
    espace en trop restent deux valeurs séparées ici, justement pour que ça se voie. Renvoie une
    liste de (valeur, longueur, nombre d'occurrences), triée alphabétiquement pour que deux
    variantes très proches (ex: avec/sans espace final) se retrouvent l'une à côté de l'autre."""
    if not colonne:
        return []
    compte = {}
    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        valeur = ws_valeurs.cell(row=ligne, column=colonne).value
        if valeur is None or str(valeur) == "":
            continue
        valeur = str(valeur)
        compte[valeur] = compte.get(valeur, 0) + 1
    return sorted(((v, len(v), n) for v, n in compte.items()), key=lambda t: t[0])


def genere_rapport(ws, ws_valeurs, nom_feuille, ligne_debut, col_designation, col_unite_ref,
                    col_poids_ref, fournisseurs, categories=None, col_lot=None):
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

    # Valeurs uniques de Lot, en premier et triées par ordre CROISSANT du numéro de lot (pas
    # alphabétique : sinon '12.' passerait avant '4.') — espaces comptés, aucune normalisation,
    # pour repérer à l'œil des variantes qui devraient être identiques mais ne le sont pas
    # exactement (espace en trop, casse différente...).
    valeurs_lot = calcule_valeurs_uniques(ws_valeurs, ligne_debut, col_lot)

    def _cle_tri_lot(item):
        valeur = item[0]
        m = re.match(r"\s*(\d+)", valeur)
        return (int(m.group(1)), valeur) if m else (float("inf"), valeur)

    valeurs_lot.sort(key=_cle_tri_lot)
    lignes_sortie.append(f"## Valeurs uniques — Lot ({len(valeurs_lot)}, espaces comptés)")
    lignes_sortie.append("")
    if not valeurs_lot:
        lignes_sortie.append("_(aucune colonne Lot repérée)_")
    else:
        lignes_sortie.append("| Valeur | Longueur | Occurrences |")
        lignes_sortie.append("|---|---|---|")
        for valeur, longueur, occurrences in valeurs_lot:
            lignes_sortie.append(f"| `{valeur}` | {longueur} | {occurrences} |")
    lignes_sortie.append("")

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

    for f in fournisseurs:
        valeurs_fourn = calcule_valeurs_uniques(ws_valeurs, ligne_debut, f["cond"] - 1)
        lignes_sortie.append(
            f"## Valeurs uniques — {f['nom']} ({len(valeurs_fourn)}, espaces comptés)")
        lignes_sortie.append("")
        if not valeurs_fourn:
            lignes_sortie.append("_(aucune valeur)_")
        else:
            lignes_sortie.append("| Valeur | Longueur | Occurrences |")
            lignes_sortie.append("|---|---|---|")
            for valeur, longueur, occurrences in valeurs_fourn:
                lignes_sortie.append(f"| `{valeur}` | {longueur} | {occurrences} |")
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


def derive_unite_ref_vide(ws_valeurs, ligne, fournisseurs):
    """Quand l'unité de référence est vide, tente de la déduire d'un Conditionnement propre
    (motif reconnu, pas une recherche floue) dont la quantité est EXACTEMENT confirmée par le
    ratio Prix de commande ÷ Prix de référence d'un fournisseur rempli. Si plusieurs
    fournisseurs sont remplis et se contredisent (bases différentes), on ne déduit rien —
    mieux vaut laisser vide à compléter que de deviner sur un désaccord."""
    bases_confirmees = set()
    for f in fournisseurs:
        conditionnement = ws_valeurs.cell(row=ligne, column=f["cond"]).value
        prix_ref = ws_valeurs.cell(row=ligne, column=f["prix_ref"]).value
        prix_cmd = ws_valeurs.cell(row=ligne, column=f["prix_cmd"]).value
        if not conditionnement or not prix_ref or not prix_cmd:
            continue
        info = determine_base_et_quantite(nettoie_conditionnement(conditionnement))
        if not info:
            continue
        base, quantite, _mot = info
        if not quantite or quantite <= 0:
            continue
        try:
            ratio = prix_cmd / prix_ref
        except (TypeError, ZeroDivisionError):
            continue
        if abs(ratio - quantite) <= motor.TOLERANCE_RATIO * quantite:
            bases_confirmees.add(base)
    if len(bases_confirmees) == 1:
        return bases_confirmees.pop()
    return None


def traite_feuille(ws, ws_valeurs, info, indice_prix=None):
    gele_formules_externes(ws, ws_valeurs)
    ligne_debut = info["ligne_debut_donnees"]
    col_marche = info["col_marche"]
    col_lot = info["col_lot"]
    col_designation = info["col_designation"]
    col_unite_ref = info["col_unite_ref"]
    col_poids_ref = info["col_poids_ref"]
    col_nb_pieces = info["col_nb_pieces"]
    fournisseurs = info["fournisseurs"]
    for f in fournisseurs:
        f["_col_unite_ref"] = col_unite_ref  # utilisé par audite_fournisseur

    # Colonnes de prix toujours au même format monétaire (2 décimales, "€" à la fin), même
    # quand la valeur n'a pas été modifiée par l'audit — les fichiers sources mélangent des
    # formats différents d'une ligne à l'autre.
    for ligne in range(ligne_debut, ws_valeurs.max_row + 1):
        for f in fournisseurs:
            ws.cell(row=ligne, column=f["prix_ref"]).number_format = '0.00"€"'
            ws.cell(row=ligne, column=f["prix_cmd"]).number_format = '0.00"€"'

    stats = {"lignes_avec_incoherence": 0, "marques": 0, "unite_ref_vide": 0,
              "lignes_dupliquees": 0, "copies_possibles": 0, "lots_normalises": 0,
              "marche_lot_rempli": 0, "marche_lot_signale": 0, "lignes_vides_a_supprimer": 0,
              "lignes_demasquees": desocculte_lignes_et_colonnes(ws)}

    # Calculé tôt (pas seulement en fin de fonction) pour que signale_designation_dupliquee()
    # puisse exclure les lignes déjà couvertes par le doublon de ligne COMPLÈTE (repeint en
    # rouge plus bas, avec ce même dict). Note : la détection par Référence+Conditionnement+
    # Prix (en ignorant la désignation) n'alimente plus ce rouge — elle faisait doublon avec
    # signale_copies_possibles_fournisseur()/signale_designation_dupliquee() (rose), qui
    # couvrent déjà ce cas sans supposer à tort que la désignation ne diffère que légèrement
    # (ex: 'BIO' vs 'LR' sont des produits différents, pas une faute de frappe).
    doublons = detecte_lignes_dupliquees(ws_valeurs, ligne_debut, col_designation,
                                         info["derniere_colonne_cle_doublon"])

    stats["marche_lot_rempli"], stats["marche_lot_signale"] = comble_marche_lot(
        ws, ws_valeurs, ligne_debut, col_designation, col_marche, col_lot)

    stats["lots_normalises"] = normalise_lots(ws, ws_valeurs, ligne_debut, col_lot)

    for fournisseur in fournisseurs:
        stats["copies_possibles"] += signale_copies_possibles_fournisseur(
            ws, ws_valeurs, ligne_debut, fournisseur)

    stats["copies_possibles"] += signale_designation_dupliquee(
        ws, ws_valeurs, ligne_debut, col_designation, fournisseurs, set(doublons), col_lot)

    lignes_a_supprimer = set()
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
                lignes_a_supprimer.add(ligne)
            continue

        unite_ref = ws_valeurs.cell(row=ligne, column=col_unite_ref).value
        poids_ref = ws_valeurs.cell(row=ligne, column=col_poids_ref).value if col_poids_ref else None
        lot_valor = ws_valeurs.cell(row=ligne, column=col_lot).value if col_lot else None

        # "Poids à l'unité de référence" écrit en texte libre (ex: 'seau de 5 kg', '100 g')
        # au lieu d'un nombre : réutilise le même parseur que pour le Conditionnement (même
        # vocabulaire de récipients) pour en extraire la valeur numérique. Sans ça, le texte
        # est traité comme "déjà rempli" (il n'est pas vide), ce qui masquait silencieusement
        # une vraie incohérence de type.
        poids_texte_marque = False
        if poids_ref and not isinstance(poids_ref, (int, float)):
            resultat_texte = determine_base_et_quantite(nettoie_conditionnement(str(poids_ref)))
            cel_poids = ws.cell(row=ligne, column=col_poids_ref)
            if resultat_texte and resultat_texte[0] in ("kg", "l", "litre"):
                quantite_texte = round(resultat_texte[1], 3)
                marque(cel_poids, motor.COULEUR_UNITE_TRANSFORMEE, None,
                       f"Ce texte semble vouloir dire {quantite_texte} kg/L.")
                poids_ref = quantite_texte
            else:
                marque(cel_poids, motor.COULEUR_A_VERIFIER, None,
                       f"Poids à l'unité de référence ('{poids_ref}') est du texte, pas un "
                       "nombre, et on n'arrive pas à deviner lequel — à convertir à la main "
                       "en kg/L.")
            stats["marques"] += 1
            poids_texte_marque = True

        if not unite_ref:
            base_deduite = derive_unite_ref_vide(ws_valeurs, ligne, fournisseurs)
            if base_deduite:
                # Conditionnement + ratio de prix confirment exactement la base — mais dans
                # cette version, on ne remplit rien : on utilise cette valeur pour la suite de
                # l'analyse de la ligne SANS l'écrire, et on commente juste la case.
                unite_ref = base_deduite.upper()
                marque(ws.cell(row=ligne, column=col_unite_ref), motor.COULEUR_UNITE_TRANSFORMEE,
                       None,
                       f"D'après l'emballage et la comparaison des prix, l'unité de référence "
                       f"semble être '{unite_ref}'.")
                stats["marques"] += 1
            else:
                marque(ws.cell(row=ligne, column=col_unite_ref), motor.COULEUR_A_VERIFIER, None,
                       "Cette case est vide — sans elle, impossible de comparer le prix entre "
                       "fournisseurs.")
                stats["unite_ref_vide"] += 1
                stats["marques"] += 1
                continue

        ligne_a_un_probleme = poids_texte_marque
        for fournisseur in fournisseurs:
            n = audite_fournisseur(ws, ws_valeurs, ligne, designation, unite_ref, poids_ref,
                                    fournisseur, col_poids_ref, lot_valor, indice_prix,
                                    col_nb_pieces)
            if n:
                ligne_a_un_probleme = True
                stats["marques"] += n
        if ligne_a_un_probleme:
            stats["lignes_avec_incoherence"] += 1

    # Appliqué en dernier pour que le rouge des lignes dupliquées ne soit jamais recouvert par
    # une autre couleur posée cellule par cellule plus haut (doublons calculé plus haut).
    stats["lignes_dupliquees"] = marque_lignes_dupliquees(ws, doublons, col_designation)

    for ligne, (autres, _designations_identiques) in doublons.items():
        if ligne != min([ligne] + autres):
            lignes_a_supprimer.add(ligne)

    # Tout en dernier : lit l'état FINAL des cellules (après tout remplissage automatique
    # ci-dessus), et ignore les lignes déjà entièrement repeintes en rouge (doublon complet ou
    # résiduelle sans désignation ni prix) pour ne pas recouvrir ce rouge.
    stats["champs_obligatoires_optionnels"] = signale_champs_obligatoires_optionnels(
        ws, ligne_debut, col_marche, col_lot, col_designation, col_unite_ref, fournisseurs,
        lignes_a_supprimer)

    applique_formatage_uniforme(ws, ligne_debut)

    return stats



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
                                  info["col_poids_ref"], info["fournisseurs"], categories,
                                  info["col_lot"])
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

            total_marques = (stats["marques"] + stats["lignes_dupliquees"]
                              + stats["copies_possibles"] + stats["champs_obligatoires_optionnels"])
            print(f"    Lignes/colonnes masquées réaffichées : {stats['lignes_demasquees']}")
            print(f"    Marché/Lot recopiés (lignes voisines identiques) : {stats['marche_lot_rempli']}")
            print(f"    Marché/Lot signalés (à préciser manuellement) : {stats['marche_lot_signale']}")
            print(f"    Lignes vides (ni désignation ni prix) peintes en rouge : {stats['lignes_vides_a_supprimer']}")
            print(f"    Numéros de lot normalisés : {stats['lots_normalises']}")
            print(f"    Unité de référence vide : {stats['unite_ref_vide']}")
            print(f"    Lignes dupliquées peintes en rouge : {stats['lignes_dupliquees']}")
            print(f"    Références en double : {stats['copies_possibles']}")
            print(f"    Champs obligatoires/optionnels vides signalés : {stats['champs_obligatoires_optionnels']}")
            print(f"    Lignes avec incohérence signalée : {stats['lignes_avec_incoherence']}")
            print(f"    Total de cellules marquées : {total_marques}")

            nom_rapport = f"{nom_base} - rapport_commenté.md" if len(wb.sheetnames) <= 2 else \
                f"{nom_base} - {nom_feuille} - rapport_commenté.md"
            with open(os.path.join(DOSSIER_SORTIE, nom_rapport), "w", encoding="utf-8") as f:
                f.write(resultat["rapport"])

        chemin_sortie = os.path.join(DOSSIER_SORTIE, f"{nom_base} - audit_commenté.xlsx")
        wb.save(chemin_sortie)
        print(f"  Fichier enregistré : {chemin_sortie}")

    print("\nTerminé. Rien n'a été modifié dans les fichiers sources, à part le 'X' de Bio et "
          "le nom du label quand ils sont trouvés dans le texte. Tout le reste n'est que "
          "surlignage + commentaire en langage simple, à valider par une personne.")


if __name__ == "__main__":
    main()
