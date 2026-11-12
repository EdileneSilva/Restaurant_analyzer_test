"""
Prétraitement — fichiers reçus dans `a_formater/`, quatre familles distinctes (aucune déjà
prise en charge par pretraitement.py/REGAL ou pretraitement_grada.py/GRADA) :

- "BPU DQE" (accords-cadres Lots 1 à 5 + Epicerie) : une feuille "Page de Garde" (nom du
  fournisseur, titre du lot) + une feuille BPU avec en-têtes variables d'un lot à l'autre
  (colonnes présentes/absentes, libellés légèrement différents) — tout est donc repéré
  dynamiquement par texte d'en-tête, comme pretraitement_grada.py.
- Passion Froid et Episaveurs : tableau à plat, un seul fournisseur (le nom du fichier lui-
  même), en-têtes fixes en ligne 1.
- BPUR FRESCA : BPU avec des lignes de section au milieu des données (ex: "CONDIMENTS",
  "PRODUITS FESTIFS") qui servent de sous-lot — même logique que la famille GRADA
  (pg._lots_par_section), réutilisée telle quelle.

Les correspondances colonne par colonne ont été établies dans mapping_colonnes_a_formater.xlsx
(vérifiées fichier par fichier, pas devinées). Aucune des 9 sources ne distingue unité/prix de
RÉFÉRENCE d'unité/prix de COMMANDE (une seule colonne de prix chacune) : la commande reprend
donc toujours la référence telle quelle (voir ce classeur, lignes 15-16, vides partout).

Réutilise l'écriture canonique (pt.ecrit_classeur_canonique) et les fonctions utilitaires de
pretraitement.py et pretraitement_grada.py SANS LES MODIFIER — même convention que
pretraitement_auto.py.

Lancement CLI : lit tous les fichiers de a_formater/ (DOSSIER_ENTREE), écrit un .xlsx par
fichier dans a_formater_transforme/ (DOSSIER_SORTIE) — PAS directement dans pret_pour_audit/ :
audit_general.construit_indice_prix() agrège les prix de TOUS les .xlsx de pret_pour_audit/
pour détecter les écarts de prix anormaux, donc y déposer ces 9 fichiers sans relecture
changerait silencieusement l'indice de prix partagé et, avec lui, les résultats déjà obtenus
sur les 7 fichiers de référence (et les totaux figés dans les tests de non-régression). Une
fois chaque fichier relu et jugé correct, le déplacer à la main vers pret_pour_audit/.
"""

import os
import re

import openpyxl
from openpyxl.comments import Comment

import pretraitement as pt
import pretraitement_grada as pg
from causes_format import CauseFormatInvalide

DOSSIER_ENTREE = "a_formater"
DOSSIER_SORTIE = "a_formater_transforme"

RE_LOT_TITRE = re.compile(r"lot\s*\d+.*", re.I)

# Valeurs de remplissage qui veulent dire « aucun label » plutôt qu'un vrai nom de label —
# à ignorer avant pg._bio_et_label, sous peine d'écrire ce texte tel quel dans Label sur la
# quasi-totalité des lignes (vérifié : "non imposée" apparaît sur 314 des 388 lignes
# d'Epicerie ; "NA"/"#N/A" sur la colonne Label d'Episaveurs).
VALEURS_SANS_LABEL = {"non imposee", "non eligible", "na", "#n/a"}


def _label_ou_rien(valeur_brute):
    if valeur_brute and pt._normalise(valeur_brute) in VALEURS_SANS_LABEL:
        return None
    return valeur_brute


# pg._bio_et_label (GRADA) ne reconnaît que le sigle "AB" comme signe bio — adapté à la
# colonne EGALIM d'Episaveurs/BPUR FRESCA, mais PAS à la famille BPU DQE (Lots 1-5 +
# Epicerie), où le signe bio est le mot "Bio" écrit en toutes lettres dans la même colonne
# que les autres labels (BBC, CEE2, AOP...). Sans ce repérage, "Bio" finirait comme texte
# dans Label, sans jamais remplir la colonne Bio — et l'auto-détection de
# audit_general.py (RE_BIO_DETECTE) ne regarde QUE la Désignation/Désignation fournisseur,
# jamais une colonne Label déjà écrite : elle ne rattraperait pas l'oubli.
RE_BIO_MOT = re.compile(r"\bbio(?:logique)?\b", re.I)


def _bio_et_label_bpu(valeur_brute):
    valeur = _label_ou_rien(valeur_brute)
    if not valeur:
        return None, None
    if RE_BIO_MOT.search(str(valeur)):
        reste = RE_BIO_MOT.sub("", str(valeur)).strip(" ,;/").strip()
        return "X", (reste or None)
    return pg._bio_et_label(valeur)


def _lot_depuis_titre(texte):
    """Repère « LOT N ... » dans un texte de titre (qui peut contenir plein d'espaces de
    remplissage avant/après) et renvoie ce passage, espaces multiples réduits à un seul."""
    if not texte:
        return None
    m = RE_LOT_TITRE.search(str(texte))
    if not m:
        return None
    return re.sub(r"\s+", " ", m.group(0)).strip().upper()


RE_LOT_NUMERO = re.compile(r"lot\s*(\d+)", re.I)
RE_LOT_DESCRIPTION_SANS_NUMERO = re.compile(r"\blot\s+(?!\d)([a-zà-ÿ].{2,})", re.I)


def _lot_numero(texte):
    if not texte:
        return None
    m = RE_LOT_NUMERO.search(str(texte))
    return m.group(1) if m else None


def _lot_description_sans_numero(texte):
    """Repère « LOT <description, sans chiffre collé derrière> » — ex: la page de garde du
    Lot 5 dit "LOT DESSERTS ET GLACES" (pas de numéro), alors que la feuille BPU elle-même ne
    donne que "... - LOT 5" (numéro sans description) : aucune des deux seule ne suffit, voir
    la recombinaison faite dans transforme_bpu_dqe."""
    if not texte:
        return None
    m = RE_LOT_DESCRIPTION_SANS_NUMERO.search(str(texte))
    if not m:
        return None
    return re.sub(r"\s+", " ", m.group(1)).strip().upper()


def _feuille_page_de_garde(wb):
    for nom in wb.sheetnames:
        if "garde" in pt._normalise(nom):
            return wb[nom]
    return None


def _fournisseur_depuis_page_de_garde(wb):
    """Cherche « Nom de l'entreprise » sur la feuille de garde et renvoie la valeur trouvée
    plus loin sur la même ligne — une seule valeur par fichier (pas une colonne par produit),
    confirmé présent sur les 6 fichiers de la famille BPU DQE."""
    ws = _feuille_page_de_garde(wb)
    if ws is None:
        return None
    for r in range(1, min(20, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            if "nom de l'entreprise" in pt._normalise(ws.cell(row=r, column=c).value):
                for c2 in range(c + 1, ws.max_column + 1):
                    v = ws.cell(row=r, column=c2).value
                    if v:
                        return str(v).strip()
    return None


def _texte_titre_page_de_garde(wb):
    """Renvoie le texte de la cellule de la feuille de garde qui mentionne « lot » (ligne ~4,
    texte long du type « Accord-cadre à bons de commande pour: ... LOT N <description> ») —
    PAS la feuille "Taux de remise", qui contient une erreur de copier-coller sur au moins un
    fichier source (Lot 4 y est étiqueté "LOT 3 - LEGUMES", vérifié le 2026-10-29) : la feuille
    de garde est la source fiable. Renvoyé tel quel (pas encore découpé en numéro/description) :
    voir `_lot_depuis_titre`/`_lot_numero`/`_lot_description_sans_numero`."""
    ws = _feuille_page_de_garde(wb)
    if ws is None:
        return None
    for r in range(1, min(10, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v and re.search(r"\blot\b", str(v), re.I):
                return str(v)
    return None


def _trouve_ligne_entete_bpu(ws, max_lignes=8):
    derniere_ligne_cherchee = min(max_lignes, ws.max_row)
    for r in range(1, derniere_ligne_cherchee + 1):
        textes = [pt._normalise(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        if "unite" in textes and any(t.startswith("libell") for t in textes if t):
            return r
    raise CauseFormatInvalide(
        "ENTETE_MODIFIE", f"Ligne d'en-tête introuvable dans « {ws.title} »", onglet=ws.title,
        cellule=f"lignes 1 à {derniere_ligne_cherchee}",
    )


def _colonnes_entete(ws, ligne_entete):
    """{texte d'en-tête normalisé: [colonnes]} — une liste par texte car certains en-têtes
    sont répétés à l'identique (ex: "Format" x2 dans le fichier Epicerie, souhaité puis
    proposé) ; `_col` ci-dessous choisit alors la DERNIÈRE occurrence par défaut."""
    cols = {}
    for c in range(1, ws.max_column + 1):
        t = pt._normalise(ws.cell(row=ligne_entete, column=c).value)
        if t:
            cols.setdefault(t, []).append(c)
    return cols


def _col(cols, texte, which=-1):
    cible = pt._normalise(texte)
    lst = cols.get(cible)
    if not lst:
        lst = sorted(c for k, v in cols.items() if cible in k for c in v)
    return lst[which] if lst else None


def _derniere_ligne_numero(ws, ligne_debut, col_numero, limite=600):
    """Comme pg._derniere_ligne_avec_reference, mais pour une colonne numérique (N° Art) —
    utilisé pour borner la zone de données avant les lignes de pied de page (signature,
    date...) qui suivent parfois juste après, sans ligne vide entre les deux."""
    derniere = ligne_debut - 1
    for r in range(ligne_debut, min(ws.max_row, ligne_debut + limite) + 1):
        if isinstance(ws.cell(row=r, column=col_numero).value, (int, float)):
            derniere = r
    return derniere


def _arrondi_prix(valeur):
    """Un prix n'a jamais de sens à plus de deux décimales (centimes) — certaines sources
    arrivent avec plus de décimales (imprécision de flottant ou saisie d'origine), donc
    arrondi systématiquement avant écriture, dans les quatre familles de fichiers."""
    return round(valeur, 2) if valeur is not None else None


def _commande_miroir_reference(unite_ref, prix_ref):
    """Les 9 fichiers de a_formater/ n'ont qu'UNE SEULE colonne de prix/unité (pas de
    distinction référence/commande comme REGAL) : rien dans la source ne dit ce que serait
    réellement l'unité/le prix de commande, donc ils restent vides plutôt que d'y recopier
    la référence (qui serait une valeur inventée, pas une donnée lue) — laissés vides,
    `audit_general.py` les signalera lui-même comme champs obligatoires manquants, ce qui
    est la réalité : cette donnée n'existe simplement pas dans ces fichiers."""
    return None, None


# pt.traite_origine (pretraitement.py) ne reconnaît une valeur comme pays que si elle
# correspond EXACTEMENT à un pays connu — une valeur source avec plusieurs pays (ex: "France,
# Espagne", ou "BELGIQUE / FRANCE" chez FRESCA) ne matche rien du tout et tombe dans son cas
# par défaut, qui la remplace silencieusement par "France" avec un commentaire (vérifié sur le
# fichier Epicerie : "France, Espagne", "Portugal, Pays-Bas", "Allemagne, Suisse, France,
# Hongrie"... deviennent tous "France"). Idem pour un pays connu de l'utilisateur mais absent
# de PAYS_ETRANGERS_CONNUS (Inde, Indonésie, Equateur, USA, Thaïlande, Grèce, Rép. Tchèque).
# Remplacé ici, pour les familles BPU DQE et BPUR FRESCA, par une logique dédiée : plusieurs
# pays séparés par "," ou "/" ou " - " => "UE" si tous les pays listés sont dans l'Union
# européenne, "Hors UE" si aucun ne l'est (PAS "NON UE" : audit_general.ORIGINES_LARGES ne
# reconnaît que "hors ue" parmi les 4 grandes catégories — "non ue" aurait été rejeté comme
# origine non reconnue ; bug trouvé le 2026-10-30 lors d'un test manuel), et si mélange (un pays
# UE + un pays hors UE), les deux noms sont gardés tels quels plutôt que résumés — demandé
# explicitement le 2026-10-30 (avant : "UE" par défaut, jugé trompeur). Un seul pays reste
# inchangé (délégué à pt.traite_origine s'il
# n'est pas l'un des ajouts ci-dessous), sauf Indonésie/Thaïlande, toujours réécrits avec leur
# accent. Chez FRESCA spécifiquement, beaucoup de valeurs sont des zones de pêche ("Océan
# Pacifique (FAO 61/67)"...) et pas des pays du tout : jamais découpées par cette logique,
# repérées par les mots FAO/Océan/Atlantique/Pacifique/Indien et laissées telles quelles.
PAYS_UE = {
    "allemagne", "autriche", "belgique", "bulgarie", "chypre", "croatie", "danemark",
    "espagne", "estonie", "finlande", "france", "grece", "hollande", "hongrie", "irlande",
    "italie", "lettonie", "lituanie", "luxembourg", "malte", "pays bas", "pologne",
    "portugal", "rep tcheque", "republique tcheque", "tchequie", "roumanie", "slovaquie",
    "slovenie", "suede",
}
PAYS_HORS_UE = {
    "suisse", "chine", "russie", "cote divoire", "equateur", "el salvador", "ile maurice",
    "turquie", "royaume uni", "uk", "ecosse", "chili", "perou", "egypte", "asie", "maroc",
    "inde", "indonesie", "thailande", "usa", "bresil", "mexique", "vietnam", "serbie",
    "ukraine", "norvege", "liechtenstein", "nouvelle zelande", "macedoine", "bosnie herzegovine",
    "afrique du sud", "bolivie", "tunisie",
}

RE_ZONE_PECHE = re.compile(r"\bfao\b|\bocean\b|\boceans\b|\batlantique\b|\bpacifique\b|\bindien\b", re.I)
RE_PARENTHESE = re.compile(r"\s*\([^)]*\)\s*")

# Codes de région/département français vus dans "Origine matière première" (famille MP
# ARDECHE 24, marché départemental) — pas des pays, mais pas non plus "non reconnu" : ce sont
# des valeurs françaises, donc équivalentes à "France" au sens du format cible (`Origine` n'a
# pas de valeur "région"). Vérifié le 2026-10-29 sur les 4 premiers fichiers ("ARDECHE", "AURA",
# "ISERE", "SAVOIE/HAUTE-SAVOIE", "REUNION") puis étendu le 2026-10-30 avec les 8 fichiers
# supplémentaires (Agricourt, DADS x3, Passion Froid x2, SYSCO FRANCE, Vivalya), qui apportent
# plusieurs départements de plus (AIN, DROME, RHONE, VAUCLUSE, BOUCHE(S) DU RHONE, HAUTE LOIRE)
# et une valeur composée SANS séparateur ("DROME ARDECHE", contrairement à "DROME/ARDECHE"
# ailleurs dans les mêmes fichiers — incohérence de saisie de la source, pas corrigée, juste
# reconnue telle quelle comme un 2e jeton possible). Singulier/pluriel ("BOUCHE"/"BOUCHES") et
# accents (RHÔNE/RHÖNE/RHONE) coexistent dans la source — gérés ici en listant les deux formes
# (les accents sont déjà neutralisés par ailleurs via pt._normalise, pas le singulier/pluriel).
REGIONS_FRANCE = {
    "ardeche", "aura", "isere", "savoie", "haute savoie", "reunion",
    "ain", "drome", "rhone", "vaucluse", "bouche du rhone", "bouches du rhone", "haute loire",
    "drome ardeche",
}

# Sous-ensemble de REGIONS_FRANCE qui appartient réellement à la région Auvergne-Rhône-Alpes
# (AURA), celle de l'Ardèche — utilisé par _local_ardeche (Local = X). Vaucluse et Bouches-du-
# Rhône sont bien des départements français (donc Origine = "France"), mais appartiennent à
# PACA (Provence-Alpes-Côte d'Azur), pas à AURA : jamais "locaux" pour ce marché départemental,
# même si fréquemment combinés avec des départements AURA dans la même valeur source (ex:
# "ISERE/VAUCLUSE" — reste non local, cf. _local_ardeche qui exige TOUS les jetons locaux).
REGIONS_LOCALES_ARDECHE = REGIONS_FRANCE - {
    "reunion", "vaucluse", "bouche du rhone", "bouches du rhone",
}


def _normalise_pays(texte):
    t = pt._normalise(texte)
    t = re.sub(r"[.'’]", "", t)
    t = t.replace("-", " ")
    return re.sub(r"\s+", " ", t).strip()


def _origine_bpu(valeur_brute):
    if not valeur_brute or not str(valeur_brute).strip():
        return None, False, None

    texte = str(valeur_brute).strip()
    norm_complet = pt._normalise(texte)

    if norm_complet == "ue":
        return "UE", False, None
    if "non ue" in norm_complet or "hors ue" in norm_complet:
        return texte, False, None  # ex: "UE et non UE" : mélange assumé, laissé tel quel
    if RE_ZONE_PECHE.search(texte):
        return texte, False, None  # zone de pêche (FAO/Océan...), pas une liste de pays
    if norm_complet == "import":
        # Vu 2 fois (Agricourt, 2026-10-30) : "Import" dit explicitement que ce N'EST PAS de
        # la matière première française, mais sans préciser quel pays — le pt.traite_origine
        # par défaut (plus bas) forcerait "France" avec un commentaire, ce qui serait FAUX ici
        # (le contraire de ce que dit la source) ; gardé tel quel plutôt que de deviner UE/non
        # UE sans base, avec un commentaire pour que le pays exact soit confirmé si besoin.
        return texte, False, (
            f"Donnée source : « {valeur_brute} » — indique une origine non française, mais "
            "sans préciser quel pays ; laissé tel quel plutôt que remplacé par défaut."
        )

    jetons = [j.strip() for j in re.split(r"\s*/\s*|\s+-\s+|,\s*", texte) if j.strip()]

    if jetons and all(_normalise_pays(j) in REGIONS_FRANCE | {"france"} for j in jetons):
        return "France", False, None  # région(s) française(s) seule(s) ou combinée(s)

    if len(jetons) > 1:
        est_ue = [_normalise_pays(j) in PAYS_UE for j in jetons]
        if all(est_ue):
            return "UE", False, None
        if not any(est_ue):
            return "Hors UE", False, None
        return texte, False, None  # mélange UE/hors UE : les deux noms gardés, cf. docstring

    pays_norm = _normalise_pays(RE_PARENTHESE.sub(" ", texte).strip())
    if pays_norm == "indonesie":
        return "Indonésie", False, None
    if pays_norm == "thailande":
        return "Thaïlande", False, None
    if pays_norm in PAYS_UE or pays_norm in PAYS_HORS_UE:
        return texte, False, None
    if re.search(r"\bue\b", norm_complet):
        return texte, False, None  # ex: "UE et Pays de l'Est", "UE et Pays-Tiers" : phrase
        # mélangée non découpable proprement, laissée telle quelle plutôt que remplacée

    return pt.traite_origine(valeur_brute)


def transforme_bpu_dqe(chemin_source, chemin_sortie, nom_feuille_bpu, fournisseur_force=None,
                        lot_force=None):
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    ws = wb[nom_feuille_bpu]

    fournisseur = fournisseur_force or _fournisseur_depuis_page_de_garde(wb) or ws.title

    texte_garde = _texte_titre_page_de_garde(wb)
    texte_feuille = ws.cell(row=1, column=1).value
    lot_fichier = _lot_depuis_titre(texte_garde) or _lot_depuis_titre(texte_feuille)
    # Aucune des deux sources n'est toujours complète à elle seule : sur le Lot 5, la page de
    # garde dit "LOT DESSERTS ET GLACES" (description sans numéro) alors que le titre de la
    # feuille BPU ne donne que "... - LOT 5" (numéro sans description) — ni l'une ni l'autre
    # seule ne suffit, donc numéro et description sont recombinés quand le premier essai n'a
    # trouvé qu'un numéro nu ("LOT 5"), jamais quand il a déjà numéro + description (Lots 1-4).
    if lot_fichier and re.fullmatch(r"lot\s*\d+", lot_fichier, re.I):
        numero = _lot_numero(lot_fichier)
        description = (_lot_description_sans_numero(texte_garde)
                        or _lot_description_sans_numero(texte_feuille))
        if numero and description:
            lot_fichier = f"LOT {numero} {description}"

    ligne_entete = _trouve_ligne_entete_bpu(ws)
    cols = _colonnes_entete(ws, ligne_entete)

    col_designation = _col(cols, "libell", which=0)
    col_desig_fourn = _col(cols, "libell", which=-1)
    col_ref = _col(cols, "reference")
    col_unite = _col(cols, "unite")
    col_prix = _col(cols, "pu ht")
    col_cond = _col(cols, "conditionnement")
    col_egalim = _col(cols, "obligatoire") or _col(cols, "labellis")
    col_origine = _col(cols, "provenance")
    col_numero = _col(cols, "n°", which=0)
    # "Famille" par ligne (ex: "CONSERVE DE LEGUMES", "CONDIMENTS/ASSAISONNEMENTS") : utilisé
    # comme Lot UNIQUEMENT quand le fichier entier ne correspond pas à un numéro de lot précis
    # (cas Epicerie — pas de "LOT N" trouvable sur la page de garde ni dans le titre,
    # contrairement aux lots 1-5 Viandes/Poissons/... qui en ont un). Comparé aux vraies
    # valeurs de Lot des fichiers de référence (pret_pour_audit/, ex: "16. Produits
    # d'épicerie coopératifs", "3. Légumes 4ème gamme fermiers ou OP") : leur niveau de
    # détail correspond à "Famille" (19 valeurs dans ce fichier), pas à un seul mot
    # générique "Epicerie" qui serait anormalement grossier par comparaison.
    col_famille = _col(cols, "famille") if lot_fichier is None else None

    manquantes = [nom for nom, val in {
        "N°": col_numero, "libellé": col_designation, "unité": col_unite, "PU HT": col_prix,
    }.items() if val is None]
    if manquantes:
        raise CauseFormatInvalide(
            "COLONNE_OBLIGATOIRE_ABSENTE",
            f"Colonnes essentielles introuvables dans « {ws.title} » ({chemin_source}) : "
            f"{manquantes} absentes — en-têtes vus = {sorted(cols)}",
            onglet=ws.title, colonne=", ".join(manquantes), cellule=f"ligne {ligne_entete}",
        )

    ligne_debut = ligne_entete + 1
    ligne_fin = _derniere_ligne_numero(ws, ligne_debut, col_numero)

    lignes = []
    for r in range(ligne_debut, ligne_fin + 1):
        if not isinstance(ws.cell(row=r, column=col_numero).value, (int, float)):
            continue
        designation = ws.cell(row=r, column=col_designation).value
        if not designation:
            continue

        lot = lot_force or lot_fichier or (
            ws.cell(row=r, column=col_famille).value if col_famille else None) or ws.title

        desig_fourn = ws.cell(row=r, column=col_desig_fourn).value
        reference = ws.cell(row=r, column=col_ref).value if col_ref else None
        unite_brute = ws.cell(row=r, column=col_unite).value
        prix = _arrondi_prix(pg._nombre(ws.cell(row=r, column=col_prix).value))
        cond = ws.cell(row=r, column=col_cond).value if col_cond else None
        egalim = ws.cell(row=r, column=col_egalim).value if col_egalim else None
        origine_brute = ws.cell(row=r, column=col_origine).value if col_origine else None

        unite_ref, poids_kg, commentaire_unite = pt.traite_unite_reference(
            unite_brute, None, cond, prix, None, None
        )
        unite_cmd, prix_cmd = _commande_miroir_reference(unite_ref, prix)
        cond_a_ecrire = pt.reecrit_conditionnement(cond, unite_cmd, unite_ref)

        bio_x, label = _bio_et_label_bpu(egalim)
        if col_origine:
            valeur_origine, _, commentaire_origine = _origine_bpu(origine_brute)
        else:
            valeur_origine, commentaire_origine = None, None

        lignes.append(pg._ligne_vide_de_base(
            fournisseur=fournisseur, lot=lot, designation=designation, unite_ref=unite_ref,
            poids=poids_kg, commentaire_unite=commentaire_unite, cond=cond_a_ecrire,
            desig_fourn=desig_fourn, ref=reference, prix_ref=prix, unite_cmd=unite_cmd,
            prix_cmd=prix_cmd, local=None, bio=bio_x, label=label, origine=valeur_origine,
            commentaire_origine=commentaire_origine, marche=None,
        ))

    lignes.sort(key=lambda l: (pt._normalise(l["lot"]), pt._normalise(l["designation"])))
    return pt.ecrit_classeur_canonique(lignes, chemin_sortie, "BPU DQE")


def transforme_passion_froid(chemin_source, chemin_sortie):
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    # "SANS DOUBLONS" (préférée : déjà dédupliquée) sinon "GLOBAL" en secours, si un futur
    # envoi de ce fichier ne contient plus l'onglet dédupliqué — même disposition de colonnes
    # dans les deux onglets (vérifié le 2026-11-10 sur le fichier reçu), donc même lecture.
    # Demandé par l'utilisateur, qui a choisi "SANS DOUBLONS" par défaut mais veut que l'outil
    # reste utilisable si cet onglet venait à manquer.
    if "SANS DOUBLONS" in wb.sheetnames:
        ws = wb["SANS DOUBLONS"]
    elif "GLOBAL" in wb.sheetnames:
        ws = wb["GLOBAL"]
    else:
        raise CauseFormatInvalide(
            "ONGLET_MANQUANT",
            f"Ni l'onglet « SANS DOUBLONS » ni l'onglet « GLOBAL » n'ont été trouvés dans "
            f"« {chemin_source} » — onglets présents : {wb.sheetnames}.",
            onglet=None, colonne=None, cellule=None,
        )

    lignes = []
    for r in range(2, ws.max_row + 1):
        designation = ws.cell(row=r, column=4).value
        if not designation:
            continue
        gamme = ws.cell(row=r, column=2).value
        reference = ws.cell(row=r, column=3).value
        unite_brute = ws.cell(row=r, column=5).value
        prix = _arrondi_prix(pg._nombre(ws.cell(row=r, column=6).value))
        cond = ws.cell(row=r, column=7).value

        unite_ref, poids_kg, commentaire_unite = pt.traite_unite_reference(
            unite_brute, None, cond, prix, None, None
        )
        unite_cmd, prix_cmd = _commande_miroir_reference(unite_ref, prix)
        cond_a_ecrire = pt.reecrit_conditionnement(cond, unite_cmd, unite_ref)

        lignes.append(pg._ligne_vide_de_base(
            fournisseur="Passion Froid", lot=gamme, designation=designation, unite_ref=unite_ref,
            poids=poids_kg, commentaire_unite=commentaire_unite, cond=cond_a_ecrire,
            # Désignation fournisseur laissée vide : la source n'a qu'une seule colonne de
            # désignation ("Libellé"), pas de colonne fournisseur distincte — la dupliquer
            # aurait juste répété la même valeur dans les deux colonnes, sans info en plus.
            # Demandé par l'utilisateur le 2026-11-09.
            desig_fourn=None, ref=reference, prix_ref=prix, unite_cmd=unite_cmd,
            prix_cmd=prix_cmd, local=None, bio=None, label=None, origine=None,
            commentaire_origine=None, marche=None,
        ))

    lignes.sort(key=lambda l: (pt._normalise(l["lot"]), pt._normalise(l["designation"])))
    return pt.ecrit_classeur_canonique(lignes, chemin_sortie, "PASSION FROID")


# Colonne "Label" brute d'Episaveurs : plusieurs tags séparés par virgule (ex: "Bio, Légumes
# de France", "Appell. Origine Protégée AOP, Bio"), jamais traitée par pg._bio_et_label (qui ne
# reconnaît que le sigle "AB", pas le mot "Bio", ni cette structure à plusieurs tags) — d'où le
# Bio manquant signalé le 2026-10-29. Chaque tag est classé séparément :
# - un tag contenant "bio" (Bio, Bio AB, Conversion BIO...) => Bio = X, retiré du Label ;
# - un tag "Produit en ..." (ex: "Produit en Bretagne") => le texte source ne dit pas
#   explicitement "local", mais c'est le signal le plus probable : Local = X, avec un
#   commentaire pour que ce soit vérifié par une personne plutôt qu'affirmé sans vérification ;
# - un tag qui est un nom complet suivi de son sigle (ex: "Indication géo. protégée IGP",
#   "Appell. Origine Protégée AOP/AOC", "Marine Stewardship Council MSC") => seul le sigle est
#   gardé ("IGP", "AOP", "AOC", "MSC") ;
# - les autres (RSPO Segregated, Rainforest Alliance, Légumes de France, Max Havelaar...)
#   gardés tels quels, faute de sigle redondant à en extraire.
RE_SIGLE_FINAL = re.compile(r"\b([A-Z]{2,4})\s*$")
RE_PRODUIT_EN = re.compile(r"^produit en\b", re.I)
RE_BIO_TAG = re.compile(r"\bbio\b", re.I)


def _traite_label_episaveurs(valeur_brute):
    if not valeur_brute:
        return None, None, None, None

    bio_x = None
    local_x = None
    commentaire_local = None
    restants = []

    for tag in str(valeur_brute).split(","):
        tag = tag.strip()
        if not tag or pt._normalise(tag) in VALEURS_SANS_LABEL:
            continue
        if RE_BIO_TAG.search(tag):
            bio_x = "X"
            continue
        if RE_PRODUIT_EN.match(tag):
            local_x = "X"
            commentaire_local = (
                f"Donnée source (colonne Label) : « {tag} » — marqué Local sur la base de ce "
                "texte, à confirmer."
            )
            continue
        m = RE_SIGLE_FINAL.search(tag)
        restants.append(m.group(1) if m and m.group(1) != tag else tag)

    label = ", ".join(restants) if restants else None
    return bio_x, label, local_x, commentaire_local


def _ajoute_commentaires_local(chemin_sortie, lignes):
    """pt.ecrit_classeur_canonique n'ajoute de commentaire que sur Unité/Origine, pas sur
    Local — ajouté ici après coup, en se basant sur l'ordre d'écriture (ligne i de `lignes`
    -> rangée i+3 du classeur, cf. pt.ecrit_classeur_canonique)."""
    a_commenter = [(i, l["commentaire_local"]) for i, l in enumerate(lignes) if l.get("commentaire_local")]
    if not a_commenter:
        return
    wb = openpyxl.load_workbook(chemin_sortie)
    ws = wb.active
    for i, commentaire in a_commenter:
        ws.cell(row=i + 3, column=pt.COL_LOCAL).comment = Comment(commentaire, "audit_general")
    wb.save(chemin_sortie)


def transforme_episaveurs(chemin_source, chemin_sortie):
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    ws = wb.worksheets[0]

    lignes = []
    for r in range(2, ws.max_row + 1):
        desig_fourn = ws.cell(row=r, column=2).value
        if not desig_fourn:
            continue
        groupe = ws.cell(row=r, column=6).value
        reference = ws.cell(row=r, column=1).value
        unite_brute = ws.cell(row=r, column=3).value
        cond = ws.cell(row=r, column=4).value
        poids_uf = ws.cell(row=r, column=5).value
        prix = _arrondi_prix(pg._nombre(ws.cell(row=r, column=7).value))
        label_brut = ws.cell(row=r, column=11).value

        unite_ref, poids_kg, commentaire_unite = pt.traite_unite_reference(
            unite_brute, None, cond, prix, None, None
        )
        if poids_kg is None and unite_ref not in ("KG", "L") and poids_uf:
            poids_kg = pg._nombre(poids_uf)  # "Poids UF" est déjà en kg, pas en grammes
        unite_cmd, prix_cmd = _commande_miroir_reference(unite_ref, prix)
        cond_a_ecrire = pt.reecrit_conditionnement(cond, unite_cmd, unite_ref)

        bio_x, label, local_x, commentaire_local = _traite_label_episaveurs(label_brut)

        ligne = pg._ligne_vide_de_base(
            fournisseur="Episaveurs", lot=groupe, designation=desig_fourn, unite_ref=unite_ref,
            poids=poids_kg, commentaire_unite=commentaire_unite, cond=cond_a_ecrire,
            desig_fourn=desig_fourn, ref=reference, prix_ref=prix, unite_cmd=unite_cmd,
            prix_cmd=prix_cmd, local=local_x, bio=bio_x, label=label, origine=None,
            commentaire_origine=None, marche=None,
        )
        ligne["commentaire_local"] = commentaire_local
        lignes.append(ligne)

    lignes.sort(key=lambda l: (pt._normalise(l["lot"]), pt._normalise(l["designation"])))
    nb_lignes = pt.ecrit_classeur_canonique(lignes, chemin_sortie, "EPISAVEURS")
    _ajoute_commentaires_local(chemin_sortie, lignes)
    return nb_lignes


def _tronque_titre_lot_fresca(lot_global):
    """Le titre de lot BPUR FRESCA a 2 tirets (ex: "LOT 1 – PRODUITS SURGELES –
    CONVENTIONNEL ET SIQO") : seuls les 2 premiers segments (jusqu'au "–" qui précède
    "CONVENTIONNEL ET SIQO") forment le vrai nom du lot — le reste est une mention
    administrative sans intérêt pour l'audit. Demandé explicitement par le métier."""
    segments = re.split(r"\s*[–-]\s*", lot_global)
    if len(segments) >= 2:
        return " – ".join(segments[:2])
    return lot_global


def transforme_bpur_fresca(chemin_source, chemin_sortie, fournisseur="FRESCA"):
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    ws = wb.worksheets[0]

    lot_global = _tronque_titre_lot_fresca(
        pg.extrait_titre_lot(ws.cell(row=3, column=1).value) or ws.title
    )

    ligne_entete = 6
    col_famille, col_numero, col_intitule, col_desig_fourn = 1, 2, 3, 4
    col_origine, col_label_egalim, col_autre_label = 6, 7, 8
    col_cond, col_unite, col_prix = 12, 13, 14

    ligne_debut = ligne_entete + 1
    ligne_fin = pg._derniere_ligne_avec_reference(ws, ligne_debut, col_numero)
    # _lots_par_section repère une "section" (ex: "CONDIMENTS") par ligne, mais elle n'entre PAS
    # dans le Lot écrit en sortie (seulement ignorée ci-dessous, via _) : ajouter la section au
    # Lot créerait, pour un même numéro de lot, autant de noms de Lot différents que de
    # sections — demandé explicitement le 2026-10-30 ("um mesmo nome de lote não pode ter
    # diferentes nomes"). Le Lot reste donc `lot_global` seul, identique pour toutes les lignes.
    lots = pg._lots_par_section(ws, ligne_debut, ligne_fin, col_famille, col_numero)

    lignes = []
    for r, _section in lots.items():
        designation = ws.cell(row=r, column=col_intitule).value
        if not designation:
            continue
        desig_fourn = ws.cell(row=r, column=col_desig_fourn).value
        unite_brute = ws.cell(row=r, column=col_unite).value
        prix = _arrondi_prix(pg._nombre(ws.cell(row=r, column=col_prix).value))
        cond = ws.cell(row=r, column=col_cond).value
        origine_brute = ws.cell(row=r, column=col_origine).value
        autre_label = ws.cell(row=r, column=col_autre_label).value

        unite_ref, poids_kg, commentaire_unite = pt.traite_unite_reference(
            unite_brute, None, cond, prix, None, None
        )
        unite_cmd, prix_cmd = _commande_miroir_reference(unite_ref, prix)
        cond_a_ecrire = pt.reecrit_conditionnement(cond, unite_cmd, unite_ref)

        bio_x, label = pg._bio_et_label(autre_label)
        valeur_origine, _, commentaire_origine = _origine_bpu(origine_brute)

        lignes.append(pg._ligne_vide_de_base(
            fournisseur=fournisseur, lot=lot_global,
            designation=designation, unite_ref=unite_ref, poids=poids_kg,
            commentaire_unite=commentaire_unite, cond=cond_a_ecrire, desig_fourn=desig_fourn,
            ref=None, prix_ref=prix, unite_cmd=unite_cmd, prix_cmd=prix_cmd, local=None,
            bio=bio_x, label=label, origine=valeur_origine,
            commentaire_origine=commentaire_origine, marche=None,
        ))

    lignes.sort(key=lambda l: (pt._normalise(l["lot"]), pt._normalise(l["designation"])))
    return pt.ecrit_classeur_canonique(lignes, chemin_sortie, "BPUR FRESCA")


# Famille "MP ARDECHE 24" (DADS, Episaveurs, LDC, Passion Froid — 4 fichiers reçus le
# 2026-10-29, un fournisseur différent par fichier malgré le nom partagé avec Passion
# Froid/Episaveurs déjà traités plus haut : format totalement différent). Tableau à plat,
# en-têtes fixes en ligne 1 ("Id Produit", "Réf", "Nom technique", "Signes", "Egalim",
# "Groupe", "Marque", "Pays de transformation", "Origine matière première", "Cond.", "Prix",
# "Fact.", "UC", "Commentaire", "PU"). Décisions de mapping confirmées avec l'utilisateur le
# 2026-10-29 (voir mapping_colonnes_a_formater.xlsx) :
# - Lot et Fournisseur viennent du nom du fichier (pas d'une colonne).
# - "UC" (KG/LT/BT/UN) est l'unité de COMMANDE, pas de référence — confirmé par "PU" (= Prix
#   ÷ Fact., vérifié cohérent sur toutes les lignes), qui est TOUJOURS le prix d'UNE unité UC.
# - L'unité de référence (BPU) réelle ne se cherche QUE quand l'unité de commande n'est pas
#   déjà KG/L (cas LDC notamment, UC toujours KG : pas de conversion à chercher, PU est déjà
#   le prix au kg) : dans ce cas on regarde la fin du texte de "Nom technique" — un poids/
#   volume net ("sachet 630 g", "poche 1 kg", "1 lt") permet de convertir le prix en KG/L ; un
#   code de contenant en N/M ("bt 3/1", "bt 4/4") veut toujours dire BOITE (confirmé par
#   l'utilisateur, et déjà un format accepté — voir unites_commande_non_reconnues.md) ; une
#   fourchette ("300/350 g") n'est PAS convertie (pas de valeur unique fiable, pas de
#   conversion inventée) — dans ce dernier cas l'unité de référence reste égale à l'unité de
#   commande. "Cond." sert de multiplicateur ("N X ...") dans le Conditionnement reconstruit,
#   jamais dans un calcul de prix — voir `_unite_reference_ardeche` pour le détail des deux cas
#   (Cond. exprimé dans l'unité de commande elle-même quand elle n'est pas KG/L, confirmé par
#   DADS "Yaourt ... (x48) 125 g" / Cond.=48 ; Cond. exprimé en KG/L quand elle l'est déjà).
# - Origine = "Pays de transformation" (demandé par le métier le 2026-11-09, après retour sur
#   le mapping de colonnes — avant cette date, "Origine matière première" était utilisée à sa
#   place). "Origine matière première" reste malgré tout lue séparément, UNIQUEMENT pour
#   calculer Local (_local_ardeche) : c'est elle qui porte les régions ARDECHE/AURA/ISERE/
#   SAVOIE (voir REGIONS_FRANCE ci-dessous) quand la matière première est française, alors que
#   "Pays de transformation" ne décrit que le lieu de transformation/conditionnement (jamais une
#   région française, toujours un pays) — les deux colonnes sont donc complémentaires, pas
#   interchangeables : Origine vient de l'une, Local de l'autre.
# - "Signes" : "AB" et tout tag contenant le mot "bio" (ex: "Bio Part.", vu le 2026-10-30) sont
#   des signes bio ("ABE" l'accompagne toujours, jamais seul, aucune source ne le confirme
#   comme label officiel séparé — traité comme faisant partie du même signal bio) ; "null"
#   (texte littéral, artefact d'export vu le 2026-10-30 dans plusieurs fichiers DADS) est du
#   bruit, retiré sans devenir un Label. Les autres sigles vont en Label tels quels (recherchés
#   le 2026-10-29, complétés le 2026-10-30) :
#   AOP/IGP/HVE/LR = labels connus, CE2 = Certification Environnementale niveau 2, ASC =
#   Aquaculture Stewardship Council (confirmé, équivalent MSC pour l'aquaculture — fichier
#   SYSCO FRANCE, poisson), RUP = Région Ultrapériphérique (confirmé, SIQO officiel pour les
#   produits des territoires français d'outre-mer — "Banane ... RUP", Vivalya), VBF = Viande
#   Bovine Française (confirmé), VOF = Viande Ovine Française (confirmé, fichier DADS, sur de
#   l'agneau uniquement), VVF = Viande de Veau Française (confirmé, fichier DADS/Passion
#   Froid, sur du veau uniquement), LPF = Le Porc Français (confirmé, label officiel filière
#   porcine — fichier DADS/Passion Froid, sur du porc uniquement), VF = probablement Volaille
#   Française (fichier LDC, volailles uniquement — pas de confirmation officielle trouvée),
#   ODF = probablement Oeufs de France (n'apparaît que sur des produits à base d'oeuf), BZH =
#   Bretagne (n'apparaît que sur des produits bretons), VGN = probablement Vegan (biscuits
#   sans ingrédient animal évident — pas de confirmation officielle trouvée, mais abréviation
#   courante). GG (saumon), MGDF (magret de canard), FFL (poivre), CER (vanille, brownie) :
#   sens inconnu malgré recherche, gardés tels quels sans interprétation — comme "16" (2
#   occurrences, produits Cantadou, déjà inexpliqué le 2026-10-29).

RE_ARDECHE_FOURCHETTE = re.compile(r"\d+\s*/\s*\d+\s*(?:kg|g|lt|l|ml|cl)\.?\s*$", re.I)
RE_ARDECHE_BOITE = re.compile(r"\bbte?\.?\s*(\d+\s*/\s*\d+)\s*$", re.I)
RE_ARDECHE_POIDS = re.compile(r"(\d+[,.]?\d*)\s*(kg|g)\.?\s*$", re.I)
RE_ARDECHE_VOLUME = re.compile(r"(\d+[,.]?\d*)\s*(lt|l|ml|cl)\.?\s*$", re.I)
RE_ARDECHE_CONTENANT_MOT = re.compile(
    r"\b(sachet|sac|poche|bte?|boite|flacon|pot|barquette|seau|brique|tube)\b", re.I
)
RE_SIGNE_ABE = re.compile(r"\babe\b", re.I)


def _nombre_entier_cond(cond):
    """"Cond." ("1,000", "25,000"...) -> entier le plus proche (jamais 0 : round() arrondit
    0,5 à 0 par convention bancaire, qui n'a pas de sens ici comme multiplicateur), ou 1 si
    vide/illisible — sert uniquement à préfixer le Conditionnement ("N X ..."), jamais le
    calcul de prix (voir le commentaire de section : Cond. est indépendant du prix, qui est
    toujours par UC)."""
    valeur = pg._nombre(cond)
    return max(1, int(round(valeur))) if valeur else 1


def _texte_conditionnement(n, quantite_texte, conteneur, unite_individuelle=False):
    """Assemble le Conditionnement final à partir d'une quantité déjà formatée (via
    pt._formate_quantite) — ne passe PAS par pt.reecrit_conditionnement : cette dernière
    réanalyse le texte avec sa propre regex (format "N x qté"), qui ne reconnaît que les cas
    où `n` ET `conteneur` sont établis autrement (REGAL) ; ici les deux sont déjà connus avec
    certitude, donc assemblés directement, même logique de rendu qu'elle ("CONTENANT DE N X
    QTE", ou juste la quantité si n=1 et pas de contenant).

    `unite_individuelle=True` dit que `conteneur` nomme l'unité vendue ELLE-MÊME (ex: "PIECE",
    "UNITÉ", mais aussi un mot de contenant trouvé dans la désignation comme "POT" quand le
    code source est "UN" — un pot = une part, pas un carton qui en regrouperait plusieurs) :
    au pluriel avec le nombre devant, "N POTS DE qté" plutôt que "POT DE N X qté", qui se
    lirait à tort comme "un pot contenant N×qté" plutôt que "N pots de qté chacun" — et ce
    texte, repris tel quel par l'audit (qui cherche des motifs "N X qté" dans le
    Conditionnement pour vérifier le prix), créait un faux signalement d'incohérence. Trouvé
    sur "Yaourt ... pot carton (x48) 125 g" (le "carton" regroupe 48 "pot", chacun = 1 part de
    125 g — le mot reconnu dans la désignation est "pot", pas "carton", absent de la liste de
    mots de contenant reconnus) — demandé le 2026-11-09. `unite_individuelle=False` (le
    défaut) garde le sens inverse : `conteneur` nomme un contenant qui REGROUPE n unités plus
    petites (ex: "BOITE DE 6 X 800 G")."""
    if conteneur and unite_individuelle and n > 1:
        mot_pluriel = "UNITÉS" if conteneur == "UNITÉ" else f"{conteneur}S"
        return f"{n} {mot_pluriel} DE {quantite_texte}"
    if conteneur:
        return f"{conteneur} DE {n} X {quantite_texte}" if n > 1 else f"{conteneur} DE {quantite_texte}"
    return f"{n} X {quantite_texte}" if n > 1 else quantite_texte


def _unite_reference_ardeche(designation, mot_cond, prix_cmd, cond, base_uc):
    """Cherche dans la fin de `designation` de quoi déduire une unité de référence plus
    fiable que l'unité de commande telle quelle, et le Conditionnement qui va avec — voir le
    commentaire de section ci-dessus pour la règle. Renvoie (unite_ref, poids_ou_none,
    prix_ref, commentaire_ou_none, conditionnement_ou_none).

    `mot_cond` (voir _mot_contenant_ardeche) est le mot utilisé pour le Conditionnement et,
    ici, l'Unité de référence — PAS la colonne Unité de commande elle-même (qui recopie le
    code source tel quel, voir _unite_commande_ardeche) : ex: "BRIQUE", "SEAU", "UNITÉ".
    `base_uc` est le code source normalisé ("kg"/"l"/"un"/"autre") et décide SEUL si Cond. doit
    multiplier le prix — jamais `mot_cond`, pour ne pas faire dépendre un calcul de prix du
    vocabulaire trouvé dans un texte libre. Séparation demandée par l'utilisateur le
    2026-11-09 (trouvé sur "Lait fermenté ... brique 1 lt", code source "LT", où le mot
    "brique" disparaissait avant alors qu'il n'aurait dû changer que l'affichage, jamais le
    calcul) ; et la colonne Unité de commande a ensuite été remise à recopier le code source
    tel quel, toute incohérence avec `mot_cond`/le Conditionnement étant désormais détectée
    par l'audit (verifie_unite_commande_etendue), pas devinée ici.

    Le 2e élément (poids à l'unité de référence) est TOUJOURS None pour cette famille : cette
    colonne n'est prévue que pour Unité de référence = PC (format_cible.md ligne 31), jamais
    le cas ici (KG, L, BOITE N/M, ou `mot_cond` tel quel) — le détail de poids par pièce,
    quand il est connu, va dans le texte du Conditionnement, pas dans cette colonne.

    Si le code source est déjà KG ou L, aucune conversion n'est cherchée : le prix de commande
    (PU) est déjà le prix par KG/L (confirmé par l'utilisateur), donc Unité de référence = KG/L
    et Prix de référence = Prix de commande directement. Seul le Conditionnement est
    reconstruit à partir du nombre brut "Cond." (déjà exprimé dans cette même unité — ex: LDC,
    "Cond." = poids total du colis en kg), précédé de `mot_cond` s'il en nomme un
    (ex: "SEAU DE 10L")."""
    if base_uc in ("kg", "l"):
        valeur = pg._nombre(cond)
        unite_ref_finale = "KG" if base_uc == "kg" else "L"
        quantite_texte = pt._formate_quantite(valeur, base_uc) if valeur else None
        cond_ecrit = (f"{mot_cond} DE {quantite_texte}"
                      if quantite_texte and mot_cond != unite_ref_finale else quantite_texte)
        return unite_ref_finale, None, prix_cmd, None, cond_ecrit

    texte = str(designation or "").strip()
    # Ici "Cond." est exprimé dans l'unité de commande elle-même (ex: nombre de boîtes, ou
    # nombre de pièces — confirmé par l'exemple DADS "Yaourt ... (x48) 125 g" / Cond.=48,000)
    # : utilisable tel quel comme multiplicateur "N X ...", sans division par un poids-pièce.
    n = _nombre_entier_cond(cond)
    # Jamais multiplié par Cond. : base source "un" (une unité sans ambiguïté, quel que soit
    # `mot_cond`), ou code de boîte en N/M ("BOITE 3/1" — la boîte EST l'unité vendue, Cond.
    # ne compte que combien de boîtes, jamais un poids à l'intérieur de CETTE boîte-ci).
    code_boite = mot_cond.startswith("BOITE ") and "/" in mot_cond
    jamais_multiplier = base_uc == "un" or code_boite

    if not RE_ARDECHE_FOURCHETTE.search(texte):
        if code_boite:
            cond_ecrit = _texte_conditionnement(n, mot_cond, None)
            return mot_cond, None, prix_cmd, None, cond_ecrit

        m = RE_ARDECHE_POIDS.search(texte)
        if m:
            poids_kg = float(m.group(1).replace(",", ".")) * (0.001 if m.group(2).lower() == "g" else 1)
            if poids_kg:
                # Quand l'unité vendue est un contenant qui groupe plusieurs pièces (BOITE,
                # SACHET...), le poids trouvé en fin de désignation ne décrit qu'UNE pièce
                # dedans — le contenant vendu pèse Cond fois ce poids, donc Cond entre dans la
                # conversion. Quand la base source est "un" (une seule unité vendue à la
                # fois), Cond ne compte que combien d'unités sont groupées dans le même
                # sachet/plateau source, sans rapport avec le prix de CETTE unité-ci —
                # confirmé le 2026-10-30 : "Quenelle nature (x16) 120 g", Cond=16, UC=UN,
                # PU=0,550€ (prix d'UNE quenelle de 120g, pas de 16) ; à l'inverse
                # "Brioche... 800 g", Cond=6, UC=BT (mot BOITE générique), PU=6,716€ est le
                # prix de la BOITE de 6, donc 6*0,8=4,8 kg, pas 0,8 kg.
                poids_total_kg = poids_kg if jamais_multiplier else poids_kg * n
                prix_kg = round(prix_cmd / poids_total_kg, 3) if prix_cmd is not None else None
                cond_ecrit = _texte_conditionnement(
                    n, pt._formate_quantite(poids_kg, "kg"), mot_cond, jamais_multiplier)
                return "KG", None, prix_kg, None, cond_ecrit

        m = RE_ARDECHE_VOLUME.search(texte)
        if m:
            facteurs = {"l": 1, "lt": 1, "ml": 0.001, "cl": 0.01}
            unite_vol = m.group(2).lower()
            volume_l = float(m.group(1).replace(",", ".")) * facteurs[unite_vol]
            if volume_l:
                # Même logique que pour le poids ci-dessus (contenant vs unité vendue seule).
                volume_total_l = volume_l if jamais_multiplier else volume_l * n
                prix_l = round(prix_cmd / volume_total_l, 3) if prix_cmd is not None else None
                cond_ecrit = _texte_conditionnement(
                    n, pt._formate_quantite(volume_l, "l"), mot_cond, jamais_multiplier)
                return "L", None, prix_l, None, cond_ecrit

    if jamais_multiplier and n and n > 1:
        # Rien d'exploitable comme poids/volume (ex: "Oeuf frais plein air 53/63 (x30)"), mais
        # Cond. reste un pur compte ici (jamais un multiplicateur de prix) : au moins dire
        # "N <mot>S" plutôt qu'un nombre brut sans unité — demandé le 2026-11-09.
        mot_pluriel = "UNITÉS" if mot_cond == "UNITÉ" else f"{mot_cond}S"
        return mot_cond, None, prix_cmd, None, f"{n} {mot_pluriel}"

    # Rien d'exploitable (fourchette, ou aucun poids/volume/code boîte en fin de texte, et pas
    # un simple compte d'unités) : l'unité de référence reste `mot_cond`, sans conversion
    # inventée, et le Conditionnement reste le "Cond." brut de la source.
    return mot_cond, None, prix_cmd, None, None


def _base_uc_ardeche(uc_brut):
    """Catégorie SOURCE du code "UC" brut (KG/LT/UN/BT...), utilisée uniquement par
    _unite_reference_ardeche pour décider si Cond. doit multiplier le prix — jamais pour
    choisir le mot affiché (voir _unite_commande_ardeche). "kg"/"l"/"un" sont des bases sans
    ambiguïté (Cond. n'est jamais un multiplicateur de prix) ; tout le reste ("autre", ex: le
    code "BT") reste ambigu, résolu au cas par cas dans _unite_reference_ardeche."""
    norm = pt._normalise(uc_brut)
    if norm == "kg":
        return "kg"
    if norm == "lt":
        return "l"
    if norm == "un":
        return "un"
    return "autre"


def _mot_contenant_ardeche(uc_brut, designation):
    """Mot utilisé pour enrichir le Conditionnement et, dans certains cas déjà confirmés
    (contenant imbriqué, voir format_cible.md), l'Unité de référence (BPU) elle-même —
    JAMAIS la colonne Unité de commande, qui recopie le code source tel quel (voir
    _unite_commande_ardeche, juste en dessous). Le mot de contenant réellement écrit dans la
    désignation (seau, brique, pot, sachet...) ou un code de boîte en N/M ("bt 3/1", "bt
    4/4", gardé complet — "BOITE 3/1", pas juste "BOITE", confirmé le 2026-10-30, format déjà
    accepté : voir unites_commande_non_reconnues.md) sont préférés quand ils existent, même
    quand le code source était déjà "KG"/"LT"/"UN" (ex: "Lait fermenté ... brique 1 lt", code
    source "LT" -> mot "BRIQUE" ; "Crème fraîche ... seau 10 lt", code source "LT" -> "SEAU").
    Sans mot trouvé, replie sur le mot générique du code source (KG/L/UNITÉ, ou le code brut
    tel quel pour un code ambigu comme "BT")."""
    m = RE_ARDECHE_BOITE.search(str(designation or ""))
    if m:
        return f"BOITE {m.group(1).replace(' ', '')}"
    m = RE_ARDECHE_CONTENANT_MOT.search(str(designation or ""))
    if m:
        mot = pt._normalise(m.group(1))
        return "BOITE" if mot in ("bt", "bte", "boite") else mot.upper()
    return _unite_commande_ardeche(uc_brut)


def _unite_commande_ardeche(uc_brut):
    """Unité de commande FINALE (colonne "UNITE COMMANDE") : reflet direct du code source
    "UC", seulement traduit vers le vocabulaire canonique (KG->KG, LT->L, UN->UNITÉ — une
    traduction du MÊME code, pas une supposition) — jamais deviné à partir du texte libre de
    la Désignation, contrairement au Conditionnement et, dans certains cas, l'Unité de
    référence (voir _mot_contenant_ardeche). Une éventuelle incohérence entre cette valeur et
    ce que suggère le Conditionnement (ex: UC="BT" alors que la désignation dit "brique") est
    désormais du ressort de l'audit (verifie_unite_commande_etendue), pas de la mise en forme
    — demandé explicitement par l'utilisateur le 2026-11-09 : "recuperemos a coluna UC como
    ela é"."""
    base = _base_uc_ardeche(uc_brut)
    if base == "kg":
        return "KG"
    if base == "l":
        return "L"
    if base == "un":
        return "UNITÉ"
    return uc_brut


# REGIONS_LOCALES_ARDECHE : voir sa définition à côté de REGIONS_FRANCE plus haut dans ce
# fichier — confirmé par l'utilisateur le 2026-10-30 ("ardeche e tudo mais" = la région
# Auvergne-Rhône-Alpes à laquelle l'Ardèche appartient administrativement).
def _local_ardeche(origine_brute):
    """Local = X quand "Origine matière première" ne contient QUE des codes régionaux locaux
    (seul(s) ou combinés, ex: "SAVOIE/HAUTE-SAVOIE") — jamais pour "France" tout court (trop
    générique pour ce marché départemental) ni pour "REUNION" (française mais pas locale)."""
    if not origine_brute or not str(origine_brute).strip():
        return None
    jetons = [j.strip() for j in re.split(r"\s*/\s*|\s+-\s+|,\s*", str(origine_brute).strip()) if j.strip()]
    if jetons and all(_normalise_pays(j) in REGIONS_LOCALES_ARDECHE for j in jetons):
        return "X"
    return None


RE_SIGNE_BIO_MOT = re.compile(r"\bbio\b", re.I)
RE_SIGNE_NULL = re.compile(r"^null$", re.I)


def _bio_et_label_ardeche(signes_brut):
    """"Signes" (ex: "ABE,AB,VF", "Ardèche,null,VBF", "ABE,Bio Part.") :
    - "null" (texte littéral, vu le 2026-10-30 dans plusieurs fichiers DADS — artefact
      d'export, pas un signe réel) est retiré silencieusement ;
    - "ABE" est retiré : accompagne toujours un autre signal bio ("AB", ou "Bio Part." vu le
      2026-10-30) sans jamais apparaître comme signal bio indépendant ;
    - tout tag contenant le mot "bio" (pas seulement le sigle exact "AB" — ex: "Bio Part.",
      vu sur des produits partiellement bio) déclenche Bio = X et est retiré du Label, même
      logique que pour Episaveurs (RE_BIO_TAG) ;
    - le reste est délégué à pg._bio_et_label (reconnaît le sigle "AB")."""
    if not signes_brut:
        return None, None
    bio_mot = None
    tags = []
    for tag in str(signes_brut).split(","):
        tag = tag.strip()
        if not tag or RE_SIGNE_ABE.fullmatch(tag) or RE_SIGNE_NULL.fullmatch(tag):
            continue
        if RE_SIGNE_BIO_MOT.search(tag):
            bio_mot = "X"
            continue
        tags.append(tag)
    bio_ab, label = pg._bio_et_label(", ".join(tags) if tags else None)
    return (bio_mot or bio_ab), label


COL_ARDECHE_REF, COL_ARDECHE_DESIGNATION, COL_ARDECHE_SIGNES = 2, 3, 4
COL_ARDECHE_PAYS_TRANSFORMATION = 8
COL_ARDECHE_ORIGINE_MP = 9
COL_ARDECHE_COND, COL_ARDECHE_PRIX, COL_ARDECHE_FACT, COL_ARDECHE_UC = 10, 11, 12, 13
COL_ARDECHE_PU = 15


def _nombre_ardeche(valeur):
    """"Prix"/"PU" sont écrits avec un symbole € (ex: "4,300 €") que pg._nombre ne gère pas
    (il attend juste une virgule décimale) — retiré ici avant de lui déléguer la conversion."""
    if isinstance(valeur, str):
        valeur = valeur.replace("€", "").strip()
    return pg._nombre(valeur)


def transforme_ardeche(chemin_source, chemin_sortie, fournisseur, lot):
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    ws = wb.worksheets[0]

    lignes = []
    for r in range(2, ws.max_row + 1):
        designation = ws.cell(row=r, column=COL_ARDECHE_DESIGNATION).value
        if not designation:
            continue
        reference = ws.cell(row=r, column=COL_ARDECHE_REF).value
        signes = ws.cell(row=r, column=COL_ARDECHE_SIGNES).value
        pays_transformation = ws.cell(row=r, column=COL_ARDECHE_PAYS_TRANSFORMATION).value
        origine_mp_brute = ws.cell(row=r, column=COL_ARDECHE_ORIGINE_MP).value
        cond = ws.cell(row=r, column=COL_ARDECHE_COND).value
        uc_brut = ws.cell(row=r, column=COL_ARDECHE_UC).value
        # Prix de commande = "Prix" tel quel (recopié de la source, sans division) : si le
        # Conditionnement dit "10 L", le prix doit refléter ces 10 L, pas un prix ramené à 1 L
        # — l'incohérence entre Unité de commande (ex: "L" alors que le prix couvre tout le
        # Conditionnement) est volontairement laissée telle quelle ici, à corriger plus tard
        # côté audit, pas à la formatage (décidé avec l'utilisateur le 2026-11-09, après un
        # premier essai en Prix ÷ Fact. qui ne correspondait plus à la colonne "Prix" source).
        #
        # Prix de référence (BPU, par KG/L/PC), lui, reste calculé à partir de "Prix" ÷
        # "Fact." (pas "Prix" seul) : "Fact." n'est PAS toujours 1 (vérifié sur les 12 fichiers
        # MP ARDECHE 24 : va de 0,04 à 200 selon le fichier) — sans cette division, le Prix de
        # référence serait faux d'un facteur "Fact." à chaque fois qu'il diffère de 1 (ex: ce
        # même exemple, Cond.=10L/Fact.=10 : prix au litre réel = 65,4€ ÷ 10 = 6,54€, pas 65,4€).
        prix_brut = _nombre_ardeche(ws.cell(row=r, column=COL_ARDECHE_PRIX).value)
        fact = _nombre_ardeche(ws.cell(row=r, column=COL_ARDECHE_FACT).value) or 1
        prix_commande = _arrondi_prix(prix_brut) if prix_brut is not None else None
        prix_pour_bpu = _arrondi_prix(prix_brut / fact) if prix_brut is not None else None

        unite_cmd = _unite_commande_ardeche(uc_brut)
        mot_cond = _mot_contenant_ardeche(uc_brut, designation)
        unite_ref, poids_ou_volume, prix_ref, commentaire_unite, cond_propose = _unite_reference_ardeche(
            designation, mot_cond, prix_pour_bpu, cond, _base_uc_ardeche(uc_brut)
        )
        cond_a_ecrire = cond_propose if cond_propose else cond

        bio_x, label = _bio_et_label_ardeche(signes)
        # Origine = "Pays de transformation" (demandé par le métier le 2026-11-09, pas
        # "Origine matière première" comme avant — voir le commentaire de section plus haut,
        # resté correct pour expliquer pourquoi "Origine matière première" est gardée ci-dessous
        # UNIQUEMENT pour Local). "Pays de transformation" ne contient jamais de région
        # française (toujours un pays), donc ne casse pas _origine_bpu.
        valeur_origine, _, commentaire_origine = _origine_bpu(pays_transformation)
        local_x = _local_ardeche(origine_mp_brute)

        lignes.append(pg._ligne_vide_de_base(
            fournisseur=fournisseur, lot=lot, designation=designation, unite_ref=unite_ref,
            poids=poids_ou_volume, commentaire_unite=commentaire_unite, cond=cond_a_ecrire,
            # Désignation fournisseur laissée vide : la source n'a qu'une seule colonne de
            # désignation ("Nom technique"), pas de colonne fournisseur distincte — la
            # dupliquer aurait juste répété la même valeur dans les deux colonnes, sans info en
            # plus. Demandé par l'utilisateur le 2026-11-09 (remarqué sur DADS/Passion Froid MP
            # ARDECHE 24, mais s'applique aux 12 fichiers de cette famille).
            desig_fourn=None, ref=reference, prix_ref=prix_ref, unite_cmd=unite_cmd,
            prix_cmd=prix_commande, local=local_x, bio=bio_x, label=label, origine=valeur_origine,
            commentaire_origine=commentaire_origine, marche=None,
        ))

    lignes.sort(key=lambda l: pt._normalise(l["designation"]))
    return pt.ecrit_classeur_canonique(lignes, chemin_sortie, "MP ARDECHE 24")


# Un seul fichier à la fois peut présenter une particularité propre à lui (page de garde
# absente, feuille BPU nommée différemment...) : plutôt que de tout déduire dynamiquement
# avec le risque de se tromper en silence, chaque fichier connu est répertorié ici
# explicitement, comme nom_feuille_bpu dans transforme_fichier_grada_auto (pretraitement_auto.py).
FICHIERS_BPU_DQE = {
    "05_BPU DQE Taux de remise_2023_0498_ Lot 1 Viandes et volailles.xlsx": "BPU Lot 1",
    "06_BPU DQE Taux de remise_2023_0498_Lot 2 Poissons.xlsx": "BPU Lot 2",
    "6-BPU DQE Taux de remise_2023_0498_Lot 3 Legumes V2.xlsx": "BPU Lot 3",
    "08_BPU DQE Taux de remise_2023_0498_Lot 4 PC et Entrees.xlsx": "BPU Lot 4",
    "09_BPU DQE Taux de remise_2023_0498_Lot 5 Glaces et desserts.xlsx": "BPU Lot 5",
    "13-BPU_DQE_2023_0276_DEC_CA_EPICERIE V2.xlsx": "BPU",
}
# Lot 4 n'a pas de titre de lot fiable sur sa page de garde pour ce cas précis (vérifié :
# "FOURNITURE ET LIVRAISON ... - LOT 4" y est bien présent, donc en réalité pas besoin de
# forcer — gardé explicite malgré tout pour qu'un futur changement de ce fichier ne passe
# pas inaperçu).
LOT_FORCE = {}

FICHIER_PASSION_FROID = "Passion froid - Tarif général Mise en place V2 1.xlsx"
FICHIER_EPISAVEURS_XLS = "Episaveurs - Tarifs Catalogue complémentaire Septembre 2024 V2.xls"
FICHIER_BPUR_FRESCA = "1. BPUR FRESCA LOT 1 SURGELES CONV SIQO.xlsx"

# Fournisseur et Lot : voir le commentaire de section au-dessus de transforme_ardeche — pas de
# colonne dédiée, répertorié par fichier comme pour FICHIERS_BPU_DQE. Les noms de fichier DADS
# contiennent des entités HTML littérales (D&#x27;Ard&#xE8;che) — un défaut du fichier source,
# pas corrigé ici (seule la clé du dict doit matcher le nom de fichier exact) : le Fournisseur
# écrit en sortie est le nom propre, pas la clé. Un des 3 fichiers DADS ajoutés le 2026-10-30 a
# en plus un nom DIFFÉREMMENT cassé ("DADS  et de Saison", double espace, "D'Ardèche" absent du
# nom de fichier bien que présent dans le nom de la feuille) — même traitement, clé copiée
# telle quelle, Fournisseur écrit = le nom propre comme les autres.
#
# Numéro de lot stocké séparément de la description (pas une chaîne déjà assemblée) pour
# pouvoir le préfixer d'un zéro de façon uniforme (LOT 01 à LOT 09) via _lot_ardeche : les 12
# lots connus à ce jour vont de 1 à 12, donc un tri/affichage texte de "LOT 1"/"LOT 10"/"LOT 2"
# serait dans le désordre sans ce zéro — décidé le 2026-10-30, appliqué aussi aux 4 fichiers
# déjà traités (LOT 1/5/7 deviennent LOT 01/05/07 ; LOT 06 était déjà correct).
def _lot_ardeche(numero, description):
    return f"LOT {int(numero):02d} - {description}"


FICHIERS_ARDECHE = {
    "DADS D&#x27;Ard&#xE8;che et de Saison MP ARDECHE 24 - Lot 06  Crèmerie et produits "
    "frais Du vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx":
        ("DADS D'Ardèche et de Saison", 6, "CRÈMERIE ET PRODUITS FRAIS"),
    "Episaveurs MP ARDECHE 24 - Lot 1 Epicerie, boisson Du mardi 7 mai 2024 Au mardi 11 "
    "avril 2028.xlsx":
        ("Episaveurs", 1, "EPICERIE, BOISSON"),
    "LDC MP ARDECHE 24 - Lot 7 Volaille Du mercredi 1 janvier 2025 Au jeudi 15 avril 2027.xlsx":
        ("LDC", 7, "VOLAILLE"),
    "Passion Froid MP ARDECHE 24 - Lot 5 Autres produits de crèmerie Du mercredi 1 janvier "
    "2025 Au mercredi 14 avril 2027.xlsx":
        ("Passion Froid", 5, "AUTRES PRODUITS DE CRÈMERIE"),
    # Ajoutés le 2026-10-30 (8 nouveaux fichiers reçus le même jour dans a_formater/) :
    "Agricourt MP ARDECHE 24 - Lot 2 Epicerie BIO Du mardi 15 avril 2025 Au samedi 10 avril "
    "2032 (1).xlsx":
        ("Agricourt", 2, "EPICERIE BIO"),
    "DADS  et de Saison MP ARDECHE 24 - Lot 3 Jus de fruits en bouteilles, produits à base de "
    "chataîgne, compotes et vin Du vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx":
        ("DADS D'Ardèche et de Saison", 3,
         "JUS DE FRUITS EN BOUTEILLES, PRODUITS À BASE DE CHATAÎGNE, COMPOTES ET VIN"),
    "DADS D&#x27;Ard&#xE8;che et de Saison MP ARDECHE 24 - Lot 8 Boucherie saucisserie Du "
    "vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx":
        ("DADS D'Ardèche et de Saison", 8, "BOUCHERIE SAUCISSERIE"),
    "DADS D&#x27;Ard&#xE8;che et de Saison MP ARDECHE 24 - Lot 11 Poisson frais d'eau douce "
    "Du vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx":
        ("DADS D'Ardèche et de Saison", 11, "POISSON FRAIS D'EAU DOUCE"),
    "Passion Froid MP ARDECHE 24 - Lot 4 Surgelés Du mercredi 5 février 2025 Au jeudi 22 "
    "juillet 2027.xlsx":
        ("Passion Froid", 4, "SURGELÉS"),
    "Passion Froid MP ARDECHE 24 - Lot 9 Charcuterie Traiteur Du lundi 19 août 2024 Au jeudi "
    "13 avril 2028.xlsx":
        ("Passion Froid", 9, "CHARCUTERIE TRAITEUR"),
    "SYSCO FRANCE MP ARDECHE 24 - Lot 10 Produits frais de la mer Du lundi 27 janvier 2025 Au "
    "mardi 13 avril 2027.xlsx":
        ("SYSCO France", 10, "PRODUITS FRAIS DE LA MER"),
    "Vivalya MP ARDECHE 24 - Lot 12 Fruits et légumes frais et 4ème et 5ème gamme dont BIO Du "
    "vendredi 12 avril 2024 Au mardi 11 avril 2028.xlsx":
        ("Vivalya", 12, "FRUITS ET LÉGUMES FRAIS ET 4ÈME ET 5ÈME GAMME DONT BIO"),
}


# Reconnaissance d'UN SEUL fichier de a_formater/ (US-06 : pour brancher ces 5 familles sur
# pretraitement_auto.detecte_format()/transforme_automatiquement(), le point d'entrée unique).
# Par NOM DE FICHIER EXACT, pas par en-tête, contrairement à REGAL/GRADA/canonique : seule BPU
# DQE a une vraie validation de colonnes par texte (_trouve_ligne_entete_bpu) — Passion Froid,
# Episaveurs et BPUR FRESCA lisent des colonnes à position FIXE, sans aucune vérification de
# texte d'en-tête (voir leurs transforme_* respectifs) ; inventer une détection de contenu pour
# ces trois-là introduirait un risque de faux positif non validé. Se fier au nom de fichier,
# déjà répertorié explicitement pour ces 5 familles, est la même philosophie que les registres
# ci-dessus ("plutôt que de tout déduire dynamiquement avec le risque de se tromper en
# silence") — limite connue : un fichier JAMAIS VU, même strictement au même format, n'est pas
# reconnu tant qu'il n'est pas ajouté à un registre.
def _sans_extension_xlsx(nom):
    """Retire UNIQUEMENT une vraie extension .xlsx/.xls terminale, par opposition à
    os.path.splitext qui coupe au DERNIER point du nom quel qu'il soit — ce qui casse un nom
    déjà sans extension mais contenant un point ailleurs (ex: "1. BPUR FRESCA LOT 1 SURGELES
    CONV SIQO", le "1." étant un numéro de lot, pas une extension : os.path.splitext le
    tronquait à tort en juste "1", faisant échouer la reconnaissance de ce fichier précis -
    bug trouvé le 2026-10-30 lors d'un test manuel)."""
    if nom.lower().endswith((".xlsx", ".xls")):
        return os.path.splitext(nom)[0]
    return nom


def trouve_entree_connue(nom_fichier_ou_chemin):
    """Retrouve la clé exacte d'un des 5 registres (FICHIERS_BPU_DQE, FICHIER_PASSION_FROID,
    FICHIER_EPISAVEURS_XLS, FICHIER_BPUR_FRESCA, FICHIERS_ARDECHE) à partir d'un chemin ou d'un
    nom de fichier, en comparant SANS l'extension : un fichier déposé en mémoire (app.py,
    BytesIO) arrive à pretraitement_auto.py déjà amputé de son extension d'origine (voir
    nom_fournisseur dans transforme_automatiquement), donc la comparer telle quelle à une clé
    de registre qui la garde ("....xlsx") ne matcherait jamais. Renvoie la clé si reconnue,
    sinon None."""
    stem = pt._normalise(_sans_extension_xlsx(os.path.basename(str(nom_fichier_ou_chemin))))
    toutes_cles = (
        list(FICHIERS_BPU_DQE) + [FICHIER_PASSION_FROID, FICHIER_EPISAVEURS_XLS, FICHIER_BPUR_FRESCA]
        + list(FICHIERS_ARDECHE)
    )
    for cle in toutes_cles:
        if pt._normalise(_sans_extension_xlsx(cle)) == stem:
            return cle
    return None


# ============================================================================
# "TR__BPU_pour_intégration_Webgerest_CD45" — 13 fichiers (un par fournisseur/lot), reçus le
# 2026-11-12. PAS ENCORE enregistrée dans trouve_entree_connue/transforme_fichier_connu : par
# choix explicite de l'utilisatrice, cette famille reste à part pour l'instant (seule la
# transformation est demandée, pas encore le branchement automatique ni l'audit).
#
# Règle de correspondance établie à partir de MODELE_IMPORT_MARCHE EDILENE.xlsx (ligne 2 =
# en-tête canonique, ligne 3 = légende du nom utilisé dans les fichiers bruts) et vérifiée sur
# les 13 fichiers réels : feuille "BPU - LOT N", en-tête en ligne 3, données à partir de la
# ligne 5 (la ligne 4 est un texte-modèle pour les colonnes A à P, mais sert de VRAI sous-
# en-tête pour le bloc des 14 allergènes — voir _COL_ALLERGENES_DEPART ci-dessous). Colonnes A
# à M à position fixe, identique dans les 13 fichiers ; à partir de la colonne N, l'ensemble et
# l'ordre de (Bio/Label eligible EGAlim/Provenance ou Origine/Allergènes/Possibilité de
# déconditionner) varient d'un fichier à l'autre — repérées par texte d'en-tête, jamais par
# position.
# ============================================================================

def _est_nombre_bpu(valeur):
    return isinstance(valeur, (int, float))


# Sigles trouvés dans "Unité de vente" (colonne J), confirmés UN PAR UN en croisant avec le
# Conditionnement/Désignation de la même ligne sur les fichiers réels (pas devinés) — voir
# catalogue_erreurs.md. Mot complet/bien orthographié seulement : PAS encore la substitution
# officielle de CONTENANT_ABREGE (audit_general.py) — demandé explicitement par l'utilisatrice
# le 2026-11-12, tant qu'elle n'a pas la permission de faire ce choix métier.
SIGLES_UNITE_VENTE_BPU = {
    "ba": "Barquette", "bid": "Bidon", "bte": "Boîte", "ct": "Carton",
    "paq": "Paquet", "pch": "Poche", "sea": "Seau", "pl": "Plateau",
    "pi": "Pièce", "st": "Sachet", "sh": "Sachet", "bcl": "Bocal",
    "col": "Colis", "flc": "Flacon", "brq": "Brique",
    # Déjà en toutes lettres dans certains fichiers (pas des sigles à proprement parler),
    # gardés ici pour une seule recherche normalisée : juste la bonne casse.
    "sac": "Sac", "kg": "KG", "pc": "Pièce", "pu": "Pièce", "pot": "Pot",
}

# "BIB" (Bag-in-Box, confirmé via "Bib de 5 L"/"Bib 10kg" dans le Conditionnement) n'a aucun
# équivalent dans unites_commande_non_reconnues.md — décision métier en attente, voir
# catalogue_erreurs.md. Pas dans SIGLES_UNITE_VENTE_BPU : traité par repli (mot du
# Conditionnement recopié tel quel) comme n'importe quel mot non reconnu.

RE_PREMIER_MOT_BPU = re.compile(r"^\s*([A-Za-zÀ-ÿ]+)")


def _mot_recipient_conditionnement_bpu(conditionnement):
    """Premier mot du Conditionnement (ex: 'Bib 10kg' -> 'Bib', 'Carton de 6 boutelles' ->
    'Carton') — utilisé seulement pour le cas où 'Unité de vente' n'est qu'un multiplicateur
    nu (voir _unite_et_prix_vente_bpu). Recherche volontairement simple (premier mot
    alphabétique) : pas de vocabulaire de récipients à importer ici (éviterait d'importer
    audit_general.py, dangereux — voir catalogue_erreurs.md, bug n°41)."""
    if not conditionnement:
        return None
    m = RE_PREMIER_MOT_BPU.match(str(conditionnement))
    return m.group(1).capitalize() if m else None


def _unite_reference_et_poids_bpu(unite_min_brute):
    """'Unité minimum' (colonne H) -> (Unité de référence, Poids). Vu dans les 13 fichiers :
    'Kg'/'L'/'Pièce' presque partout (-> KG/L/PC, pas de poids connu par pièce à ce stade) ;
    APAGEH Lot 12 écrit '1 kg' partout (équivalent à 'Kg') ; POMONA Lot 7 a quelques lignes où
    c'est déjà un nom de contenant ('Colis', 'Boîte') — recopié tel quel pour l'instant (pas de
    règle PC/PAQUET appliquée ici, qui est du ressort de l'audit, pas de cette transformation)."""
    norm = pt._normalise(unite_min_brute)
    norm = re.sub(r"^1\s*", "", norm)  # "1 kg" -> "kg" (APAGEH)
    if norm in ("kg", "kilo", "kilogramme"):
        return "KG", None
    if norm in ("l", "litre"):
        return "L", None
    if norm in ("piece", "pièce", "pc"):
        return "PC", None
    return str(unite_min_brute).strip(), None


def _unite_et_prix_vente_bpu(unite_ref, prix_min, unite_vente_brute, prix_vente_brute, conditionnement):
    """'Unité de vente'/'Prix de l'unité de vente' (colonnes J/I) -> (unite_cmd, prix_cmd,
    commentaire_ou_None). Trois cas, en plus du cas normal :
    1. Les deux vides : "SI DIFFERENTE DE L'UNITE MINIMUM" (mode d'emploi du fichier source) —
       miroir direct de l'unité/prix minimum.
    2. Inversion confirmée sur BPF Lot 5 (100% des lignes utilisées, 44/44) : le sigle d'unité
       est dans la colonne Prix, le prix est dans la colonne Unité — corrigé en les
       échangeant avant de continuer.
    3. Multiplicateur nu confirmé sur LOGIFRAIS Lot 13 (42/71 lignes) : 'Unité de vente' n'est
       qu'un nombre (ex: 10, 24) — le vrai nom d'unité vient du Conditionnement, pas de cette
       colonne."""
    if prix_vente_brute is None and unite_vente_brute is None:
        return unite_ref, prix_min, None

    prix_vente, unite_vente = prix_vente_brute, unite_vente_brute
    commentaire = None
    if not _est_nombre_bpu(prix_vente) and _est_nombre_bpu(unite_vente):
        # Cas 2 : inversion.
        prix_vente, unite_vente = unite_vente, prix_vente
        commentaire = (
            "Prix et Unité de vente trouvés inversés dans le fichier source (le sigle d'unité "
            "était dans la colonne Prix, le prix dans la colonne Unité) — corrigé "
            "automatiquement avant transformation."
        )

    if _est_nombre_bpu(unite_vente):
        # Cas 3 (ou cas 2 non reconnu) : pas de mot d'unité du tout, juste un nombre -> on le
        # cherche dans le Conditionnement.
        mot = _mot_recipient_conditionnement_bpu(conditionnement)
        if mot:
            texte = (
                f"Unité de vente du fichier source n'est qu'un multiplicateur ({unite_vente}) "
                f"sans nom d'unité — repris du Conditionnement ('{conditionnement}')."
            )
            return mot, prix_vente, (commentaire + " " + texte if commentaire else texte)
        texte = (
            f"Unité de vente du fichier source n'est qu'un multiplicateur ({unite_vente}) sans "
            "nom d'unité, et aucun mot n'a été trouvé dans le Conditionnement pour la "
            "déduire — à compléter manuellement."
        )
        return None, prix_vente, (commentaire + " " + texte if commentaire else texte)

    sigle = pt._normalise(unite_vente)
    mot_connu = SIGLES_UNITE_VENTE_BPU.get(sigle)
    if mot_connu:
        return mot_connu, prix_vente, commentaire
    # Pas un sigle connu : déjà écrit en toutes lettres dans le fichier source (ex: "Barquette
    # 1,5 kg", "colis de 24 unités") — recopié tel quel.
    return str(unite_vente).strip(), prix_vente, commentaire


def _trouve_colonne_entete_bpu(ws, ligne, fragment):
    fragment = pt._normalise(fragment)
    for c in range(1, ws.max_column + 1):
        v = pt._normalise(ws.cell(row=ligne, column=c).value)
        if fragment in v:
            return c
    return None


ALLERGENES_ATTENDUS_BPU = ["Gluten", "Crustacés", "Oeufs", "Poissons", "Arachides", "Soja",
                           "Lait", "Fruits à coques", "Céleri", "Moutarde", "Sésame",
                           "Sulfites", "Lupins", "Mollusques"]

RE_FOURNISSEUR_NOM_FICHIER_BPU = re.compile(r"^(.*?)\s*-\s*lot\s*\d+", re.I)

# "Local" déduit de la Provenance : demandé par l'utilisatrice le 2026-11-12 pour cette famille
# de fichiers (CD45 = Conseil départemental du Loiret) — toute mention de "Loire" (Loiret,
# Indre-et-Loire, Pays de la Loire, "maraîchers loirétains"...) OU "Loir et Cher"/"Loir-et-Cher"
# (département voisin, nom différent de "Loire" mais inclus explicitement à sa demande) marque
# Local = 'X'. Recherche sur le texte normalisé (sans accents/casse), car la Provenance de ces
# fichiers est souvent une liste d'ingrédients ("Artichauts - Espagne; eau - Espagne..."), donc
# "loire"/"loir et cher" peuvent apparaître n'importe où dans le texte, pas juste au début.
RE_ORIGINE_LOCALE_BPU = re.compile(r"loire|loir\s*-?\s*et\s*-?\s*cher", re.I)


def _local_depuis_provenance_bpu(origine):
    if not origine:
        return None
    return "X" if RE_ORIGINE_LOCALE_BPU.search(pt._normalise(origine)) else None


# Même convention que _lot_ardeche ci-dessus (LOT 01 à LOT 09, zéro devant pour un tri/affichage
# texte correct) — demandé explicitement par l'utilisatrice pour cette famille aussi, les lots 1
# à 9 existant bien ici (Marché, colonne A, ex: "LOT 1 - VIANDES SURGELEES").
RE_MARCHE_LOT_NUMERO_BPU = re.compile(r"^(lot)\s*(\d+)(.*)$", re.I)


def _formate_marche_bpu(marche_brut):
    if not marche_brut:
        return marche_brut
    m = RE_MARCHE_LOT_NUMERO_BPU.match(str(marche_brut).strip())
    if not m:
        return marche_brut
    mot, numero, reste = m.groups()
    return f"{mot.upper()} {int(numero):02d}{reste}"


def _fournisseur_depuis_nom_fichier_bpu(nom_fichier):
    """Le nom du fournisseur n'est écrit NULLE PART dans ces fichiers (ni Page de Garde, ni
    cellule) — seulement dans le nom du fichier lui-même (ex: 'BPF - LOT 11 - BPU DQE -
    ...xlsx' -> 'BPF'), comme pour la famille MP ARDECHE 24. Tolère l'espace manquant avant
    'LOT' vu sur un des 13 fichiers ('SOCOPA -LOT 9')."""
    nom = _sans_extension_xlsx(os.path.basename(str(nom_fichier)))
    m = RE_FOURNISSEUR_NOM_FICHIER_BPU.match(nom)
    return m.group(1).strip() if m else nom


def transforme_bpu_dqe_webgerest_cd45(chemin_source, chemin_sortie, nom_fichier_pour_fournisseur=None):
    """Transforme un fichier de la famille "TR__BPU_pour_intégration_Webgerest_CD45" (13
    fichiers, un par fournisseur/lot). `nom_fichier_pour_fournisseur` : nom à utiliser pour en
    déduire le fournisseur (utile quand chemin_source est déjà un objet en mémoire, sans nom de
    fichier exploitable) — par défaut, celui de chemin_source.

    Renvoie (nb_lignes, avertissements) : `avertissements` est la liste des lignes où
    'Unité de vente' a demandé une correction (inversion, multiplicateur nu) ou n'a pas pu être
    déduite du tout — `ecrit_classeur_canonique` ne sait poser un commentaire que sur la cellule
    Unité de RÉFÉRENCE, pas sur l'Unité de commande, donc ce n'est pas écrit dans le fichier de
    sortie pour l'instant, juste renvoyé ici pour revue."""
    wb = openpyxl.load_workbook(chemin_source, data_only=True)
    nom_aba_bpu = next((n for n in wb.sheetnames if pt._normalise(n).startswith("bpu")), None)
    if not nom_aba_bpu:
        raise CauseFormatInvalide(
            "ONGLET_MANQUANT", f"Aucun onglet 'BPU - LOT N' trouvé dans {chemin_source}.",
        )
    ws = wb[nom_aba_bpu]

    fournisseur = _fournisseur_depuis_nom_fichier_bpu(
        nom_fichier_pour_fournisseur or chemin_source)

    manquantes = [nom for nom in ("desig", "unite min", "prix")
                  if _trouve_colonne_entete_bpu(ws, 3, nom.split()[0]) is None]
    if manquantes:
        raise CauseFormatInvalide(
            "COLONNE_OBLIGATOIRE_ABSENTE",
            f"Colonnes essentielles introuvables dans « {nom_aba_bpu} » ({chemin_source}).",
            onglet=nom_aba_bpu, colonne=", ".join(manquantes), cellule="ligne 3",
        )

    col_origine = (_trouve_colonne_entete_bpu(ws, 3, "provenance")
                   or _trouve_colonne_entete_bpu(ws, 3, "origine"))
    col_bio = _trouve_colonne_entete_bpu(ws, 3, "bio")
    col_label = _trouve_colonne_entete_bpu(ws, 3, "label eligible")
    col_allergenes = _trouve_colonne_entete_bpu(ws, 3, "allergen")

    lignes = []
    avertissements = []
    for r in range(5, ws.max_row + 1):
        designation = ws.cell(row=r, column=3).value
        if not designation:
            continue
        marche = _formate_marche_bpu(ws.cell(row=r, column=1).value)
        famille = ws.cell(row=r, column=2).value
        desig_fourn = ws.cell(row=r, column=4).value
        reference = ws.cell(row=r, column=5).value
        tva = ws.cell(row=r, column=6).value
        prix_min = ws.cell(row=r, column=7).value
        unite_min_brute = ws.cell(row=r, column=8).value
        prix_vente_brut = ws.cell(row=r, column=9).value
        unite_vente_brute = ws.cell(row=r, column=10).value
        conditionnement = ws.cell(row=r, column=13).value

        unite_ref, poids = _unite_reference_et_poids_bpu(unite_min_brute)
        unite_cmd, prix_cmd, commentaire_vente = _unite_et_prix_vente_bpu(
            unite_ref, _arrondi_prix(prix_min), unite_vente_brute, prix_vente_brut, conditionnement)
        if commentaire_vente:
            avertissements.append((r, designation, commentaire_vente))

        # Vérifié sur les 13 fichiers réels : la seule valeur utilisée dans cette colonne est
        # le mot "BIO" lui-même (jamais "OUI"/"X" malgré le menu déroulant source qui propose
        # les deux) — toute valeur non vide y suffit, donc testée au sens large plutôt que sur
        # une liste de mots précis qui aurait pu en rater d'autres.
        bio = "X" if col_bio and ws.cell(row=r, column=col_bio).value else None
        label = ws.cell(row=r, column=col_label).value if col_label else None
        origine = ws.cell(row=r, column=col_origine).value if col_origine else None
        local = _local_depuis_provenance_bpu(origine)

        allergenes = {nom: None for nom in pt.ALLERGENES}
        if col_allergenes:
            for i, nom in enumerate(ALLERGENES_ATTENDUS_BPU):
                v = ws.cell(row=r, column=col_allergenes + i).value
                if v and pt._normalise(v) in ("x", "o", "oui"):
                    allergenes[pt.ALLERGENES[i]] = "o"

        ligne = pg._ligne_vide_de_base(
            fournisseur=fournisseur, lot=famille, designation=designation, unite_ref=unite_ref,
            poids=poids, commentaire_unite=None, cond=conditionnement,
            desig_fourn=desig_fourn, ref=reference, prix_ref=_arrondi_prix(prix_min),
            unite_cmd=unite_cmd, prix_cmd=_arrondi_prix(prix_cmd), local=local, bio=bio,
            label=label, origine=origine, commentaire_origine=None, marche=marche,
        )
        ligne["allergenes"] = allergenes
        ligne["tva"] = tva
        lignes.append(ligne)

    nb_lignes = pt.ecrit_classeur_canonique(lignes, chemin_sortie, nom_aba_bpu)
    return nb_lignes, avertissements


def transforme_fichier_connu(chemin_source, chemin_sortie, cle_registre):
    """Transforme UN SEUL fichier déjà reconnu par trouve_entree_connue — le pendant, pour les
    5 familles de a_formater/, de transforme_fichier_grada_auto (pretraitement_auto.py) pour
    GRADA. `chemin_source` peut être un chemin sur disque (CLI) ou déjà un .xlsx en mémoire
    (BytesIO) si pretraitement_auto.py a dû le convertir depuis un .xls avant d'arriver ici."""
    if cle_registre in FICHIERS_BPU_DQE:
        return transforme_bpu_dqe(
            chemin_source, chemin_sortie, FICHIERS_BPU_DQE[cle_registre],
            lot_force=LOT_FORCE.get(cle_registre),
        )
    if cle_registre == FICHIER_PASSION_FROID:
        return transforme_passion_froid(chemin_source, chemin_sortie)
    if cle_registre == FICHIER_EPISAVEURS_XLS:
        # Le .xls Episaveurs a besoin de SA PROPRE conversion LibreOffice (colonne par
        # colonne, lue ensuite comme un .xlsx) : si pretraitement_auto.py a déjà converti le
        # fichier plus haut (cas d'un .xls déposé via l'app), chemin_source est déjà un .xlsx
        # en mémoire, à lire directement ; sinon (CLI, chemin sur disque vers le vrai .xls),
        # passer par la conversion dédiée.
        if isinstance(chemin_source, str) and chemin_source.lower().endswith(".xls"):
            return _convertit_et_transforme_episaveurs(chemin_source, chemin_sortie)
        return transforme_episaveurs(chemin_source, chemin_sortie)
    if cle_registre == FICHIER_BPUR_FRESCA:
        return transforme_bpur_fresca(chemin_source, chemin_sortie)
    if cle_registre in FICHIERS_ARDECHE:
        fournisseur, numero, description = FICHIERS_ARDECHE[cle_registre]
        return transforme_ardeche(
            chemin_source, chemin_sortie, fournisseur, _lot_ardeche(numero, description)
        )
    raise CauseFormatInvalide(
        "FORMAT_NON_SUPPORTE", f"« {cle_registre} » n'est dans aucun registre connu de a_formater/.",
    )


def _convertit_et_transforme_episaveurs(chemin_source, chemin_sortie):
    import pretraitement_auto as pa  # import local : évite un import circulaire au chargement du module
    with open(chemin_source, "rb") as f:
        xlsx_en_memoire = pa.convertit_xls_en_xlsx(f)
    return transforme_episaveurs(xlsx_en_memoire, chemin_sortie)


def transforme_dossier_a_formater(dossier_entree=DOSSIER_ENTREE, dossier_sortie=DOSSIER_SORTIE):
    os.makedirs(dossier_sortie, exist_ok=True)
    resultats = {}

    for nom_fichier, nom_feuille in FICHIERS_BPU_DQE.items():
        chemin_source = os.path.join(dossier_entree, nom_fichier)
        if not os.path.exists(chemin_source):
            continue
        nom_sortie = os.path.splitext(nom_fichier)[0].strip() + " (formaté).xlsx"
        chemin_sortie = os.path.join(dossier_sortie, nom_sortie)
        resultats[nom_fichier] = transforme_bpu_dqe(
            chemin_source, chemin_sortie, nom_feuille, lot_force=LOT_FORCE.get(nom_fichier)
        )

    chemin_pf = os.path.join(dossier_entree, FICHIER_PASSION_FROID)
    if os.path.exists(chemin_pf):
        chemin_sortie = os.path.join(dossier_sortie, "Passion Froid (formaté).xlsx")
        resultats[FICHIER_PASSION_FROID] = transforme_passion_froid(chemin_pf, chemin_sortie)

    chemin_ep = os.path.join(dossier_entree, FICHIER_EPISAVEURS_XLS)
    if os.path.exists(chemin_ep):
        chemin_sortie = os.path.join(dossier_sortie, "Episaveurs (formaté).xlsx")
        resultats[FICHIER_EPISAVEURS_XLS] = _convertit_et_transforme_episaveurs(chemin_ep, chemin_sortie)

    chemin_fresca = os.path.join(dossier_entree, FICHIER_BPUR_FRESCA)
    if os.path.exists(chemin_fresca):
        chemin_sortie = os.path.join(dossier_sortie, "BPUR FRESCA Lot 1 (formaté).xlsx")
        resultats[FICHIER_BPUR_FRESCA] = transforme_bpur_fresca(chemin_fresca, chemin_sortie)

    for nom_fichier, (fournisseur, numero, description) in FICHIERS_ARDECHE.items():
        chemin_source = os.path.join(dossier_entree, nom_fichier)
        if not os.path.exists(chemin_source):
            continue
        # Le numéro de lot fait partie du nom de sortie : depuis le 2026-10-30, plusieurs
        # lots partagent le même fournisseur (ex: DADS a les lots 03/06/08/11) — sans lui,
        # ils s'écraseraient silencieusement les uns les autres au même nom de fichier.
        nom_sortie = f"{fournisseur} MP ARDECHE 24 Lot {numero:02d} (formaté).xlsx"
        chemin_sortie = os.path.join(dossier_sortie, nom_sortie)
        resultats[nom_fichier] = transforme_ardeche(
            chemin_source, chemin_sortie, fournisseur, _lot_ardeche(numero, description)
        )

    return resultats


if __name__ == "__main__":
    for nom_fichier, nb_lignes in transforme_dossier_a_formater().items():
        print(f"{nom_fichier} -> {nb_lignes} lignes")
