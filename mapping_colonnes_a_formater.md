# Mapping colonnes — fichiers à formater (hors REGAL / GRADA / canonique)

Version Markdown (suivie par git) de `mapping_colonnes_a_formater.xlsx` (non suivi : fichier
`.xlsx`, ignoré par `.gitignore`). Correspondance entre le format canonique (IANORD) et les
formats fournisseurs encore à transformer, un par un, colonne par colonne.

**Légende** : nom de colonne = équivalent trouvé dans le fichier source. `—` = aucune colonne
équivalente trouvée (case vide + fond jaune dans le classeur d'origine). Une note sous le
tableau signale un choix ambigu déjà tranché, ou encore ouvert.

À relire et mettre à jour manuellement si le classeur `.xlsx` est modifié directement (ce
fichier n'est pas régénéré automatiquement depuis le `.xlsx`).

## BPU DQE (6 feuilles : lots 1 à 5 + Epicerie)

| Colonne du format canonique (IANORD) | Lot 1 (Viandes et volailles) | Lot 2 (Poissons) | Lot 3 (Légumes) | Lot 4 (PC et Entrées) | Lot 5 (Glaces et desserts) | Epicerie |
|---|---|---|---|---|---|---|
| Marché | — | — | — | — | — | — |
| Lot | Titre | Titre | Titre | Titre | Titre | Titre |
| Désignation de l'article | Libellés | Libellés | Libellés | Libellés | Libellés | Libellé Produit |
| Unité de référence (BPU) | Unité | Unité | Unité | Unité | Unité | Unité |
| Poids à l'unité de référence | Format proposé | Format proposé | Format proposé | Format proposé | Format proposé | Format (proposé) |
| Nombre de Pièces dans conditionnement | — | — | — | — | — | — |
| Fournisseur (nom) | — | — | — | — | — | — |
| Conditionnement | Conditionnement proposé | Conditionnement proposé | Conditionnement proposé | Conditionnement proposé | Conditionnement proposé | Conditionnement (proposé) |
| Désignation fournisseur | Libellé Produit Fournisseur | Libellé Produit Fournisseur | Libellé Produit Fournisseur | Libellé Produit Fournisseur | Libellé Produit Fournisseur | Libellé produit (fournisseur) |
| Référence produit | Références Fournisseurs | Références Fournisseurs | Références Fournisseurs | Références Fournisseurs | Références Fournisseurs | Références (fournisseur) |
| PRIX UNITE REFERENCE (BPU) | PU HT | PU HT | PU HT | PU HT | PU HT | PU HT |
| UNITE COMMANDE | — | — | — | — | — | — |
| PRIX UNITE COMMANDE | — | — | — | — | — | — |
| Local (X si local) | — | — | — | — | — | — |
| Bio (X si bio) | chercher dans egalim | — | — | — | — | — |
| Circuit court | — | — | — | — | — | — |
| Label | OBLIGATOIRE | OBLIGATOIRE | OBLIGATOIRE | OBLIGATOIRE | — | Démarche labellisée proposée (EGALIM) |
| Origine | — | — | — | — | — | Provenance |
| Indice I0 | — | — | — | — | — | — |

Aucune note (aucun commentaire) sur ces 6 colonnes dans le classeur d'origine — mapping jugé
non ambigu. `Fournisseur (nom)` et `Lot` viennent en réalité du nom du fichier / de la page de
garde, pas d'une colonne du tableau de données (voir `pretraitement_a_formater.py`,
`_fournisseur_depuis_page_de_garde` / `_lot_depuis_page_de_garde` / `_lot_depuis_titre`).

## Passion Froid / Episaveurs / BPUR FRESCA / MP ARDECHE 24

| Colonne du format canonique (IANORD) | Passion Froid | Episaveurs | BPUR FRESCA (Lot 1 Surgelés Conv/SIQO) | MP ARDECHE 24 (DADS / Episaveurs / LDC / Passion Froid / Agricourt / SYSCO France / Vivalya — 12 fichiers, 2026-10-30) |
|---|---|---|---|---|
| Marché | — | — | — | — |
| Lot | Gamme | Groupe | Titre | Titre¹ |
| Désignation de l'article | Libellé | Désignation fournisseur | Intitulé | Nom technique |
| Unité de référence (BPU) | Unité de facturation | Unité de Facturation | Unité de facturation | UC² |
| Poids à l'unité de référence | — | Poids UF | Calibre proposé | — ³ |
| Nombre de Pièces dans conditionnement | — | — | — | — |
| Fournisseur (nom) | Titre | Titre | Titre | Titre⁴ |
| Conditionnement | Conditionnement minimum | Conditionnement | Conditionnement proposé | Cond.⁵ |
| Désignation fournisseur | Libellé | Désignation fournisseur | Dénomination du produit par le candidat | Nom technique⁶ |
| Référence produit | Code Article | Code produit | — | Réf |
| PRIX UNITE REFERENCE (BPU) | Prix de vente Brut H.T | Prix de vente Brut INTEGRATION | Montant HT UNITAIRE | PU⁷ |
| UNITE COMMANDE | — | — | — | UC¹² |
| PRIX UNITE COMMANDE | — | — | — | Prix¹³ |
| Local (X si local) | — | — | — | Origine matière première⁸ |
| Bio (X si bio) | — | label | Label qualité éligible égalim | Signes⁹ |
| Circuit court | — | — | — | — |
| Label | — | Label | Autre label | Signes¹⁰ |
| Origine | — | — | Origine de la matière première | Pays de transformation¹¹ |
| Indice I0 | — | — | — | — |

### Notes (MP ARDECHE 24 — toutes datées, confirmées par le métier)

1. **Lot** — Vient du nom du fichier (ex: « Lot 06  Crèmerie et produits frais »), pas d'une
   colonne — même logique que la famille BPU DQE.
2. **Unité de référence (BPU) — UC** — Codes abrégés : KG, LT, BT, UN. Correspondance
   confirmée : KG→KG, LT→L, UN→UNITÉ (jusqu'au 2026-11-09 : « UN→PIECE » ; changé à la demande
   de l'utilisateur, qui préfère « UNITÉ »). Pour l'Unité de référence et le Conditionnement
   affichés (mais **pas** pour la colonne UNITE COMMANDE elle-même — voir note 12, qui recopie
   le code source tel quel), un mot de contenant réellement écrit dans la désignation est en
   plus utilisé quand il existe (ex: « sachet », « pot », « seau », « brique »), sinon secours
   sur le code lui-même (ex: « BT » sans aucun mot trouvé dans la désignation reste « BT », pas
   deviné). Un code de contenant en N/M en fin de désignation (« bt 3/1 », « bt 4/4 ») prime sur
   tout : le mot retenu devient « BOITE N/M » au complet (pas juste « BOITE »), format déjà
   accepté — voir `unites_commande_non_reconnues.md`.
3. **Poids à l'unité de référence** — Pas de colonne poids par pièce directe. Le texte de
   « Nom technique » contient parfois un poids en fin de texte (ex: « (x25) 100 g »), mais le
   format n'est pas assez régulier (parfois une fourchette « 90/110 g ») pour l'extraire de
   façon fiable sans risquer d'inventer une valeur — laissé vide.
4. **Fournisseur (nom)** — Préfixe du nom de fichier, avant « MP ARDECHE 24 » (DADS,
   Episaveurs, LDC, Passion Froid, Agricourt, SYSCO France, Vivalya) — ces fichiers ne sont PAS
   le même fournisseur que les colonnes « Passion Froid »/« Episaveurs » déjà traitées plus à
   gauche (format totalement différent, juste un nom en commun pour 2 d'entre eux). Plusieurs
   fichiers peuvent partager le même fournisseur avec un Lot différent (DADS : lots
   03/06/08/11 ; Passion Froid : lots 04/05/09) — le nom de sortie inclut donc le numéro de lot
   pour ne pas s'écraser.
5. **Conditionnement — Cond.** — « Cond. » est un nombre nu, exprimé dans l'unité de commande
   elle-même : poids/volume total du colis quand UC est déjà KG/L (ex: LDC, Cond.=14,240 →
   « 14,24KG », ou « SEAU DE 10L » si un mot de contenant est trouvé dans la désignation), ou
   nombre de pièces/contenants dans les autres cas (ex: DADS « Yaourt ... pot carton (x48)
   125 g », Cond.=48 → « 48 POTS DE 125G » ; confirmé cohérent avec le « (xN) » du Nom
   technique). Quand UC code une unité vendue individuellement (UN, ou un code boîte N/M), le
   nombre devient un compte au pluriel plutôt qu'un multiplicateur de prix — « 48 POTS DE
   125G », jamais « POT DE 48 X 125G », qui se lirait à tort comme « un pot contenant 48×125g »
   plutôt que « 48 pots de 125g chacun » (et ferait calculer un prix au kilo 48× trop petit) ;
   quand aucun poids/volume n'est trouvé, juste « 30 UNITÉS » (ex: « Oeuf ... (x30) »). Détail
   demandé le 2026-11-09, après avoir vu ce cas précis sur le fichier DADS Lot 06. La colonne
   « Commentaire » (ex: « COLIS DE 4 POCHES », vue chez LDC) n'est PAS utilisée : présente sur
   trop peu de lignes pour être fiable, et l'info utile (poids/N) est déjà couverte par Cond. +
   le texte de Nom technique.
6. **Désignation fournisseur** — Même colonne que Désignation de l'article : pas de 2e champ
   dédié fournisseur, comme pour Passion Froid/Episaveurs déjà traités.
7. **PRIX UNITE REFERENCE (BPU) — PU** — = Prix ÷ Fact. (vérifié cohérent partout) : c'est le
   prix d'UNE unité de COMMANDE (UC), pas forcément celui de l'unité de référence (BPU) — les
   deux ne coïncident que quand UC est déjà KG/L. Dans les autres cas (UC = boîte/pièce/
   sachet...), le prix de référence est recalculé en divisant PU par le poids/volume net trouvé
   en fin de Nom technique. Cette division par Fact. ne sert QUE pour ce calcul interne du prix
   de référence — la colonne PRIX UNITE COMMANDE (note 13) recopie « Prix » brut, sans la
   diviser par Fact.
8. **Local (X si local) — Origine matière première** — DÉCIDÉ (2026-10-30, confirmé par
   l'utilisateur) : Local = X quand « Origine matière première » ne contient QUE des codes de
   la région Auvergne-Rhône-Alpes (AURA) : ARDECHE, AURA, ISERE, SAVOIE/HAUTE-SAVOIE, AIN,
   DROME, RHONE, HAUTE LOIRE (liste étendue le 2026-10-30 avec les 8 nouveaux fichiers).
   VAUCLUSE et BOUCHES DU RHONE sont bien françaises mais PAS locales (région PACA, pas AURA) —
   jamais Local même combinées à un département AURA (ex: « ISERE/VAUCLUSE » reste non local).
   « France » tout court ne compte pas non plus comme local (trop générique). Implémenté dans
   `_local_ardeche`/`REGIONS_LOCALES_ARDECHE` (`pretraitement_a_formater.py`). Depuis le
   2026-11-09, c'est la SEULE utilisation de « Origine matière première » pour cette famille de
   fichiers — l'Origine BPU elle-même vient désormais de « Pays de transformation » (note 11),
   une colonne différente.
9. **Bio (X si bio) — Signes** — Tag « AB » = Bio, ainsi que tout tag contenant le mot « bio »
   (ex: « Bio Part. », vu sur des produits partiellement bio, 2026-10-30). « ABE » apparaît
   toujours accolé à un autre signal bio (jamais seul) — traité comme partie du signal bio, pas
   gardé comme Label séparé. « null » (texte littéral, vu dans plusieurs fichiers DADS) est un
   artefact d'export sans signification — retiré silencieusement, ni Bio ni Label.
10. **Label — Signes** — Tags restants après Bio : AOP, HVE, LR, IGP, CE2 = labels/
    certifications connus. Confirmés par recherche le 2026-10-30 : VBF = Viande Bovine
    Française, VOF = Viande Ovine Française, VVF = Viande de Veau Française, LPF = Le Porc
    Français (label officiel filière porcine), ASC = Aquaculture Stewardship Council
    (équivalent MSC pour l'aquaculture), RUP = Région Ultrapériphérique (SIQO officiel,
    produits d'outre-mer, ex: banane). Probables mais non confirmés officiellement : VF
    (Volaille Française, fichier LDC), FLF (Fruits et Légumes de France), ODF (Oeufs de
    France), BZH (Bretagne), VGN (Vegan). Sens inconnu malgré recherche, gardés tels quels : GG
    (saumon), MGDF (magret de canard), FFL (poivre), CER (vanille/brownie), « 16 » (code
    numérique, produits Cantadou).
11. **Origine — Pays de transformation** — Changé le 2026-11-09 (demandé par le métier) :
    jusque-là (décision du 2026-10-30, maintenant dépassée) c'était « Origine matière
    première » qui alimentait l'Origine BPU — désormais cette colonne n'alimente plus que
    « Local » (note 8), et l'Origine BPU vient de « Pays de transformation », qui ne contient
    jamais un nom de région française (toujours un pays entier, ex: « FRANCE », « UE »), donc
    ne déclenche jamais la reconnaissance de région de `_origine_bpu`/`REGIONS_FRANCE`
    (gardée malgré tout dans le code, car réutilisée par d'autres formats). Conséquence directe :
    le cas particulier « IMPORT » documenté ici avant (vu 2x chez Agricourt, dans « Origine
    matière première ») ne touche plus l'Origine BPU — ces mêmes lignes ont « FRANCE »/« UE »
    en « Pays de transformation », donc une Origine propre sans ambiguïté ; « IMPORT » continue
    d'exister dans « Origine matière première », mais cette colonne ne sert plus qu'à décider
    Local (note 8), où elle ne correspond de toute façon à aucune région AURA.

12. **UNITE COMMANDE — UC** — Depuis le 2026-11-09 (demandé explicitement par l'utilisateur :
    « recuperemos a coluna UC como ela é »), cette colonne recopie le code source "UC" tel
    quel, SEULEMENT traduit vers le vocabulaire canonique (KG→KG, LT→L, UN→UNITÉ — une
    traduction du même code, jamais une supposition à partir du texte libre de la désignation).
    Avant cette date, un mot de contenant trouvé dans la désignation (ex: « pot », « seau »)
    pouvait remplacer le code brut ici aussi, comme pour l'Unité de référence/le Conditionnement
    (note 2) — source d'un vrai bug trouvé le même jour : un « UN » sans ambiguïté (prix déjà
    par pièce) se faisait écraser par un mot de contenant coïncidant dans le texte (ex: « pot »
    dans « Yaourt ... pot carton (x48) 125 g »), avec le risque de multiplier le prix par le
    nombre de pièces du conditionnement alors qu'il ne le fallait pas. Toute incohérence entre
    cette valeur brute et ce que le Conditionnement/la désignation suggèrent est désormais du
    ressort de l'audit (`verifie_unite_commande_etendue`, dans `audit_general.py`/
    `audit_general_commentaires.py`), pas de la mise en forme.

13. **PRIX UNITE COMMANDE — Prix** — Recopie la colonne « Prix » brute de la source, SANS la
    diviser par « Fact. » (contrairement au Prix de référence, note 7) : demandé le 2026-11-09,
    après un premier essai en Prix ÷ Fact. qui ne correspondait plus à ce que dit la colonne
    « Prix » source. Une éventuelle incohérence entre cette valeur et l'Unité de commande/le
    Conditionnement (ex: Unité de commande = « L » alors que le prix couvre tout un
    Conditionnement de plusieurs litres) est, là aussi, volontairement laissée telle quelle ici
    — à l'audit de la détecter, pas à la mise en forme de la corriger en silence.
