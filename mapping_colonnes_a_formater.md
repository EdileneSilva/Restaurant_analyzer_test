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
| UNITE COMMANDE | — | — | — | — |
| PRIX UNITE COMMANDE | — | — | — | — |
| Local (X si local) | — | — | — | Origine matière première⁸ |
| Bio (X si bio) | — | label | Label qualité éligible égalim | Signes⁹ |
| Circuit court | — | — | — | — |
| Label | — | Label | Autre label | Signes¹⁰ |
| Origine | — | — | Origine de la matière première | Origine matière première¹¹ |
| Indice I0 | — | — | — | — |

### Notes (MP ARDECHE 24 — toutes datées, confirmées par le métier)

1. **Lot** — Vient du nom du fichier (ex: « Lot 06  Crèmerie et produits frais »), pas d'une
   colonne — même logique que la famille BPU DQE.
2. **Unité de référence (BPU) — UC** — Codes abrégés : KG, LT, BT, UN. Correspondance
   confirmée (2026-10-30) : KG→KG, LT→L, UN→PIECE, BT→pas toujours « bouteille »
   (sachet/sac/boîte selon les cas) — le mot de contenant réellement écrit dans la désignation
   est utilisé quand il existe (ex: « sachet », « pot »), sinon secours sur le code lui-même.
   Un code de contenant en N/M en fin de désignation (« bt 3/1 », « bt 4/4 ») prime sur tout :
   Unité de commande = « BOITE N/M » au complet (pas juste « BOITE »), format déjà accepté —
   voir `unites_commande_non_reconnues.md`.
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
   « 14,24KG »), ou nombre de pièces/boîtes dans les autres cas (ex: DADS « Yaourt ... (x48)
   125 g », Cond.=48 → « 48 X 125G » ; confirmé cohérent avec le « (xN) » du Nom technique). La
   colonne « Commentaire » (ex: « COLIS DE 4 POCHES », vue chez LDC) n'est PAS utilisée :
   présente sur trop peu de lignes pour être fiable, et l'info utile (poids/N) est déjà
   couverte par Cond. + le texte de Nom technique.
6. **Désignation fournisseur** — Même colonne que Désignation de l'article : pas de 2e champ
   dédié fournisseur, comme pour Passion Froid/Episaveurs déjà traités.
7. **PRIX UNITE REFERENCE (BPU) — PU** — = Prix ÷ Fact. (vérifié cohérent partout) : c'est le
   prix d'UNE unité de COMMANDE (UC), pas forcément celui de l'unité de référence (BPU) — les
   deux ne coïncident que quand UC est déjà KG/L. Dans les autres cas (UC = boîte/pièce/
   sachet...), le prix de référence est recalculé en divisant PU par le poids/volume net trouvé
   en fin de Nom technique.
8. **Local (X si local) — Origine matière première** — DÉCIDÉ (2026-10-30, confirmé par
   l'utilisateur) : Local = X quand « Origine matière première » ne contient QUE des codes de
   la région Auvergne-Rhône-Alpes (AURA) : ARDECHE, AURA, ISERE, SAVOIE/HAUTE-SAVOIE, AIN,
   DROME, RHONE, HAUTE LOIRE (liste étendue le 2026-10-30 avec les 8 nouveaux fichiers).
   VAUCLUSE et BOUCHES DU RHONE sont bien françaises (Origine = France) mais PAS locales
   (région PACA, pas AURA) — jamais Local même combinées à un département AURA (ex:
   « ISERE/VAUCLUSE » reste non local). « France » tout court ne compte pas non plus comme
   local (trop générique). Implémenté dans `_local_ardeche`/`REGIONS_LOCALES_ARDECHE`
   (`pretraitement_a_formater.py`).
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
11. **Origine — Origine matière première** — DÉCIDÉ (2026-10-30, confirmé par l'utilisateur) :
    « Origine matière première » retenue, pas « Pays de transformation ». Reconnaît aussi tous
    les départements AURA ET les départements PACA limitrophes vus dans les fichiers (AIN,
    DROME, RHONE, VAUCLUSE, BOUCHE(S) DU RHONE, HAUTE LOIRE, REUNION) comme « France » — voir
    `REGIONS_FRANCE` (`pretraitement_a_formater.py`). Cas particulier « IMPORT » (vu 2x,
    Agricourt) : laissé tel quel avec un commentaire de vérification plutôt que remplacé par
    « France » par défaut, ce qui serait faux (le mot dit explicitement que ce n'est PAS
    français).
