"""
Cause structurée pour signaler qu'un fichier n'est pas au format attendu (US-06, critères 2
et 3) : code d'erreur + emplacement (onglet/colonne/cellule), au lieu d'un message brut en
texte libre qui ne fait aujourd'hui que planter l'appelant (`raise ValueError(...)`).

Fichier séparé, à part des scripts qui l'utilisent (pretraitement.py, pretraitement_grada.py,
pretraitement_a_formater.py, pretraitement_auto.py), pour que les deux premiers puissent
l'importer sans dépendre de pretraitement_auto.py (qui, lui, importe les quatre) — éviter tout
import circulaire.

Les 4 catégories minimales demandées par US-06, plus une au-delà du minimum :
ONGLET_MANQUANT, ENTETE_MODIFIE, COLONNE_OBLIGATOIRE_ABSENTE, FORMAT_NON_SUPPORTE,
SECTION_INTROUVABLE — voir detection_format_causes.md pour le détail, un exemple réel par
catégorie, et la correspondance avec les `raise ValueError` déjà existants avant ce changement.
"""


class CauseFormatInvalide(ValueError):
    """Toujours une ValueError (compatible avec le code existant qui attraperait encore
    `except ValueError`), mais porte en plus les champs structurés demandés par US-06.2."""

    def __init__(self, code, message, onglet=None, colonne=None, cellule=None):
        super().__init__(message)
        self.code = code
        self.onglet = onglet
        self.colonne = colonne
        self.cellule = cellule

    def cause(self):
        """{code, message, onglet, colonne, cellule} — forme exposée à l'appelant (ex:
        app.py), prête à afficher ou à sérialiser en JSON."""
        return {
            "code": self.code,
            "message": str(self),
            "onglet": self.onglet,
            "colonne": self.colonne,
            "cellule": self.cellule,
        }
