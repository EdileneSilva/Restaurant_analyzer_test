# Test manuel — cohérence formatage + audit dans Streamlit

Ce document répond à US-08, critère 3 (« Le POC est testé par le métier sur des fichiers réels »
— et, dans une moindre mesure, à l'équivalent côté audit seul, US-05.3), restés tous les deux
sans trace dans ce dépôt jusqu'ici. C'est un guide de test MANUEL, à exécuter par une personne
dans le navigateur — rien ici n'est automatisé ni vérifié par du code.

## User Story

> En tant que chargé(e) de marché qui valide le POC, je veux déposer des fichiers réels variés
> dans l'application Streamlit et vérifier moi-même, fichier par fichier, que le résultat du
> formatage ET de l'audit sont cohérents avec ce que j'attends du fichier source, afin de
> confirmer que la chaîne complète (détection → formatage → audit) est fiable avant de la
> considérer validée par le métier.

**Critères d'acceptation :**
1. Pour chaque fichier testé (voir l'échantillon ci-dessous), le statut annoncé par l'app
   correspond à ce qu'on attend (pas de "format non reconnu" pour un fichier déjà couvert, pas
   de transformation pour un fichier déjà canonique).
2. Le fichier **formaté** (téléchargeable avant l'audit) a les bonnes valeurs dans les bonnes
   colonnes — vérifié en comparant quelques lignes au fichier source original, pas juste en
   confiance.
3. Le fichier **audité**, dans CHACUN des deux modes (complet / commentaires seulement), signale
   des incohérences qui font sens (pas de faux positif évident), et la différence entre les deux
   modes correspond à ce qui est annoncé (le mode complet corrige un peu plus, silencieusement).
4. Un retour écrit (ce qui marche, ce qui semble faux, ce qui est ambigu) est noté quelque part
   — ce document peut servir de brouillon de notes, à recopier ensuite dans Jira/ailleurs.

## Échantillon suggéré (pas besoin de tester les 21 fichiers de `a_formater/` un par un)

Un représentant par famille suffit pour une première passe — n'étendre à plus de fichiers de la
même famille que si quelque chose semble faux sur le premier :

- [ ] Un fichier déjà canonique (`pret_pour_audit/` ou `a_inspecter/`) — doit aller direct à
      l'audit, sans aucune note de transformation.
- [ ] Un fichier BPU DQE (ex: Lot 1 ou Lot 5 — celui-ci a le cas particulier du nom de Lot
      recombiné depuis 2 sources, voir `pretraitement_a_formater.py`).
- [ ] `Passion froid - Tarif général...xlsx`.
- [ ] `Episaveurs - Tarifs Catalogue complémentaire...xls` (le seul `.xls` du lot — vérifie au
      passage la conversion automatique).
- [ ] `1. BPUR FRESCA LOT 1...xlsx`.
- [ ] Un fichier "MP ARDECHE 24" avec unité de référence dérivée d'un poids/volume en fin de
      désignation (ex: LDC, ou DADS Lot 06 — comparer le prix BPU recalculé avec un calcul à la
      main sur 2-3 lignes).
- [ ] Un fichier "MP ARDECHE 24" avec des codes boîte (ex: Episaveurs MP ARDECHE, lignes "bt
      3/1"/"bt 4/4" — vérifier que l'Unité de référence devient bien "BOITE 3/1" complet).
- [ ] Un fichier jamais vu (pas dans les registres — ex: un des fichiers "MP ARDECHE 24" lu lors
      des 8 derniers reçus mais SANS être renommé comme les 12 déjà connus, ou n'importe quel
      autre fichier hors format) — doit afficher une cause structurée (code + message), pas
      planter l'app ni afficher un message vide.

## Pour chaque fichier déposé

### A. Détection + formatage
- [ ] Le statut annoncé (note bleue "reconnu comme format «...» et transformé" / rien si déjà
      canonique / erreur rouge avec cause) correspond à ce qu'on attend du fichier.
- [ ] Télécharger le bouton **« ... - formaté (avant audit).xlsx »** et l'ouvrir.
- [ ] Choisir 3-5 lignes au hasard et comparer au fichier source original :
  - Désignation, Unité de référence, Prix de référence, Conditionnement : cohérents avec la
    ligne source ?
  - Lot et Fournisseur corrects (surtout pour les familles où ça vient du nom de fichier, pas
    d'une colonne — BPU DQE et MP ARDECHE 24) ?
  - Origine/Local/Bio/Label cohérents avec ce que dit le fichier source ?
- [ ] Rien d'inventé : une case vide dans le fichier formaté doit correspondre à une case qu'on
      ne pouvait vraiment pas déduire du fichier source (pas à un oubli).

### B. Audit — mode "complet (avec corrections)"
- [ ] Lancer l'audit complet, regarder le résumé par feuille (lignes avec incohérence, cellules
      marquées).
- [ ] Ouvrir 2-3 des catégories dépliables qui ont des éléments : le texte du commentaire
      explique-t-il clairement le problème et ce qu'il faudrait corriger ?
- [ ] Vérifier la ligne "Marché/Lot recopiés automatiquement" : les chiffres semblent-ils
      raisonnables par rapport à la taille du fichier ?
- [ ] Télécharger **« ... - audit.xlsx »**, l'ouvrir, et vérifier que le surlignage/les
      commentaires affichés à l'écran sont bien dans le fichier téléchargé (pas juste à l'écran).

### C. Audit — mode "commentaires seulement"
- [ ] Relancer sur le MÊME fichier avec le 2e bouton.
- [ ] Comparer avec le résultat du mode complet : les incohérences signalées sont globalement
      les mêmes, mais est-ce que ce mode a bien corrigé MOINS de choses silencieusement (ex:
      "Poids à l'unité de référence" en suggestion commentée plutôt que rempli directement) ?
- [ ] Télécharger ce 2e fichier audité et vérifier qu'il est bien différent du fichier du mode
      complet (pas le même fichier renvoyé deux fois par erreur).

## Notes libres (à remplir pendant le test)

_Ce qui fonctionne :_

_Ce qui semble faux ou surprenant (fichier, ligne, ce qui cloche) :_

_Questions à trancher avec l'équipe :_
