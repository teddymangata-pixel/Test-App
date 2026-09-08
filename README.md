# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — une **feuille unique** (`Programme de vols`) couvrant
l'année d'exploitation **avril 2027 → mars 2028**, pilotée par deux listes déroulantes.

## Utilisation

| Cellule | Rôle |
|---|---|
| `C4` | **Filtre Mois** — sélectionne le tableau de saisie lu par la grille et le récapitulatif |
| `C5` | **Filtre Route** (12 routes) — pilote uniquement la grille |
| `C44:H45` | **Sommaire cliquable** : un clic sur un mois saute à son tableau de saisie |
| lignes 47 → 513 | **12 tableaux de saisie**, un par mois |
| `B521:F560` | **Vols additionnels** : vols datés qui s'ajoutent au programme du mois |

Le filtre Mois **ne modifie aucune structure de saisie** : les 12 tableaux restent en place, il
désigne seulement celui qui alimente l'affichage.

## La saisie : 12 tableaux identiques, un par mois

Chaque tableau contient les 12 routes ; **chaque route occupe 3 lignes** — les **vols 1, 2 et 3 de
la journée — et chaque cellule J1→J7 porte le **type avion** opéré ce jour-là (liste déroulante
77W / 778 / A320, cellule vide = pas de vol).

| Route / vol du jour | J1 | J2 | J3 | J4 | J5 | J6 | J7 | Rot./sem. | Type(s) avion |
|---|---|---|---|---|---|---|---|---|---|
| **CDGRUN** | 77W | 77W | 77W | 77W | 77W | 77W | 77W | 10 | 77W(7), 778(2), A320(1) |
| + vol 2 du jour | 778 | | | | 778 | | | | |
| + vol 3 du jour | A320 | | | | | | | | |

- **La fréquence peut être totalement différente d'un mois à l'autre** : chaque mois a son propre
  tableau, indépendant des autres.
- **2 voire 3 vols le même jour** : renseigner plusieurs lignes sur la même colonne, avec le même
  type avion ou des types différents (ici lundi = 77W + 778 + A320, affiché `77W/778/A320`).
- Le tableau du **mois filtré** est signalé : son titre passe en **vert** avec la mention
  `<<< MOIS ACTUELLEMENT FILTRE`.
- Chaque tableau affiche en automatique, par route, les **rotations/semaine** et les **types avion**
  utilisés ; un lien `^ Retour en haut` ramène à la grille.
- **Reprendre un mois sur un autre** : sélectionner la zone des 7 jours d'un tableau (36 lignes),
  copier, coller dans le tableau cible — les 12 tableaux ont exactement la même forme.

### Où se trouve chaque tableau

| N° | Mois | Ligne de titre | Lignes de saisie |
|---|---|---|---|
| 1 | avril 2027 | 47 | 49 → 84 |
| 2 | mai 2027 | 86 | 88 → 123 |
| 3 | juin 2027 | 125 | 127 → 162 |
| 4 | juillet 2027 | 164 | 166 → 201 |
| 5 | aout 2027 | 203 | 205 → 240 |
| 6 | septembre 2027 | 242 | 244 → 279 |
| 7 | octobre 2027 | 281 | 283 → 318 |
| 8 | novembre 2027 | 320 | 322 → 357 |
| 9 | decembre 2027 | 359 | 361 → 396 |
| 10 | janvier 2028 | 398 | 400 → 435 |
| 11 | fevrier 2028 | 437 | 439 → 474 |
| 12 | mars 2028 | 476 | 478 → 513 |

## Vols additionnels

Une ligne = un vol supplémentaire daté (charter, renfort ponctuel, fret…), avec sa route, son type
avion, un nombre de rotations et un commentaire. Il s'ajoute au programme du mois, sur n'importe
quel jour. Pour une modification qui vaut sur tout un mois, modifier le tableau du mois.

Lecture d'une cellule de la grille :

| Affichage | Signification |
|---|---|
| `778` | un vol du programme du mois |
| `77W/778` | deux vols le même jour, deux types différents |
| `77W/778/A320` | trois vols le même jour |
| `778 (+A320)` | programme + un vol additionnel |
| `(+A320)` / `(+2x778)` | vol additionnel seul, sur un jour hors programme |

Les cellules contenant un vol additionnel sont entourées d'une **bordure rouge**.

## Structure de la feuille

- **Lignes 4-5** : les deux filtres. Colonnes N/O : repères calculés (1er/dernier jour du mois,
  lundi de la semaine 1, numéro du tableau lu, type(s) avion, rotations/semaine et total du mois
  pour la route filtrée).
- **Lignes 7-19** : grille **semaines (verticales) × jours J1→J7 (horizontaux)**, 6 blocs de
  semaines. Chaque semaine occupe deux lignes : les **vols du jour** (colorés selon la route) et,
  en sous-info, les **dates réelles**. Jours hors mois et jours sans vol restent vides.
- **Ligne 20** : occurrences de chaque jour de semaine dans le mois (semaines partielles incluses).
- **Lignes 23-37** : **récapitulatif mensuel toutes routes** — Route, Type(s) avion, Semaine type du
  mois, Rotations/semaine, Nb de semaines/occurrences, Total rotations du mois, dont programme du
  mois, dont vols additionnels, Total 77W / 778 / A320 — plus le total général.
- **Lignes 41-513** : les 12 tableaux de saisie. **Lignes 516-560** : vols additionnels.
- **Colonnes N→U** : blocs de calcul automatique (programme du mois recopié pour les 12 routes,
  programme de la route filtrée) ; **colonnes W→Z** : listes de référence. Hors zone d'impression.

## Règles de calcul

- Le tableau lu = celui du mois filtré (`INDEX` sur la zone des 12 tableaux, décalage de 39 lignes
  par mois).
- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont affichées, donc comptées (semaines partielles incluses).
- Pour chaque jour opéré, le nombre d'occurrences de ce jour dans le mois est compté à partir des
  dates réelles de la grille, **les vols 1, 2 et 3 comptés séparément**.
- `Total rotations du mois` = **programme du mois + vols additionnels** datés du mois.
- Code couleur : une couleur fixe par route, appliquée à la grille, au récapitulatif et aux
  12 tableaux de saisie.

## Exemples validés

| Mois filtré | Tableau | Route | Rot./sem. | Total du mois |
|---|---|---|---|---|
| avril 2027 | 1 | NOSRUN | 3 | 13 + 2 vols add. = **15** |
| juillet 2027 | 4 | NOSRUN | 4 | **18** |
| aout 2027 | 5 | CDGRUN (77W 7/7 + 778 7/7) | 14 | **62** (31 en 77W, 31 en 778) |
| decembre 2027 | 9 | NOSRUN | 5 | **21** |
| mars 2028 | 12 | TMMRUN | 2 | **9** |

Total toutes routes : 203 rotations en avril 2027, 209 en juillet, 219 en août, 214 en décembre,
206 en mars 2028.

## Données livrées

Les routes et les types avion sont ceux de la demande. Les 12 tableaux sont pré-remplis avec le
même programme de base, **valeurs d'exemple à ajuster** — sauf `NOSRUN`, qui reprend l'exemple
fourni (3/7 en A320 sur J1/J4/J6). Trois mois montrent une saisonnalité : `NOSRUN` en juillet (4/7)
et en décembre (5/7), `CDGRUN` en août (77W 7/7 + 778 7/7). Les 3 vols additionnels sont également
des exemples à supprimer. Les cellules de saisie sont en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
