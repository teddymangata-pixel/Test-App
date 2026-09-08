# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — une **feuille unique** (`Programme de vols`) couvrant
l'année d'exploitation **avril 2027 → mars 2028**, pilotée par deux listes déroulantes.

## Utilisation

| Cellule | Rôle |
|---|---|
| `C4` | **Filtre Mois** (avril 2027 → mars 2028) — pilote la grille *et* le récapitulatif |
| `C5` | **Filtre Route** (12 routes) — pilote uniquement la grille |
| `C45:I56` | **Semaines types** : pour chaque route, le **type avion** opéré chaque jour J1→J7 (vide = pas d'opération) |
| `B62:F101` | **Vols additionnels** : vols datés qui s'ajoutent à la semaine type |

Tout le reste est calculé par formule : changer un filtre, une semaine type ou un vol additionnel
recalcule immédiatement la grille, les occurrences et les totaux, pour les 12 mois.

## Plusieurs types avion par route

La semaine type se saisit **un type avion par jour** (liste déroulante 77W / 778 / A320, cellule
vide = pas d'opération). Une même route peut donc mélanger plusieurs appareils dans la semaine —
par exemple `CDGRUN` : 77W les J1/J3/J5/J7 et 778 les J2/J4/J6, affiché `77W(4), 778(3)` dans le
récapitulatif. Les colonnes **Total 77W / Total 778 / Total A320** ventilent les rotations du mois
par type (semaine type + vols additionnels).

## Vols additionnels

Une ligne = un vol supplémentaire daté (charter, renfort saisonnier, fret…), avec sa route, son
type avion, un nombre de rotations et un commentaire libre. Il peut tomber sur un jour déjà opéré
ou sur un jour hors semaine type, et il est compté dans le mois auquel appartient sa date.

Lecture d'une cellule de la grille :

| Affichage | Signification |
|---|---|
| `778` | vol de la semaine type |
| `778 (+A320)` | vol de la semaine type + un vol additionnel en A320 |
| `(+A320)` | vol additionnel seul, sur un jour hors semaine type |
| `778 (+2x778)` | semaine type + un vol additionnel de 2 rotations |
| `77W (+2 vols)` | semaine type + plusieurs vols additionnels le même jour |

Les cellules contenant un vol additionnel sont entourées d'une **bordure rouge**.

## Structure de la feuille

- **Lignes 4-5** : les deux filtres. Colonnes N/O : repères calculés (1er/dernier jour du mois,
  lundi de la semaine 1, type(s) avion, rotations/semaine et total du mois pour la route filtrée).
- **Lignes 7-19** : grille **semaines (verticales) × jours J1→J7 (horizontaux)**, 6 blocs de
  semaines. Chaque semaine occupe deux lignes : le **type avion opéré** (coloré selon la route) et,
  en sous-info, les **dates réelles** du calendrier civil. Jours hors mois et jours sans opération
  restent vides.
- **Ligne 20** : nombre d'occurrences de chaque jour de semaine dans le mois (semaines partielles
  incluses) — c'est la base du calcul des totaux.
- **Lignes 23-37** : **récapitulatif mensuel toutes routes** — Route, Type(s) avion, Semaine type,
  Rotations/semaine, Nb de semaines/occurrences, Total rotations du mois, dont semaine type, dont
  vols additionnels, Total 77W, Total 778, Total A320 — plus la ligne de total général.
- **Lignes 41-56** : semaines types par route + table des mois + liste des types avion.
- **Lignes 59-101** : table des vols additionnels (40 lignes de saisie, extensible en recopiant les
  formules de la colonne J).

## Règles de calcul

- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont affichées, donc comptées (semaines partielles de début/fin de mois incluses).
- Pour chaque jour d'opération de la semaine type, le nombre d'occurrences du jour dans le mois est
  compté à partir des dates réelles de la grille.
- `Total rotations du mois` = **semaine type** (`SUMPRODUCT` occurrences × jours opérés) **+ vols
  additionnels** du mois (`SUMIFS` bornée par le 1er et le dernier jour du mois).
- Code couleur : une couleur fixe par route, en mise en forme conditionnelle, appliquée à la
  grille, au récapitulatif et aux semaines types.

## Exemple validé

`NOSRUN` — semaine type 3/7 en A320 sur J1, J4, J6 :

| Mois filtré | Semaines | Occurrences J1 / J4 / J6 | Semaine type | Vols add. | Total |
|---|---|---|---|---|---|
| avril 2027 | 5 (dont 2 partielles) | 4 / 5 / 4 | **13** | 2 (lignes d'exemple) | **15** |
| mai 2027 | 6 (dont 2 partielles) | 5 / 4 / 5 | **14** | 0 | **14** |

## Données livrées

Les routes et les types avion sont ceux de la demande. Les semaines types et les 3 vols
additionnels sont des **valeurs d'exemple à ajuster ou supprimer** — sauf `NOSRUN`, dont la semaine
type reprend l'exemple fourni (3/7 en A320 sur J1/J4/J6). Les cellules de saisie sont en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
