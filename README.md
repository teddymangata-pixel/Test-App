# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — une **feuille unique** (`Programme de vols`) couvrant
l'année d'exploitation **avril 2027 → mars 2028**, pilotée par deux listes déroulantes.

## Utilisation

| Cellule | Rôle |
|---|---|
| `C4` | **Filtre Mois** (avril 2027 → mars 2028) — pilote la grille *et* le récapitulatif |
| `C5` | **Filtre Route** (12 routes) — pilote uniquement la grille |
| `C43:J54` | **Zone de paramètres** : semaine type de chaque route (type avion + jours J1→J7) |

Tout le reste est calculé par formule : changer un filtre ou une semaine type recalcule
immédiatement la grille, les occurrences et les totaux, pour les 12 mois.

## Structure de la feuille

- **Lignes 4-5** : les deux filtres. Colonnes K/L : repères calculés (1er/dernier jour du mois,
  lundi de la semaine 1, type avion, rotations/semaine et total du mois pour la route filtrée).
- **Lignes 7-19** : grille **semaines (verticales) × jours J1→J7 (horizontaux)**, 6 blocs de
  semaines. Chaque semaine occupe deux lignes : le **type avion opéré** (coloré selon la route)
  et, en sous-info, les **dates réelles** du calendrier civil. Les jours hors mois et les jours
  sans opération restent vides.
- **Ligne 20** : nombre d'occurrences de chaque jour de semaine dans le mois (semaines
  partielles incluses) — c'est la base du calcul des totaux.
- **Lignes 23-37** : **récapitulatif mensuel toutes routes** (Route, Type avion, Semaine type,
  Rotations/semaine, Nb de semaines/occurrences, Total rotations du mois) + ligne de total général.
- **Lignes 40-54** : zone de paramètres (semaines types, table des mois, liste des types avion).

## Règles de calcul

- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont affichées, donc comptées (semaines partielles de début/fin de mois incluses).
- Pour chaque jour d'opération de la semaine type, le nombre d'occurrences du jour dans le mois est
  compté à partir des dates réelles de la grille.
- `Total rotations du mois` = somme des occurrences des jours d'opération de la route
  (`SUMPRODUCT` occurrences × jours de la semaine type).
- Code couleur : une couleur fixe par route, en mise en forme conditionnelle, appliquée à la
  grille, au récapitulatif et à la zone de paramètres.

## Exemple validé

`NOSRUN` — semaine type 3/7 en A320 sur J1, J4, J6 :

| Mois filtré | Semaines | Occurrences J1 / J4 / J6 | Total rotations |
|---|---|---|---|
| avril 2027 | 5 (dont 2 partielles) | 4 / 5 / 4 | **13** |
| mai 2027 | 6 (dont 2 partielles) | 5 / 4 / 5 | **14** |

## Données livrées

Les routes et les types avion sont ceux de la demande. Les semaines types sont des **valeurs
d'exemple à ajuster** dans la zone de paramètres — sauf `NOSRUN`, qui reprend l'exemple fourni
(3/7 en A320 sur J1/J4/J6). Les cellules de saisie sont en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
