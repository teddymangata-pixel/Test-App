# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — une **feuille unique** (`Programme de vols`) couvrant
l'année d'exploitation **avril 2027 → mars 2028**, pilotée par deux listes déroulantes.

## Utilisation

| Cellule | Rôle |
|---|---|
| `C4` | **Filtre Mois** (avril 2027 → mars 2028) — pilote la grille, le récapitulatif et le programme appliqué |
| `C5` | **Filtre Route** (12 routes) — pilote uniquement la grille |
| `B46:J125` | **Programmes hebdomadaires** : la saisie principale, mois par mois |
| `B133:F172` | **Vols additionnels** : vols datés qui s'ajoutent au programme du mois |

## La saisie : un programme hebdomadaire par route **et par mois**

La fréquence n'est plus figée pour l'année : elle se saisit **mois par mois**, dans une table où
**une ligne = une route + un mois + le type avion opéré chaque jour J1→J7** (cellule vide = pas de
vol ce jour-là).

| Route | Mois du programme | J1 | J2 | J3 | J4 | J5 | J6 | J7 | Rot./sem. | Statut |
|---|---|---|---|---|---|---|---|---|---|---|
| NOSRUN | TOUS | A320 | | | A320 | | A320 | | 3 | |
| NOSRUN | juillet 2027 | A320 | | A320 | A320 | | A320 | | 4 | APPLIQUE |
| NOSRUN | decembre 2027 | A320 | A320 | | A320 | | A320 | A320 | 5 | |

- **Changer la fréquence d'un mois** = ajouter une ligne à ce mois. Elle **remplace** le programme
  par défaut de la route pour ce mois-là.
- **`TOUS`** est un mois spécial : programme par défaut, appliqué **uniquement** aux mois qui n'ont
  aucune ligne propre pour cette route. Il évite de retaper 12 fois un programme identique ; si un
  mois doit être différent, on lui écrit sa ligne. Rien n'oblige à l'utiliser : on peut saisir les
  12 mois explicitement.
- **Plusieurs vols le même jour** = plusieurs lignes pour le même couple route/mois, avec le même
  type avion ou des types différents (ex. `CDGRUN` en août : une ligne 77W 7/7 + une ligne 778 7/7
  → 14 rotations par semaine, affichées `77W/778` dans la grille). Aucune limite de vols par jour.
- La colonne **Statut** passe en **vert `APPLIQUE`** sur les lignes réellement utilisées par le mois
  filtré, et grise les autres : on voit immédiatement ce qui pilote l'affichage.
- L'en-tête de la table porte les **flèches de filtre Excel** : filtrer sur un mois ou une route
  pour ne saisir que la tranche concernée. Copier/coller une ligne puis changer le mois est le
  chemin le plus rapide pour créer le programme d'un nouveau mois.

## Vols additionnels

Une ligne = un vol supplémentaire daté (charter, renfort ponctuel, fret…), avec sa route, son type
avion, un nombre de rotations et un commentaire. Il s'ajoute au programme du mois, sur n'importe
quel jour. Pour une modification qui vaut sur tout un mois, écrire plutôt une ligne de programme.

Lecture d'une cellule de la grille :

| Affichage | Signification |
|---|---|
| `778` | un vol du programme hebdomadaire |
| `77W/778` | deux vols le même jour, deux types différents |
| `77W/778/A320` | trois vols le même jour |
| `2xA320` | deux vols du même type le même jour |
| `778 (+A320)` | programme + un vol additionnel |
| `(+A320)` / `(+2x778)` | vol additionnel seul, sur un jour hors programme |

Les cellules contenant un vol additionnel sont entourées d'une **bordure rouge**.

## Structure de la feuille

- **Lignes 4-5** : les deux filtres. Colonnes N/O : repères calculés (1er/dernier jour du mois,
  lundi de la semaine 1, **programme appliqué**, type(s) avion, rotations/semaine, total du mois
  pour la route filtrée).
- **Lignes 7-19** : grille **semaines (verticales) × jours J1→J7 (horizontaux)**, 6 blocs de
  semaines. Chaque semaine occupe deux lignes : les **vols du jour** (colorés selon la route) et,
  en sous-info, les **dates réelles**. Jours hors mois et jours sans vol restent vides.
- **Ligne 20** : occurrences de chaque jour de semaine dans le mois (semaines partielles incluses).
- **Lignes 23-37** : **récapitulatif mensuel toutes routes** — Route, Type(s) avion, Semaine type du
  mois, Rotations/semaine, Nb de semaines/occurrences, Total rotations du mois, **Programme
  appliqué**, dont programme hebdo, dont vols additionnels, Total 77W / 778 / A320 — plus le total
  général.
- **Lignes 41-125** : table des programmes hebdomadaires (80 lignes de saisie).
- **Lignes 128-172** : table des vols additionnels (40 lignes de saisie).
- **Colonnes N→U** : blocs de calcul automatique (programme de la route filtrée, nombre de vols par
  route et par jour) ; **colonnes W→AA** : listes de référence. Hors zone d'impression.

## Règles de calcul

- Programme retenu pour une route = les lignes du **mois filtré** s'il en existe, sinon les lignes
  **`TOUS`** (colonne `Programme appliqué` du récapitulatif).
- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont affichées, donc comptées (semaines partielles incluses).
- Pour chaque jour opéré, le nombre d'occurrences de ce jour dans le mois est compté à partir des
  dates réelles de la grille, **chaque vol de la journée compté séparément**.
- `Total rotations du mois` = **programme hebdomadaire + vols additionnels** datés du mois.
- Code couleur : une couleur fixe par route, en mise en forme conditionnelle, appliquée à la
  grille, au récapitulatif et à la table des programmes.

## Exemples validés

| Mois filtré | Route | Programme appliqué | Rot./sem. | Total du mois |
|---|---|---|---|---|
| avril 2027 | NOSRUN | TOUS (3/7 A320 J1/J4/J6) | 3 | 13 + 2 vols add. = **15** |
| juillet 2027 | NOSRUN | juillet 2027 (4/7) | 4 | **18** |
| decembre 2027 | NOSRUN | decembre 2027 (5/7) | 5 | **21** |
| aout 2027 | CDGRUN | aout 2027 (77W 7/7 + 778 7/7) | 14 | **62** (31 en 77W, 31 en 778) |

Total toutes routes : 203 rotations en avril 2027, 209 en juillet, 219 en août, 214 en décembre.

## Données livrées

Les routes et les types avion sont ceux de la demande. Les 17 lignes de programme `TOUS`, les
4 lignes de programme mensuel et les 3 vols additionnels sont des **valeurs d'exemple à ajuster ou
supprimer** — sauf le programme `TOUS` de `NOSRUN`, qui reprend l'exemple fourni (3/7 en A320 sur
J1/J4/J6). Les cellules de saisie sont en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
