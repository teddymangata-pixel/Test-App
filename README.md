# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — **quatre feuilles**, année d'exploitation
**avril 2027 → mars 2028**, types avion **77W / 787 / A320**.

| Feuille | Rôle |
|---|---|
| **Programme de vols** | lecture : filtres Mois + Route, grille mensuelle, récapitulatif. Aucune saisie de programme. |
| **Saisie programme** | les 12 tableaux mensuels — la saisie du programme régulier. |
| **Vols additionnels** | les vols supplémentaires datés. Livrée **vide**. |
| **Analyse capacite** | sièges offerts. Seule saisie : les hypothèses de capacité. |

Des liens en haut de chaque feuille permettent de circuler entre elles.

## Feuille « Saisie programme »

Chaque route occupe **3 lignes = les 3 types avion**. Dans la case d'un jour, on saisit **un
caractère par rotation** (aller-retour) qui **part** ce jour-là ; ce caractère dit **quand cette
rotation rentre** :

| Caractère | Retour | | Code | Signification |
|---|---|---|---|---|
| `0` (ou `x`) | le jour même | | `0` | 1 rotation, A/R dans la journée |
| `1` | à J+1 | | `1` | 1 rotation qui rentre le lendemain |
| `2` | à J+2 | | `01` | **2 rotations : l'une rentre le jour même, l'autre le lendemain** |
| `3` | à J+3 | | `00` | 2 rotations rentrant toutes deux le jour même |

Case vide = aucun vol. Une rotation → cellule **verte** ; plusieurs rotations → cellule **orange**.
Maximum 4 rotations par jour, par route et par type ; une validation refuse tout autre caractère.
La cellule `C6` rappelle le mois filtré sur la feuille de lecture, et le tableau correspondant a son
titre **en vert**.

| N° | Mois | Ligne de titre | Lignes de saisie |
|---|---|---|---|
| 1 | avril 2027 | 11 | 13 → 48 |
| 2 | mai 2027 | 50 | 52 → 87 |
| 3 | juin 2027 | 89 | 91 → 126 |
| 4 | juillet 2027 | 128 | 130 → 165 |
| 5 | aout 2027 | 167 | 169 → 204 |
| 6 | septembre 2027 | 206 | 208 → 243 |
| 7 | octobre 2027 | 245 | 247 → 282 |
| 8 | novembre 2027 | 284 | 286 → 321 |
| 9 | decembre 2027 | 323 | 325 → 360 |
| 10 | janvier 2028 | 362 | 364 → 399 |
| 11 | fevrier 2028 | 401 | 403 → 438 |
| 12 | mars 2028 | 440 | 442 → 477 |

Les 12 tableaux sont indépendants : la fréquence peut être différente chaque mois. Pour reprendre un
mois sur un autre, copier la zone des 7 jours (36 lignes) et la coller dans le tableau cible.

## Feuille « Vols additionnels »

Une ligne = un vol supplémentaire daté (charter, renfort ponctuel, fret, seconde rotation
exceptionnelle) : date, route, type avion, nombre de rotations, commentaire. Il s'ajoute au
programme du mois concerné, apparaît dans la grille entre parenthèses avec un `+` et une **bordure
rouge**, et alimente la colonne « dont vols additionnels » du récapitulatif.

> Cette table est livrée **vide** : tout ce qui y est saisi est compté. Pour une rotation récurrente
> sur tout un mois, ajouter plutôt un caractère dans la case du jour, sur « Saisie programme ».

## Feuille « Programme de vols » (lecture)

| Cellule | Rôle |
|---|---|
| `C4` | **Filtre Mois** — désigne le tableau de saisie lu |
| `C5` | **Filtre Route** — pilote uniquement la grille |

### Lecture d'une cellule de la grille

La cellule a **deux lignes** :

| Affichage | Signification |
|---|---|
| `77W/787/2xA320` | 4 rotations partent ce jour-là : une 77W, une 787 et **deux** A320 |
| `< 77W/A320` (2ᵉ ligne) | les **retours** du jour, de rotations parties les jours précédents |
| `787 (+A320)` — bordure rouge | programme + un vol saisi sur « Vols additionnels » |

### Repères calculés (colonnes N/O)

1er et dernier jour du mois, lundi de la semaine 1, numéro du tableau lu, type(s) avion,
rotations/semaine, puis `O10` **Total rotations du mois** et `O11`/`O12`/`O13` **dont 77W / 787 /
A320**, pour la route filtrée.

### Récapitulatif (lignes 23-37)

Route, Type(s) avion, Semaine type du mois, Rotations/semaine, Nb de semaines/occurrences, Total
rotations du mois, dont programme du mois, dont vols additionnels, Total 77W / 787 / A320, plus le
total général. Il ne dépend que du filtre Mois.

## Feuille « Analyse capacite »

Hypothèses (seules cellules saisissables) : 77W **438** sièges, 787 **262**, A320 **174**, et
`Legs comptés par rotation` = `1` (mettre `2` pour compter l'aller **et** le retour).

Contenu : synthèse annuelle par route · sièges offerts par mois et par route · rotations par mois et
par route · rotations et sièges par type avion et par mois · blocs automatiques (calendrier des
occurrences, détail route × type × mois).

## Règles de calcul

- Le tableau lu = celui du mois filtré (`INDEX` sur la zone des 12 tableaux, décalage de 39 lignes
  par mois).
- **Chaque caractère saisi = une rotation**, comptée le jour de son départ quel que soit son retour :
  le nombre de rotations d'un jour est la longueur du code (`LEN`).
- Jour de retour = jour de départ + le chiffre du caractère, avec report sur la semaine suivante
  quand le décalage franchit le dimanche.
- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont comptées (semaines partielles incluses).
- `Total rotations du mois` = **programme du mois + vols additionnels** datés du mois.
- Sièges = rotations par type × capacité du type × legs comptés.

## Exemples validés (avril 2027)

| Route | Programme | Rot./sem. | Total du mois |
|---|---|---|---|
| CDGRUN | 77W 7/7 `1` · 787 `1` J1/J5 · A320 `01` J1 | 11 | **47** (30 / 9 / 8) |
| RRGRUN | A320 6 jours dont `00` le jeudi | 7 | **31** |
| NOSRUN | A320 `0` J1/J4/J6 | 3 | **13** |

Réseau : **198 rotations en avril 2027** (34 en 77W, 44 en 787, 120 en A320) → **47 300 sièges** ;
sur l'année **2 430 rotations** et **581 220 sièges**. Totaux recoupés entre les feuilles.

## Données livrées

Les 12 tableaux sont pré-remplis avec le même programme de base, **valeurs d'exemple à ajuster** —
sauf `NOSRUN`, qui reprend l'exemple fourni (3/7 en A320 sur J1/J4/J6). Deux exemples illustrent les
codes multi-rotations : `CDGRUN` A320 le lundi (`01`) et `RRGRUN` le jeudi (`00`). Trois mois
montrent une saisonnalité : `NOSRUN` en juillet (4/7) et décembre (5/7), `CDGRUN` en août. La table
des vols additionnels est vide. Les cellules de saisie sont en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
