# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — deux feuilles couvrant l'année d'exploitation
**avril 2027 → mars 2028** :

| Feuille | Rôle |
|---|---|
| **Programme de vols** | saisie et visualisation du programme (12 tableaux mensuels, grille, récapitulatif). Aucune notion de sièges. |
| **Analyse capacite** | interprétation automatique du programme en **sièges offerts**. Seule saisie : les hypothèses. |

Types avion : **77W**, **787**, **A320**.

## Feuille 1 — Programme de vols

| Cellule | Rôle |
|---|---|
| `C4` | **Filtre Mois** — sélectionne le tableau de saisie lu par la grille et le récapitulatif |
| `C5` | **Filtre Route** — pilote uniquement la grille |
| `C45:H46` | **Sommaire cliquable** : un clic sur un mois saute à son tableau |
| lignes 48 → 514 | **12 tableaux de saisie**, un par mois |
| `B522:F561` | **Vols additionnels** : vols datés qui s'ajoutent au programme du mois |

### Saisie : un caractère par rotation, le caractère dit quand elle rentre

Chaque route occupe **3 lignes = les 3 types avion**. Dans la case d'un jour, on saisit **un
caractère par rotation** (aller-retour) qui **part** ce jour-là ; ce caractère indique **quand cette
rotation rentre** :

| Caractère | Retour |
|---|---|
| `0` (ou `x`) | le jour même |
| `1` | à J+1 |
| `2` | à J+2 |
| `3` | à J+3 |

| Code saisi | Signification |
|---|---|
| `0` | 1 rotation, aller-retour dans la journée |
| `1` | 1 rotation qui rentre le lendemain |
| `01` | **2 rotations le même jour : la première rentre le jour même, la seconde le lendemain** |
| `00` | 2 rotations rentrant toutes deux le jour même |
| `11` | 2 rotations rentrant toutes deux à J+1 |
| *(vide)* | aucun vol |

Une cellule contenant une rotation devient **verte** ; une cellule à **plusieurs rotations** devient
**orange** pour se repérer d'un coup d'œil. Maximum 4 rotations par jour, par route et par type ;
une validation refuse tout autre caractère.

| Route | Type | J1 | J2 | … | J7 | Rot./sem. par type | Rot./sem. route |
|---|---|---|---|---|---|---|---|
| **CDGRUN** | 77W | 1 | 1 | … | 1 | 7 | 11 |
| | 787 | 1 | | | | 2 | |
| | A320 | **01** | | | | 2 | |

Les 12 tableaux sont indépendants : la fréquence peut être différente chaque mois. Celui du mois
filtré a son titre **en vert**. Pour reprendre un mois sur un autre : copier la zone des 7 jours
(36 lignes) et la coller dans le tableau cible.

### Lecture d'une cellule de la grille

La cellule a **deux lignes** :

| Affichage | Signification |
|---|---|
| `77W/787/2xA320` | 4 rotations partent ce jour-là : une 77W, une 787 et **deux** A320 |
| `< 77W/A320` (2ᵉ ligne) | les **retours** du jour : une 77W et une A320 parties les jours précédents |
| `787 (+A320)` | programme + un vol additionnel |
| `(+A320)` / `(+2x787)` | vol additionnel seul |

Exemple `CDGRUN` en avril, lundi `01` sur la ligne A320 : la grille affiche
`77W/787/2xA320` puis `< 77W` le lundi, et `77W` puis `< 77W/787/A320` le mardi — l'un des deux
A320 est rentré dans la journée, l'autre rentre le mardi. Les cellules contenant un vol additionnel
sont entourées d'une **bordure rouge**.

### Repères calculés (colonnes N/O)

1er et dernier jour du mois, lundi de la semaine 1, numéro du tableau lu, type(s) avion,
rotations/semaine, puis :

| Cellule | Contenu |
|---|---|
| `O10` | **Total rotations du mois** pour la route filtrée |
| `O11` / `O12` / `O13` | **dont 77W / dont 787 / dont A320** |

### Où se trouve chaque tableau

| N° | Mois | Ligne de titre | Lignes de saisie |
|---|---|---|---|
| 1 | avril 2027 | 48 | 50 → 85 |
| 2 | mai 2027 | 87 | 89 → 124 |
| 3 | juin 2027 | 126 | 128 → 163 |
| 4 | juillet 2027 | 165 | 167 → 202 |
| 5 | aout 2027 | 204 | 206 → 241 |
| 6 | septembre 2027 | 243 | 245 → 280 |
| 7 | octobre 2027 | 282 | 284 → 319 |
| 8 | novembre 2027 | 321 | 323 → 358 |
| 9 | decembre 2027 | 360 | 362 → 397 |
| 10 | janvier 2028 | 399 | 401 → 436 |
| 11 | fevrier 2028 | 438 | 440 → 475 |
| 12 | mars 2028 | 477 | 479 → 514 |

### Récapitulatif mensuel (lignes 23-37)

Route, Type(s) avion, Semaine type du mois, Rotations/semaine, Nb de semaines/occurrences, Total
rotations du mois, dont programme du mois, dont vols additionnels, Total 77W / 787 / A320, plus le
total général. Il ne dépend que du filtre Mois.

## Feuille 2 — Analyse capacite

Lecture automatique de la feuille 1, croisée avec les hypothèses (**seules cellules saisissables**) :

| Hypothèse | Valeur |
|---|---|
| 77W | 438 sièges |
| 787 | 262 sièges |
| A320 | 174 sièges |
| Legs comptés par rotation | `1` |

> `Legs comptés par rotation` = 1 : l'aller-retour compte pour un vol, les sièges offerts sont ceux
> d'un sens. Mettre `2` pour compter l'aller **et** le retour.

Contenu : synthèse annuelle par route · sièges offerts par mois et par route · rotations par mois et
par route · rotations et sièges par type avion et par mois · blocs automatiques (calendrier des
occurrences, détail route × type × mois).

## Règles de calcul

- Le tableau lu = celui du mois filtré (`INDEX` sur la zone des 12 tableaux, décalage de 39 lignes
  par mois).
- **Chaque caractère saisi = une rotation**, comptée le jour de son départ quel que soit son retour :
  le nombre de rotations d'un jour est la longueur du code (`LEN`).
- Le jour de retour d'une rotation = son jour de départ + le chiffre du caractère, avec report sur
  la semaine suivante quand le décalage franchit le dimanche.
- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont comptées (semaines partielles incluses).
- `Total rotations du mois` = **programme du mois + vols additionnels** datés du mois.
- Sièges = rotations par type × capacité du type × legs comptés.

## Exemples validés

| Mois | Route | Programme | Rot./sem. | Total du mois |
|---|---|---|---|---|
| avril 2027 | CDGRUN | 77W 7/7 `1` · 787 `1` J1/J5 · A320 `01` J1 | 11 | **49** (30 / 11 / 8) |
| avril 2027 | RRGRUN | A320 6 jours dont `00` le jeudi | 7 | **31** |
| avril 2027 | NOSRUN | A320 `0` J1/J4/J6 | 3 | 13 + 2 vols add. = **15** |
| aout 2027 | CDGRUN | 77W 7/7 + 787 7/7 | 14 | **62** |

Réseau : **202 rotations en avril 2027** (34 en 77W, 46 en 787, 122 en A320) soit **48 172 sièges**,
et sur l'année **2 434 rotations** pour **582 092 sièges**. Les totaux de la feuille 2 sont recoupés
mois par mois avec ceux de la feuille 1.

## Données livrées

Les routes et les types avion sont ceux de la demande. Les 12 tableaux sont pré-remplis avec le même
programme de base, **valeurs d'exemple à ajuster** — sauf `NOSRUN`, qui reprend l'exemple fourni
(3/7 en A320 sur J1/J4/J6). Deux exemples illustrent les nouveaux codes : `CDGRUN` A320 le lundi
(`01`, deux rotations à retours différents) et `RRGRUN` le jeudi (`00`, double rotation du même
type). Trois mois montrent une saisonnalité : `NOSRUN` en juillet (4/7) et décembre (5/7), `CDGRUN`
en août. Les cellules de saisie sont en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
