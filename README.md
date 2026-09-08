# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — deux feuilles couvrant l'année d'exploitation
**avril 2027 → mars 2028** :

| Feuille | Rôle |
|---|---|
| **Programme de vols** | saisie et visualisation du programme (12 tableaux mensuels, grille, récapitulatif). Aucune notion de sièges. |
| **Analyse capacite** | interprétation automatique du programme en **sièges offerts**. Seule saisie : les capacités. |

Types avion : **77W**, **787**, **A320**.

## Feuille 1 — Programme de vols

| Cellule | Rôle |
|---|---|
| `C4` | **Filtre Mois** — sélectionne le tableau de saisie lu par la grille et le récapitulatif |
| `C5` | **Filtre Route** — pilote uniquement la grille |
| `C45:H46` | **Sommaire cliquable** : un clic sur un mois saute à son tableau |
| lignes 48 → 514 | **12 tableaux de saisie**, un par mois |
| `B522:F561` | **Vols additionnels** : vols datés qui s'ajoutent au programme du mois |

### Saisie : une case cochée = une rotation (un aller-retour)

Dans chaque tableau mensuel, une route occupe **3 lignes nommées 77W, 787 et A320**. On **coche le
jour de l'aller** de chaque rotation.

| Route | Type | J1 | J2 | J3 | J4 | J5 | J6 | J7 | Retour J+n | Rot./sem. par type | Rot./sem. route |
|---|---|---|---|---|---|---|---|---|---|---|---|
| **CDGRUN** | 77W | X | X | X | X | X | X | X | 1 | 7 | 10 |
| | 787 | X | | | | X | | | 1 | 2 | |
| | A320 | X | | | | | | | 1 | 1 | |

- **Cocher** : taper `x` — la case devient **verte cochée `X`**. **Décocher** : `Suppr`.
  Sélection + `x` + `Ctrl+Entrée` coche une plage entière ; le copier/coller fonctionne aussi.
- **Aller-retour non simultané** → colonne **`Retour J+n`** : `0` = A/R dans la journée, `1` = retour
  le lendemain, `2` = le surlendemain, `3` = J+3. Le décalage se règle **par route et par type**,
  et il est propre à chaque mois.
- La **rotation reste comptée une seule fois**, le jour de l'aller, quel que soit le décalage : pas
  de double comptage dans le récapitulatif ni dans l'analyse de capacité.
- Cocher **2 ou 3 lignes le même jour** = 2 ou 3 rotations ce jour-là avec des types différents.
  Une case = 1 rotation ; pour deux rotations du **même type** le même jour, utiliser la table des
  vols additionnels (colonne `Nb de rotations`).
- Les 12 tableaux sont indépendants : la fréquence peut être différente chaque mois. Celui du mois
  filtré a son titre **en vert**.

### Lecture d'une cellule de la grille

La cellule a **deux lignes** :

| Affichage | Signification |
|---|---|
| `787` | une rotation qui **part** ce jour-là (A/R dans la journée si `Retour J+0`) |
| `77W/787/A320` | trois rotations qui partent le même jour |
| `< 77W` (2ᵉ ligne) | **retour** d'une rotation partie J-n |
| `787 (+A320)` | rotation du programme + un vol additionnel |
| `(+A320)` / `(+2x787)` | vol additionnel seul |

Exemple `CDGRUN` (retour J+1) un lundi : `77W/787/A320` puis `< 77W` — trois rotations partent, et
le 77W parti le dimanche rentre ce jour-là. Les cellules contenant un vol additionnel sont
entourées d'une **bordure rouge**.

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

Lecture automatique de la feuille 1 (les 12 tableaux + les vols additionnels), croisée avec les
capacités. **Aucune saisie sauf les hypothèses** (cellules bleues) :

| Hypothèse | Valeur |
|---|---|
| 77W | 438 sièges |
| 787 | 262 sièges |
| A320 | 174 sièges |
| Legs comptés par rotation | `1` |

> `Legs comptés par rotation` = 1 : l'aller-retour compte pour un vol, les sièges offerts sont ceux
> d'un sens. Mettre `2` pour compter l'aller **et** le retour.

Contenu :

1. **Synthèse annuelle par route** : rotations, sièges offerts, sièges/rotation, part du réseau.
2. **Sièges offerts par mois et par route** (12 mois × 12 routes + totaux).
3. **Rotations par mois et par route**.
4. **Rotations et sièges par type avion et par mois**, avec le siège moyen par rotation.
5. Blocs automatiques : calendrier (occurrences de chaque jour par mois) et détail par route, type
   et mois.

## Règles de calcul

- Le tableau lu = celui du mois filtré (`INDEX` sur la zone des 12 tableaux, décalage de 39 lignes
  par mois).
- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont comptées (semaines partielles incluses).
- Pour chaque jour d'aller coché, le nombre d'occurrences de ce jour dans le mois est compté à
  partir des dates réelles, **chaque type avion compté séparément**.
- Le jour de retour affiché = jour d'aller décalé de `Retour J+n`, avec report sur la semaine
  suivante quand le décalage franchit le dimanche.
- `Total rotations du mois` = **programme du mois + vols additionnels** datés du mois.
- Sièges = rotations par type × capacité du type × legs comptés.

## Exemples validés

| Mois filtré | Route | Rot./sem. | Total du mois |
|---|---|---|---|
| avril 2027 | NOSRUN (A320 3/7, retour J+0) | 3 | 13 + 2 vols add. = **15** |
| avril 2027 | CDGRUN (77W 7/7 + 787 2/7 + A320 1/7, retour J+1) | 10 | **45** (30 / 11 / 4) |
| aout 2027 | CDGRUN (77W 7/7 + 787 7/7) | 14 | **62** (31 / 31 / 0) |

Réseau : **193 rotations en avril 2027** (34 en 77W, 46 en 787, 113 en A320) soit **46 606 sièges**,
et sur l'année **2 334 rotations** pour **564 692 sièges**. Le passage au retour décalé ne change
aucun total : la rotation est comptée le jour de son aller.

## Données livrées

Les routes et les types avion sont ceux de la demande. Les 12 tableaux sont pré-remplis avec le
même programme de base, **valeurs d'exemple à ajuster** — sauf `NOSRUN`, qui reprend l'exemple
fourni (3/7 en A320 sur J1/J4/J6). Décalages retour livrés : `1` pour CDGRUN, CDGDZA et BKKRUN
(long-courriers), `0` ailleurs — à ajuster. Trois mois montrent une saisonnalité : `NOSRUN` en
juillet (4/7) et décembre (5/7), `CDGRUN` en août (77W 7/7 + 787 7/7). Les cellules de saisie sont
en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
