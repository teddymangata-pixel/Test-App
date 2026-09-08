# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — deux feuilles couvrant l'année d'exploitation
**avril 2027 → mars 2028** :

| Feuille | Rôle |
|---|---|
| **Programme de vols** | saisie et visualisation du programme (12 tableaux mensuels, grille, récapitulatif). Aucune notion de sièges. |
| **Analyse capacite** | interprétation automatique du programme en **sièges offerts**. Seule saisie : les capacités. |

## Feuille 1 — Programme de vols

| Cellule | Rôle |
|---|---|
| `C4` | **Filtre Mois** — sélectionne le tableau de saisie lu par la grille et le récapitulatif |
| `C5` | **Filtre Route** — pilote uniquement la grille |
| `C44:H45` | **Sommaire cliquable** : un clic sur un mois saute à son tableau |
| lignes 47 → 513 | **12 tableaux de saisie**, un par mois |
| `B521:F560` | **Vols additionnels** : vols datés qui s'ajoutent au programme du mois |

Le filtre Mois **ne modifie aucune structure de saisie** : il désigne seulement le tableau lu.

### Saisie : 3 lignes par route = les 3 types avion, et on coche les jours

Dans chaque tableau mensuel, une route occupe **3 lignes nommées 77W, 778 et A320**. Il suffit de
**cocher** les jours d'opération de chaque type.

| Route | Type avion | J1 | J2 | J3 | J4 | J5 | J6 | J7 | Rot./sem. par type | Rot./sem. route |
|---|---|---|---|---|---|---|---|---|---|---|
| **CDGRUN** | 77W | X | X | X | X | X | X | X | 7 | 10 |
| | 778 | X | | | | X | | | 2 | |
| | A320 | X | | | | | | | 1 | |

- **Cocher** : taper `x` dans la case — elle devient une **case verte cochée `X`**.
  **Décocher** : touche `Suppr`. Pour cocher plusieurs cases d'un coup : sélectionner la plage,
  taper `x` puis `Ctrl+Entrée`. Copier/coller fonctionne aussi.
- Cocher **2 ou 3 lignes le même jour** = 2 ou 3 rotations ce jour-là avec des types différents
  (ici lundi = 77W + 778 + A320, affiché `77W/778/A320` dans la grille).
- **Une case cochée = 1 rotation.** Pour deux rotations du **même type** le même jour, utiliser la
  table des vols additionnels (colonne `Nb de rotations`).
- Les 12 tableaux sont indépendants : la fréquence peut être différente chaque mois. Celui du mois
  filtré a son titre **en vert** (`<<< MOIS ACTUELLEMENT FILTRE`).
- Chaque route affiche automatiquement ses **rotations/semaine par type** et son **total route**.

### Repères calculés (colonnes N/O)

1er et dernier jour du mois, lundi de la semaine 1, numéro du tableau lu, type(s) avion,
rotations/semaine, puis :

| Cellule | Contenu |
|---|---|
| `O10` | **Total rotations du mois** pour la route filtrée |
| `O11` / `O12` / `O13` | **dont 77W / dont 778 / dont A320** |

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

### Lecture d'une cellule de la grille

| Affichage | Signification |
|---|---|
| `778` | un vol du programme du mois |
| `77W/778` | deux vols le même jour, deux types différents |
| `77W/778/A320` | trois vols le même jour |
| `778 (+A320)` | programme + un vol additionnel |
| `(+A320)` / `(+2x778)` | vol additionnel seul, sur un jour hors programme |

Les cellules contenant un vol additionnel sont entourées d'une **bordure rouge**.

### Récapitulatif mensuel (lignes 23-37)

Route, Type(s) avion, Semaine type du mois, Rotations/semaine, Nb de semaines/occurrences, Total
rotations du mois, dont programme du mois, dont vols additionnels, Total 77W / 778 / A320, plus le
total général. Il ne dépend que du filtre Mois.

## Feuille 2 — Analyse capacite

Lecture automatique de la feuille 1 (les 12 tableaux + les vols additionnels), croisée avec les
capacités. **Aucune saisie sauf les sièges par type** (cellules bleues) :

| Type avion | Sièges par rotation |
|---|---|
| 77W | 438 |
| 778 | 262 |
| A320 | 174 |

> La valeur 262 a été donnée pour le **787** ; elle est portée par le code type `778` de la
> feuille 1. Si `778` désigne autre chose chez vous, corriger la cellule `C7`.

Contenu :

1. **Synthèse annuelle par route** : rotations, sièges offerts, sièges/rotation, part du réseau.
2. **Sièges offerts par mois et par route** (12 mois × 12 routes + totaux).
3. **Rotations par mois et par route**.
4. **Rotations et sièges par type avion et par mois**, avec le siège moyen par rotation.
5. Blocs automatiques : calendrier (occurrences de chaque jour par mois) et détail par route, type
   et mois.

Toute modification du programme ou d'une capacité met à jour l'ensemble.

## Règles de calcul

- Le tableau lu = celui du mois filtré (`INDEX` sur la zone des 12 tableaux, décalage de 39 lignes
  par mois).
- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont comptées (semaines partielles incluses).
- Pour chaque jour coché, le nombre d'occurrences de ce jour dans le mois est compté à partir des
  dates réelles, **chaque type avion compté séparément**.
- `Total rotations du mois` = **programme du mois + vols additionnels** datés du mois.
- Sièges = rotations par type × capacité du type.

## Exemples validés

| Mois filtré | Tableau | Route | Rot./sem. | Total du mois |
|---|---|---|---|---|
| avril 2027 | 1 | NOSRUN | 3 | 13 + 2 vols add. = **15** (dont A320 : 15) |
| avril 2027 | 1 | CDGRUN (77W 7/7 + 778 2/7 + A320 1/7) | 10 | **45** (30 / 11 / 4) |
| aout 2027 | 5 | CDGRUN (77W 7/7 + 778 7/7) | 14 | **62** (31 / 31 / 0) |

Réseau : **193 rotations en avril 2027** (34 en 77W, 46 en 778, 113 en A320) soit **46 606 sièges**,
et sur l'année **2 334 rotations** pour **564 692 sièges** (241 sièges par rotation en moyenne).
Les totaux de la feuille 2 sont recoupés avec ceux de la feuille 1 mois par mois.

## Données livrées

Les routes et les types avion sont ceux de la demande. Les 12 tableaux sont pré-remplis avec le
même programme de base, **valeurs d'exemple à ajuster** — sauf `NOSRUN`, qui reprend l'exemple
fourni (3/7 en A320 sur J1/J4/J6). Trois mois montrent une saisonnalité : `NOSRUN` en juillet (4/7)
et en décembre (5/7), `CDGRUN` en août (77W 7/7 + 778 7/7). Les 3 vols additionnels sont aussi des
exemples à supprimer. Les cellules de saisie sont en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
