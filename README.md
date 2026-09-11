# Programme de vols mensuel automatisé (Excel)

`Programme_vols_mensuel_2027-2028.xlsx` — **quatre feuilles**, année d'exploitation
**avril 2027 → mars 2028**, types avion **77W / 787 / A320**.

| Feuille | Rôle |
|---|---|
| **Programme de vols** | lecture : filtres Mois + Route, grille mensuelle, récapitulatif |
| **Saisie programme** | les 12 tableaux mensuels — le programme régulier |
| **Vols additionnels** | les écarts datés : **ajouts (+) et retraits (−)**. Livrée vide |
| **Analyse capacite** | sièges offerts. Seule saisie : les hypothèses |

## La règle de comptage

**Une rotation = un aller-retour.** Elle produit un **départ** le jour où elle part et une
**arrivée** le jour où elle rentre. Le total d'un mois vaut :

> **rotations du mois = (départs du mois + arrivées du mois) / 2**

- une rotation entièrement dans le mois compte **1** ;
- une rotation partie en fin de mois et rentrée le mois suivant compte **0,5 de chaque côté**.

Exemple vérifié — **CDGRUN en 787, avril 2027** : 9 départs (lundis et vendredis) mais seulement
8 arrivées, car le départ du vendredi 30/04 rentre le 1ᵉʳ mai → **8,5 rotations**, et non 9.

Les arrivées du début de mois issues du mois précédent sont reprises du **tableau de ce mois
précédent** : un départ du 30 avril apparaît en arrivée dans le décompte de mai.

Le récapitulatif affiche les trois colonnes — **Départs du mois**, **Arrivées du mois** et
**Total rotations du mois** — pour que le calcul soit vérifiable ligne par ligne.

## Feuille « Saisie programme »

Chaque route occupe **3 lignes = les 3 types avion**. Dans la case d'un jour, **un caractère par
rotation** qui **part** ce jour-là ; le caractère dit **quand elle rentre** :

| Caractère | Retour | | Code | Signification |
|---|---|---|---|---|
| `0` (ou `x`) | le jour même | | `0` | 1 rotation, A/R dans la journée |
| `1` | à J+1 | | `1` | 1 rotation qui rentre le lendemain |
| `2` | à J+2 | | `01` | 2 rotations : l'une rentre le jour même, l'autre le lendemain |
| `3` | à J+3 | | `00` | 2 rotations rentrant toutes deux le jour même |

Case vide = aucun vol. Une rotation → cellule **verte** ; plusieurs → cellule **orange**. Maximum
4 rotations par jour, par route et par type. La cellule `C6` rappelle le mois filtré, et le tableau
correspondant a son titre **en vert**.

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

## Feuille « Vols additionnels » — ajouts et régulations

Une ligne = un écart ponctuel au programme :

| Colonne | Rôle |
|---|---|
| Date du vol (départ) | du **25/03/2027** au 31/03/2028 |
| Route · Type avion | listes déroulantes |
| **Nb de rotations** | **positif = ajout**, **négatif = retrait** (régulation, annulation) |
| **Retour J+n** | comme dans le programme : `0` le jour même, `1` le lendemain… |
| Commentaire | libre |

- Un **retrait** (`-1`) se soustrait du programme de ce jour-là ; la ligne est surlignée en rose et
  la grille encadre la cellule en **rouge** avec `(-A320)`. Un **ajout** surligne la ligne en vert
  clair et encadre la cellule de la grille en **vert** avec `(+A320)`. Si une même cellule porte un
  ajout et un retrait, la bordure rouge l'emporte.
- La plage de dates commence avant le 1ᵉʳ avril **exprès** : un vol parti le **31/03/2027** avec
  `Retour J+1` rentre le 01/04 → il apporte **0,5 rotation à avril** (son arrivée), son départ
  tombant hors de l'année. La grille affiche `< (+787)` le 1ᵉʳ avril.

## Feuille « Programme de vols » (lecture)

`C4` = filtre Mois (désigne le tableau lu) · `C5` = filtre Route (grille seulement).

### Lecture d'une cellule de la grille

| Affichage | Signification |
|---|---|
| `77W/787/2xA320` (1ʳᵉ ligne) | les rotations qui **partent** ce jour-là |
| `< 77W/A320` (2ᵉ ligne) | les **retours** du jour |
| `(+787)` — **bordure verte** | vol **ajouté** sur « Vols additionnels » |
| `(-A320)` — **bordure rouge** | vol **retiré** (régulation) sur « Vols additionnels » |

### Repères (colonnes N/O)

1er et dernier jour du mois, lundi de la semaine 1, numéro du tableau lu, type(s) avion,
rotations/semaine, `O10` **Total rotations du mois** et `O11`/`O12`/`O13` **dont 77W / 787 / A320**,
pour la route filtrée.

### Récapitulatif (lignes 23-37, colonnes B→N)

Route · Type(s) avion · Semaine type du mois · Rotations/semaine · Nb de semaines/occurrences ·
**Total rotations du mois** · **Départs du mois** · **Arrivées du mois** · dont programme du mois ·
dont vols additionnels · Rotations 77W / 787 / A320 · plus le total général.

## Feuille « Analyse capacite »

Hypothèses (seules cellules saisissables) : 77W **438** sièges, 787 **262**, A320 **174**, et
`Legs comptés par rotation` = `1` (mettre `2` pour compter l'aller **et** le retour).

Contenu : synthèse annuelle par route · sièges offerts par mois et par route · rotations par mois et
par route · rotations et sièges par type avion et par mois · trois blocs automatiques (départs du
programme, queues de mois, rotations) et le calendrier des occurrences.

## Règles de calcul

- Le tableau lu = celui du mois filtré (`INDEX` sur la zone des 12 tableaux, décalage de 39 lignes
  par mois).
- Nombre de rotations d'un jour = longueur du code (`LEN`) ; jour de retour = jour de départ + le
  chiffre du caractère, avec report sur la semaine suivante quand le décalage franchit le dimanche.
- **Arrivées du mois = départs du mois − rotations parties dans les derniers jours du mois +
  rotations parties dans les derniers jours du mois précédent.** Ces deux « queues de mois » sont
  calculées dans les colonnes X et Y de la matrice (feuille de lecture) et dans le bloc AUTO 2/3
  (feuille d'analyse).
- Semaine 1 = semaine calendaire contenant le 1er du mois ; seules les dates réellement comprises
  dans le mois sont comptées (semaines partielles incluses).
- Sièges = rotations par type × capacité × legs comptés.

## Exemples validés (avril 2027)

| Route | Programme | Départs | Arrivées | Rotations |
|---|---|---|---|---|
| CDGRUN — 787 | `1` sur J1 et J5 | 9 | 8 | **8,5** |
| CDGRUN — 77W | `1` 7 jours sur 7 | 30 | 29 | **29,5** |
| CDGRUN — A320 | `01` le lundi | 8 | 8 | **8** |
| CDGRUN (total) | | 47 | 45 | **46** |
| NOSRUN | `0` sur J1/J4/J6 | 13 | 13 | **13** |

Réseau : **196,5 rotations en avril 2027** (33,5 en 77W, 43 en 787, 120 en A320) → **46 819
sièges** ; sur l'année **2 428,5 rotations** et **580 739 sièges**. Totaux recoupés entre les
feuilles.

Écarts testés : un vol `31/03/2027 · CDGRUN · 787 · +1 · Retour J+1` porte CDGRUN à **46,5**
rotations en avril (arrivées 45 → 46) ; un retrait `19/04/2027 · NOSRUN · A320 · −1` ramène NOSRUN
à **12**.

## Données livrées

Les 12 tableaux sont pré-remplis avec le même programme de base, **valeurs d'exemple à ajuster** —
sauf `NOSRUN`, qui reprend l'exemple fourni (3/7 en A320 sur J1/J4/J6). Exemples de codes
multi-rotations : `CDGRUN` A320 le lundi (`01`) et `RRGRUN` le jeudi (`00`). Saisonnalité :
`NOSRUN` en juillet (4/7) et décembre (5/7), `CDGRUN` en août. La table des vols additionnels est
vide. Les cellules de saisie sont en **bleu**.

## Régénérer le fichier

```bash
pip install openpyxl
python tools/build_programme_vols.py
```
