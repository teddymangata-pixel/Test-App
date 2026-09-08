# -*- coding: utf-8 -*-
"""Genere le classeur 'Programme de vols mensuel' (feuille unique, 12 tableaux de saisie).

Saisie : 12 tableaux identiques, un par mois (avril 2027 -> mars 2028). Chaque tableau contient
les 12 routes, chacune sur NB_SLOTS lignes (vol 1 / vol 2 / vol 3 de la journee) x 7 jours J1-J7,
la cellule portant le type avion opere. Le filtre Mois ne change aucune structure : il selectionne
le tableau lu par la grille et le recapitulatif.
Vols additionnels : table de vols dates qui s'ajoutent au programme du mois.
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

OUT = "/home/user/Test-App/Programme_vols_mensuel_2027-2028.xlsx"
F = "Arial"
NB_SLOTS = 3            # lignes de type avion par route (vols possibles le meme jour)
_ = None

ROUTES = [  # (route, couleur, police blanche ?)
    ("CDGRUN", "9DC3E6", False), ("CDGDZA", "2E75B6", True),  ("BKKRUN", "FFE699", False),
    ("DZARUN", "F4B183", False), ("MRURUN", "C6E0B4", False), ("NOSRUN", "FFD966", False),
    ("RUNTNR", "D9D2E9", False), ("JNBRUN", "A9D08E", False), ("CPTRUN", "8FAADC", False),
    ("DIERUN", "FFC7CE", False), ("RRGRUN", "B7DEE8", False), ("TMMRUN", "D5A6BD", False),
]
TYPES = ["77W", "778", "A320"]
MOIS = [("avril 2027",2027,4),("mai 2027",2027,5),("juin 2027",2027,6),("juillet 2027",2027,7),
        ("aout 2027",2027,8),("septembre 2027",2027,9),("octobre 2027",2027,10),
        ("novembre 2027",2027,11),("decembre 2027",2027,12),("janvier 2028",2028,1),
        ("fevrier 2028",2028,2),("mars 2028",2028,3)]
JOURS = ["J1 - Lundi","J2 - Mardi","J3 - Mercredi","J4 - Jeudi","J5 - Vendredi","J6 - Samedi","J7 - Dimanche"]

# Programme de base recopie dans les 12 tableaux : {route: [ligne vol1, vol2, vol3]}
BASE = {
    "CDGRUN": [["77W","77W","77W","77W","77W","77W","77W"], ["778",_,_,_,"778",_,_], ["A320",_,_,_,_,_,_]],
    "CDGDZA": [[_,"778",_,_,"778",_,"77W"], [_]*7, [_]*7],
    "BKKRUN": [[_,_,"778",_,_,"778",_], [_]*7, [_]*7],
    "DZARUN": [["A320","A320","A320","A320","A320",_,_], [_,_,"778",_,_,_,_], [_]*7],
    "MRURUN": [["A320",_,"A320",_,"A320",_,"A320"], [_]*7, [_]*7],
    "NOSRUN": [["A320",_,_,"A320",_,"A320",_], [_]*7, [_]*7],
    "RUNTNR": [[_,"A320",_,"A320",_,"A320",_], [_]*7, [_]*7],
    "JNBRUN": [["778",_,_,_,"778",_,_], [_]*7, [_]*7],
    "CPTRUN": [[_,_,_,"778",_,_,_], [_]*7, [_]*7],
    "DIERUN": [[_,"A320",_,_,_,"A320",_], [_]*7, [_]*7],
    "RRGRUN": [["A320","A320","A320","A320","A320","A320",_], [_,_,_,_,"A320",_,_], [_,_,_,_,"A320",_,_]],
    "TMMRUN": [[_,_,"A320",_,_,_,"A320"], [_]*7, [_]*7],
}
# Exemples de saisonnalite : programme different pour certains mois. {(mois, route): [3 lignes]}
SAISON = {
    ("juillet 2027", "NOSRUN"):  [["A320",_,"A320","A320",_,"A320",_], [_]*7, [_]*7],
    ("decembre 2027", "NOSRUN"): [["A320","A320",_,"A320",_,"A320","A320"], [_]*7, [_]*7],
    ("aout 2027", "CDGRUN"):     [["77W"]*7, ["778"]*7, [_]*7],
}

VOLS_ADD = [
    (datetime.date(2027,4,7),  "NOSRUN", "A320", 1, "EXEMPLE a supprimer - vol supplementaire un mercredi (jour hors programme)"),
    (datetime.date(2027,4,10), "NOSRUN", "A320", 1, "EXEMPLE a supprimer - rotation doublee un samedi (jour deja opere)"),
    (datetime.date(2027,4,20), "CDGRUN", "778",  2, "EXEMPLE a supprimer - 2 rotations fret supplementaires"),
]

NB_WEEKS, GRID_C0 = 6, 3
ROW_HDR, ROW_W1 = 7, 8
ROW_OCC = ROW_W1 + 2*NB_WEEKS                     # 20
ROW_NBW = ROW_OCC + 1                             # 21
ROW_REC_TITLE, ROW_REC_HDR, ROW_REC0 = 23, 24, 25
ROW_REC_TOT = ROW_REC0 + len(ROUTES)              # 37
ROW_SEC_TITLE = 41                                # titre section saisie
ROW_NAV = 44                                      # sommaire : 2 lignes (44-45)
BLK0 = 47                                         # 1re ligne de titre de tableau
NB_ROWS_BLK = NB_SLOTS*len(ROUTES)                # 36 lignes de donnees
STEP = NB_ROWS_BLK + 3                            # titre + entete + donnees + ligne vide = 39
DATA0 = BLK0 + 2                                  # 1re ligne de donnees du 1er tableau
DATA_LAST = BLK0 + (len(MOIS)-1)*STEP + 1 + NB_ROWS_BLK
ROW_ADD_TITLE = DATA_LAST + 3
ROW_ADD_HDR, ROW_ADD0 = ROW_ADD_TITLE + 4, ROW_ADD_TITLE + 5
NB_ADD_ROWS = 40
ROW_ADD_END = ROW_ADD0 + NB_ADD_ROWS - 1
# zone auto a droite (N..U)
R_REP0 = 4
R_HELP_TITLE, R_HELP_HDR, R_HELP_TXT, R_HELP_S0 = 12, 13, 14, 15
R_MAT_TITLE, R_MAT_HDR, R_MAT0 = 19, 20, 21
HELP_C0, REF0 = 15, 4

wb = Workbook()
ws = wb.active
ws.title = "Programme de vols"
ws.sheet_view.showGridLines = False

def cell(ref, value=None, bold=False, size=10, color="000000", fill=None, halign="center",
         valign="center", italic=False, wrap=False, fmt=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = Font(name=F, size=size, bold=bold, color=color, italic=italic)
    c.alignment = Alignment(horizontal=halign, vertical=valign, wrap_text=wrap)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        c.number_format = fmt
    if border:
        c.border = border
    return c

thin = Side(style="thin", color="A6A6A6")
med = Side(style="medium", color="404040")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
NAVY, LIGHT, GREY, BROWN, RED, ACTIVE = "1F3864", "D9E2F3", "F2F2F2", "833C0C", "C00000", "375623"
BLUE_IN = "0000FF"

BIG = f"$C${DATA0}:$I${DATA_LAST}"                       # les 12 tableaux, zone continue
MAT = f"$O${R_MAT0}:$U${R_MAT0+NB_ROWS_BLK-1}"           # programme du mois filtre (auto)
ADD_D = f"$B${ROW_ADD0}:$B${ROW_ADD_END}"
ADD_R = f"$C${ROW_ADD0}:$C${ROW_ADD_END}"
ADD_T = f"$D${ROW_ADD0}:$D${ROW_ADD_END}"
ADD_N = f"$E${ROW_ADD0}:$E${ROW_ADD_END}"
ADD_K = f"$H${ROW_ADD0}:$H${ROW_ADD_END}"
OCC = f"$C${ROW_OCC}:$I${ROW_OCC}"
hc = lambda j: get_column_letter(HELP_C0+j)
blk_title = lambda k: BLK0 + k*STEP
mat_row = lambda i, s: R_MAT0 + NB_SLOTS*i + s

# ---------------------------------------------------------------- titre
ws.merge_cells("B1:I1")
cell("B1", "PROGRAMME DE VOLS MENSUEL - ANNEE D'EXPLOITATION AVRIL 2027 / MARS 2028",
     bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
ws.row_dimensions[1].height = 30
ws.merge_cells("B2:I2")
cell("B2", "Feuille unique : les deux filtres ci-dessous (Mois et Route) pilotent la grille, le recapitulatif ne suit que le filtre "
           "Mois. La saisie se fait dans 12 tableaux identiques, un par mois : chaque route y occupe 3 lignes (vols 1, 2 et 3 de la "
           "journee) et chaque cellule porte le type avion opere ce jour-la. Le filtre Mois ne modifie aucune structure de saisie : "
           "il selectionne le tableau lu. Tout est calcule par formule.",
     size=9, italic=True, halign="left", wrap=True, color="404040")
ws.row_dimensions[2].height = 26

# ---------------------------------------------------------------- filtres
cell("B4", "FILTRE  MOIS", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
ws.merge_cells("C4:D4")
cell("C4", "avril 2027", bold=True, size=12, color=BLUE_IN, fill="FFF2CC",
     border=Border(left=med, right=med, top=med, bottom=med))
cell("B5", "FILTRE  ROUTE", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
ws.merge_cells("C5:D5")
cell("C5", "NOSRUN", bold=True, size=12, color=BLUE_IN, fill="FFF2CC",
     border=Border(left=med, right=med, top=med, bottom=med))
cell("E4", "<-- liste deroulante : selectionne le tableau de saisie lu (la structure ne change pas)", size=9, italic=True, halign="left", color="808080")
cell("E5", "<-- liste deroulante : pilote la grille uniquement", size=9, italic=True, halign="left", color="808080")
ws["C4"].comment = Comment("Liste deroulante : avril 2027 -> mars 2028.", "Modele")
ws["C5"].comment = Comment("Liste deroulante : les 12 routes exploitees.", "Modele")

# ---------------------------------------------------------------- reperes calcules (N/O)
ws.merge_cells("N3:O3")
cell("N3", "REPERES CALCULES", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left")
W, X = f"$W${REF0}:$W${REF0+11}", f"$X${REF0}:$X${REF0+11}"
reperes = [
    ("1er jour du mois",     f"=INDEX({X},MATCH($C$4,{W},0))", "DD/MM/YYYY"),
    ("Dernier jour du mois", "=EOMONTH($O$4,0)", "DD/MM/YYYY"),
    ("Lundi semaine 1",      "=$O$4-WEEKDAY($O$4,3)", "DD/MM/YYYY"),
    ("Tableau de saisie lu", f"=MATCH($C$4,{W},0)", "0"),
    ("Type(s) avion",        f'=IFERROR(INDEX($C${ROW_REC0}:$C${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "General"),
    ("Rotations / semaine",  f'=IFERROR(INDEX($E${ROW_REC0}:$E${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "0"),
    ("Total rotations mois", f'=IFERROR(INDEX($G${ROW_REC0}:$G${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "0"),
]
for i, (lab, fml, fmt) in enumerate(reperes):
    r = R_REP0 + i
    cell(f"N{r}", lab, size=9, halign="left", fill=GREY, border=box)
    cell(f"O{r}", fml, size=9, bold=True, fmt=fmt, border=box)
MIDX = "$O$7"   # numero du tableau lu

# ---------------------------------------------------------------- auto : programme du mois filtre (36 lignes)
ws.merge_cells(f"N{R_MAT_TITLE}:{hc(6)}{R_MAT_TITLE}")
cell(f"N{R_MAT_TITLE}", "PROGRAMME DU MOIS FILTRE, TOUTES ROUTES (auto - recopie du tableau du mois)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"N{R_MAT_HDR}", "Route / vol", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_MAT_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    for s in range(NB_SLOTS):
        r = mat_row(i, s)
        cell(f"N{r}", f"{route} - vol {s+1}", size=8, halign="left", border=box)
        for j in range(7):
            rel = f"({MIDX}-1)*{STEP}+{NB_SLOTS*i+s+1}"
            cell(f"{hc(j)}{r}", f'=IFERROR(INDEX({BIG},{rel},{j+1})&"","")', size=8, border=box)

# ---------------------------------------------------------------- auto : programme de la route filtree
ws.merge_cells(f"N{R_HELP_TITLE}:{hc(6)}{R_HELP_TITLE}")
cell(f"N{R_HELP_TITLE}", "PROGRAMME DE LA ROUTE FILTREE (auto - alimente la grille)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"N{R_HELP_HDR}", "Jour", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_HELP_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
cell(f"N{R_HELP_TXT}", "Vols du jour (affichage grille)", bold=True, size=8, halign="left", fill=GREY, border=box)
for s in range(NB_SLOTS):
    cell(f"N{R_HELP_S0+s}", f"Vol {s+1}", size=8, halign="left", fill=GREY, border=box)
ROUTE_IDX = f'MATCH($C$5,$Z${REF0}:$Z${REF0+11},0)'
for j in range(7):
    c = hc(j)
    for s in range(NB_SLOTS):
        cell(f"{c}{R_HELP_S0+s}",
             f'=IFERROR(INDEX({MAT},({ROUTE_IDX}-1)*{NB_SLOTS}+{s+1},{j+1})&"","")', size=9, border=box)
    sref = [f"{c}${R_HELP_S0+s}" for s in range(NB_SLOTS)]
    parts = [sref[0]]
    for si in range(1, NB_SLOTS):
        before = (f'{sref[0]}=""' if si == 1 else "AND(" + ",".join(f'{x}=""' for x in sref[:si]) + ")")
        parts.append(f'IF({sref[si]}="","",IF({before},"","/")&{sref[si]})')
    cell(f"{c}{R_HELP_TXT}", "=" + "&".join(parts), bold=True, size=9, border=box)

# ---------------------------------------------------------------- grille semaine x jour
ws.merge_cells(f"B{ROW_HDR-1}:I{ROW_HDR-1}")
cell(f"B{ROW_HDR-1}", "GRILLE MENSUELLE  -  type avion opere par semaine et par jour (filtres Mois + Route)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
cell(f"B{ROW_HDR}", "Semaine", bold=True, size=10, color="FFFFFF", fill="2F5597", border=box)
for j in range(7):
    cell(f"{get_column_letter(GRID_C0+j)}{ROW_HDR}", JOURS[j], bold=True, size=10,
         color="FFFFFF", fill="2F5597", border=box)

type_rows, date_rows = [], []
for k in range(1, NB_WEEKS+1):
    rt, rd = ROW_W1 + 2*(k-1), ROW_W1 + 2*(k-1) + 1
    type_rows.append(rt); date_rows.append(rd)
    first, last = f"{get_column_letter(GRID_C0)}{rd}", f"{get_column_letter(GRID_C0+6)}{rd}"
    cell(f"B{rt}", f"Semaine {k}", bold=True, size=10, fill=LIGHT, halign="left", border=box)
    cell(f"B{rd}", f'=IF(COUNT({first}:{last})=0,"","du "&TEXT(MIN({first}:{last}),"DD/MM")&'
                   f'" au "&TEXT(MAX({first}:{last}),"DD/MM"))',
         size=8, italic=True, color="595959", fill=LIGHT, halign="left", border=box)
    ws.row_dimensions[rt].height = 22
    ws.row_dimensions[rd].height = 14
    for j in range(7):
        col = get_column_letter(GRID_C0+j)
        off = 7*(k-1) + j
        d = f"{col}{rd}"
        cell(d, f'=IF(OR($O$6+{off}<$O$4,$O$6+{off}>$O$5),"",$O$6+{off})',
             size=8, italic=True, color="595959", fmt="DD/MM", border=box)
        nrot = f"SUMIFS({ADD_N},{ADD_R},$C$5,{ADD_D},{d})"
        nrows = f"COUNTIFS({ADD_R},$C$5,{ADD_D},{d})"
        type1 = f'IFERROR(INDEX({ADD_T},MATCH({d}&"|"&$C$5,{ADD_K},0)),"")'
        lbl = f'IF({nrows}=1,IF({nrot}=1,{type1},{nrot}&"x"&{type1}),{nrot}&" vols")'
        cell(f"{col}{rt}",
             f'=IF({d}="","",TRIM(${hc(j)}${R_HELP_TXT}&IF({nrot}=0,""," (+"&{lbl}&")")))',
             bold=True, size=10, border=box)

cell(f"B{ROW_OCC}", "Occurrences du jour dans le mois", bold=True, size=9, halign="left", fill="E7E6E6", border=box)
for j in range(7):
    col = get_column_letter(GRID_C0+j)
    cell(f"{col}{ROW_OCC}", "=COUNT(" + ",".join(f"{col}{r}" for r in date_rows) + ")",
         bold=True, size=10, fill="E7E6E6", border=box)
cell(f"B{ROW_NBW}", "Nb de semaines calendaires du mois", bold=True, size=9, halign="left", fill="E7E6E6", border=box)
cell(f"C{ROW_NBW}", "=" + "+".join(
        f'IF(COUNT({get_column_letter(GRID_C0)}{r}:{get_column_letter(GRID_C0+6)}{r})>0,1,0)' for r in date_rows),
     bold=True, size=10, fill="E7E6E6", border=box)
ws.merge_cells(f"D{ROW_NBW}:I{ROW_NBW}")
cell(f"D{ROW_NBW}", 'Lecture : "778" = 1 vol du programme  |  "77W/778/A320" = 2 ou 3 vols le meme jour  |  "778 (+A320)" = '
                    'programme + vol additionnel (bordure rouge)  |  "(+A320)" = vol additionnel seul.',
     size=8, italic=True, color="595959", halign="left")

# ---------------------------------------------------------------- recapitulatif mensuel
ws.merge_cells(f"B{ROW_REC_TITLE}:L{ROW_REC_TITLE}")
cell(f"B{ROW_REC_TITLE}", "RECAPITULATIF MENSUEL - TOUTES ROUTES  (depend uniquement du filtre Mois)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
rec_hdr = ["Route", "Type(s) avion", "Semaine type du mois (jours J1-J7)", "Rotations/semaine",
           "Nb de semaines/occurrences dans le mois", "Total rotations du mois",
           "dont programme du mois", "dont vols additionnels", "Total 77W", "Total 778", "Total A320"]
for i, h in enumerate(rec_hdr):
    cell(f"{get_column_letter(2+i)}{ROW_REC_HDR}", h, bold=True, size=9, color="FFFFFF",
         fill="2F5597", wrap=True, border=box)
ws.row_dimensions[ROW_REC_HDR].height = 44

for i in range(len(ROUTES)):
    r = ROW_REC0 + i
    rows = [mat_row(i, s) for s in range(NB_SLOTS)]
    blk = f"$O${rows[0]}:$U${rows[-1]}"
    add = lambda extra="": f'SUMIFS({ADD_N},{ADD_R},$B{r}{extra},{ADD_D},">="&$O$4,{ADD_D},"<="&$O$5)'
    prods = {t: "+".join(f'SUMPRODUCT({OCC},--($O${x}:$U${x}="{t}"))' for x in rows) for t in TYPES}
    cell(f"B{r}", f"=$Z${REF0+i}", bold=True, size=10, border=box)
    tl = "&".join(f'IF(COUNTIF({blk},"{t}")>0,"{t}("&COUNTIF({blk},"{t}")&") ","")' for t in TYPES)
    cell(f"C{r}", f'=IF(E{r}=0,"-",SUBSTITUTE(TRIM({tl})," ",", "))', size=9, wrap=True, border=box)
    jl = "&".join(f'IF(SUMPRODUCT(--(${hc(j)}${rows[0]}:${hc(j)}${rows[-1]}<>""))>0,"J{j+1} ","")' for j in range(7))
    cell(f"D{r}", f'=IF(E{r}=0,"aucune operation",SUBSTITUTE(TRIM({jl})," ",", "))', size=9, wrap=True, border=box)
    cell(f"E{r}", f'=SUMPRODUCT(--({blk}<>""))', size=10, fmt="0", border=box)
    cell(f"F{r}", f"=$C${ROW_NBW}", size=10, fmt="0", border=box)
    cell(f"G{r}", f"=H{r}+I{r}", bold=True, size=11, fmt="0", border=box)
    cell(f"H{r}", "=" + "+".join(f'SUMPRODUCT({OCC},--($O${x}:$U${x}<>""))' for x in rows), size=10, fmt="0", border=box)
    cell(f"I{r}", f"={add()}", size=10, fmt="0", border=box)
    for t, col in zip(TYPES, ("J", "K", "L")):
        cell(f"{col}{r}", f'={prods[t]}+{add(f",{ADD_T},"+chr(34)+t+chr(34))}', size=9, fmt="0", border=box)
    ws.row_dimensions[r].height = 24

cell(f"B{ROW_REC_TOT}", "TOTAL GENERAL", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
for c in ("C", "D"):
    cell(f"{c}{ROW_REC_TOT}", "", fill=NAVY, border=box)
for c in ("E", "F", "G", "H", "I", "J", "K", "L"):
    f = f"=$C${ROW_NBW}" if c == "F" else f"=SUM({c}{ROW_REC0}:{c}{ROW_REC_TOT-1})"
    cell(f"{c}{ROW_REC_TOT}", f, bold=True, size=11 if c == "G" else 10,
         color="FFFFFF", fill=NAVY, fmt="0", border=box)

r = ROW_REC_TOT + 1
ws.merge_cells(f"B{r}:L{r+1}")
cell(f"B{r}", "Methode : le recapitulatif lit le tableau de saisie du mois filtre. Pour chaque jour opere, le nombre d'occurrences "
              "de ce jour dans le mois est compte a partir des dates reelles de la grille (ligne \"Occurrences du jour dans le "
              "mois\", semaines partielles incluses), chaque vol de la journee (vols 1, 2 et 3) etant compte separement. Total "
              "rotations du mois = programme du mois + vols additionnels dates ; les colonnes Total 77W / 778 / A320 le ventilent "
              "par type avion.",
     size=8, italic=True, color="595959", halign="left", wrap=True)

# ---------------------------------------------------------------- 1) les 12 tableaux de saisie
ws.merge_cells(f"B{ROW_SEC_TITLE}:L{ROW_SEC_TITLE}")
cell(f"B{ROW_SEC_TITLE}", "1) SAISIE DU PROGRAMME : 12 TABLEAUX, UN PAR MOIS",
     bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
ws.merge_cells(f"B{ROW_SEC_TITLE+1}:L{ROW_SEC_TITLE+2}")
cell(f"B{ROW_SEC_TITLE+1}",
     "Cellules en BLEU = saisie utilisateur. Les 12 tableaux ci-dessous sont identiques : chaque route y occupe 3 lignes (vols 1, 2 "
     "et 3 de la journee) et chaque cellule J1-J7 porte le type avion opere ce jour-la (liste deroulante 77W / 778 / A320, vide = "
     "pas de vol). Renseigner plusieurs lignes le meme jour = 2 ou 3 rotations, avec le meme type ou des types differents. La "
     "frequence peut donc etre totalement differente d'un mois a l'autre. Le tableau du mois filtre est signale en vert. Pour "
     "reprendre un mois sur un autre : selectionner la zone des 7 jours d'un tableau (36 lignes), copier, coller dans le tableau "
     "cible. Valeurs livrees = exemples a ajuster.",
     size=8, italic=True, color=BROWN, halign="left", wrap=True)
ws.row_dimensions[ROW_SEC_TITLE+1].height = 46

cell(f"B{ROW_NAV}", "Aller au mois :", bold=True, size=9, halign="left", fill=GREY, border=box)
cell(f"B{ROW_NAV+1}", "(clic sur un mois)", size=8, italic=True, halign="left", color="808080", fill=GREY, border=box)
for k, (lab, y, m) in enumerate(MOIS):
    ref = f"{get_column_letter(3 + k % 6)}{ROW_NAV + k // 6}"
    c = cell(ref, lab, size=9, bold=True, color="0563C1", fill=GREY, border=box)
    c.hyperlink = Hyperlink(ref=ref, location=f"'{ws.title}'!B{blk_title(k)}", tooltip=f"Aller au tableau {lab}")

prog_hdr = ["Route / vol du jour"] + JOURS + ["Rot./sem. (auto)", "Type(s) avion (auto)"]
for k, (lab, y, m) in enumerate(MOIS):
    t = blk_title(k)
    ws.merge_cells(f"B{t}:I{t}")
    cell(f"B{t}", f'="TABLEAU {k+1}/12  -  PROGRAMME DE {lab.upper()}"&IF($C$4="{lab}","          <<< MOIS ACTUELLEMENT FILTRE","")',
         bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
    ws.row_dimensions[t].height = 20
    c = cell(f"K{t}", "^ Retour en haut", size=8, italic=True, color="0563C1", halign="right")
    c.hyperlink = Hyperlink(ref=f"K{t}", location=f"'{ws.title}'!B1", tooltip="Retour en haut de la feuille")
    for i, h in enumerate(prog_hdr):
        cell(f"{get_column_letter(2+i)}{t+1}", h, bold=True, size=9, color="FFFFFF", fill="A6774B", wrap=True, border=box)
    ws.row_dimensions[t+1].height = 26
    for i, (route, colr, white) in enumerate(ROUTES):
        p = t + 2 + NB_SLOTS*i
        prog = SAISON.get((lab, route), BASE[route])
        for s in range(NB_SLOTS):
            r = p + s
            cell(f"B{r}", route if s == 0 else f"      + vol {s+1} du jour",
                 bold=(s == 0), size=10 if s == 0 else 8, italic=(s > 0),
                 color="000000" if s == 0 else "595959", halign="left", border=box)
            for j in range(7):
                cell(f"{get_column_letter(3+j)}{r}", prog[s][j], size=10, color=BLUE_IN, border=box)
        span = f"C{p}:I{p+NB_SLOTS-1}"
        for col in ("J", "K"):
            ws.merge_cells(f"{col}{p}:{col}{p+NB_SLOTS-1}")
        cell(f"J{p}", f"=COUNTA({span})", size=10, bold=True, fmt="0", border=box)
        tl = "&".join(f'IF(COUNTIF({span},"{t2}")>0,"{t2}("&COUNTIF({span},"{t2}")&") ","")' for t2 in TYPES)
        cell(f"K{p}", f'=IF(COUNTA({span})=0,"-",SUBSTITUTE(TRIM({tl})," ",", "))', size=9, border=box)
        for s in range(1, NB_SLOTS):
            for col in ("J", "K"):
                cell(f"{col}{p+s}", border=box)
ws[f"B{DATA0}"].comment = Comment(
    "12 tableaux identiques, un par mois.\n"
    "3 lignes par route = les vols 1, 2 et 3 de la journee.\n"
    "Cellule vide = pas de vol ce jour-la.\n"
    "Copier/coller la zone des 7 jours pour reprendre un mois sur un autre.", "Modele")

# ---------------------------------------------------------------- 2) vols additionnels
ws.merge_cells(f"B{ROW_ADD_TITLE}:L{ROW_ADD_TITLE}")
cell(f"B{ROW_ADD_TITLE}", "2) VOLS ADDITIONNELS (en plus du programme du mois)",
     bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
ws.merge_cells(f"B{ROW_ADD_TITLE+1}:L{ROW_ADD_TITLE+2}")
cell(f"B{ROW_ADD_TITLE+1}", "Une ligne = un vol supplementaire date (charter, renfort ponctuel, fret...). Il s'ajoute au programme du mois, "
                            "sur n'importe quel jour, y compris un jour deja opere ou un jour hors programme. Il apparait dans la grille "
                            "entre parentheses avec un + et une bordure rouge, et il est compte dans le recapitulatif du mois concerne. "
                            "Pour une modification recurrente sur tout un mois, modifier plutot le tableau du mois.",
     size=8, italic=True, color=BROWN, halign="left", wrap=True)
ws.row_dimensions[ROW_ADD_TITLE+1].height = 26
for i, h in enumerate(["Date du vol", "Route", "Type avion", "Nb de rotations", "Commentaire (libre)"]):
    cell(f"{get_column_letter(2+i)}{ROW_ADD_HDR}", h, bold=True, size=9, color="FFFFFF", fill=BROWN, wrap=True, border=box)
cell(f"H{ROW_ADD_HDR}", "Cle technique (auto)", bold=True, size=8, color="FFFFFF", fill="A6A6A6", wrap=True, border=box)
for i in range(NB_ADD_ROWS):
    r = ROW_ADD0 + i
    v = VOLS_ADD[i] if i < len(VOLS_ADD) else None
    cell(f"B{r}", v[0] if v else None, size=10, color=BLUE_IN, fmt="DD/MM/YYYY", border=box)
    cell(f"C{r}", v[1] if v else None, size=10, color=BLUE_IN, border=box)
    cell(f"D{r}", v[2] if v else None, size=10, color=BLUE_IN, border=box)
    cell(f"E{r}", v[3] if v else None, size=10, color=BLUE_IN, fmt="0", border=box)
    cell(f"F{r}", v[4] if v else None, size=9, italic=True, color="595959", halign="left", border=box)
    cell(f"H{r}", f'=IF(OR($B{r}="",$C{r}=""),"",$B{r}&"|"&$C{r})', size=8, color="A6A6A6", border=box)
ws[f"B{ROW_ADD0}"].comment = Comment(
    "Date du vol (annee avril 2027 - mars 2028).\n"
    "Nb de rotations : 1 par defaut, 2 pour une double rotation le meme jour.\n"
    "Les 3 premieres lignes sont des exemples : les supprimer.", "Modele")

# ---------------------------------------------------------------- listes de reference (W..Z)
for col, h in (("W", "Mois (filtre)"), ("X", "1er jour"), ("Y", "Types avion"), ("Z", "Routes")):
    cell(f"{col}{REF0-1}", h, bold=True, size=8, color="FFFFFF", fill="808080", wrap=True, border=box)
for i, (lab, y, m) in enumerate(MOIS):
    cell(f"W{REF0+i}", lab, size=8, halign="left", border=box)
    cell(f"X{REF0+i}", f"=DATE({y},{m},1)", size=8, fmt="DD/MM/YYYY", border=box)
for i, t in enumerate(TYPES):
    cell(f"Y{REF0+i}", t, size=8, border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    cell(f"Z{REF0+i}", route, size=8, border=box)

# ---------------------------------------------------------------- listes deroulantes
def add_dv(dv, ranges):
    ws.add_data_validation(dv)
    for rng in (ranges if isinstance(ranges, (list, tuple)) else [ranges]):
        dv.add(rng)

L_MOIS, L_TYPE = f"=$W${REF0}:$W${REF0+11}", f"=$Y${REF0}:$Y${REF0+len(TYPES)-1}"
L_ROUTE = f"=$Z${REF0}:$Z${REF0+11}"
add_dv(DataValidation(type="list", formula1=L_MOIS, allow_blank=False), "C4")
add_dv(DataValidation(type="list", formula1=L_ROUTE, allow_blank=False), "C5")
add_dv(DataValidation(type="list", formula1=L_TYPE, allow_blank=True,
                      error="Choisir un type avion, ou laisser vide s'il n'y a pas de vol ce jour-la.",
                      errorTitle="Type avion invalide"),
       [f"C{blk_title(k)+2}:I{blk_title(k)+1+NB_ROWS_BLK}" for k in range(len(MOIS))])
add_dv(DataValidation(type="list", formula1=L_ROUTE, allow_blank=True), f"C{ROW_ADD0}:C{ROW_ADD_END}")
add_dv(DataValidation(type="list", formula1=L_TYPE, allow_blank=True), f"D{ROW_ADD0}:D{ROW_ADD_END}")
add_dv(DataValidation(type="whole", operator="between", formula1="1", formula2="20", allow_blank=True,
                      error="Nombre de rotations : entier entre 1 et 20.", errorTitle="Valeur invalide"),
       f"E{ROW_ADD0}:E{ROW_ADD_END}")
add_dv(DataValidation(type="date", operator="between", formula1="DATE(2027,4,1)", formula2="DATE(2028,3,31)",
                      allow_blank=True, errorStyle="warning", errorTitle="Date hors annee d'exploitation",
                      error="Date hors de l'annee avril 2027 - mars 2028 : le vol ne sera visible dans aucun mois."),
       f"B{ROW_ADD0}:B{ROW_ADD_END}")

# ---------------------------------------------------------------- mise en forme conditionnelle
grid_ranges = " ".join(f"{get_column_letter(GRID_C0)}{r}:{get_column_letter(GRID_C0+6)}{r}" for r in type_rows)
c0 = get_column_letter(GRID_C0)
ws.conditional_formatting.add(
    grid_ranges,
    FormulaRule(formula=[f'ISNUMBER(SEARCH("(+",{c0}{ROW_W1}))'],
                border=Border(left=Side(style="medium", color=RED), right=Side(style="medium", color=RED),
                              top=Side(style="medium", color=RED), bottom=Side(style="medium", color=RED)),
                stopIfTrue=False))
blk_route_ranges = " ".join(f"B{blk_title(k)+2}:B{blk_title(k)+1+NB_ROWS_BLK}" for k in range(len(MOIS)))
for route, colr, white in ROUTES:
    fc = "FFFFFF" if white else "000000"
    ws.conditional_formatting.add(
        grid_ranges,
        FormulaRule(formula=[f'AND($C$5="{route}",{c0}{ROW_W1}<>"")'],
                    fill=PatternFill("solid", fgColor=colr), font=Font(name=F, size=10, bold=True, color=fc),
                    stopIfTrue=True))
    ws.conditional_formatting.add(
        f"B{ROW_REC0}:B{ROW_REC_TOT-1}",
        FormulaRule(formula=[f'$B{ROW_REC0}="{route}"'],
                    fill=PatternFill("solid", fgColor=colr),
                    font=Font(name=F, size=10, bold=True, color=fc), stopIfTrue=True))
    conds = ",".join([f'$B{DATA0}="{route}"'] + [f'$B{DATA0-s}="{route}"' for s in range(1, NB_SLOTS)])
    ws.conditional_formatting.add(
        blk_route_ranges,
        FormulaRule(formula=[f'OR({conds})'], fill=PatternFill("solid", fgColor=colr),
                    font=Font(name=F, size=10, bold=True, color=fc), stopIfTrue=True))
# tableau du mois filtre : titre en vert
for k, (lab, y, m) in enumerate(MOIS):
    t = blk_title(k)
    ws.conditional_formatting.add(
        f"B{t}:K{t}",
        FormulaRule(formula=[f'$C$4="{lab}"'], fill=PatternFill("solid", fgColor=ACTIVE),
                    font=Font(name=F, size=11, bold=True, color="FFFFFF"), stopIfTrue=True))

# ---------------------------------------------------------------- mise en page
widths = {"A": 2, "B": 30, "J": 18, "K": 18, "L": 14, "M": 3, "N": 26, "V": 3,
          "W": 18, "X": 12, "Y": 10, "Z": 10}
for j in range(7):
    widths[get_column_letter(GRID_C0+j)] = 18
for j in range(7):
    widths[get_column_letter(HELP_C0+j)] = 10
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = "C8"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_area = f"B1:L{ROW_REC_TOT+2}"

wb.save(OUT)
print(f"ecrit -> {OUT}  (tableaux lignes {BLK0} a {DATA_LAST}, vols add. {ROW_ADD0}-{ROW_ADD_END})")
