# -*- coding: utf-8 -*-
"""Genere le classeur 'Programme de vols mensuel' (2 feuilles).

Feuille 1 "Programme de vols" : 12 tableaux de saisie (un par mois). Chaque route y occupe
3 lignes = les 3 types avion (77W / 778 / A320) ; on coche les jours d'operation J1-J7.
Aucune notion de sieges sur cette feuille.
Feuille 2 "Analyse capacite" : interpretation automatique de la feuille 1 avec les capacites
(sieges par rotation) : sieges offerts par mois, par route et par type avion.
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils import get_column_letter as gl
from openpyxl.comments import Comment

OUT = "/home/user/Test-App/Programme_vols_mensuel_2027-2028.xlsx"
F = "Arial"
X = "x"          # marque de case cochee

ROUTES = [
    ("CDGRUN", "9DC3E6", False), ("CDGDZA", "2E75B6", True),  ("BKKRUN", "FFE699", False),
    ("DZARUN", "F4B183", False), ("MRURUN", "C6E0B4", False), ("NOSRUN", "FFD966", False),
    ("RUNTNR", "D9D2E9", False), ("JNBRUN", "A9D08E", False), ("CPTRUN", "8FAADC", False),
    ("DIERUN", "FFC7CE", False), ("RRGRUN", "B7DEE8", False), ("TMMRUN", "D5A6BD", False),
]
TYPES = ["77W", "778", "A320"]
SIEGES = {"77W": 438, "778": 262, "A320": 174}
SIEGES_NOTE = {"77W": "Valeur fournie (77W = 777-300ER).",
               "778": "Valeur fournie pour le 787 (Dreamliner) ; code type 778 de la feuille 1.",
               "A320": "Valeur fournie (A320)."}
MOIS = [("avril 2027",2027,4),("mai 2027",2027,5),("juin 2027",2027,6),("juillet 2027",2027,7),
        ("aout 2027",2027,8),("septembre 2027",2027,9),("octobre 2027",2027,10),
        ("novembre 2027",2027,11),("decembre 2027",2027,12),("janvier 2028",2028,1),
        ("fevrier 2028",2028,2),("mars 2028",2028,3)]
JOURS = ["J1 - Lundi","J2 - Mardi","J3 - Mercredi","J4 - Jeudi","J5 - Vendredi","J6 - Samedi","J7 - Dimanche"]

# Programme de base recopie dans les 12 tableaux : {route: {type: [J1..J7 coche ?]}}
BASE = {
    "CDGRUN": {"77W": [1,1,1,1,1,1,1], "778": [1,0,0,0,1,0,0], "A320": [1,0,0,0,0,0,0]},
    "CDGDZA": {"77W": [0,0,0,0,0,0,1], "778": [0,1,0,0,1,0,0], "A320": [0]*7},
    "BKKRUN": {"77W": [0]*7,           "778": [0,0,1,0,0,1,0], "A320": [0]*7},
    "DZARUN": {"77W": [0]*7,           "778": [0,0,1,0,0,0,0], "A320": [1,1,1,1,1,0,0]},
    "MRURUN": {"77W": [0]*7,           "778": [0]*7,           "A320": [1,0,1,0,1,0,1]},
    "NOSRUN": {"77W": [0]*7,           "778": [0]*7,           "A320": [1,0,0,1,0,1,0]},
    "RUNTNR": {"77W": [0]*7,           "778": [0]*7,           "A320": [0,1,0,1,0,1,0]},
    "JNBRUN": {"77W": [0]*7,           "778": [1,0,0,0,1,0,0], "A320": [0]*7},
    "CPTRUN": {"77W": [0]*7,           "778": [0,0,0,1,0,0,0], "A320": [0]*7},
    "DIERUN": {"77W": [0]*7,           "778": [0]*7,           "A320": [0,1,0,0,0,1,0]},
    "RRGRUN": {"77W": [0]*7,           "778": [0]*7,           "A320": [1,1,1,1,1,1,0]},
    "TMMRUN": {"77W": [0]*7,           "778": [0]*7,           "A320": [0,0,1,0,0,0,1]},
}
# Exemples de saisonnalite : programme different pour certains mois {(mois, route): {type: [...]}}
SAISON = {
    ("juillet 2027", "NOSRUN"):  {"77W": [0]*7, "778": [0]*7, "A320": [1,0,1,1,0,1,0]},
    ("decembre 2027", "NOSRUN"): {"77W": [0]*7, "778": [0]*7, "A320": [1,1,0,1,0,1,1]},
    ("aout 2027", "CDGRUN"):     {"77W": [1]*7, "778": [1]*7, "A320": [0]*7},
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
ROW_SEC_TITLE, ROW_NAV, BLK0 = 41, 44, 47
NB_T = len(TYPES)
NB_ROWS_BLK = NB_T*len(ROUTES)                    # 36
STEP = NB_ROWS_BLK + 3                            # 39
DATA0 = BLK0 + 2                                  # 49
DATA_LAST = BLK0 + (len(MOIS)-1)*STEP + 1 + NB_ROWS_BLK      # 513
ROW_ADD_TITLE = DATA_LAST + 3                     # 516
ROW_ADD_HDR, ROW_ADD0 = ROW_ADD_TITLE + 4, ROW_ADD_TITLE + 5
NB_ADD_ROWS = 40
ROW_ADD_END = ROW_ADD0 + NB_ADD_ROWS - 1          # 560
SAI_C0 = 4                                        # colonne D = J1 dans les tableaux de saisie
R_REP0 = 4                                        # reperes : lignes 4..13
R_HELP_TITLE, R_HELP_HDR, R_HELP_TXT, R_HELP_S0 = 15, 16, 17, 18
R_MAT_TITLE, R_MAT_HDR, R_MAT0 = 22, 23, 24
HELP_C0, REF0 = 15, 4

wb = Workbook()
ws = wb.active
ws.title = "Programme de vols"
ws.sheet_view.showGridLines = False

def C(sheet, ref, value=None, bold=False, size=10, color="000000", fill=None, halign="center",
      valign="center", italic=False, wrap=False, fmt=None, border=None):
    c = sheet[ref]
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
cell = lambda *a, **k: C(ws, *a, **k)

thin = Side(style="thin", color="A6A6A6")
med = Side(style="medium", color="404040")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
NAVY, LIGHT, GREY, BROWN, RED, ACTIVE, CHECK = "1F3864", "D9E2F3", "F2F2F2", "833C0C", "C00000", "375623", "70AD47"
BLUE_IN = "0000FF"
CHK_FMT = ';;;"X"'                          # toute saisie texte s'affiche en X centre

BIG = f"$D${DATA0}:$J${DATA_LAST}"                             # les 12 tableaux (7 jours)
MAT = f"$O${R_MAT0}:$U${R_MAT0+NB_ROWS_BLK-1}"                 # programme du mois filtre (auto)
ADD_D, ADD_R = f"$B${ROW_ADD0}:$B${ROW_ADD_END}", f"$C${ROW_ADD0}:$C${ROW_ADD_END}"
ADD_T, ADD_N = f"$D${ROW_ADD0}:$D${ROW_ADD_END}", f"$E${ROW_ADD0}:$E${ROW_ADD_END}"
ADD_K = f"$H${ROW_ADD0}:$H${ROW_ADD_END}"
OCC = f"$C${ROW_OCC}:$I${ROW_OCC}"
hc = lambda j: gl(HELP_C0+j)
blk_title = lambda k: BLK0 + k*STEP
mat_row = lambda i, t: R_MAT0 + NB_T*i + TYPES.index(t)

# ---------------------------------------------------------------- titre + filtres
ws.merge_cells("B1:I1")
cell("B1", "PROGRAMME DE VOLS MENSUEL - ANNEE D'EXPLOITATION AVRIL 2027 / MARS 2028",
     bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
ws.row_dimensions[1].height = 30
ws.merge_cells("B2:I2")
cell("B2", "Feuille unique de programmation : les deux filtres ci-dessous (Mois et Route) pilotent la grille, le recapitulatif ne "
           "suit que le filtre Mois. La saisie se fait dans 12 tableaux identiques, un par mois : chaque route y occupe 3 lignes, "
           "une par type avion (77W / 778 / A320), et il suffit de COCHER les jours d'operation. La feuille \"Analyse capacite\" "
           "traduit ce programme en sieges offerts.",
     size=9, italic=True, halign="left", wrap=True, color="404040")
ws.row_dimensions[2].height = 26
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

# ---------------------------------------------------------------- reperes calcules
ws.merge_cells("N3:O3")
cell("N3", "REPERES CALCULES", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left")
W, Xr = f"$W${REF0}:$W${REF0+11}", f"$X${REF0}:$X${REF0+11}"
rec_col = lambda c: f'IFERROR(INDEX(${c}${ROW_REC0}:${c}${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")'
reperes = [("1er jour du mois", f"=INDEX({Xr},MATCH($C$4,{W},0))", "DD/MM/YYYY", False),
           ("Dernier jour du mois", "=EOMONTH($O$4,0)", "DD/MM/YYYY", False),
           ("Lundi semaine 1", "=$O$4-WEEKDAY($O$4,3)", "DD/MM/YYYY", False),
           ("Tableau de saisie lu", f"=MATCH($C$4,{W},0)", "0", False),
           ("Type(s) avion", f"={rec_col('C')}", "General", False),
           ("Rotations / semaine", f"={rec_col('E')}", "0", False),
           ("Total rotations du mois", f"={rec_col('G')}", "0", True),
           ("        dont 77W", f"={rec_col('J')}", "0", False),
           ("        dont 778", f"={rec_col('K')}", "0", False),
           ("        dont A320", f"={rec_col('L')}", "0", False)]
for i, (lab, fml, fmt, strong) in enumerate(reperes):
    r = R_REP0 + i
    cell(f"N{r}", lab, size=9, bold=strong, halign="left", fill="FFF2CC" if strong else GREY, border=box)
    cell(f"O{r}", fml, size=9, bold=True, fmt=fmt, fill="FFF2CC" if strong else None, border=box)
MIDX = "$O$7"

# ---------------------------------------------------------------- auto : programme du mois filtre
ws.merge_cells(f"N{R_MAT_TITLE}:{hc(6)}{R_MAT_TITLE}")
cell(f"N{R_MAT_TITLE}", "PROGRAMME DU MOIS FILTRE, TOUTES ROUTES (auto - recopie du tableau du mois)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"N{R_MAT_HDR}", "Route / type", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_MAT_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    for ti, t in enumerate(TYPES):
        r = mat_row(i, t)
        cell(f"N{r}", f"{route} - {t}", size=8, halign="left", border=box)
        for j in range(7):
            rel = f"({MIDX}-1)*{STEP}+{NB_T*i+ti+1}"
            cell(f"{hc(j)}{r}", f'=IFERROR(INDEX({BIG},{rel},{j+1})&"","")', size=8, border=box)

# ---------------------------------------------------------------- auto : programme de la route filtree
ws.merge_cells(f"N{R_HELP_TITLE}:{hc(6)}{R_HELP_TITLE}")
cell(f"N{R_HELP_TITLE}", "PROGRAMME DE LA ROUTE FILTREE (auto - alimente la grille)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"N{R_HELP_HDR}", "Jour", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_HELP_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
cell(f"N{R_HELP_TXT}", "Vols du jour (affichage grille)", bold=True, size=8, halign="left", fill=GREY, border=box)
for ti, t in enumerate(TYPES):
    cell(f"N{R_HELP_S0+ti}", f"{t} opere ?", size=8, halign="left", fill=GREY, border=box)
RIDX = f'MATCH($C$5,$Z${REF0}:$Z${REF0+11},0)'
for j in range(7):
    c = hc(j)
    for ti, t in enumerate(TYPES):
        cell(f"{c}{R_HELP_S0+ti}",
             f'=IFERROR(INDEX({MAT},({RIDX}-1)*{NB_T}+{ti+1},{j+1})&"","")', size=9, border=box)
    ref = [f"{c}${R_HELP_S0+ti}" for ti in range(NB_T)]
    parts = [f'IF({ref[0]}="","","{TYPES[0]}")']
    for ti in range(1, NB_T):
        before = (f'{ref[0]}=""' if ti == 1 else "AND(" + ",".join(f'{x}=""' for x in ref[:ti]) + ")")
        parts.append(f'IF({ref[ti]}="","",IF({before},"","/")&"{TYPES[ti]}")')
    cell(f"{c}{R_HELP_TXT}", "=" + "&".join(parts), bold=True, size=9, border=box)

# ---------------------------------------------------------------- grille
ws.merge_cells(f"B{ROW_HDR-1}:I{ROW_HDR-1}")
cell(f"B{ROW_HDR-1}", "GRILLE MENSUELLE  -  type avion opere par semaine et par jour (filtres Mois + Route)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
cell(f"B{ROW_HDR}", "Semaine", bold=True, size=10, color="FFFFFF", fill="2F5597", border=box)
for j in range(7):
    cell(f"{gl(GRID_C0+j)}{ROW_HDR}", JOURS[j], bold=True, size=10, color="FFFFFF", fill="2F5597", border=box)
type_rows, date_rows = [], []
for k in range(1, NB_WEEKS+1):
    rt, rd = ROW_W1 + 2*(k-1), ROW_W1 + 2*(k-1) + 1
    type_rows.append(rt); date_rows.append(rd)
    first, last = f"{gl(GRID_C0)}{rd}", f"{gl(GRID_C0+6)}{rd}"
    cell(f"B{rt}", f"Semaine {k}", bold=True, size=10, fill=LIGHT, halign="left", border=box)
    cell(f"B{rd}", f'=IF(COUNT({first}:{last})=0,"","du "&TEXT(MIN({first}:{last}),"DD/MM")&'
                   f'" au "&TEXT(MAX({first}:{last}),"DD/MM"))',
         size=8, italic=True, color="595959", fill=LIGHT, halign="left", border=box)
    ws.row_dimensions[rt].height = 22
    ws.row_dimensions[rd].height = 14
    for j in range(7):
        col, off = gl(GRID_C0+j), 7*(k-1) + j
        d = f"{col}{rd}"
        cell(d, f'=IF(OR($O$6+{off}<$O$4,$O$6+{off}>$O$5),"",$O$6+{off})',
             size=8, italic=True, color="595959", fmt="DD/MM", border=box)
        nrot = f"SUMIFS({ADD_N},{ADD_R},$C$5,{ADD_D},{d})"
        nrows = f"COUNTIFS({ADD_R},$C$5,{ADD_D},{d})"
        t1 = f'IFERROR(INDEX({ADD_T},MATCH({d}&"|"&$C$5,{ADD_K},0)),"")'
        lbl = f'IF({nrows}=1,IF({nrot}=1,{t1},{nrot}&"x"&{t1}),{nrot}&" vols")'
        cell(f"{col}{rt}", f'=IF({d}="","",TRIM(${hc(j)}${R_HELP_TXT}&IF({nrot}=0,""," (+"&{lbl}&")")))',
             bold=True, size=10, border=box)
cell(f"B{ROW_OCC}", "Occurrences du jour dans le mois", bold=True, size=9, halign="left", fill="E7E6E6", border=box)
for j in range(7):
    col = gl(GRID_C0+j)
    cell(f"{col}{ROW_OCC}", "=COUNT(" + ",".join(f"{col}{r}" for r in date_rows) + ")",
         bold=True, size=10, fill="E7E6E6", border=box)
cell(f"B{ROW_NBW}", "Nb de semaines calendaires du mois", bold=True, size=9, halign="left", fill="E7E6E6", border=box)
cell(f"C{ROW_NBW}", "=" + "+".join(f'IF(COUNT({gl(GRID_C0)}{r}:{gl(GRID_C0+6)}{r})>0,1,0)' for r in date_rows),
     bold=True, size=10, fill="E7E6E6", border=box)
ws.merge_cells(f"D{ROW_NBW}:I{ROW_NBW}")
cell(f"D{ROW_NBW}", 'Lecture : "778" = 1 vol du programme  |  "77W/778/A320" = 2 ou 3 vols le meme jour  |  "778 (+A320)" = '
                    'programme + vol additionnel (bordure rouge)  |  "(+A320)" = vol additionnel seul.',
     size=8, italic=True, color="595959", halign="left")

# ---------------------------------------------------------------- recapitulatif
ws.merge_cells(f"B{ROW_REC_TITLE}:L{ROW_REC_TITLE}")
cell(f"B{ROW_REC_TITLE}", "RECAPITULATIF MENSUEL - TOUTES ROUTES  (depend uniquement du filtre Mois)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
for i, h in enumerate(["Route", "Type(s) avion", "Semaine type du mois (jours J1-J7)", "Rotations/semaine",
                       "Nb de semaines/occurrences dans le mois", "Total rotations du mois",
                       "dont programme du mois", "dont vols additionnels", "Total 77W", "Total 778", "Total A320"]):
    cell(f"{gl(2+i)}{ROW_REC_HDR}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
ws.row_dimensions[ROW_REC_HDR].height = 44
for i in range(len(ROUTES)):
    r = ROW_REC0 + i
    rows = {t: mat_row(i, t) for t in TYPES}
    blk = f"$O${rows[TYPES[0]]}:$U${rows[TYPES[-1]]}"
    add = lambda extra="": f'SUMIFS({ADD_N},{ADD_R},$B{r}{extra},{ADD_D},">="&$O$4,{ADD_D},"<="&$O$5)'
    nsem = {t: f'SUMPRODUCT(--($O${rows[t]}:$U${rows[t]}<>""))' for t in TYPES}
    prod = {t: f'SUMPRODUCT({OCC},--($O${rows[t]}:$U${rows[t]}<>""))' for t in TYPES}
    cell(f"B{r}", f"=$Z${REF0+i}", bold=True, size=10, border=box)
    tl = "&".join(f'IF({nsem[t]}>0,"{t}("&{nsem[t]}&") ","")' for t in TYPES)
    cell(f"C{r}", f'=IF(E{r}=0,"-",SUBSTITUTE(TRIM({tl})," ",", "))', size=9, wrap=True, border=box)
    jl = "&".join(f'IF(SUMPRODUCT(--(${hc(j)}${rows[TYPES[0]]}:${hc(j)}${rows[TYPES[-1]]}<>""))>0,"J{j+1} ","")'
                  for j in range(7))
    cell(f"D{r}", f'=IF(E{r}=0,"aucune operation",SUBSTITUTE(TRIM({jl})," ",", "))', size=9, wrap=True, border=box)
    cell(f"E{r}", f'=SUMPRODUCT(--({blk}<>""))', size=10, fmt="0", border=box)
    cell(f"F{r}", f"=$C${ROW_NBW}", size=10, fmt="0", border=box)
    cell(f"G{r}", f"=H{r}+I{r}", bold=True, size=11, fmt="0", border=box)
    cell(f"H{r}", "=" + "+".join(prod[t] for t in TYPES), size=10, fmt="0", border=box)
    cell(f"I{r}", f"={add()}", size=10, fmt="0", border=box)
    for t, col in zip(TYPES, ("J", "K", "L")):
        cell(f"{col}{r}", f'={prod[t]}+{add(f",{ADD_T},"+chr(34)+t+chr(34))}', size=9, fmt="0", border=box)
    ws.row_dimensions[r].height = 24
cell(f"B{ROW_REC_TOT}", "TOTAL GENERAL", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
for c in ("C", "D"):
    cell(f"{c}{ROW_REC_TOT}", "", fill=NAVY, border=box)
for c in ("E", "F", "G", "H", "I", "J", "K", "L"):
    f = f"=$C${ROW_NBW}" if c == "F" else f"=SUM({c}{ROW_REC0}:{c}{ROW_REC_TOT-1})"
    cell(f"{c}{ROW_REC_TOT}", f, bold=True, size=11 if c == "G" else 10, color="FFFFFF", fill=NAVY, fmt="0", border=box)
r = ROW_REC_TOT + 1
ws.merge_cells(f"B{r}:L{r+1}")
cell(f"B{r}", "Methode : le recapitulatif lit le tableau de saisie du mois filtre. Pour chaque jour coche, le nombre d'occurrences "
              "de ce jour dans le mois est compte a partir des dates reelles de la grille (ligne \"Occurrences du jour dans le "
              "mois\", semaines partielles incluses), chaque type avion etant compte separement. Total rotations du mois = "
              "programme du mois + vols additionnels dates. Une case cochee = 1 rotation ; pour deux rotations du meme type le "
              "meme jour, utiliser la table des vols additionnels.",
     size=8, italic=True, color="595959", halign="left", wrap=True)

# ---------------------------------------------------------------- 12 tableaux de saisie
ws.merge_cells(f"B{ROW_SEC_TITLE}:L{ROW_SEC_TITLE}")
cell(f"B{ROW_SEC_TITLE}", "1) SAISIE DU PROGRAMME : 12 TABLEAUX, UN PAR MOIS  -  COCHER LES JOURS D'OPERATION",
     bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
ws.merge_cells(f"B{ROW_SEC_TITLE+1}:L{ROW_SEC_TITLE+2}")
cell(f"B{ROW_SEC_TITLE+1}",
     "COCHER UNE CASE : taper x dans la case du jour (elle devient une case verte cochee X). DECOCHER : touche Suppr. Pour cocher "
     "plusieurs cases d'un coup : selectionner la plage, taper x puis Ctrl+Entree. Chaque route occupe 3 lignes = les 3 types "
     "avion (77W / 778 / A320) : cocher la ligne du type qui opere. Cocher deux ou trois lignes le meme jour = 2 ou 3 rotations ce "
     "jour-la avec des types differents (une case cochee = 1 rotation ; pour 2 rotations du meme type le meme jour, utiliser la "
     "table des vols additionnels). Les 12 tableaux sont independants : la frequence peut donc etre differente chaque mois. Le "
     "tableau du mois filtre est signale en vert. Pour reprendre un mois sur un autre : copier la zone des 7 jours (36 lignes) et "
     "la coller dans le tableau cible. Valeurs livrees = exemples a ajuster.",
     size=8, italic=True, color=BROWN, halign="left", wrap=True)
ws.row_dimensions[ROW_SEC_TITLE+1].height = 58
cell(f"B{ROW_NAV}", "Aller au mois :", bold=True, size=9, halign="left", fill=GREY, border=box)
cell(f"B{ROW_NAV+1}", "(clic sur un mois)", size=8, italic=True, halign="left", color="808080", fill=GREY, border=box)
for k, (lab, y, m) in enumerate(MOIS):
    ref = f"{gl(3 + k % 6)}{ROW_NAV + k // 6}"
    c = cell(ref, lab, size=9, bold=True, color="0563C1", fill=GREY, border=box)
    c.hyperlink = Hyperlink(ref=ref, location=f"'{ws.title}'!B{blk_title(k)}", tooltip=f"Aller au tableau {lab}")

for k, (lab, y, m) in enumerate(MOIS):
    t0 = blk_title(k)
    ws.merge_cells(f"B{t0}:J{t0}")
    cell(f"B{t0}", f'="TABLEAU {k+1}/12  -  PROGRAMME DE {lab.upper()}"&IF($C$4="{lab}","          <<< MOIS ACTUELLEMENT FILTRE","")',
         bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
    ws.row_dimensions[t0].height = 20
    c = cell(f"K{t0}", "^ Retour en haut", size=8, italic=True, color="0563C1", halign="right")
    c.hyperlink = Hyperlink(ref=f"K{t0}", location=f"'{ws.title}'!B1", tooltip="Retour en haut de la feuille")
    for i, h in enumerate(["Route", "Type avion"] + JOURS + ["Rot./sem. par type", "Rot./sem. route"]):
        cell(f"{gl(2+i)}{t0+1}", h, bold=True, size=9, color="FFFFFF", fill="A6774B", wrap=True, border=box)
    ws.row_dimensions[t0+1].height = 26
    for i, (route, colr, white) in enumerate(ROUTES):
        p = t0 + 2 + NB_T*i
        prog = SAISON.get((lab, route), BASE[route])
        ws.merge_cells(f"B{p}:B{p+NB_T-1}")
        cell(f"B{p}", route, bold=True, size=11, halign="left", border=box)
        ws.merge_cells(f"L{p}:L{p+NB_T-1}")
        cell(f"L{p}", f"=SUM(K{p}:K{p+NB_T-1})", bold=True, size=11, fmt="0", border=box)
        for ti, t in enumerate(TYPES):
            r = p + ti
            cell(f"C{r}", t, bold=True, size=10, fill="F2E4D8", border=box)
            for j in range(7):
                cell(f"{gl(SAI_C0+j)}{r}", X if prog[t][j] else None,
                     bold=True, size=11, color=BLUE_IN, fmt=CHK_FMT, border=box)
            cell(f"K{r}", f"=COUNTA(D{r}:J{r})", size=9, fmt="0", border=box)
        for s in range(1, NB_T):
            cell(f"B{p+s}", border=box)
            cell(f"L{p+s}", border=box)
ws[f"B{DATA0}"].comment = Comment(
    "12 tableaux identiques, un par mois.\n"
    "3 lignes par route = les 3 types avion 77W / 778 / A320.\n"
    "Cocher un jour : taper x. Decocher : touche Suppr.\n"
    "Copier/coller la zone des 7 jours pour reprendre un mois sur un autre.", "Modele")

# ---------------------------------------------------------------- vols additionnels
ws.merge_cells(f"B{ROW_ADD_TITLE}:L{ROW_ADD_TITLE}")
cell(f"B{ROW_ADD_TITLE}", "2) VOLS ADDITIONNELS (en plus du programme du mois)",
     bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
ws.merge_cells(f"B{ROW_ADD_TITLE+1}:L{ROW_ADD_TITLE+2}")
cell(f"B{ROW_ADD_TITLE+1}", "Une ligne = un vol supplementaire date (charter, renfort ponctuel, fret, seconde rotation du meme type le "
                            "meme jour...). Il s'ajoute au programme du mois, sur n'importe quel jour. Il apparait dans la grille entre "
                            "parentheses avec un + et une bordure rouge, et il est compte dans le recapitulatif du mois concerne.",
     size=8, italic=True, color=BROWN, halign="left", wrap=True)
ws.row_dimensions[ROW_ADD_TITLE+1].height = 26
for i, h in enumerate(["Date du vol", "Route", "Type avion", "Nb de rotations", "Commentaire (libre)"]):
    cell(f"{gl(2+i)}{ROW_ADD_HDR}", h, bold=True, size=9, color="FFFFFF", fill=BROWN, wrap=True, border=box)
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

# ---------------------------------------------------------------- listes de reference
for col, h in (("W", "Mois (filtre)"), ("X", "1er jour"), ("Y", "Types avion"), ("Z", "Routes")):
    cell(f"{col}{REF0-1}", h, bold=True, size=8, color="FFFFFF", fill="808080", wrap=True, border=box)
for i, (lab, y, m) in enumerate(MOIS):
    cell(f"W{REF0+i}", lab, size=8, halign="left", border=box)
    cell(f"X{REF0+i}", f"=DATE({y},{m},1)", size=8, fmt="DD/MM/YYYY", border=box)
for i, t in enumerate(TYPES):
    cell(f"Y{REF0+i}", t, size=8, border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    cell(f"Z{REF0+i}", route, size=8, border=box)

# ---------------------------------------------------------------- validations
def add_dv(dv, ranges):
    ws.add_data_validation(dv)
    for rng in (ranges if isinstance(ranges, (list, tuple)) else [ranges]):
        dv.add(rng)
L_MOIS, L_TYPE = f"=$W${REF0}:$W${REF0+11}", f"=$Y${REF0}:$Y${REF0+len(TYPES)-1}"
L_ROUTE = f"=$Z${REF0}:$Z${REF0+11}"
add_dv(DataValidation(type="list", formula1=L_MOIS, allow_blank=False), "C4")
add_dv(DataValidation(type="list", formula1=L_ROUTE, allow_blank=False), "C5")
add_dv(DataValidation(type="list", formula1='"x,X"', allow_blank=True, showDropDown=True,
                      errorTitle="Case a cocher", error="Taper x pour cocher la case, ou touche Suppr pour la decocher."),
       [f"D{blk_title(k)+2}:J{blk_title(k)+1+NB_ROWS_BLK}" for k in range(len(MOIS))])
add_dv(DataValidation(type="list", formula1=L_ROUTE, allow_blank=True), f"C{ROW_ADD0}:C{ROW_ADD_END}")
add_dv(DataValidation(type="list", formula1=L_TYPE, allow_blank=True), f"D{ROW_ADD0}:D{ROW_ADD_END}")
add_dv(DataValidation(type="whole", operator="between", formula1="1", formula2="20", allow_blank=True,
                      error="Nombre de rotations : entier entre 1 et 20.", errorTitle="Valeur invalide"),
       f"E{ROW_ADD0}:E{ROW_ADD_END}")
add_dv(DataValidation(type="date", operator="between", formula1="DATE(2027,4,1)", formula2="DATE(2028,3,31)",
                      allow_blank=True, errorStyle="warning", errorTitle="Date hors annee d'exploitation",
                      error="Date hors de l'annee avril 2027 - mars 2028 : le vol ne sera visible dans aucun mois."),
       f"B{ROW_ADD0}:B{ROW_ADD_END}")

# ---------------------------------------------------------------- mises en forme conditionnelles
grid_ranges = " ".join(f"{gl(GRID_C0)}{r}:{gl(GRID_C0+6)}{r}" for r in type_rows)
c0 = gl(GRID_C0)
ws.conditional_formatting.add(
    grid_ranges,
    FormulaRule(formula=[f'ISNUMBER(SEARCH("(+",{c0}{ROW_W1}))'],
                border=Border(left=Side(style="medium", color=RED), right=Side(style="medium", color=RED),
                              top=Side(style="medium", color=RED), bottom=Side(style="medium", color=RED)),
                stopIfTrue=False))
chk_ranges = " ".join(f"D{blk_title(k)+2}:J{blk_title(k)+1+NB_ROWS_BLK}" for k in range(len(MOIS)))
ws.conditional_formatting.add(
    chk_ranges,
    FormulaRule(formula=[f'D{DATA0}<>""'], fill=PatternFill("solid", fgColor=CHECK),
                font=Font(name=F, size=11, bold=True, color="FFFFFF"), stopIfTrue=True))
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
        FormulaRule(formula=[f'$B{ROW_REC0}="{route}"'], fill=PatternFill("solid", fgColor=colr),
                    font=Font(name=F, size=10, bold=True, color=fc), stopIfTrue=True))
    ws.conditional_formatting.add(
        blk_route_ranges,
        FormulaRule(formula=[f'$B{DATA0}="{route}"'], fill=PatternFill("solid", fgColor=colr),
                    font=Font(name=F, size=11, bold=True, color=fc), stopIfTrue=True))
for k, (lab, y, m) in enumerate(MOIS):
    t0 = blk_title(k)
    ws.conditional_formatting.add(
        f"B{t0}:K{t0}",
        FormulaRule(formula=[f'$C$4="{lab}"'], fill=PatternFill("solid", fgColor=ACTIVE),
                    font=Font(name=F, size=11, bold=True, color="FFFFFF"), stopIfTrue=True))

widths = {"A": 2, "B": 26, "C": 16, "K": 16, "L": 16, "M": 3, "N": 26, "V": 3,
          "W": 18, "X": 12, "Y": 10, "Z": 10}
for j in range(7):
    widths[gl(GRID_C0+j)] = 16
for j in range(7):
    widths[gl(HELP_C0+j)] = 10
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = "C8"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_area = f"B1:L{ROW_REC_TOT+2}"

# ================================================================= FEUILLE 2 : ANALYSE CAPACITE
an = wb.create_sheet("Analyse capacite")
an.sheet_view.showGridLines = False
A = lambda *a, **k: C(an, *a, **k)
S1 = f"'{ws.title}'!"
R_HYP, R_SYN, R_SIE, R_ROT, R_TYP, R_CAL, R_DET = 5, 12, 28, 45, 62, 78, 93
mcol = lambda m: gl(3+m)                      # C..N : les 12 mois
rcol = lambda i: gl(3+i)                      # C..N : les 12 routes
det_row = lambda i, t: R_DET + 1 + NB_T*i + TYPES.index(t)
cal_row = lambda m: R_CAL + 1 + m
occ_row = lambda m: f"$C${cal_row(m)}:$I${cal_row(m)}"
seat = {t: f"$C${R_HYP+1+ti}" for ti, t in enumerate(TYPES)}

an.merge_cells("B1:J1")
A("B1", "ANALYSE DE CAPACITE - SIEGES OFFERTS (avril 2027 / mars 2028)",
  bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
an.row_dimensions[1].height = 30
an.merge_cells("B2:J2")
A("B2", "Cette feuille interprete automatiquement la feuille \"Programme de vols\" (les 12 tableaux mensuels et les vols "
        "additionnels) en la croisant avec les capacites ci-dessous. Aucune saisie ici, hormis les sieges par type avion. "
        "Toute modification du programme met a jour cette feuille.",
  size=9, italic=True, halign="left", wrap=True, color="404040")
an.row_dimensions[2].height = 26

an.merge_cells(f"B{R_HYP-1}:F{R_HYP-1}")
A(f"B{R_HYP-1}", "HYPOTHESES DE CAPACITE (sieges par rotation)", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
for i, h in enumerate(["Type avion", "Sieges par rotation", "Remarque"]):
    A(f"{gl(2+i)}{R_HYP}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", border=box)
for ti, t in enumerate(TYPES):
    r = R_HYP + 1 + ti
    A(f"B{r}", t, bold=True, size=10, border=box)
    A(f"C{r}", SIEGES[t], bold=True, size=10, color=BLUE_IN, fill="FFF2CC", fmt="#,##0", border=box)
    A(f"D{r}", SIEGES_NOTE[t], size=8, italic=True, color="595959", halign="left", border=box)
an.merge_cells(f"D{R_HYP+1}:F{R_HYP+1}"); an.merge_cells(f"D{R_HYP+2}:F{R_HYP+2}"); an.merge_cells(f"D{R_HYP+3}:F{R_HYP+3}")

# --- synthese annuelle par route
an.merge_cells(f"B{R_SYN-1}:F{R_SYN-1}")
A(f"B{R_SYN-1}", "SYNTHESE ANNUELLE PAR ROUTE", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
for i, h in enumerate(["Route", "Rotations sur l'annee", "Sieges offerts sur l'annee", "Sieges / rotation",
                       "Part du reseau (sieges)"]):
    A(f"{gl(2+i)}{R_SYN}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
an.row_dimensions[R_SYN].height = 30
for i, (route, colr, white) in enumerate(ROUTES):
    r = R_SYN + 1 + i
    col = rcol(i)
    A(f"B{r}", route, bold=True, size=10,
      fill=colr, color="FFFFFF" if white else "000000", border=box)
    A(f"C{r}", f"=SUM({col}${R_ROT+1}:{col}${R_ROT+len(MOIS)})", size=10, fmt="#,##0", border=box)
    A(f"D{r}", f"=SUM({col}${R_SIE+1}:{col}${R_SIE+len(MOIS)})", bold=True, size=10, fmt="#,##0", border=box)
    A(f"E{r}", f'=IFERROR(D{r}/C{r},"-")', size=10, fmt="#,##0", border=box)
    A(f"F{r}", f"=IFERROR(D{r}/$D${R_SYN+len(ROUTES)+1},0)", size=10, fmt="0.0%", border=box)
rt = R_SYN + len(ROUTES) + 1
A(f"B{rt}", "TOTAL RESEAU", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
for c in ("C", "D"):
    A(f"{c}{rt}", f"=SUM({c}{R_SYN+1}:{c}{rt-1})", bold=True, size=11, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)
A(f"E{rt}", f'=IFERROR(D{rt}/C{rt},"-")', bold=True, size=10, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)
A(f"F{rt}", f"=IFERROR(D{rt}/$D${rt},0)", bold=True, size=10, color="FFFFFF", fill=NAVY, fmt="0.0%", border=box)

# --- tableaux mois x routes
for base, titre, fmt in ((R_SIE, "SIEGES OFFERTS PAR MOIS ET PAR ROUTE", "#,##0"),
                         (R_ROT, "ROTATIONS PAR MOIS ET PAR ROUTE", "#,##0")):
    an.merge_cells(f"B{base-1}:O{base-1}")
    A(f"B{base-1}", titre, bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
    A(f"B{base}", "Mois", bold=True, size=9, color="FFFFFF", fill="2F5597", border=box)
    for i, (route, colr, white) in enumerate(ROUTES):
        A(f"{rcol(i)}{base}", route, bold=True, size=9, fill=colr, color="FFFFFF" if white else "000000", border=box)
    A(f"O{base}", "TOTAL", bold=True, size=9, color="FFFFFF", fill="2F5597", border=box)
    for m, (lab, y, mm) in enumerate(MOIS):
        r = base + 1 + m
        A(f"B{r}", f"={S1}$W${REF0+m}", size=9, halign="left", border=box)
        for i, (route, colr, white) in enumerate(ROUTES):
            d = mcol(m)
            if base == R_ROT:
                f = f"=SUM({d}{det_row(i,TYPES[0])}:{d}{det_row(i,TYPES[-1])})"
            else:
                f = "=" + "+".join(f"{seat[t]}*{d}{det_row(i,t)}" for t in TYPES)
            A(f"{rcol(i)}{r}", f, size=9, fmt=fmt, border=box)
        A(f"O{r}", f"=SUM(C{r}:N{r})", bold=True, size=9, fmt=fmt, border=box)
    r = base + 1 + len(MOIS)
    A(f"B{r}", "TOTAL ANNEE", bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left", border=box)
    for i in range(len(ROUTES) + 1):
        c = rcol(i) if i < len(ROUTES) else "O"
        A(f"{c}{r}", f"=SUM({c}{base+1}:{c}{base+len(MOIS)})", bold=True, size=9,
          color="FFFFFF", fill=NAVY, fmt=fmt, border=box)

# --- par type avion et par mois
an.merge_cells(f"B{R_TYP-1}:K{R_TYP-1}")
A(f"B{R_TYP-1}", "ROTATIONS ET SIEGES PAR TYPE AVION ET PAR MOIS", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
hdr = ["Mois"] + [f"Rotations {t}" for t in TYPES] + ["Total rotations"] + [f"Sieges {t}" for t in TYPES] + \
      ["Total sieges", "Sieges / rotation"]
for i, h in enumerate(hdr):
    A(f"{gl(2+i)}{R_TYP}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
an.row_dimensions[R_TYP].height = 30
DET_T = f"$P${R_DET+1}:$P${R_DET+NB_ROWS_BLK}"
for m, (lab, y, mm) in enumerate(MOIS):
    r = R_TYP + 1 + m
    d = mcol(m)
    A(f"B{r}", f"={S1}$W${REF0+m}", size=9, halign="left", border=box)
    for ti, t in enumerate(TYPES):
        A(f"{gl(3+ti)}{r}", f'=SUMIF({DET_T},"{t}",${d}${R_DET+1}:${d}${R_DET+NB_ROWS_BLK})',
          size=9, fmt="#,##0", border=box)
    A(f"F{r}", f"=SUM(C{r}:E{r})", bold=True, size=9, fmt="#,##0", border=box)
    for ti, t in enumerate(TYPES):
        A(f"{gl(7+ti)}{r}", f"={gl(3+ti)}{r}*{seat[t]}", size=9, fmt="#,##0", border=box)
    A(f"J{r}", f"=SUM(G{r}:I{r})", bold=True, size=9, fmt="#,##0", border=box)
    A(f"K{r}", f'=IFERROR(J{r}/F{r},"-")', size=9, fmt="#,##0", border=box)
r = R_TYP + 1 + len(MOIS)
A(f"B{r}", "TOTAL ANNEE", bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left", border=box)
for i in range(2, 10):
    c = gl(i+1)
    A(f"{c}{r}", f"=SUM({c}{R_TYP+1}:{c}{r-1})", bold=True, size=9, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)
A(f"K{r}", f'=IFERROR(J{r}/F{r},"-")', bold=True, size=9, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)

# --- calendrier auto
an.merge_cells(f"B{R_CAL-1}:L{R_CAL-1}")
A(f"B{R_CAL-1}", "CALENDRIER AUTO : occurrences de chaque jour de semaine dans le mois (semaines partielles incluses)",
  bold=True, size=9, color="FFFFFF", fill="808080", halign="left")
for i, h in enumerate(["Mois", "J1", "J2", "J3", "J4", "J5", "J6", "J7", "1er jour", "Dernier jour", "Nb de jours"]):
    A(f"{gl(2+i)}{R_CAL}", h, bold=True, size=8, fill=GREY, border=box)
for m, (lab, y, mm) in enumerate(MOIS):
    r = cal_row(m)
    A(f"B{r}", f"={S1}$W${REF0+m}", size=8, halign="left", border=box)
    A(f"J{r}", f"={S1}$X${REF0+m}", size=8, fmt="DD/MM/YYYY", border=box)
    A(f"K{r}", f"=EOMONTH($J{r},0)", size=8, fmt="DD/MM/YYYY", border=box)
    A(f"L{r}", f"=$K{r}-$J{r}+1", size=8, fmt="0", border=box)
    for j in range(7):
        off = f"MOD({j+1}-WEEKDAY($J{r},2),7)"
        A(f"{gl(3+j)}{r}", f"=IF({off}>=$L{r},0,INT(($L{r}-1-{off})/7)+1)", size=8, fmt="0", border=box)

# --- detail auto : rotations par route, type et mois
an.merge_cells(f"B{R_DET-1}:N{R_DET-1}")
A(f"B{R_DET-1}", "DETAIL AUTO : rotations par route, type avion et mois (programme du mois + vols additionnels)",
  bold=True, size=9, color="FFFFFF", fill="808080", halign="left")
A(f"B{R_DET}", "Route / type", bold=True, size=8, fill=GREY, border=box)
for m, (lab, y, mm) in enumerate(MOIS):
    A(f"{mcol(m)}{R_DET}", f"={S1}$W${REF0+m}", bold=True, size=8, fill=GREY, wrap=True, border=box)
A(f"O{R_DET}", "Route", bold=True, size=8, fill=GREY, border=box)
A(f"P{R_DET}", "Type", bold=True, size=8, fill=GREY, border=box)
ADD_D2, ADD_R2 = f"{S1}{ADD_D}", f"{S1}{ADD_R}"
ADD_T2, ADD_N2 = f"{S1}{ADD_T}", f"{S1}{ADD_N}"
for i, (route, colr, white) in enumerate(ROUTES):
    for ti, t in enumerate(TYPES):
        r = det_row(i, t)
        A(f"B{r}", f"{route} - {t}", size=8, halign="left", border=box)
        A(f"O{r}", route, size=8, border=box)
        A(f"P{r}", t, size=8, border=box)
        for m, (lab, y, mm) in enumerate(MOIS):
            rel = m*STEP + NB_T*i + ti + 1
            cal = cal_row(m)
            A(f"{mcol(m)}{r}",
              f'=SUMPRODUCT({occ_row(m)},--(INDEX({S1}{BIG},{rel},0)<>""))'
              f'+SUMIFS({ADD_N2},{ADD_R2},"{route}",{ADD_T2},"{t}",'
              f'{ADD_D2},">="&$J${cal},{ADD_D2},"<="&$K${cal})',
              size=8, fmt="0", border=box)

for col, w in {"A": 2, "B": 26, "O": 14, "P": 10}.items():
    an.column_dimensions[col].width = w
for i in range(12):
    an.column_dimensions[gl(3+i)].width = 13
an.freeze_panes = "C3"
an.page_setup.orientation = "landscape"
an.page_setup.fitToWidth = 1
an.sheet_properties.pageSetUpPr.fitToPage = True
an.print_area = f"B1:O{R_TYP+len(MOIS)+1}"

wb.save(OUT)
print(f"ecrit -> {OUT}  (tableaux {BLK0}-{DATA_LAST}, vols add. {ROW_ADD0}-{ROW_ADD_END})")
