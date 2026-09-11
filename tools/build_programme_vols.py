# -*- coding: utf-8 -*-
"""Genere le classeur 'Programme de vols mensuel' (2 feuilles).

Feuille 1 "Programme de vols" : 12 tableaux de saisie (un par mois). Chaque route y occupe
3 lignes = les 3 types avion (77W / 787 / A320). Dans une case de jour, on saisit UN CARACTERE
PAR ROTATION, et ce caractere indique le decalage du retour de cette rotation :
   0 (ou x) = retour le jour meme     1 = retour a J+1     2 = J+2     3 = J+3
Exemples : "0" = 1 rotation A/R dans la journee ; "1" = 1 rotation qui rentre le lendemain ;
"01" = 2 rotations le meme jour, la premiere rentre le jour meme, la seconde le lendemain ;
"00" = 2 rotations rentrant toutes deux le jour meme. Aucune notion de sieges sur cette feuille.
Feuille 2 "Analyse capacite" : interpretation automatique avec les capacites (sieges/rotation).
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
MAX_ROT = 4          # rotations maximum par jour, par route et par type
OFFSETS = "0123"     # decalages de retour autorises (x = alias de 0)

ROUTES = [
    ("CDGRUN", "9DC3E6", False), ("CDGDZA", "2E75B6", True),  ("BKKRUN", "FFE699", False),
    ("DZARUN", "F4B183", False), ("MRURUN", "C6E0B4", False), ("NOSRUN", "FFD966", False),
    ("RUNTNR", "D9D2E9", False), ("JNBRUN", "A9D08E", False), ("CPTRUN", "8FAADC", False),
    ("DIERUN", "FFC7CE", False), ("RRGRUN", "B7DEE8", False), ("TMMRUN", "D5A6BD", False),
]
TYPES = ["77W", "787", "A320"]
SIEGES = {"77W": 438, "787": 262, "A320": 174}
SIEGES_NOTE = {"77W": "Valeur fournie (77W = 777-300ER).",
               "787": "Valeur fournie (787 Dreamliner).",
               "A320": "Valeur fournie (A320)."}
MOIS = [("avril 2027",2027,4),("mai 2027",2027,5),("juin 2027",2027,6),("juillet 2027",2027,7),
        ("aout 2027",2027,8),("septembre 2027",2027,9),("octobre 2027",2027,10),
        ("novembre 2027",2027,11),("decembre 2027",2027,12),("janvier 2028",2028,1),
        ("fevrier 2028",2028,2),("mars 2028",2028,3)]
JOURS = ["J1 - Lundi","J2 - Mardi","J3 - Mercredi","J4 - Jeudi","J5 - Vendredi","J6 - Samedi","J7 - Dimanche"]
_ = None

# Programme de base recopie dans les 12 tableaux : {route: {type: [codes J1..J7]}}
BASE = {
    "CDGRUN": {"77W": ["1","1","1","1","1","1","1"], "787": ["1",_,_,_,"1",_,_], "A320": ["01",_,_,_,_,_,_]},
    "CDGDZA": {"77W": [_,_,_,_,_,_,"1"],             "787": [_,"1",_,_,"1",_,_], "A320": [_]*7},
    "BKKRUN": {"77W": [_]*7,                         "787": [_,_,"1",_,_,"1",_], "A320": [_]*7},
    "DZARUN": {"77W": [_]*7,                         "787": [_,_,"0",_,_,_,_],   "A320": ["0","0","0","0","0",_,_]},
    "MRURUN": {"77W": [_]*7,                         "787": [_]*7,               "A320": ["0",_,"0",_,"0",_,"0"]},
    "NOSRUN": {"77W": [_]*7,                         "787": [_]*7,               "A320": ["0",_,_,"0",_,"0",_]},
    "RUNTNR": {"77W": [_]*7,                         "787": [_]*7,               "A320": [_,"0",_,"0",_,"0",_]},
    "JNBRUN": {"77W": [_]*7,                         "787": ["0",_,_,_,"0",_,_], "A320": [_]*7},
    "CPTRUN": {"77W": [_]*7,                         "787": [_,_,_,"0",_,_,_],   "A320": [_]*7},
    "DIERUN": {"77W": [_]*7,                         "787": [_]*7,               "A320": [_,"0",_,_,_,"0",_]},
    "RRGRUN": {"77W": [_]*7,                         "787": [_]*7,               "A320": ["0","0","0","00","0","0",_]},
    "TMMRUN": {"77W": [_]*7,                         "787": [_]*7,               "A320": [_,_,"0",_,_,_,"0"]},
}
SAISON = {
    ("juillet 2027", "NOSRUN"):  {"77W": [_]*7, "787": [_]*7, "A320": ["0",_,"0","0",_,"0",_]},
    ("decembre 2027", "NOSRUN"): {"77W": [_]*7, "787": [_]*7, "A320": ["0","0",_,"0",_,"0","0"]},
    ("aout 2027", "CDGRUN"):     {"77W": ["1"]*7, "787": ["1"]*7, "A320": [_]*7},
}
VOLS_ADD = []   # livre vide : toute ligne ajoutee ici compte dans les totaux du mois

NB_WEEKS, GRID_C0 = 6, 3
NB_T = len(TYPES)
NB_ROWS_BLK = NB_T*len(ROUTES)                    # 36 lignes de donnees par tableau mensuel
STEP = NB_ROWS_BLK + 3                            # titre + entete + donnees + ligne vide
# --- feuille 1 "Programme de vols" (pilotage et lecture, aucune saisie de programme)
ROW_HDR, ROW_W1 = 7, 8
ROW_OCC = ROW_W1 + 2*NB_WEEKS                     # 20
ROW_NBW = ROW_OCC + 1                             # 21
ROW_REC_TITLE, ROW_REC_HDR, ROW_REC0 = 23, 24, 25
ROW_REC_TOT = ROW_REC0 + len(ROUTES)              # 37
R_REP0 = 4                                        # reperes : lignes 4..13
R_HELP_TITLE, R_HELP_HDR, R_HELP_TXT, R_HELP_TXT_R = 15, 16, 17, 18
R_CODE0, R_DEP0, R_RET0 = 19, 22, 25
R_MAT_TITLE, R_MAT_HDR, R_MAT0 = 29, 30, 31
HELP_C0, REF0 = 15, 4
# --- feuille 2 "Saisie programme"
S_MIRROR, S_NAV, BLK0 = 6, 8, 11                  # mois filtre, sommaire, 1er titre de tableau
DATA0 = BLK0 + 2
DATA_LAST = BLK0 + (len(MOIS)-1)*STEP + 1 + NB_ROWS_BLK
SAI_C0 = 4                                        # colonne D = J1
# --- feuille 3 "Vols additionnels"
ROW_ADD_HDR, ROW_ADD0, NB_ADD_ROWS = 6, 7, 60
ROW_ADD_END = ROW_ADD0 + NB_ADD_ROWS - 1

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

BIG = f"$D${DATA0}:$J${DATA_LAST}"                              # les 12 tableaux (7 jours)
MAT = f"$O${R_MAT0}:$U${R_MAT0+NB_ROWS_BLK-1}"                  # programme du mois filtre (auto)
ADD_D, ADD_R = f"$B${ROW_ADD0}:$B${ROW_ADD_END}", f"$C${ROW_ADD0}:$C${ROW_ADD_END}"
ADD_T, ADD_N = f"$D${ROW_ADD0}:$D${ROW_ADD_END}", f"$E${ROW_ADD0}:$E${ROW_ADD_END}"
ADD_K = f"$H${ROW_ADD0}:$H${ROW_ADD_END}"
OCC = f"$C${ROW_OCC}:$I${ROW_OCC}"
hc = lambda j: gl(HELP_C0+j)
blk_title = lambda k: BLK0 + k*STEP
mat_row = lambda i, t: R_MAT0 + NB_T*i + TYPES.index(t)

S1N, S2N, S3N, S4N = "Programme de vols", "Saisie programme", "Vols additionnels", "Analyse capacite"
ws.title = S1N
sa = wb.create_sheet(S2N)
va = wb.create_sheet(S3N)
an = wb.create_sheet(S4N)
for sh in (sa, va, an):
    sh.sheet_view.showGridLines = False
S1R, S2R, S3R = f"'{S1N}'!", f"'{S2N}'!", f"'{S3N}'!"
SA = lambda *a, **k: C(sa, *a, **k)
VA = lambda *a, **k: C(va, *a, **k)
A = lambda *a, **k: C(an, *a, **k)

BIG = f"{S2R}$D${DATA0}:$J${DATA_LAST}"
MAT = f"$O${R_MAT0}:$U${R_MAT0+NB_ROWS_BLK-1}"
ADD_D, ADD_R = f"{S3R}$B${ROW_ADD0}:$B${ROW_ADD_END}", f"{S3R}$C${ROW_ADD0}:$C${ROW_ADD_END}"
ADD_T, ADD_N = f"{S3R}$D${ROW_ADD0}:$D${ROW_ADD_END}", f"{S3R}$E${ROW_ADD0}:$E${ROW_ADD_END}"
ADD_K = f"{S3R}$H${ROW_ADD0}:$H${ROW_ADD_END}"
OCC = f"$C${ROW_OCC}:$I${ROW_OCC}"
hc = lambda j: gl(HELP_C0+j)
blk_title = lambda k: BLK0 + k*STEP
mat_row = lambda i, t: R_MAT0 + NB_T*i + TYPES.index(t)
CODE_AIDE = ("0 (ou x) = retour le jour meme  |  1 = retour a J+1  |  2 = J+2  |  3 = J+3.  "
             "Un caractere par rotation : \"0\" = 1 rotation A/R dans la journee, \"1\" = 1 rotation qui rentre le lendemain, "
             "\"01\" = 2 rotations dont une rentre le jour meme et l'autre le lendemain, \"00\" = 2 rotations rentrant le jour meme.")

# ================================================================= FEUILLE 1 : PROGRAMME DE VOLS
ws.merge_cells("B1:I1")
cell("B1", "PROGRAMME DE VOLS MENSUEL - ANNEE D'EXPLOITATION AVRIL 2027 / MARS 2028",
     bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
ws.row_dimensions[1].height = 30
ws.merge_cells("B2:I2")
cell("B2", "Feuille de lecture : les deux filtres ci-dessous pilotent la grille ; le recapitulatif ne suit que le filtre Mois. "
           "La saisie se fait sur les feuilles \"Saisie programme\" (12 tableaux, un par mois) et \"Vols additionnels\". "
           "La feuille \"Analyse capacite\" traduit le programme en sieges offerts.",
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
cell("E4", "<-- listes deroulantes. Le filtre Mois designe le tableau de saisie lu.", size=9, italic=True, halign="left", color="808080")
for ref, lab, dest in (("E5", ">> Aller a la saisie du programme", f"'{S2N}'!A1"),
                       ("G5", ">> Aller aux vols additionnels", f"'{S3N}'!A1")):
    c = cell(ref, lab, size=9, bold=True, color="0563C1", halign="left")
    c.hyperlink = Hyperlink(ref=ref, location=dest, tooltip=lab)
ws["C4"].comment = Comment("Liste deroulante : avril 2027 -> mars 2028.", "Modele")
ws["C5"].comment = Comment("Liste deroulante : les 12 routes exploitees.", "Modele")

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
           ("        dont 787", f"={rec_col('K')}", "0", False),
           ("        dont A320", f"={rec_col('L')}", "0", False)]
for i, (lab, fml, fmt, strong) in enumerate(reperes):
    r = R_REP0 + i
    cell(f"N{r}", lab, size=9, bold=strong, halign="left", fill="FFF2CC" if strong else GREY, border=box)
    cell(f"O{r}", fml, size=9, bold=True, fmt=fmt, fill="FFF2CC" if strong else None, border=box)
MIDX = "$O$7"

ws.merge_cells(f"N{R_MAT_TITLE}:{hc(6)}{R_MAT_TITLE}")
cell(f"N{R_MAT_TITLE}", "PROGRAMME DU MOIS FILTRE, TOUTES ROUTES (auto - recopie de la feuille Saisie programme)",
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

ws.merge_cells(f"N{R_HELP_TITLE}:{hc(6)}{R_HELP_TITLE}")
cell(f"N{R_HELP_TITLE}", "PROGRAMME DE LA ROUTE FILTREE (auto - alimente la grille)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"N{R_HELP_HDR}", "Jour", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_HELP_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
cell(f"N{R_HELP_TXT}", "Departs du jour (grille)", bold=True, size=8, halign="left", fill=GREY, border=box)
cell(f"N{R_HELP_TXT_R}", "Retours du jour (grille)", bold=True, size=8, halign="left", fill=GREY, border=box)
for ti, t in enumerate(TYPES):
    cell(f"N{R_CODE0+ti}", f"{t} : code saisi", size=8, halign="left", fill=GREY, border=box)
    cell(f"N{R_DEP0+ti}", f"{t} : rotations au depart", size=8, halign="left", fill=GREY, border=box)
    cell(f"N{R_RET0+ti}", f"{t} : retours ce jour", size=8, halign="left", fill=GREY, border=box)
RIDX = f'MATCH($C$5,$Z${REF0}:$Z${REF0+11},0)'
for j in range(7):
    c = hc(j)
    for ti, t in enumerate(TYPES):
        cell(f"{c}{R_CODE0+ti}", f'=IFERROR(INDEX({MAT},({RIDX}-1)*{NB_T}+{ti+1},{j+1})&"","")', size=8, border=box)
        cell(f"{c}{R_DEP0+ti}", f"=LEN({c}{R_CODE0+ti})", size=8, fmt="0", border=box)
        terms = []
        for k in range(1, len(OFFSETS)):
            src = f"{hc((j - k) % 7)}${R_CODE0+ti}"
            terms.append(f'LEN({src})-LEN(SUBSTITUTE({src},"{k}",""))')
        cell(f"{c}{R_RET0+ti}", "=" + "+".join(terms), size=8, fmt="0", border=box)
    dep = '&" "&'.join(f'IF({c}${R_DEP0+ti}=0,"",IF({c}${R_DEP0+ti}=1,"{t}",{c}${R_DEP0+ti}&"x{t}"))'
                       for ti, t in enumerate(TYPES))
    ret = '&" "&'.join(f'IF({c}${R_RET0+ti}=0,"",IF({c}${R_RET0+ti}=1,"{t}",{c}${R_RET0+ti}&"x{t}"))'
                       for ti, t in enumerate(TYPES))
    cell(f"{c}{R_HELP_TXT}", f'=SUBSTITUTE(TRIM({dep})," ","/")', bold=True, size=9, border=box)
    cell(f"{c}{R_HELP_TXT_R}", f'=SUBSTITUTE(TRIM({ret})," ","/")', size=9, border=box)

ws.merge_cells(f"B{ROW_HDR-1}:I{ROW_HDR-1}")
cell(f"B{ROW_HDR-1}", "GRILLE MENSUELLE  -  rotations par semaine et par jour (filtres Mois + Route)",
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
    ws.row_dimensions[rt].height = 30
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
        AL = f'TRIM(${hc(j)}${R_HELP_TXT}&IF({nrot}=0,""," (+"&{lbl}&")"))'
        RE = f"${hc(j)}${R_HELP_TXT_R}"
        cell(f"{col}{rt}", f'=IF({d}="","",{AL}&IF({RE}="","",IF({AL}="","",CHAR(10))&"< "&{RE}))',
             bold=True, size=10, wrap=True, border=box)
cell(f"B{ROW_OCC}", "Occurrences du jour dans le mois", bold=True, size=9, halign="left", fill="E7E6E6", border=box)
for j in range(7):
    col = gl(GRID_C0+j)
    cell(f"{col}{ROW_OCC}", "=COUNT(" + ",".join(f"{col}{r}" for r in date_rows) + ")",
         bold=True, size=10, fill="E7E6E6", border=box)
cell(f"B{ROW_NBW}", "Nb de semaines calendaires du mois", bold=True, size=9, halign="left", fill="E7E6E6", border=box)
cell(f"C{ROW_NBW}", "=" + "+".join(f'IF(COUNT({gl(GRID_C0)}{r}:{gl(GRID_C0+6)}{r})>0,1,0)' for r in date_rows),
     bold=True, size=10, fill="E7E6E6", border=box)
ws.merge_cells(f"D{ROW_NBW}:I{ROW_NBW}")
cell(f"D{ROW_NBW}", 'Lecture : 1re ligne = les rotations qui PARTENT ce jour-la ("2xA320" = deux rotations A320)  |  2e ligne '
                    '"< 787" = les RETOURS de rotations parties les jours precedents  |  "(+A320)" encadre en rouge = vol saisi '
                    'sur la feuille Vols additionnels.',
     size=8, italic=True, color="595959", halign="left")

ws.merge_cells(f"B{ROW_REC_TITLE}:L{ROW_REC_TITLE}")
cell(f"B{ROW_REC_TITLE}", "RECAPITULATIF MENSUEL - TOUTES ROUTES  (depend uniquement du filtre Mois)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
for i, h in enumerate(["Route", "Type(s) avion", "Semaine type du mois (jours J1-J7)", "Rotations/semaine",
                       "Nb de semaines/occurrences dans le mois", "Total rotations du mois",
                       "dont programme du mois", "dont vols additionnels", "Total 77W", "Total 787", "Total A320"]):
    cell(f"{gl(2+i)}{ROW_REC_HDR}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
ws.row_dimensions[ROW_REC_HDR].height = 44
for i in range(len(ROUTES)):
    r = ROW_REC0 + i
    rows = {t: mat_row(i, t) for t in TYPES}
    add = lambda extra="": f'SUMIFS({ADD_N},{ADD_R},$B{r}{extra},{ADD_D},">="&$O$4,{ADD_D},"<="&$O$5)'
    nsem = {t: f'SUMPRODUCT(LEN($O${rows[t]}:$U${rows[t]}))' for t in TYPES}
    prod = {t: f'SUMPRODUCT({OCC},LEN($O${rows[t]}:$U${rows[t]}))' for t in TYPES}
    cell(f"B{r}", f"=$Z${REF0+i}", bold=True, size=10, border=box)
    tl = "&".join(f'IF({nsem[t]}>0,"{t}("&{nsem[t]}&") ","")' for t in TYPES)
    cell(f"C{r}", f'=IF(E{r}=0,"-",SUBSTITUTE(TRIM({tl})," ",", "))', size=9, wrap=True, border=box)
    jl = "&".join(f'IF(SUMPRODUCT(LEN(${hc(j)}${rows[TYPES[0]]}:${hc(j)}${rows[TYPES[-1]]}))>0,"J{j+1} ","")'
                  for j in range(7))
    cell(f"D{r}", f'=IF(E{r}=0,"aucune operation",SUBSTITUTE(TRIM({jl})," ",", "))', size=9, wrap=True, border=box)
    cell(f"E{r}", "=" + "+".join(nsem[t] for t in TYPES), size=10, fmt="0", border=box)
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
cell(f"B{r}", "Methode : le recapitulatif lit le tableau du mois filtre sur la feuille \"Saisie programme\". Chaque caractere "
              "saisi dans une case = une rotation (un aller-retour), comptee le jour de son depart quel que soit le decalage de "
              "son retour. Pour chaque jour opere, le nombre d'occurrences de ce jour dans le mois est compte a partir des dates "
              "reelles de la grille (semaines partielles incluses), chaque type avion etant compte separement. Total rotations du "
              "mois = programme du mois + vols additionnels dates du mois.",
     size=8, italic=True, color="595959", halign="left", wrap=True)

for col, h in (("W", "Mois (filtre)"), ("X", "1er jour"), ("Y", "Types avion"), ("Z", "Routes")):
    cell(f"{col}{REF0-1}", h, bold=True, size=8, color="FFFFFF", fill="808080", wrap=True, border=box)
for i, (lab, y, m) in enumerate(MOIS):
    cell(f"W{REF0+i}", lab, size=8, halign="left", border=box)
    cell(f"X{REF0+i}", f"=DATE({y},{m},1)", size=8, fmt="DD/MM/YYYY", border=box)
for i, t in enumerate(TYPES):
    cell(f"Y{REF0+i}", t, size=8, border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    cell(f"Z{REF0+i}", route, size=8, border=box)

def add_dv(sheet, dv, ranges):
    sheet.add_data_validation(dv)
    for rng in (ranges if isinstance(ranges, (list, tuple)) else [ranges]):
        dv.add(rng)
add_dv(ws, DataValidation(type="list", formula1=f"=$W${REF0}:$W${REF0+11}", allow_blank=False), "C4")
add_dv(ws, DataValidation(type="list", formula1=f"=$Z${REF0}:$Z${REF0+11}", allow_blank=False), "C5")

grid_ranges = " ".join(f"{gl(GRID_C0)}{r}:{gl(GRID_C0+6)}{r}" for r in type_rows)
c0 = gl(GRID_C0)
ws.conditional_formatting.add(
    grid_ranges,
    FormulaRule(formula=[f'ISNUMBER(SEARCH("(+",{c0}{ROW_W1}))'],
                border=Border(left=Side(style="medium", color=RED), right=Side(style="medium", color=RED),
                              top=Side(style="medium", color=RED), bottom=Side(style="medium", color=RED)),
                stopIfTrue=False))
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
widths = {"A": 2, "B": 26, "M": 3, "N": 26, "V": 3, "W": 18, "X": 12, "Y": 10, "Z": 10}
for j in range(7):
    widths[gl(GRID_C0+j)] = 16
for j in range(7):
    widths[gl(HELP_C0+j)] = 10
widths.update({"J": 14, "K": 14, "L": 14})
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = "C8"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_area = f"B1:L{ROW_REC_TOT+2}"

# ================================================================= FEUILLE 2 : SAISIE PROGRAMME
sa.merge_cells("B1:L1")
SA("B1", "SAISIE DU PROGRAMME - 12 TABLEAUX, UN PAR MOIS", bold=True, size=16, color="FFFFFF", fill=BROWN, halign="left")
sa.row_dimensions[1].height = 30
sa.merge_cells("B2:L4")
SA("B2", "Cellules en BLEU = saisie utilisateur. Chaque route occupe 3 lignes = les 3 types avion. Dans la case d'un jour, on "
         "saisit UN CARACTERE PAR ROTATION (aller-retour) qui PART ce jour-la, et ce caractere indique QUAND CETTE ROTATION "
         "RENTRE :  " + CODE_AIDE + "  Case vide = aucun vol. Une rotation : fond vert ; plusieurs rotations : fond orange. "
         "Les 12 tableaux sont independants, la frequence peut donc etre differente chaque mois. Pour reprendre un mois sur un "
         "autre : copier la zone des 7 jours (36 lignes) et la coller dans le tableau cible.",
   size=9, italic=True, color=BROWN, halign="left", wrap=True)
sa.row_dimensions[2].height = 56
SA(f"B{S_MIRROR}", "Mois filtre sur la feuille de lecture :", bold=True, size=10, halign="left", fill=GREY, border=box)
sa.merge_cells(f"C{S_MIRROR}:D{S_MIRROR}")
SA(f"C{S_MIRROR}", f"={S1R}$C$4", bold=True, size=11, fill="FFF2CC", border=box)
c = SA(f"E{S_MIRROR}", "<< retour a la feuille Programme de vols", size=9, bold=True, color="0563C1", halign="left")
c.hyperlink = Hyperlink(ref=f"E{S_MIRROR}", location=f"'{S1N}'!B1", tooltip="Retour a la feuille de lecture")
SA(f"B{S_NAV}", "Aller au mois :", bold=True, size=9, halign="left", fill=GREY, border=box)
SA(f"B{S_NAV+1}", "(clic sur un mois)", size=8, italic=True, halign="left", color="808080", fill=GREY, border=box)
for k, (lab, y, m) in enumerate(MOIS):
    ref = f"{gl(3 + k % 6)}{S_NAV + k // 6}"
    c = SA(ref, lab, size=9, bold=True, color="0563C1", fill=GREY, border=box)
    c.hyperlink = Hyperlink(ref=ref, location=f"'{S2N}'!B{blk_title(k)}", tooltip=f"Aller au tableau {lab}")

for k, (lab, y, m) in enumerate(MOIS):
    t0 = blk_title(k)
    sa.merge_cells(f"B{t0}:J{t0}")
    SA(f"B{t0}", f'="TABLEAU {k+1}/12  -  PROGRAMME DE {lab.upper()}"&IF($C${S_MIRROR}="{lab}","          <<< MOIS ACTUELLEMENT FILTRE","")',
       bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
    sa.row_dimensions[t0].height = 20
    c = SA(f"L{t0}", "^ Haut de la feuille", size=8, italic=True, color="0563C1", halign="right")
    c.hyperlink = Hyperlink(ref=f"L{t0}", location=f"'{S2N}'!B1", tooltip="Haut de la feuille")
    for i, h in enumerate(["Route", "Type avion"] + JOURS + ["Rot./sem. par type", "Rot./sem. route"]):
        SA(f"{gl(2+i)}{t0+1}", h, bold=True, size=9, color="FFFFFF", fill="A6774B", wrap=True, border=box)
    sa.row_dimensions[t0+1].height = 26
    for i, (route, colr, white) in enumerate(ROUTES):
        p = t0 + 2 + NB_T*i
        prog = SAISON.get((lab, route), BASE[route])
        sa.merge_cells(f"B{p}:B{p+NB_T-1}")
        SA(f"B{p}", route, bold=True, size=11, halign="left",
           fill=colr, color="FFFFFF" if white else "000000", border=box)
        sa.merge_cells(f"L{p}:L{p+NB_T-1}")
        SA(f"L{p}", f"=SUM(K{p}:K{p+NB_T-1})", bold=True, size=11, fmt="0", border=box)
        for ti, t in enumerate(TYPES):
            r = p + ti
            SA(f"C{r}", t, bold=True, size=10, fill="F2E4D8", border=box)
            for j in range(7):
                SA(f"{gl(SAI_C0+j)}{r}", prog[t][j], bold=True, size=11, color=BLUE_IN, fmt="@", border=box)
            SA(f"K{r}", f"=SUMPRODUCT(LEN(D{r}:J{r}))", size=9, fmt="0", border=box)
        for sl in range(1, NB_T):
            SA(f"B{p+sl}", fill=colr, border=box)
            SA(f"L{p+sl}", border=box)
sa[f"D{DATA0}"].comment = Comment(
    "Un caractere par rotation partant ce jour-la.\n"
    "0 (ou x) = retour le jour meme, 1 = retour J+1, 2 = J+2, 3 = J+3.\n"
    "Ex. : 01 = deux rotations, une rentre le jour meme, l'autre le lendemain.\n"
    "Case vide = aucun vol.", "Modele")

anchor = f"D{DATA0}"
strip = f'SUBSTITUTE(UPPER({anchor}),"X","0")'
for ch in OFFSETS:
    strip = f'SUBSTITUTE({strip},"{ch}","")'
add_dv(sa, DataValidation(type="custom", formula1=f"=AND(LEN({anchor})<={MAX_ROT},LEN({strip})=0)",
                          allow_blank=True, errorTitle="Code de rotation invalide",
                          error="Un caractere par rotation : 0 (ou x) = retour le jour meme, 1 = J+1, 2 = J+2, 3 = J+3.\n"
                                f"Exemples : 0 | 1 | 01 | 00. Maximum {MAX_ROT} rotations par jour et par type."),
       [f"D{blk_title(k)+2}:J{blk_title(k)+1+NB_ROWS_BLK}" for k in range(len(MOIS))])

chk_ranges = " ".join(f"D{blk_title(k)+2}:J{blk_title(k)+1+NB_ROWS_BLK}" for k in range(len(MOIS)))
sa.conditional_formatting.add(
    chk_ranges, FormulaRule(formula=[f'LEN(D{DATA0})>1'], fill=PatternFill("solid", fgColor="ED7D31"),
                            font=Font(name=F, size=11, bold=True, color="FFFFFF"), stopIfTrue=True))
sa.conditional_formatting.add(
    chk_ranges, FormulaRule(formula=[f'D{DATA0}<>""'], fill=PatternFill("solid", fgColor=CHECK),
                            font=Font(name=F, size=11, bold=True, color="FFFFFF"), stopIfTrue=True))
for k, (lab, y, m) in enumerate(MOIS):
    t0 = blk_title(k)
    sa.conditional_formatting.add(
        f"B{t0}:L{t0}",
        FormulaRule(formula=[f'$C${S_MIRROR}="{lab}"'], fill=PatternFill("solid", fgColor=ACTIVE),
                    font=Font(name=F, size=11, bold=True, color="FFFFFF"), stopIfTrue=True))
for col, w in {"A": 2, "B": 16, "C": 12, "K": 14, "L": 14}.items():
    sa.column_dimensions[col].width = w
for j in range(7):
    sa.column_dimensions[gl(SAI_C0+j)].width = 13
sa.freeze_panes = "D10"
sa.page_setup.orientation = "landscape"
sa.page_setup.fitToWidth = 1
sa.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================= FEUILLE 3 : VOLS ADDITIONNELS
va.merge_cells("B1:H1")
VA("B1", "VOLS ADDITIONNELS - en plus du programme mensuel", bold=True, size=16, color="FFFFFF", fill=BROWN, halign="left")
va.row_dimensions[1].height = 30
va.merge_cells("B2:H3")
VA("B2", "Une ligne = un vol supplementaire date (charter, renfort ponctuel, fret, seconde rotation exceptionnelle...). Il "
         "s'ajoute au programme du mois concerne, sur n'importe quel jour, y compris un jour deja opere ou un jour hors "
         "programme. Sur la feuille \"Programme de vols\", il apparait dans la cellule du jour entre parentheses avec un + et une "
         "bordure rouge, et il est compte dans la colonne \"dont vols additionnels\" du recapitulatif. Pour une rotation "
         "recurrente sur tout un mois, ajouter plutot un caractere dans la case du jour, sur la feuille \"Saisie programme\". "
         "Cette table est livree VIDE : tout ce qui y est saisi est compte.",
   size=9, italic=True, color=BROWN, halign="left", wrap=True)
va.row_dimensions[2].height = 42
c = VA("B4", "<< retour a la feuille Programme de vols", size=9, bold=True, color="0563C1", halign="left")
c.hyperlink = Hyperlink(ref="B4", location=f"'{S1N}'!B1", tooltip="Retour a la feuille de lecture")
for i, h in enumerate(["Date du vol", "Route", "Type avion", "Nb de rotations", "Commentaire (libre)"]):
    VA(f"{gl(2+i)}{ROW_ADD_HDR}", h, bold=True, size=9, color="FFFFFF", fill=BROWN, wrap=True, border=box)
VA(f"H{ROW_ADD_HDR}", "Cle technique (auto)", bold=True, size=8, color="FFFFFF", fill="A6A6A6", wrap=True, border=box)
for i in range(NB_ADD_ROWS):
    r = ROW_ADD0 + i
    VA(f"B{r}", None, size=10, color=BLUE_IN, fmt="DD/MM/YYYY", border=box)
    VA(f"C{r}", None, size=10, color=BLUE_IN, border=box)
    VA(f"D{r}", None, size=10, color=BLUE_IN, border=box)
    VA(f"E{r}", None, size=10, color=BLUE_IN, fmt="0", border=box)
    VA(f"F{r}", None, size=9, italic=True, color="595959", halign="left", border=box)
    VA(f"H{r}", f'=IF(OR($B{r}="",$C{r}=""),"",$B{r}&"|"&$C{r})', size=8, color="A6A6A6", border=box)
va[f"B{ROW_ADD0}"].comment = Comment(
    "Date du vol (annee avril 2027 - mars 2028).\n"
    "Route et type avion : listes deroulantes.\n"
    "Nb de rotations : 1 par defaut, 2 pour une double rotation le meme jour.", "Modele")
VA(f"J{ROW_ADD_HDR}", "Routes", bold=True, size=8, color="FFFFFF", fill="808080", border=box)
VA(f"K{ROW_ADD_HDR}", "Types", bold=True, size=8, color="FFFFFF", fill="808080", border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    VA(f"J{ROW_ADD0+i}", route, size=8, border=box)
for i, t in enumerate(TYPES):
    VA(f"K{ROW_ADD0+i}", t, size=8, border=box)
VA(f"J{ROW_ADD0+len(ROUTES)+1}", "listes de reference - ne pas modifier", size=8, italic=True, color="808080", halign="left")
add_dv(va, DataValidation(type="list", formula1=f"=$J${ROW_ADD0}:$J${ROW_ADD0+len(ROUTES)-1}", allow_blank=True,
                          errorTitle="Route inconnue", error="Choisir une route dans la liste."),
       f"C{ROW_ADD0}:C{ROW_ADD_END}")
add_dv(va, DataValidation(type="list", formula1=f"=$K${ROW_ADD0}:$K${ROW_ADD0+len(TYPES)-1}", allow_blank=True,
                          errorTitle="Type avion inconnu", error="Choisir un type avion dans la liste."),
       f"D{ROW_ADD0}:D{ROW_ADD_END}")
add_dv(va, DataValidation(type="whole", operator="between", formula1="1", formula2="20", allow_blank=True,
                          error="Nombre de rotations : entier entre 1 et 20.", errorTitle="Valeur invalide"),
       f"E{ROW_ADD0}:E{ROW_ADD_END}")
add_dv(va, DataValidation(type="date", operator="between", formula1="DATE(2027,4,1)", formula2="DATE(2028,3,31)",
                          allow_blank=True, errorStyle="warning", errorTitle="Date hors annee d'exploitation",
                          error="Date hors de l'annee avril 2027 - mars 2028 : le vol ne sera visible dans aucun mois."),
       f"B{ROW_ADD0}:B{ROW_ADD_END}")
for col, w in {"A": 2, "B": 16, "C": 14, "D": 14, "E": 14, "F": 60, "G": 3, "H": 18, "I": 3, "J": 12, "K": 10}.items():
    va.column_dimensions[col].width = w
va.freeze_panes = f"B{ROW_ADD0}"
va.page_setup.orientation = "landscape"
va.page_setup.fitToWidth = 1
va.sheet_properties.pageSetUpPr.fitToPage = True
va.print_area = f"B1:F{ROW_ADD_END}"

# ================================================================= FEUILLE 4 : ANALYSE CAPACITE
R_HYP, R_SYN, R_SIE, R_ROT, R_TYP, R_CAL, R_DET = 5, 12, 28, 45, 62, 78, 93
mcol = lambda m: gl(3+m)
rcol = lambda i: gl(3+i)
det_row = lambda i, t: R_DET + 1 + NB_T*i + TYPES.index(t)
cal_row = lambda m: R_CAL + 1 + m
occ_row = lambda m: f"$C${cal_row(m)}:$I${cal_row(m)}"
seat = {t: f"$C${R_HYP+1+ti}" for ti, t in enumerate(TYPES)}
LEGS = f"$C${R_HYP+4}"

an.merge_cells("B1:J1")
A("B1", "ANALYSE DE CAPACITE - SIEGES OFFERTS (avril 2027 / mars 2028)",
  bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
an.row_dimensions[1].height = 30
an.merge_cells("B2:J2")
A("B2", "Cette feuille interprete automatiquement les feuilles \"Saisie programme\" et \"Vols additionnels\" en les croisant avec "
        "les capacites ci-dessous. Une rotation = un aller-retour, comptee le jour de son depart. Aucune saisie ici, hormis les "
        "hypotheses. Toute modification du programme met a jour cette feuille.",
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
    an.merge_cells(f"D{r}:F{r}")
rl = R_HYP + 4
A(f"B{rl}", "Legs comptes par rotation", bold=True, size=10, border=box)
A(f"C{rl}", 1, bold=True, size=10, color=BLUE_IN, fill="FFF2CC", fmt="0", border=box)
an.merge_cells(f"D{rl}:F{rl}")
A(f"D{rl}", "1 = l'aller-retour compte pour un vol (sieges offerts d'un sens) ; mettre 2 pour compter l'aller ET le retour.",
  size=8, italic=True, color="595959", halign="left", border=box)

an.merge_cells(f"B{R_SYN-1}:F{R_SYN-1}")
A(f"B{R_SYN-1}", "SYNTHESE ANNUELLE PAR ROUTE", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
for i, h in enumerate(["Route", "Rotations sur l'annee", "Sieges offerts sur l'annee", "Sieges / rotation",
                       "Part du reseau (sieges)"]):
    A(f"{gl(2+i)}{R_SYN}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
an.row_dimensions[R_SYN].height = 30
for i, (route, colr, white) in enumerate(ROUTES):
    r = R_SYN + 1 + i
    col = rcol(i)
    A(f"B{r}", route, bold=True, size=10, fill=colr, color="FFFFFF" if white else "000000", border=box)
    A(f"C{r}", f"=SUM({col}${R_ROT+1}:{col}${R_ROT+len(MOIS)})", size=10, fmt="#,##0", border=box)
    A(f"D{r}", f"=SUM({col}${R_SIE+1}:{col}${R_SIE+len(MOIS)})", bold=True, size=10, fmt="#,##0", border=box)
    A(f"E{r}", f'=IFERROR(D{r}/C{r},"-")', size=10, fmt="#,##0", border=box)
    A(f"F{r}", f"=IFERROR(D{r}/$D${R_SYN+len(ROUTES)+1},0)", size=10, fmt="0.0%", border=box)
rt2 = R_SYN + len(ROUTES) + 1
A(f"B{rt2}", "TOTAL RESEAU", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
for c in ("C", "D"):
    A(f"{c}{rt2}", f"=SUM({c}{R_SYN+1}:{c}{rt2-1})", bold=True, size=11, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)
A(f"E{rt2}", f'=IFERROR(D{rt2}/C{rt2},"-")', bold=True, size=10, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)
A(f"F{rt2}", f"=IFERROR(D{rt2}/$D${rt2},0)", bold=True, size=10, color="FFFFFF", fill=NAVY, fmt="0.0%", border=box)

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
        A(f"B{r}", f"={S1R}$W${REF0+m}", size=9, halign="left", border=box)
        for i, (route, colr, white) in enumerate(ROUTES):
            d = mcol(m)
            if base == R_ROT:
                f = f"=SUM({d}{det_row(i,TYPES[0])}:{d}{det_row(i,TYPES[-1])})"
            else:
                f = "=" + "+".join(f"{seat[t]}*{LEGS}*{d}{det_row(i,t)}" for t in TYPES)
            A(f"{rcol(i)}{r}", f, size=9, fmt=fmt, border=box)
        A(f"O{r}", f"=SUM(C{r}:N{r})", bold=True, size=9, fmt=fmt, border=box)
    r = base + 1 + len(MOIS)
    A(f"B{r}", "TOTAL ANNEE", bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left", border=box)
    for i in range(len(ROUTES) + 1):
        c = rcol(i) if i < len(ROUTES) else "O"
        A(f"{c}{r}", f"=SUM({c}{base+1}:{c}{base+len(MOIS)})", bold=True, size=9,
          color="FFFFFF", fill=NAVY, fmt=fmt, border=box)

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
    A(f"B{r}", f"={S1R}$W${REF0+m}", size=9, halign="left", border=box)
    for ti, t in enumerate(TYPES):
        A(f"{gl(3+ti)}{r}", f'=SUMIF({DET_T},"{t}",${d}${R_DET+1}:${d}${R_DET+NB_ROWS_BLK})', size=9, fmt="#,##0", border=box)
    A(f"F{r}", f"=SUM(C{r}:E{r})", bold=True, size=9, fmt="#,##0", border=box)
    for ti, t in enumerate(TYPES):
        A(f"{gl(7+ti)}{r}", f"={gl(3+ti)}{r}*{seat[t]}*{LEGS}", size=9, fmt="#,##0", border=box)
    A(f"J{r}", f"=SUM(G{r}:I{r})", bold=True, size=9, fmt="#,##0", border=box)
    A(f"K{r}", f'=IFERROR(J{r}/F{r},"-")', size=9, fmt="#,##0", border=box)
r = R_TYP + 1 + len(MOIS)
A(f"B{r}", "TOTAL ANNEE", bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left", border=box)
for i in range(2, 10):
    c = gl(i+1)
    A(f"{c}{r}", f"=SUM({c}{R_TYP+1}:{c}{r-1})", bold=True, size=9, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)
A(f"K{r}", f'=IFERROR(J{r}/F{r},"-")', bold=True, size=9, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)

an.merge_cells(f"B{R_CAL-1}:L{R_CAL-1}")
A(f"B{R_CAL-1}", "CALENDRIER AUTO : occurrences de chaque jour de semaine dans le mois (semaines partielles incluses)",
  bold=True, size=9, color="FFFFFF", fill="808080", halign="left")
for i, h in enumerate(["Mois", "J1", "J2", "J3", "J4", "J5", "J6", "J7", "1er jour", "Dernier jour", "Nb de jours"]):
    A(f"{gl(2+i)}{R_CAL}", h, bold=True, size=8, fill=GREY, border=box)
for m, (lab, y, mm) in enumerate(MOIS):
    r = cal_row(m)
    A(f"B{r}", f"={S1R}$W${REF0+m}", size=8, halign="left", border=box)
    A(f"J{r}", f"={S1R}$X${REF0+m}", size=8, fmt="DD/MM/YYYY", border=box)
    A(f"K{r}", f"=EOMONTH($J{r},0)", size=8, fmt="DD/MM/YYYY", border=box)
    A(f"L{r}", f"=$K{r}-$J{r}+1", size=8, fmt="0", border=box)
    for j in range(7):
        off = f"MOD({j+1}-WEEKDAY($J{r},2),7)"
        A(f"{gl(3+j)}{r}", f"=IF({off}>=$L{r},0,INT(($L{r}-1-{off})/7)+1)", size=8, fmt="0", border=box)

an.merge_cells(f"B{R_DET-1}:N{R_DET-1}")
A(f"B{R_DET-1}", "DETAIL AUTO : rotations par route, type avion et mois (programme du mois + vols additionnels)",
  bold=True, size=9, color="FFFFFF", fill="808080", halign="left")
A(f"B{R_DET}", "Route / type", bold=True, size=8, fill=GREY, border=box)
for m, (lab, y, mm) in enumerate(MOIS):
    A(f"{mcol(m)}{R_DET}", f"={S1R}$W${REF0+m}", bold=True, size=8, fill=GREY, wrap=True, border=box)
A(f"O{R_DET}", "Route", bold=True, size=8, fill=GREY, border=box)
A(f"P{R_DET}", "Type", bold=True, size=8, fill=GREY, border=box)
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
              f'=SUMPRODUCT({occ_row(m)},LEN(INDEX({BIG},{rel},0)))'
              f'+SUMIFS({ADD_N},{ADD_R},"{route}",{ADD_T},"{t}",'
              f'{ADD_D},">="&$J${cal},{ADD_D},"<="&$K${cal})',
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
print(f"ecrit -> {OUT}\n  Saisie programme : tableaux {BLK0}-{DATA_LAST}  |  Vols additionnels : {ROW_ADD0}-{ROW_ADD_END}")
