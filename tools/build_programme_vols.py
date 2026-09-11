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
CABINES = [("C", "Club"), ("W", "Confort"), ("Y", "Loisir")]
SIEGES_CAB = {"77W":  {"C": 14, "W": 40, "Y": 384},     # 438
              "787":  {"C": 18, "W":  0, "Y": 244},     # 262
              "A320": {"C":  0, "W": 12, "Y": 162}}     # 174
SIEGES = {t: sum(c.values()) for t, c in SIEGES_CAB.items()}
CAB_FILTRE = ["Toutes cabines"] + [f"{k} - {lab}" for k, lab in CABINES]
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
R_HELP_TITLE, R_HELP_HDR, R_DEPTXT = 15, 16, 17   # bloc route filtree (par jour de semaine)
R_CODE0, R_CODEP0, R_DEP0 = 18, 21, 24            # codes mois courant / mois precedent / nb departs
R_RETD_TITLE, R_RETD_HDR, R_RETD0 = 28, 29, 30    # retours par date : 3 types x 6 semaines
R_TXT_TITLE, R_TXT_HDR, R_DEPTXT0, R_RETTXT0 = 49, 50, 51, 57
R_MAT_TITLE, R_MAT_HDR, R_MAT0 = 64, 65, 66       # matrice du mois filtre (+ colonnes V et W)
HELP_C0, REF_C0, REF0 = 17, 28, 4                 # O..U pour les blocs auto, AA.. pour les listes
# --- feuille 2 "Saisie programme"
S_MIRROR, S_NAV, BLK0 = 6, 8, 11                  # mois filtre, sommaire, 1er titre de tableau
DATA0 = BLK0 + 2
DATA_LAST = BLK0 + (len(MOIS)-1)*STEP + 1 + NB_ROWS_BLK
SAI_C0 = 4                                        # colonne D = J1
# --- feuille 3 "Vols additionnels"
ROW_ADD_HDR, ROW_ADD0, NB_ADD_ROWS = 6, 7, 60
ROT_FMT = '0.0'                                   # les rotations peuvent valoir 0,5
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
GREEN = "00B050"      # bordure des ajouts
BLUE_IN = "0000FF"

BIG = f"$D${DATA0}:$J${DATA_LAST}"                              # les 12 tableaux (7 jours)
MAT = f"${gl(HELP_C0)}${R_MAT0}:${gl(HELP_C0+6)}${R_MAT0+NB_ROWS_BLK-1}"                  # programme du mois filtre (auto)
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
MAT = f"${gl(HELP_C0)}${R_MAT0}:${gl(HELP_C0+6)}${R_MAT0+NB_ROWS_BLK-1}"
ADD_D = f"{S3R}$B${ROW_ADD0}:$B${ROW_ADD_END}"     # date de depart
ADD_R = f"{S3R}$C${ROW_ADD0}:$C${ROW_ADD_END}"     # route
ADD_T = f"{S3R}$D${ROW_ADD0}:$D${ROW_ADD_END}"     # type avion
ADD_N = f"{S3R}$E${ROW_ADD0}:$E${ROW_ADD_END}"     # nb de rotations (+ ajout / - retrait)
ADD_RD = f"{S3R}$I${ROW_ADD0}:$I${ROW_ADD_END}"    # date de retour (auto)
ADD_KD = f"{S3R}$J${ROW_ADD0}:$J${ROW_ADD_END}"    # cle depart (auto)
ADD_RDA = f"{S3R}$K${ROW_ADD0}:$K${ROW_ADD_END}"   # date de retour affichee (auto, vide si J+0)
ADD_KR = f"{S3R}$L${ROW_ADD0}:$L${ROW_ADD_END}"    # cle retour (auto)
OCC = f"$C${ROW_OCC}:$I${ROW_OCC}"
hc = lambda j: gl(HELP_C0+j)
rc = lambda i: gl(REF_C0+i)
blk_title = lambda k: BLK0 + k*STEP
mat_row = lambda i, t: R_MAT0 + NB_T*i + TYPES.index(t)
rel_of = lambda i, ti: NB_T*i + ti + 1
retd_row = lambda ti, w: R_RETD0 + NB_WEEKS*ti + w
CODE_AIDE = ("0 (ou x) = retour le jour meme  |  1 = retour a J+1  |  2 = J+2  |  3 = J+3.  "
             "Un caractere par rotation : \"0\" = 1 rotation A/R dans la journee, \"1\" = 1 rotation qui rentre le lendemain, "
             "\"01\" = 2 rotations dont une rentre le jour meme et l'autre le lendemain, \"00\" = 2 rotations rentrant le jour meme.")

def cnt_k(cellref, k):
    """nombre de rotations de decalage k dans un code."""
    return f'LEN({cellref})-LEN(SUBSTITUTE({cellref},"{k}",""))'

def cnt_k0(cellref):
    return f'LEN({cellref})-LEN(SUBSTITUTE(SUBSTITUTE(UPPER({cellref}),"X","0"),"0",""))'

# ================================================================= FEUILLE 1 : PROGRAMME DE VOLS
ws.merge_cells("B1:I1")
cell("B1", "PROGRAMME DE VOLS MENSUEL - ANNEE D'EXPLOITATION AVRIL 2027 / MARS 2028",
     bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
ws.row_dimensions[1].height = 30
ws.merge_cells("B2:I2")
cell("B2", "Feuille de lecture : les deux filtres pilotent la grille ; le recapitulatif ne suit que le filtre Mois. La saisie se "
           "fait sur \"Saisie programme\" (12 tableaux) et \"Vols additionnels\" (ajouts et retraits dates). Une rotation = un "
           "aller-retour : elle compte 0,5 le mois de son depart et 0,5 le mois de son retour, donc un aller-retour a cheval sur "
           "deux mois est partage entre eux.",
     size=9, italic=True, halign="left", wrap=True, color="404040")
ws.row_dimensions[2].height = 32
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

ws.merge_cells("N3:O3")
cell("N3", "REPERES CALCULES", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left")
W_, X_ = f"${rc(0)}${REF0}:${rc(0)}${REF0+11}", f"${rc(1)}${REF0}:${rc(1)}${REF0+11}"
rec_col = lambda c: f'IFERROR(INDEX(${c}${ROW_REC0}:${c}${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")'
reperes = [("1er jour du mois", f"=INDEX({X_},MATCH($C$4,{W_},0))", "DD/MM/YYYY", False),
           ("Dernier jour du mois", "=EOMONTH($O$4,0)", "DD/MM/YYYY", False),
           ("Lundi semaine 1", "=$O$4-WEEKDAY($O$4,3)", "DD/MM/YYYY", False),
           ("Tableau de saisie lu", f"=MATCH($C$4,{W_},0)", "0", False),
           ("Type(s) avion", f"={rec_col('C')}", "General", False),
           ("Rotations / semaine", f"={rec_col('E')}", "0", False),
           ("Total rotations du mois", f"={rec_col('G')}", ROT_FMT, True),
           ("        dont 77W", f"={rec_col('L')}", ROT_FMT, False),
           ("        dont 787", f"={rec_col('M')}", ROT_FMT, False),
           ("        dont A320", f"={rec_col('N')}", ROT_FMT, False)]
for i, (lab, fml, fmt, strong) in enumerate(reperes):
    r = R_REP0 + i
    cell(f"N{r}", lab, size=9, bold=strong, halign="left", fill="FFF2CC" if strong else GREY, border=box)
    cell(f"O{r}", fml, size=9, bold=True, fmt=fmt, fill="FFF2CC" if strong else None, border=box)
MIDX, D1, DN = "$O$7", "$O$4", "$O$5"

# --- bloc auto : programme de la route filtree (par jour de semaine)
ws.merge_cells(f"P{R_HELP_TITLE}:{hc(6)}{R_HELP_TITLE}")
cell(f"P{R_HELP_TITLE}", "PROGRAMME DE LA ROUTE FILTREE (auto)", bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"P{R_HELP_HDR}", "Jour", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_HELP_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
cell(f"P{R_DEPTXT}", "Departs du jour (texte)", bold=True, size=8, halign="left", fill=GREY, border=box)
for ti, t in enumerate(TYPES):
    cell(f"P{R_CODE0+ti}", f"{t} : code du mois", size=8, halign="left", fill=GREY, border=box)
    cell(f"P{R_CODEP0+ti}", f"{t} : code du mois precedent", size=8, halign="left", fill=GREY, border=box)
    cell(f"P{R_DEP0+ti}", f"{t} : rotations au depart", size=8, halign="left", fill=GREY, border=box)
RIDX = f'MATCH($C$5,${rc(3)}${REF0}:${rc(3)}${REF0+11},0)'
for j in range(7):
    c = hc(j)
    for ti, t in enumerate(TYPES):
        cell(f"{c}{R_CODE0+ti}", f'=IFERROR(INDEX({MAT},({RIDX}-1)*{NB_T}+{ti+1},{j+1})&"","")', size=8, border=box)
        cell(f"{c}{R_CODEP0+ti}",
             f'=IFERROR(INDEX({BIG},({MIDX}-2)*{STEP}+({RIDX}-1)*{NB_T}+{ti+1},{j+1})&"","")', size=8, color="808080", border=box)
        cell(f"{c}{R_DEP0+ti}", f"=LEN({c}{R_CODE0+ti})", size=8, fmt="0", border=box)
    dep = '&" "&'.join(f'IF({c}${R_DEP0+ti}=0,"",IF({c}${R_DEP0+ti}=1,"{t}",{c}${R_DEP0+ti}&"x{t}"))'
                       for ti, t in enumerate(TYPES))
    cell(f"{c}{R_DEPTXT}", f'=SUBSTITUTE(TRIM({dep})," ","/")', bold=True, size=9, border=box)

# --- bloc auto : retours par date (le code applique est celui du mois du depart)
ws.merge_cells(f"P{R_RETD_TITLE}:{hc(6)}{R_RETD_TITLE}")
cell(f"P{R_RETD_TITLE}", "RETOURS PAR DATE (auto - un retour du debut de mois vient du tableau du mois precedent)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"P{R_RETD_HDR}", "Type / semaine", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_RETD_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
for ti, t in enumerate(TYPES):
    for w in range(NB_WEEKS):
        r = retd_row(ti, w)
        cell(f"P{r}", f"{t} - semaine {w+1}", size=8, halign="left", fill=GREY, border=box)
        for j in range(7):
            d = f"{gl(GRID_C0+j)}{ROW_W1+2*w+1}"
            terms = []
            for k in range(1, len(OFFSETS)):
                cur = f"{hc((j-k) % 7)}${R_CODE0+ti}"
                prv = f"{hc((j-k) % 7)}${R_CODEP0+ti}"
                pick = f'IF({d}-{k}<{D1},{prv},{cur})'
                terms.append(f'LEN({pick})-LEN(SUBSTITUTE({pick},"{k}",""))')
            cell(f"{hc(j)}{r}", f'=IF({d}="","",' + "+".join(terms) + ")", size=8, fmt="0", border=box)

# --- bloc auto : textes affiches dans la grille (programme + vols additionnels)
ws.merge_cells(f"P{R_TXT_TITLE}:{hc(6)}{R_TXT_TITLE}")
cell(f"P{R_TXT_TITLE}", "TEXTES DE LA GRILLE (auto - programme + vols additionnels)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"P{R_TXT_HDR}", "Semaine", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_TXT_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)

def add_label(n, nrows, t1):
    """etiquette d'un ajout (+) ou d'un retrait (-) de vols."""
    one = f'IF({n}>0,"+","-")&{t1}'
    many = f'IF({n}>0,"+","")&{n}&"x"&{t1}'
    return f'IF({nrows}=1,IF(ABS({n})=1,{one},{many}),IF({n}>0,"+","")&{n}&" vols")'

for w in range(NB_WEEKS):
    cell(f"P{R_DEPTXT0+w}", f"Departs - semaine {w+1}", size=8, halign="left", fill=GREY, border=box)
    cell(f"P{R_RETTXT0+w}", f"Retours - semaine {w+1}", size=8, halign="left", fill=GREY, border=box)
    for j in range(7):
        d = f"{gl(GRID_C0+j)}{ROW_W1+2*w+1}"
        nd = f"SUMIFS({ADD_N},{ADD_R},$C$5,{ADD_D},{d})"
        rd = f"COUNTIFS({ADD_R},$C$5,{ADD_D},{d})"
        td = f'IFERROR(INDEX({ADD_T},MATCH({d}&"|"&$C$5,{ADD_KD},0)),"")'
        cell(f"{hc(j)}{R_DEPTXT0+w}",
             f'=IF({d}="","",TRIM(${hc(j)}${R_DEPTXT}&IF({nd}=0,""," ("&{add_label(nd, rd, td)}&")")))',
             size=8, border=box)
        nr = f"SUMIFS({ADD_N},{ADD_R},$C$5,{ADD_RDA},{d})"
        rr = f"COUNTIFS({ADD_R},$C$5,{ADD_RDA},{d})"
        tr = f'IFERROR(INDEX({ADD_T},MATCH({d}&"|"&$C$5,{ADD_KR},0)),"")'
        prog = '&" "&'.join(
            f'IF({hc(j)}${retd_row(ti,w)}=0,"",IF({hc(j)}${retd_row(ti,w)}=1,"{t}",{hc(j)}${retd_row(ti,w)}&"x{t}"))'
            for ti, t in enumerate(TYPES))
        cell(f"{hc(j)}{R_RETTXT0+w}",
             f'=IF({d}="","",TRIM(SUBSTITUTE(TRIM({prog})," ","/")&IF({nr}=0,""," ("&{add_label(nr, rr, tr)}&")")))',
             size=8, border=box)

# --- matrice du mois filtre + queues de mois
ws.merge_cells(f"P{R_MAT_TITLE}:{hc(6)}{R_MAT_TITLE}")
cell(f"P{R_MAT_TITLE}", "PROGRAMME DU MOIS FILTRE, TOUTES ROUTES (auto)", bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"P{R_MAT_HDR}", "Route / type", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_MAT_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
cell(f"{gl(HELP_C0+7)}{R_MAT_HDR}", "Fin de mois (retours en M+1)", bold=True, size=8, fill=GREY, wrap=True, border=box)
cell(f"{gl(HELP_C0+8)}{R_MAT_HDR}", "Fin du mois precedent (retours en M)", bold=True, size=8, fill=GREY, wrap=True, border=box)
ws.row_dimensions[R_MAT_HDR].height = 30
for i, (route, colr, white) in enumerate(ROUTES):
    for ti, t in enumerate(TYPES):
        r = mat_row(i, t)
        cell(f"P{r}", f"{route} - {t}", size=8, halign="left", border=box)
        for j in range(7):
            cell(f"{hc(j)}{r}",
                 f'=IFERROR(INDEX({BIG},({MIDX}-1)*{STEP}+{rel_of(i,ti)},{j+1})&"","")', size=8, border=box)
        tail, head = [], []
        for k in range(1, len(OFFSETS)):
            for i2 in range(1, k+1):
                cc = f'IFERROR(INDEX(${hc(0)}${r}:${hc(6)}${r},WEEKDAY({DN}-{i2-1},2))&"","")'
                tail.append(cnt_k(cc, k))
                cp = f'IFERROR(INDEX({BIG},({MIDX}-2)*{STEP}+{rel_of(i,ti)},WEEKDAY({D1}-{i2},2))&"","")'
                head.append(cnt_k(cp, k))
        cell(f"{gl(HELP_C0+7)}{r}", "=" + "+".join(tail), size=8, fmt="0", border=box)
        cell(f"{gl(HELP_C0+8)}{r}", "=" + "+".join(head), size=8, fmt="0", border=box)

# --- listes de reference
for i, h in enumerate(("Mois (filtre)", "1er jour", "Types avion", "Routes")):
    cell(f"{rc(i)}{REF0-1}", h, bold=True, size=8, color="FFFFFF", fill="808080", wrap=True, border=box)
for i, (lab, y, m) in enumerate(MOIS):
    cell(f"{rc(0)}{REF0+i}", lab, size=8, halign="left", border=box)
    cell(f"{rc(1)}{REF0+i}", f"=DATE({y},{m},1)", size=8, fmt="DD/MM/YYYY", border=box)
for i, t in enumerate(TYPES):
    cell(f"{rc(2)}{REF0+i}", t, size=8, border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    cell(f"{rc(3)}{REF0+i}", route, size=8, border=box)

# --- grille
ws.merge_cells(f"B{ROW_HDR-1}:I{ROW_HDR-1}")
cell(f"B{ROW_HDR-1}", "GRILLE MENSUELLE  -  rotations par semaine et par jour (filtres Mois + Route)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
cell(f"B{ROW_HDR}", "Semaine", bold=True, size=10, color="FFFFFF", fill="2F5597", border=box)
for j in range(7):
    cell(f"{gl(GRID_C0+j)}{ROW_HDR}", JOURS[j], bold=True, size=10, color="FFFFFF", fill="2F5597", border=box)
type_rows, date_rows = [], []
for w in range(NB_WEEKS):
    rt, rd = ROW_W1 + 2*w, ROW_W1 + 2*w + 1
    type_rows.append(rt); date_rows.append(rd)
    first, last = f"{gl(GRID_C0)}{rd}", f"{gl(GRID_C0+6)}{rd}"
    cell(f"B{rt}", f"Semaine {w+1}", bold=True, size=10, fill=LIGHT, halign="left", border=box)
    cell(f"B{rd}", f'=IF(COUNT({first}:{last})=0,"","du "&TEXT(MIN({first}:{last}),"DD/MM")&'
                   f'" au "&TEXT(MAX({first}:{last}),"DD/MM"))',
         size=8, italic=True, color="595959", fill=LIGHT, halign="left", border=box)
    ws.row_dimensions[rt].height = 30
    ws.row_dimensions[rd].height = 14
    for j in range(7):
        col, off = gl(GRID_C0+j), 7*w + j
        d = f"{col}{rd}"
        cell(d, f'=IF(OR($O$6+{off}<$O$4,$O$6+{off}>$O$5),"",$O$6+{off})',
             size=8, italic=True, color="595959", fmt="DD/MM", border=box)
        AL, RE = f"${hc(j)}${R_DEPTXT0+w}", f"${hc(j)}${R_RETTXT0+w}"
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
                    '"< 787" = les RETOURS du jour  |  cellule encadree en VERT = vol AJOUTE "(+787)", encadree en ROUGE = vol '
                    'RETIRE "(-787)", saisis sur la feuille Vols additionnels.',
     size=8, italic=True, color="595959", halign="left")

# --- recapitulatif
ws.merge_cells(f"B{ROW_REC_TITLE}:N{ROW_REC_TITLE}")
cell(f"B{ROW_REC_TITLE}", "RECAPITULATIF MENSUEL - TOUTES ROUTES  (depend uniquement du filtre Mois)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
for i, h in enumerate(["Route", "Type(s) avion", "Semaine type du mois (jours J1-J7)", "Rotations/semaine",
                       "Nb de semaines/occurrences dans le mois", "Total rotations du mois", "Departs du mois",
                       "Arrivees du mois", "dont programme du mois", "dont vols additionnels",
                       "Rotations 77W", "Rotations 787", "Rotations A320"]):
    cell(f"{gl(2+i)}{ROW_REC_HDR}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
ws.row_dimensions[ROW_REC_HDR].height = 46
for i in range(len(ROUTES)):
    r = ROW_REC0 + i
    rows = {t: mat_row(i, t) for t in TYPES}
    dep_p = {t: f'SUMPRODUCT({OCC},LEN(${hc(0)}${rows[t]}:${hc(6)}${rows[t]}))' for t in TYPES}
    arr_p = {t: f'({dep_p[t]}-${gl(HELP_C0+7)}${rows[t]}+${gl(HELP_C0+8)}${rows[t]})' for t in TYPES}
    dep_a = {t: f'SUMIFS({ADD_N},{ADD_R},$B{r},{ADD_T},"{t}",{ADD_D},">="&{D1},{ADD_D},"<="&{DN})' for t in TYPES}
    arr_a = {t: f'SUMIFS({ADD_N},{ADD_R},$B{r},{ADD_T},"{t}",{ADD_RD},">="&{D1},{ADD_RD},"<="&{DN})' for t in TYPES}
    nsem = {t: f'SUMPRODUCT(LEN(${hc(0)}${rows[t]}:${hc(6)}${rows[t]}))' for t in TYPES}
    cell(f"B{r}", f"=${rc(3)}${REF0+i}", bold=True, size=10, border=box)
    tl = "&".join(f'IF({nsem[t]}>0,"{t}("&{nsem[t]}&") ","")' for t in TYPES)
    cell(f"C{r}", f'=IF(E{r}=0,"-",SUBSTITUTE(TRIM({tl})," ",", "))', size=9, wrap=True, border=box)
    jl = "&".join(f'IF(SUMPRODUCT(LEN(${hc(j)}${rows[TYPES[0]]}:${hc(j)}${rows[TYPES[-1]]}))>0,"J{j+1} ","")'
                  for j in range(7))
    cell(f"D{r}", f'=IF(E{r}=0,"aucune operation",SUBSTITUTE(TRIM({jl})," ",", "))', size=9, wrap=True, border=box)
    cell(f"E{r}", "=" + "+".join(nsem[t] for t in TYPES), size=10, fmt="0", border=box)
    cell(f"F{r}", f"=$C${ROW_NBW}", size=10, fmt="0", border=box)
    cell(f"G{r}", f"=(H{r}+I{r})/2", bold=True, size=11, fmt=ROT_FMT, border=box)
    cell(f"H{r}", "=" + "+".join(f"{dep_p[t]}+{dep_a[t]}" for t in TYPES), size=10, fmt="0", border=box)
    cell(f"I{r}", "=" + "+".join(f"{arr_p[t]}+{arr_a[t]}" for t in TYPES), size=10, fmt="0", border=box)
    cell(f"J{r}", "=(" + "+".join(f"{dep_p[t]}+{arr_p[t]}" for t in TYPES) + ")/2", size=10, fmt=ROT_FMT, border=box)
    cell(f"K{r}", "=(" + "+".join(f"{dep_a[t]}+{arr_a[t]}" for t in TYPES) + ")/2", size=10, fmt=ROT_FMT, border=box)
    for t, col in zip(TYPES, ("L", "M", "N")):
        cell(f"{col}{r}", f"=({dep_p[t]}+{arr_p[t]}+{dep_a[t]}+{arr_a[t]})/2", size=9, fmt=ROT_FMT, border=box)
    ws.row_dimensions[r].height = 24
cell(f"B{ROW_REC_TOT}", "TOTAL GENERAL", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
for c in ("C", "D"):
    cell(f"{c}{ROW_REC_TOT}", "", fill=NAVY, border=box)
for c in ("E", "F", "G", "H", "I", "J", "K", "L", "M", "N"):
    f = f"=$C${ROW_NBW}" if c == "F" else f"=SUM({c}{ROW_REC0}:{c}{ROW_REC_TOT-1})"
    fmt = "0" if c in ("E", "F", "H", "I") else ROT_FMT
    cell(f"{c}{ROW_REC_TOT}", f, bold=True, size=11 if c == "G" else 10, color="FFFFFF", fill=NAVY, fmt=fmt, border=box)
r = ROW_REC_TOT + 1
ws.merge_cells(f"B{r}:N{r+1}")
cell(f"B{r}", "Methode : chaque caractere saisi dans une case = une rotation, c'est-a-dire un aller-retour. Elle compte un DEPART "
              "le jour ou elle part et une ARRIVEE le jour ou elle rentre (jour du depart + le chiffre du caractere). Le total du "
              "mois vaut (departs + arrivees) / 2 : une rotation entierement dans le mois compte 1, une rotation partie en fin de "
              "mois et rentree le mois suivant compte 0,5 de chaque cote. Les arrivees du debut de mois issues du mois precedent "
              "sont reprises du tableau de ce mois precedent. Les occurrences de chaque jour sont comptees sur les dates reelles "
              "(semaines partielles incluses).",
     size=8, italic=True, color="595959", halign="left", wrap=True)

def add_dv(sheet, dv, ranges):
    sheet.add_data_validation(dv)
    for rng in (ranges if isinstance(ranges, (list, tuple)) else [ranges]):
        dv.add(rng)
add_dv(ws, DataValidation(type="list", formula1=f"=${rc(0)}${REF0}:${rc(0)}${REF0+11}", allow_blank=False), "C4")
add_dv(ws, DataValidation(type="list", formula1=f"=${rc(3)}${REF0}:${rc(3)}${REF0+11}", allow_blank=False), "C5")

grid_ranges = " ".join(f"{gl(GRID_C0)}{r}:{gl(GRID_C0+6)}{r}" for r in type_rows)
c0 = gl(GRID_C0)
def frame(color):
    s_ = Side(style="medium", color=color)
    return Border(left=s_, right=s_, top=s_, bottom=s_)
# un retrait prime sur un ajout quand la cellule porte les deux
for motif, couleur in (("(-", RED), ("(+", GREEN)):
    ws.conditional_formatting.add(
        grid_ranges,
        FormulaRule(formula=[f'ISNUMBER(SEARCH("{motif}",{c0}{ROW_W1}))'],
                    border=frame(couleur), stopIfTrue=False))
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
widths = {"A": 2, "B": 26, "J": 13, "K": 13, "L": 12, "M": 12, "N": 13, "O": 13, "P": 26}
for j in range(7):
    widths[gl(GRID_C0+j)] = 16
for j in range(7):
    widths[gl(HELP_C0+j)] = 10
widths[gl(HELP_C0+7)] = 12
widths[gl(HELP_C0+8)] = 12
for i, w in enumerate((18, 12, 10, 10)):
    widths[rc(i)] = w
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = "C8"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_area = f"B1:N{ROW_REC_TOT+2}"

# ================================================================= FEUILLE 2 : SAISIE PROGRAMME
sa.merge_cells("B1:L1")
SA("B1", "SAISIE DU PROGRAMME - 12 TABLEAUX, UN PAR MOIS", bold=True, size=16, color="FFFFFF", fill=BROWN, halign="left")
sa.row_dimensions[1].height = 30
sa.merge_cells("B2:L4")
SA("B2", "Cellules en BLEU = saisie utilisateur. Chaque route occupe 3 lignes = les 3 types avion. Dans la case d'un jour, on "
         "saisit UN CARACTERE PAR ROTATION (aller-retour) qui PART ce jour-la, et ce caractere indique QUAND CETTE ROTATION "
         "RENTRE :  " + CODE_AIDE + "  Case vide = aucun vol. Une rotation : fond vert ; plusieurs rotations : fond orange. "
         "Une rotation partie en fin de mois et rentree le mois suivant est comptee 0,5 dans chacun des deux mois. Les 12 "
         "tableaux sont independants ; pour reprendre un mois sur un autre, copier la zone des 7 jours (36 lignes).",
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
        SA(f"B{p}", route, bold=True, size=11, halign="left", fill=colr, color="FFFFFF" if white else "000000", border=box)
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
    "Ex. : 01 = deux rotations, une rentre le jour meme, l'autre le lendemain.", "Modele")
anchor = f"D{DATA0}"
strip = f'SUBSTITUTE(UPPER({anchor}),"X","0")'
for ch in OFFSETS:
    strip = f'SUBSTITUTE({strip},"{ch}","")'
add_dv(sa, DataValidation(type="custom", formula1=f"=AND(LEN({anchor})<={MAX_ROT},LEN({strip})=0)",
                          allow_blank=True, errorTitle="Code de rotation invalide",
                          error="Un caractere par rotation : 0 (ou x) = retour le jour meme, 1 = J+1, 2 = J+2, 3 = J+3.\n"
                                f"Exemples : 0 | 1 | 01 | 00. Maximum {MAX_ROT} rotations par jour et par type."),
       [f"D{blk_title(k)+2}:J{blk_title(k)+1+NB_ROWS_BLK}" for k in range(len(MOIS))])
chk = " ".join(f"D{blk_title(k)+2}:J{blk_title(k)+1+NB_ROWS_BLK}" for k in range(len(MOIS)))
sa.conditional_formatting.add(chk, FormulaRule(formula=[f'LEN(D{DATA0})>1'], fill=PatternFill("solid", fgColor="ED7D31"),
                                               font=Font(name=F, size=11, bold=True, color="FFFFFF"), stopIfTrue=True))
sa.conditional_formatting.add(chk, FormulaRule(formula=[f'D{DATA0}<>""'], fill=PatternFill("solid", fgColor=CHECK),
                                               font=Font(name=F, size=11, bold=True, color="FFFFFF"), stopIfTrue=True))
for k, (lab, y, m) in enumerate(MOIS):
    sa.conditional_formatting.add(
        f"B{blk_title(k)}:L{blk_title(k)}",
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
va.merge_cells("B1:G1")
VA("B1", "VOLS ADDITIONNELS ET REGULATIONS - ajouts (+) et retraits (-)", bold=True, size=16, color="FFFFFF", fill=BROWN, halign="left")
va.row_dimensions[1].height = 30
va.merge_cells("B2:G4")
VA("B2", "Une ligne = un ecart ponctuel par rapport au programme du mois. Nb de rotations POSITIF = vol(s) ajoute(s) : charter, "
         "renfort, fret. Nb de rotations NEGATIF = vol(s) RETIRE(S) : regulation, annulation, jour ferie... Le retrait se soustrait "
         "du programme de ce jour-la (verifier qu'il y a bien un vol a retirer). La colonne Retour J+n fonctionne comme dans le "
         "programme : 0 = retour le jour meme, 1 = le lendemain. Une rotation compte 0,5 le mois du depart et 0,5 le mois du "
         "retour ; un vol parti le 31/03/2027 et rentre le 01/04/2027 apporte donc 0,5 rotation a avril. Les dates du 25/03/2027 "
         "au 31/03/2028 sont acceptees, pour pouvoir saisir un depart anterieur a l'annee dont le retour tombe en avril. "
         "Cette table est livree VIDE : tout ce qui y est saisi est compte.",
   size=9, italic=True, color=BROWN, halign="left", wrap=True)
va.row_dimensions[2].height = 56
c = VA("B5", "<< retour a la feuille Programme de vols", size=9, bold=True, color="0563C1", halign="left")
c.hyperlink = Hyperlink(ref="B5", location=f"'{S1N}'!B1", tooltip="Retour a la feuille de lecture")
for i, h in enumerate(["Date du vol (depart)", "Route", "Type avion", "Nb de rotations (+ ajout / - retrait)",
                       "Retour J+n", "Commentaire (libre)"]):
    VA(f"{gl(2+i)}{ROW_ADD_HDR}", h, bold=True, size=9, color="FFFFFF", fill=BROWN, wrap=True, border=box)
for i, h in enumerate(["Date de retour (auto)", "Cle depart (auto)", "Date retour affichee (auto)", "Cle retour (auto)"]):
    VA(f"{gl(9+i)}{ROW_ADD_HDR}", h, bold=True, size=8, color="FFFFFF", fill="A6A6A6", wrap=True, border=box)
va.row_dimensions[ROW_ADD_HDR].height = 40
for i in range(NB_ADD_ROWS):
    r = ROW_ADD0 + i
    VA(f"B{r}", None, size=10, color=BLUE_IN, fmt="DD/MM/YYYY", border=box)
    VA(f"C{r}", None, size=10, color=BLUE_IN, border=box)
    VA(f"D{r}", None, size=10, color=BLUE_IN, border=box)
    VA(f"E{r}", None, size=10, bold=True, color=BLUE_IN, fmt="+0;-0;0", border=box)
    VA(f"F{r}", None, size=10, color=BLUE_IN, fmt="0", border=box)
    VA(f"G{r}", None, size=9, italic=True, color="595959", halign="left", border=box)
    VA(f"I{r}", f'=IF($B{r}="","",$B{r}+N($F{r}))', size=8, color="A6A6A6", fmt="DD/MM/YYYY", border=box)
    VA(f"J{r}", f'=IF(OR($B{r}="",$C{r}=""),"",$B{r}&"|"&$C{r})', size=8, color="A6A6A6", border=box)
    VA(f"K{r}", f'=IF(OR($B{r}="",N($F{r})=0),"",$B{r}+$F{r})', size=8, color="A6A6A6", fmt="DD/MM/YYYY", border=box)
    VA(f"L{r}", f'=IF(OR($K{r}="",$C{r}=""),"",$K{r}&"|"&$C{r})', size=8, color="A6A6A6", border=box)
va[f"B{ROW_ADD0}"].comment = Comment(
    "Date de depart du vol (25/03/2027 -> 31/03/2028).\n"
    "Nb de rotations : 1 pour un ajout, -1 pour un retrait, 2 ou -2 pour deux vols.\n"
    "Retour J+n : 0 = retour le jour meme, 1 = le lendemain, etc.", "Modele")
VA(f"N{ROW_ADD_HDR}", "Routes", bold=True, size=8, color="FFFFFF", fill="808080", border=box)
VA(f"O{ROW_ADD_HDR}", "Types", bold=True, size=8, color="FFFFFF", fill="808080", border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    VA(f"N{ROW_ADD0+i}", route, size=8, border=box)
for i, t in enumerate(TYPES):
    VA(f"O{ROW_ADD0+i}", t, size=8, border=box)
VA(f"N{ROW_ADD0+len(ROUTES)+1}", "listes de reference - ne pas modifier", size=8, italic=True, color="808080", halign="left")
add_dv(va, DataValidation(type="list", formula1=f"=$N${ROW_ADD0}:$N${ROW_ADD0+len(ROUTES)-1}", allow_blank=True,
                          errorTitle="Route inconnue", error="Choisir une route dans la liste."),
       f"C{ROW_ADD0}:C{ROW_ADD_END}")
add_dv(va, DataValidation(type="list", formula1=f"=$O${ROW_ADD0}:$O${ROW_ADD0+len(TYPES)-1}", allow_blank=True,
                          errorTitle="Type avion inconnu", error="Choisir un type avion dans la liste."),
       f"D{ROW_ADD0}:D{ROW_ADD_END}")
add_dv(va, DataValidation(type="custom", formula1=f"=AND(E{ROW_ADD0}<>0,ABS(E{ROW_ADD0})<=20)", allow_blank=True,
                          errorTitle="Nombre de rotations invalide",
                          error="Entier non nul entre -20 et 20. Positif = ajout, negatif = retrait."),
       f"E{ROW_ADD0}:E{ROW_ADD_END}")
add_dv(va, DataValidation(type="list", formula1='"0,1,2,3"', allow_blank=True, showDropDown=False,
                          errorTitle="Decalage retour", error="0 = retour le jour meme, 1 = le lendemain, etc."),
       f"F{ROW_ADD0}:F{ROW_ADD_END}")
add_dv(va, DataValidation(type="date", operator="between", formula1="DATE(2027,3,25)", formula2="DATE(2028,3,31)",
                          allow_blank=True, errorStyle="warning", errorTitle="Date hors perimetre",
                          error="Dates acceptees : 25/03/2027 -> 31/03/2028."),
       f"B{ROW_ADD0}:B{ROW_ADD_END}")
va.conditional_formatting.add(
    f"B{ROW_ADD0}:G{ROW_ADD_END}",
    FormulaRule(formula=[f"$E{ROW_ADD0}<0"], fill=PatternFill("solid", fgColor="FCE4E4"), stopIfTrue=True))
va.conditional_formatting.add(
    f"B{ROW_ADD0}:G{ROW_ADD_END}",
    FormulaRule(formula=[f"$E{ROW_ADD0}>0"], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=True))
for col, w in {"A": 2, "B": 18, "C": 14, "D": 14, "E": 16, "F": 12, "G": 52, "H": 3,
               "I": 16, "J": 18, "K": 16, "L": 18, "M": 3, "N": 12, "O": 10}.items():
    va.column_dimensions[col].width = w
va.freeze_panes = f"B{ROW_ADD0}"
va.page_setup.orientation = "landscape"
va.page_setup.fitToWidth = 1
va.sheet_properties.pageSetUpPr.fitToPage = True
va.print_area = f"B1:G{ROW_ADD_END}"

# ================================================================= FEUILLE 4 : ANALYSE CAPACITE
R_CAB, R_HYP, R_SYN, R_SIE, R_ROT, R_TYP, R_CAL = 5, 6, 15, 31, 48, 65, 81
R_PDEP, R_TAIL, R_ROTD = 96, 135, 174          # blocs auto : departs programme, queues de mois, rotations
mcol = lambda m: gl(3+m)
rcol = lambda i: gl(3+i)
cal_row = lambda m: R_CAL + 1 + m
occ_row = lambda m: f"$C${cal_row(m)}:$I${cal_row(m)}"
drow = lambda base, i, ti: base + 1 + NB_T*i + ti
seat = {t: f"$G${R_HYP+1+ti}" for ti, t in enumerate(TYPES)}   # sieges retenus selon le filtre cabine
LEGS = f"$C${R_HYP+5}"
CABF = f"$C${R_CAB}"

an.merge_cells("B1:J1")
A("B1", "ANALYSE DE CAPACITE - SIEGES OFFERTS (avril 2027 / mars 2028)",
  bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
an.row_dimensions[1].height = 30
an.merge_cells("B2:J2")
A("B2", "Interpretation automatique des feuilles \"Saisie programme\" et \"Vols additionnels\", croisee avec les capacites "
        "ci-dessous. Une rotation = un aller-retour : elle compte 0,5 le mois de son depart et 0,5 le mois de son retour, donc "
        "les rotations d'un mois peuvent etre fractionnaires. Aucune saisie ici, hormis les hypotheses de capacite et le filtre de cabine, qui font suivre tous les sieges de la feuille.",
  size=9, italic=True, halign="left", wrap=True, color="404040")
an.row_dimensions[2].height = 26
an.merge_cells(f"B{R_CAB-1}:H{R_CAB-1}")
A(f"B{R_CAB-1}", "HYPOTHESES DE CAPACITE (sieges par rotation, par cabine)",
  bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
A(f"B{R_CAB}", "CABINE ANALYSEE :", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
an.merge_cells(f"C{R_CAB}:D{R_CAB}")
A(f"C{R_CAB}", CAB_FILTRE[0], bold=True, size=12, color=BLUE_IN, fill="FFF2CC",
  border=Border(left=med, right=med, top=med, bottom=med))
an.merge_cells(f"E{R_CAB}:H{R_CAB}")
A(f"E{R_CAB}", "<-- liste deroulante : toutes les valeurs en sieges de cette feuille suivent la cabine choisie "
               "(C = Club, W = Confort, Y = Loisir).", size=9, italic=True, halign="left", color="808080")
for i, h in enumerate(["Type avion", "C - Club", "W - Confort", "Y - Loisir", "Total cabine",
                       "Sieges retenus (filtre)", "Remarque"]):
    A(f"{gl(2+i)}{R_HYP}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
an.row_dimensions[R_HYP].height = 30
for ti, t in enumerate(TYPES):
    r = R_HYP + 1 + ti
    A(f"B{r}", t, bold=True, size=10, border=box)
    for ci, (code, lab) in enumerate(CABINES):
        A(f"{gl(3+ci)}{r}", SIEGES_CAB[t][code], bold=True, size=10, color=BLUE_IN, fill="FFF2CC", fmt="#,##0", border=box)
    A(f"F{r}", f"=SUM(C{r}:E{r})", bold=True, size=10, fmt="#,##0", border=box)
    A(f"G{r}", f'=IF({CABF}="{CAB_FILTRE[0]}",$F{r},IF(LEFT({CABF},1)="C",$C{r},IF(LEFT({CABF},1)="W",$D{r},$E{r})))',
      bold=True, size=11, fill="E2EFDA", fmt="#,##0", border=box)
    A(f"H{r}", SIEGES_NOTE[t], size=8, italic=True, color="595959", halign="left", border=box)
rl = R_HYP + 5
A(f"B{rl}", "Legs comptes par rotation", bold=True, size=10, border=box)
A(f"C{rl}", 1, bold=True, size=10, color=BLUE_IN, fill="FFF2CC", fmt="0", border=box)
an.merge_cells(f"D{rl}:H{rl}")
A(f"D{rl}", "1 = l'aller-retour compte pour un vol (sieges d'un sens) ; mettre 2 pour compter l'aller ET le retour.",
  size=8, italic=True, color="595959", halign="left", border=box)
an.merge_cells(f"B{rl+1}:H{rl+1}")
A(f"B{rl+1}", "Les cellules bleues sont modifiables. La colonne \"Sieges retenus\" est celle qu'utilisent tous les tableaux "
              "ci-dessous : elle vaut le total du type avion quand le filtre est sur \"Toutes cabines\", et la capacite de la "
              "seule cabine choisie sinon.",
  size=8, italic=True, color="595959", halign="left", wrap=True)
add_dv(an, DataValidation(type="list", formula1='"' + ",".join(CAB_FILTRE) + '"', allow_blank=False,
                          errorTitle="Cabine inconnue", error="Choisir : toutes cabines, C, W ou Y."),
       f"C{R_CAB}")

an.merge_cells(f"B{R_SYN-1}:H{R_SYN-1}")
A(f"B{R_SYN-1}", f'="SYNTHESE ANNUELLE PAR ROUTE"&IF({CABF}="{CAB_FILTRE[0]}",""," - CABINE "&{CABF})',
  bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
for i, h in enumerate(["Route", "Rotations sur l'annee", "Sieges offerts sur l'annee", "Sieges / rotation",
                       "Part du reseau (sieges)"]):
    A(f"{gl(2+i)}{R_SYN}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
an.row_dimensions[R_SYN].height = 30
for i, (route, colr, white) in enumerate(ROUTES):
    r = R_SYN + 1 + i
    col = rcol(i)
    A(f"B{r}", route, bold=True, size=10, fill=colr, color="FFFFFF" if white else "000000", border=box)
    A(f"C{r}", f"=SUM({col}${R_ROT+1}:{col}${R_ROT+len(MOIS)})", size=10, fmt=ROT_FMT, border=box)
    A(f"D{r}", f"=SUM({col}${R_SIE+1}:{col}${R_SIE+len(MOIS)})", bold=True, size=10, fmt="#,##0", border=box)
    A(f"E{r}", f'=IFERROR(D{r}/C{r},"-")', size=10, fmt="#,##0", border=box)
    A(f"F{r}", f"=IFERROR(D{r}/$D${R_SYN+len(ROUTES)+1},0)", size=10, fmt="0.0%", border=box)
rt2 = R_SYN + len(ROUTES) + 1
A(f"B{rt2}", "TOTAL RESEAU", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
A(f"C{rt2}", f"=SUM(C{R_SYN+1}:C{rt2-1})", bold=True, size=11, color="FFFFFF", fill=NAVY, fmt=ROT_FMT, border=box)
A(f"D{rt2}", f"=SUM(D{R_SYN+1}:D{rt2-1})", bold=True, size=11, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)
A(f"E{rt2}", f'=IFERROR(D{rt2}/C{rt2},"-")', bold=True, size=10, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)
A(f"F{rt2}", f"=IFERROR(D{rt2}/$D${rt2},0)", bold=True, size=10, color="FFFFFF", fill=NAVY, fmt="0.0%", border=box)

for base, titre, fmt in ((R_SIE, "SIEGES OFFERTS PAR MOIS ET PAR ROUTE", "#,##0"),
                         (R_ROT, "ROTATIONS PAR MOIS ET PAR ROUTE", ROT_FMT)):
    an.merge_cells(f"B{base-1}:O{base-1}")
    A(f"B{base-1}", f'="{titre}"&IF({CABF}="{CAB_FILTRE[0]}",""," - CABINE "&{CABF})',
      bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
    A(f"B{base}", "Mois", bold=True, size=9, color="FFFFFF", fill="2F5597", border=box)
    for i, (route, colr, white) in enumerate(ROUTES):
        A(f"{rcol(i)}{base}", route, bold=True, size=9, fill=colr, color="FFFFFF" if white else "000000", border=box)
    A(f"O{base}", "TOTAL", bold=True, size=9, color="FFFFFF", fill="2F5597", border=box)
    for m, (lab, y, mm) in enumerate(MOIS):
        r = base + 1 + m
        A(f"B{r}", f"={S1R}${rc(0)}${REF0+m}", size=9, halign="left", border=box)
        for i, (route, colr, white) in enumerate(ROUTES):
            d = mcol(m)
            if base == R_ROT:
                f = "=" + "+".join(f"{d}{drow(R_ROTD,i,ti)}" for ti in range(NB_T))
            else:
                f = "=" + "+".join(f"{seat[t]}*{LEGS}*{d}{drow(R_ROTD,i,ti)}" for ti, t in enumerate(TYPES))
            A(f"{rcol(i)}{r}", f, size=9, fmt=fmt, border=box)
        A(f"O{r}", f"=SUM(C{r}:N{r})", bold=True, size=9, fmt=fmt, border=box)
    r = base + 1 + len(MOIS)
    A(f"B{r}", "TOTAL ANNEE", bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left", border=box)
    for i in range(len(ROUTES) + 1):
        c = rcol(i) if i < len(ROUTES) else "O"
        A(f"{c}{r}", f"=SUM({c}{base+1}:{c}{base+len(MOIS)})", bold=True, size=9,
          color="FFFFFF", fill=NAVY, fmt=fmt, border=box)

an.merge_cells(f"B{R_TYP-1}:K{R_TYP-1}")
A(f"B{R_TYP-1}", f'="ROTATIONS ET SIEGES PAR TYPE AVION ET PAR MOIS"&IF({CABF}="{CAB_FILTRE[0]}",""," - CABINE "&{CABF})',
  bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
hdr = ["Mois"] + [f"Rotations {t}" for t in TYPES] + ["Total rotations"] + [f"Sieges {t}" for t in TYPES] + \
      ["Total sieges", "Sieges / rotation"]
for i, h in enumerate(hdr):
    A(f"{gl(2+i)}{R_TYP}", h, bold=True, size=9, color="FFFFFF", fill="2F5597", wrap=True, border=box)
an.row_dimensions[R_TYP].height = 30
DET_T = f"$P${R_ROTD+1}:$P${R_ROTD+NB_ROWS_BLK}"
for m, (lab, y, mm) in enumerate(MOIS):
    r = R_TYP + 1 + m
    d = mcol(m)
    A(f"B{r}", f"={S1R}${rc(0)}${REF0+m}", size=9, halign="left", border=box)
    for ti, t in enumerate(TYPES):
        A(f"{gl(3+ti)}{r}", f'=SUMIF({DET_T},"{t}",${d}${R_ROTD+1}:${d}${R_ROTD+NB_ROWS_BLK})', size=9, fmt=ROT_FMT, border=box)
    A(f"F{r}", f"=SUM(C{r}:E{r})", bold=True, size=9, fmt=ROT_FMT, border=box)
    for ti, t in enumerate(TYPES):
        A(f"{gl(7+ti)}{r}", f"={gl(3+ti)}{r}*{seat[t]}*{LEGS}", size=9, fmt="#,##0", border=box)
    A(f"J{r}", f"=SUM(G{r}:I{r})", bold=True, size=9, fmt="#,##0", border=box)
    A(f"K{r}", f'=IFERROR(J{r}/F{r},"-")', size=9, fmt="#,##0", border=box)
r = R_TYP + 1 + len(MOIS)
A(f"B{r}", "TOTAL ANNEE", bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left", border=box)
for i in range(2, 10):
    c = gl(i+1)
    A(f"{c}{r}", f"=SUM({c}{R_TYP+1}:{c}{r-1})", bold=True, size=9, color="FFFFFF", fill=NAVY,
      fmt=ROT_FMT if i < 6 else "#,##0", border=box)
A(f"K{r}", f'=IFERROR(J{r}/F{r},"-")', bold=True, size=9, color="FFFFFF", fill=NAVY, fmt="#,##0", border=box)

an.merge_cells(f"B{R_CAL-1}:L{R_CAL-1}")
A(f"B{R_CAL-1}", "CALENDRIER AUTO : occurrences de chaque jour de semaine dans le mois (semaines partielles incluses)",
  bold=True, size=9, color="FFFFFF", fill="808080", halign="left")
for i, h in enumerate(["Mois", "J1", "J2", "J3", "J4", "J5", "J6", "J7", "1er jour", "Dernier jour", "Nb de jours"]):
    A(f"{gl(2+i)}{R_CAL}", h, bold=True, size=8, fill=GREY, border=box)
for m, (lab, y, mm) in enumerate(MOIS):
    r = cal_row(m)
    A(f"B{r}", f"={S1R}${rc(0)}${REF0+m}", size=8, halign="left", border=box)
    A(f"J{r}", f"={S1R}${rc(1)}${REF0+m}", size=8, fmt="DD/MM/YYYY", border=box)
    A(f"K{r}", f"=EOMONTH($J{r},0)", size=8, fmt="DD/MM/YYYY", border=box)
    A(f"L{r}", f"=$K{r}-$J{r}+1", size=8, fmt="0", border=box)
    for j in range(7):
        off = f"MOD({j+1}-WEEKDAY($J{r},2),7)"
        A(f"{gl(3+j)}{r}", f"=IF({off}>=$L{r},0,INT(($L{r}-1-{off})/7)+1)", size=8, fmt="0", border=box)

for base, titre in ((R_PDEP, "AUTO 1/3 : departs du programme (hors vols additionnels)"),
                    (R_TAIL, "AUTO 2/3 : rotations parties dans les derniers jours du mois (retour le mois suivant)"),
                    (R_ROTD, "AUTO 3/3 : rotations du mois = (departs + arrivees) / 2, vols additionnels compris")):
    an.merge_cells(f"B{base-1}:N{base-1}")
    A(f"B{base-1}", titre, bold=True, size=9, color="FFFFFF", fill="808080", halign="left")
    A(f"B{base}", "Route / type", bold=True, size=8, fill=GREY, border=box)
    for m, (lab, y, mm) in enumerate(MOIS):
        A(f"{mcol(m)}{base}", f"={S1R}${rc(0)}${REF0+m}", bold=True, size=8, fill=GREY, wrap=True, border=box)
    A(f"O{base}", "Route", bold=True, size=8, fill=GREY, border=box)
    A(f"P{base}", "Type", bold=True, size=8, fill=GREY, border=box)
    for i, (route, colr, white) in enumerate(ROUTES):
        for ti, t in enumerate(TYPES):
            r = drow(base, i, ti)
            A(f"B{r}", f"{route} - {t}", size=8, halign="left", border=box)
            A(f"O{r}", route, size=8, border=box)
            A(f"P{r}", t, size=8, border=box)
            for m, (lab, y, mm) in enumerate(MOIS):
                d, cal = mcol(m), cal_row(m)
                rel = m*STEP + rel_of(i, ti)
                if base == R_PDEP:
                    f = f'=SUMPRODUCT({occ_row(m)},LEN(INDEX({BIG},{rel},0)))'
                elif base == R_TAIL:
                    terms = []
                    for k in range(1, len(OFFSETS)):
                        for i2 in range(1, k+1):
                            cc = f'IFERROR(INDEX({BIG},{rel},WEEKDAY($K${cal}-{i2-1},2))&"","")'
                            terms.append(cnt_k(cc, k))
                    f = "=" + "+".join(terms)
                else:
                    prev = f"{d}{drow(R_TAIL,i,ti)-STEP*0}"
                    tail_prev = "0" if m == 0 else f"{mcol(m-1)}{drow(R_TAIL,i,ti)}"
                    dep_a = (f'SUMIFS({ADD_N},{ADD_R},"{route}",{ADD_T},"{t}",'
                             f'{ADD_D},">="&$J${cal},{ADD_D},"<="&$K${cal})')
                    arr_a = (f'SUMIFS({ADD_N},{ADD_R},"{route}",{ADD_T},"{t}",'
                             f'{ADD_RD},">="&$J${cal},{ADD_RD},"<="&$K${cal})')
                    f = (f'={d}{drow(R_PDEP,i,ti)}+({dep_a}+{arr_a}-{d}{drow(R_TAIL,i,ti)}+{tail_prev})/2')
                A(f"{d}{r}", f, size=8, fmt="0" if base != R_ROTD else ROT_FMT, border=box)
for col, w in {"A": 2, "B": 26, "O": 14, "P": 10, "Q": 3, "R": 60}.items():
    an.column_dimensions[col].width = w
for i in range(12):
    an.column_dimensions[gl(3+i)].width = 13
an.freeze_panes = "C3"
an.page_setup.orientation = "landscape"
an.page_setup.fitToWidth = 1
an.sheet_properties.pageSetUpPr.fitToPage = True
an.print_area = f"B1:O{R_TYP+len(MOIS)+1}"

wb.save(OUT)
print(f"ecrit -> {OUT}\n  Saisie {BLK0}-{DATA_LAST} | Vols add. {ROW_ADD0}-{ROW_ADD_END}")
