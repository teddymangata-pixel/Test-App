# -*- coding: utf-8 -*-
"""Genere le classeur 'Programme de vols mensuel' (feuille unique, pilotee par 2 filtres).

Semaine type : un type avion par jour J1-J7 (une route peut donc melanger plusieurs types).
Vols additionnels : table de vols dates qui s'ajoutent a la semaine type.
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

OUT = "/home/user/Test-App/Programme_vols_mensuel_2027-2028.xlsx"
F = "Arial"

# ---------------------------------------------------------------- donnees de reference
# Semaines types : valeurs d'EXEMPLE a ajuster dans la zone de parametres.
# Un type avion par jour ("" = pas d'operation) -> plusieurs types possibles sur une meme route.
# NOSRUN reprend l'exemple fourni dans la demande : 3/7 en A320 sur J1, J4, J6.
ROUTES = [
    # (route, [J1..J7], couleur, police blanche ?)
    ("CDGRUN", ["77W","778","77W","778","77W","778","77W"], "9DC3E6", False),
    ("CDGDZA", ["",   "778","",   "",   "778","",   "77W"], "2E75B6", True),
    ("BKKRUN", ["",   "",   "778","",   "",   "778",""   ], "FFE699", False),
    ("DZARUN", ["A320","A320","A320","A320","A320","",""], "F4B183", False),
    ("MRURUN", ["A320","",   "A320","",   "A320","", "A320"], "C6E0B4", False),
    ("NOSRUN", ["A320","",   "",   "A320","",   "A320",""], "FFD966", False),
    ("RUNTNR", ["",   "A320","",   "A320","",   "A320",""], "D9D2E9", False),
    ("JNBRUN", ["778","",   "",   "",   "778","",   ""   ], "A9D08E", False),
    ("CPTRUN", ["",   "",   "",   "778","",   "",   ""   ], "8FAADC", False),
    ("DIERUN", ["",   "A320","",  "",   "",   "A320",""  ], "FFC7CE", False),
    ("RRGRUN", ["A320","A320","A320","A320","A320","A320",""], "B7DEE8", False),
    ("TMMRUN", ["",   "",   "A320","",   "",   "",  "A320"], "D5A6BD", False),
]
TYPES = ["77W", "778", "A320"]
MOIS = [("avril 2027",2027,4),("mai 2027",2027,5),("juin 2027",2027,6),("juillet 2027",2027,7),
        ("aout 2027",2027,8),("septembre 2027",2027,9),("octobre 2027",2027,10),
        ("novembre 2027",2027,11),("decembre 2027",2027,12),("janvier 2028",2028,1),
        ("fevrier 2028",2028,2),("mars 2028",2028,3)]
JOURS = ["J1 - Lundi","J2 - Mardi","J3 - Mercredi","J4 - Jeudi","J5 - Vendredi","J6 - Samedi","J7 - Dimanche"]

# Vols additionnels : 3 lignes d'EXEMPLE montrant le format attendu (a supprimer).
VOLS_ADD = [
    (datetime.date(2027,4,7),  "NOSRUN", "A320", 1, "EXEMPLE a supprimer - vol supplementaire un mercredi (jour hors semaine type)"),
    (datetime.date(2027,4,10), "NOSRUN", "A320", 1, "EXEMPLE a supprimer - rotation doublee un samedi (jour deja opere)"),
    (datetime.date(2027,4,20), "CDGRUN", "778",  2, "EXEMPLE a supprimer - 2 rotations fret supplementaires"),
]

NB_WEEKS = 6            # un mois peut s'etaler sur 6 semaines calendaires
GRID_C0 = 3             # colonne C = J1
ROW_HDR = 7
ROW_W1 = 8                              # semaine 1 : ligne type = 8, ligne dates = 9
ROW_OCC = ROW_W1 + 2*NB_WEEKS           # 20 : occurrences de chaque jour dans le mois
ROW_NBW = ROW_OCC + 1                   # 21 : nb de semaines calendaires
ROW_REC_TITLE, ROW_REC_HDR, ROW_REC0 = 23, 24, 25
ROW_REC_TOT = ROW_REC0 + len(ROUTES)    # 37
ROW_PAR_TITLE, ROW_PAR_HDR, ROW_PAR0 = 41, 44, 45
ROW_PAR_END = ROW_PAR0 + len(ROUTES) - 1                # 56
ROW_ADD_TITLE, ROW_ADD_HDR, ROW_ADD0 = 59, 61, 62
NB_ADD_ROWS = 40
ROW_ADD_END = ROW_ADD0 + NB_ADD_ROWS - 1                # 101

wb = Workbook()
ws = wb.active
ws.title = "Programme de vols"
ws.sheet_view.showGridLines = False

# ---------------------------------------------------------------- styles
def cell(ref, value=None, bold=False, size=10, color="000000", fill=None,
         halign="center", valign="center", italic=False, wrap=False, fmt=None, border=None):
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
NAVY, LIGHT, GREY, BROWN = "1F3864", "D9E2F3", "F2F2F2", "833C0C"
RED = "C00000"
BLUE_IN = "0000FF"   # convention : saisies utilisateur en bleu

# plages nommees en dur (references utilisees partout)
PAR_R = f"$B${ROW_PAR0}:$B${ROW_PAR_END}"          # routes
PAR_D = f"$C${ROW_PAR0}:$I${ROW_PAR_END}"          # semaine type : un type par jour
ADD_D = f"$B${ROW_ADD0}:$B${ROW_ADD_END}"          # vols additionnels : dates
ADD_R = f"$C${ROW_ADD0}:$C${ROW_ADD_END}"          # routes
ADD_T = f"$D${ROW_ADD0}:$D${ROW_ADD_END}"          # types avion
ADD_N = f"$E${ROW_ADD0}:$E${ROW_ADD_END}"          # nb de rotations
ADD_K = f"$J${ROW_ADD0}:$J${ROW_ADD_END}"          # cle technique date|route

# ---------------------------------------------------------------- titre
ws.merge_cells("B1:I1")
cell("B1", "PROGRAMME DE VOLS MENSUEL - ANNEE D'EXPLOITATION AVRIL 2027 / MARS 2028",
     bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
ws.row_dimensions[1].height = 30
ws.merge_cells("B2:I2")
cell("B2", "Feuille unique : les deux filtres ci-dessous (Mois et Route) pilotent la grille. Le recapitulatif ne suit que le "
           "filtre Mois. Chaque route a une semaine type pouvant melanger plusieurs types avion (un type par jour), a laquelle "
           "s'ajoutent les vols additionnels dates saisis en bas de feuille. Tout est calcule par formule.",
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
cell("E4", "<-- liste deroulante : pilote la grille ET le recapitulatif", size=9, italic=True, halign="left", color="808080")
cell("E5", "<-- liste deroulante : pilote la grille uniquement", size=9, italic=True, halign="left", color="808080")
ws["C4"].comment = Comment("Liste deroulante : avril 2027 -> mars 2028.", "Modele")
ws["C5"].comment = Comment("Liste deroulante : les 12 routes exploitees.", "Modele")

# reperes calcules (colonnes N/O)
ws.merge_cells("N3:O3")
cell("N3", "REPERES CALCULES", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left")
reperes = [
    ("1er jour du mois",     f"=INDEX($O${ROW_PAR0}:$O${ROW_PAR_END},MATCH($C$4,$N${ROW_PAR0}:$N${ROW_PAR_END},0))", "DD/MM/YYYY"),
    ("Dernier jour du mois", "=EOMONTH($O$4,0)", "DD/MM/YYYY"),
    ("Lundi semaine 1",      "=$O$4-WEEKDAY($O$4,3)", "DD/MM/YYYY"),
    ("Type(s) avion",        f'=IFERROR(INDEX($C${ROW_REC0}:$C${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "General"),
    ("Rotations / semaine",  f'=IFERROR(INDEX($E${ROW_REC0}:$E${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "0"),
    ("Total rotations mois", f'=IFERROR(INDEX($G${ROW_REC0}:$G${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "0"),
]
for i, (lab, fml, fmt) in enumerate(reperes):
    r = 4 + i
    cell(f"N{r}", lab, size=9, halign="left", fill=GREY, border=box)
    cell(f"O{r}", fml, size=9, bold=True, fmt=fmt, border=box)
cell(f"N{4+len(reperes)}", "(route filtree)", size=8, italic=True, halign="left", color="808080")

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
    rt = ROW_W1 + 2*(k-1)   # ligne "type avion"
    rd = rt + 1             # ligne "dates reelles"
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
        # semaine type : type avion du jour j pour la route filtree
        st = f'IFERROR(INDEX({PAR_D},MATCH($C$5,{PAR_R},0),{j+1})&"","")'
        # vols additionnels de la route a cette date
        nrot = f"SUMIFS({ADD_N},{ADD_R},$C$5,{ADD_D},{d})"
        nrows = f"COUNTIFS({ADD_R},$C$5,{ADD_D},{d})"
        type1 = f'IFERROR(INDEX({ADD_T},MATCH({d}&"|"&$C$5,{ADD_K},0)),"")'
        lbl = f'IF({nrows}=1,IF({nrot}=1,{type1},{nrot}&"x"&{type1}),{nrot}&" vols")'
        cell(f"{col}{rt}", f'=IF({d}="","",TRIM({st}&IF({nrot}=0,""," (+"&{lbl}&")")))',
             bold=True, size=11, border=box)

# ligne occurrences + nb de semaines
cell(f"B{ROW_OCC}", "Occurrences du jour dans le mois", bold=True, size=9, halign="left", fill="E7E6E6", border=box)
for j in range(7):
    col = get_column_letter(GRID_C0+j)
    refs = ",".join(f"{col}{r}" for r in date_rows)
    cell(f"{col}{ROW_OCC}", f"=COUNT({refs})", bold=True, size=10, fill="E7E6E6", border=box)
cell(f"B{ROW_NBW}", "Nb de semaines calendaires du mois", bold=True, size=9, halign="left", fill="E7E6E6", border=box)
nbw = "+".join(f'IF(COUNT({get_column_letter(GRID_C0)}{r}:{get_column_letter(GRID_C0+6)}{r})>0,1,0)' for r in date_rows)
cell(f"C{ROW_NBW}", f"={nbw}", bold=True, size=10, fill="E7E6E6", border=box)
ws.merge_cells(f"D{ROW_NBW}:I{ROW_NBW}")
cell(f"D{ROW_NBW}", 'Lecture d\'une cellule : "778" = vol de la semaine type  |  "778 (+A320)" = semaine type + vol additionnel  |  '
                    '"(+A320)" = vol additionnel seul (bordure rouge). Semaines partielles incluses.',
     size=8, italic=True, color="595959", halign="left")

# ---------------------------------------------------------------- recapitulatif mensuel
ws.merge_cells(f"B{ROW_REC_TITLE}:L{ROW_REC_TITLE}")
cell(f"B{ROW_REC_TITLE}", "RECAPITULATIF MENSUEL - TOUTES ROUTES  (depend uniquement du filtre Mois)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
rec_hdr = ["Route", "Type(s) avion", "Semaine type (jours J1-J7)", "Rotations/semaine",
           "Nb de semaines/occurrences dans le mois", "Total rotations du mois",
           "dont semaine type", "dont vols additionnels", "Total 77W", "Total 778", "Total A320"]
for i, h in enumerate(rec_hdr):
    cell(f"{get_column_letter(2+i)}{ROW_REC_HDR}", h, bold=True, size=9, color="FFFFFF",
         fill="2F5597", wrap=True, border=box)
ws.row_dimensions[ROW_REC_HDR].height = 40

OCC = f"$C${ROW_OCC}:$I${ROW_OCC}"
for i in range(len(ROUTES)):
    r = ROW_REC0 + i
    p = ROW_PAR0 + i
    add_mois = f'SUMIFS({ADD_N},{ADD_R},$B{r},{ADD_D},">="&$O$4,{ADD_D},"<="&$O$5)'
    cell(f"B{r}", f"=$B${p}", bold=True, size=10, border=box)
    cell(f"C{r}", f"=$K${p}", size=9, border=box)                       # types + nb de jours
    jl = "&".join(f'IF($'+get_column_letter(3+j)+f'${p}<>"","J{j+1} ","")' for j in range(7))
    cell(f"D{r}", f'=IF(COUNTA($C${p}:$I${p})=0,"aucune operation",SUBSTITUTE(TRIM({jl})," ",", "))',
         size=9, border=box)
    cell(f"E{r}", f"=$J${p}", size=10, fmt="0", border=box)              # rotations / semaine
    cell(f"F{r}", f"=$C${ROW_NBW}", size=10, fmt="0", border=box)
    cell(f"G{r}", f"=H{r}+I{r}", bold=True, size=11, fmt="0", border=box)
    cell(f"H{r}", f'=SUMPRODUCT({OCC},--($C${p}:$I${p}<>""))', size=10, fmt="0", border=box)
    cell(f"I{r}", f"={add_mois}", size=10, fmt="0", border=box)
    for t, col in zip(TYPES, ("J", "K", "L")):
        cell(f"{col}{r}",
             f'=SUMPRODUCT({OCC},--($C${p}:$I${p}="{t}"))'
             f'+SUMIFS({ADD_N},{ADD_R},$B{r},{ADD_T},"{t}",{ADD_D},">="&$O$4,{ADD_D},"<="&$O$5)',
             size=9, fmt="0", border=box)

cell(f"B{ROW_REC_TOT}", "TOTAL GENERAL", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
for c in ("C", "D"):
    cell(f"{c}{ROW_REC_TOT}", "", fill=NAVY, border=box)
for c in ("E", "F", "G", "H", "I", "J", "K", "L"):
    f = f"=$C${ROW_NBW}" if c == "F" else f"=SUM({c}{ROW_REC0}:{c}{ROW_REC_TOT-1})"
    cell(f"{c}{ROW_REC_TOT}", f, bold=True, size=11 if c == "G" else 10,
         color="FFFFFF", fill=NAVY, fmt="0", border=box)

r = ROW_REC_TOT + 1
ws.merge_cells(f"B{r}:L{r+1}")
cell(f"B{r}", "Methode : pour chaque jour d'operation de la semaine type, le nombre d'occurrences de ce jour dans le mois est "
              "compte a partir des dates reelles de la grille (ligne \"Occurrences du jour dans le mois\", semaines partielles "
              "incluses). Total rotations du mois = semaine type + vols additionnels dates du mois. La colonne "
              "\"Nb de semaines/occurrences\" indique le nombre de semaines calendaires couvertes par le mois ; les colonnes "
              "Total 77W / 778 / A320 ventilent le total du mois par type avion (semaine type + vols additionnels).",
     size=8, italic=True, color="595959", halign="left", wrap=True)

# ---------------------------------------------------------------- zone de parametres : semaines types
ws.merge_cells(f"B{ROW_PAR_TITLE}:L{ROW_PAR_TITLE}")
cell(f"B{ROW_PAR_TITLE}", "1) SEMAINES TYPES PAR ROUTE (saisie unique, valable pour les 12 mois)",
     bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
ws.merge_cells(f"B{ROW_PAR_TITLE+1}:L{ROW_PAR_TITLE+2}")
cell(f"B{ROW_PAR_TITLE+1}", "Cellules en BLEU = saisie utilisateur. Pour chaque jour J1-J7, choisir dans la liste deroulante le TYPE AVION "
                            "opere ce jour-la (77W / 778 / A320), ou laisser la cellule vide s'il n'y a pas d'operation. Une meme route peut "
                            "donc utiliser plusieurs types avion dans la semaine (ex. CDGRUN : 77W les jours impairs, 778 les jours pairs). "
                            "Toute modification recalcule immediatement la grille et le recapitulatif, pour les 12 mois. Valeurs livrees = "
                            "exemples a ajuster (NOSRUN = 3/7 en A320 sur J1/J4/J6, conforme a l'exemple fourni).",
     size=8, italic=True, color=BROWN, halign="left", wrap=True)

par_hdr = ["Route"] + JOURS + ["Rot./sem. (auto)", "Type(s) avion (auto)", "Couleur"]
for i, h in enumerate(par_hdr):
    cell(f"{get_column_letter(2+i)}{ROW_PAR_HDR}", h, bold=True, size=9, color="FFFFFF", fill=BROWN, wrap=True, border=box)
cell(f"N{ROW_PAR_HDR}", "Mois (liste)", bold=True, size=9, color="FFFFFF", fill=BROWN, border=box)
cell(f"O{ROW_PAR_HDR}", "1er jour", bold=True, size=9, color="FFFFFF", fill=BROWN, border=box)
cell(f"P{ROW_PAR_HDR}", "Types avion", bold=True, size=9, color="FFFFFF", fill=BROWN, border=box)
ws.row_dimensions[ROW_PAR_HDR].height = 28

for i, (route, days, colr, white) in enumerate(ROUTES):
    r = ROW_PAR0 + i
    cell(f"B{r}", route, bold=True, size=10, halign="left", border=box)
    for j, t in enumerate(days):
        cell(f"{get_column_letter(3+j)}{r}", t if t else None, size=10, color=BLUE_IN, border=box)
    cell(f"J{r}", f"=COUNTA(C{r}:I{r})", size=10, bold=True, fmt="0", border=box)
    tl = "&".join(f'IF(COUNTIF(C{r}:I{r},"{t}")>0,"{t}("&COUNTIF(C{r}:I{r},"{t}")&") ","")' for t in TYPES)
    cell(f"K{r}", f'=IF(COUNTA(C{r}:I{r})=0,"-",SUBSTITUTE(TRIM({tl})," ",", "))', size=9, border=box)
    cell(f"L{r}", "", fill=colr, border=box)

for i, (lab, y, m) in enumerate(MOIS):
    r = ROW_PAR0 + i
    cell(f"N{r}", lab, size=9, halign="left", border=box)
    cell(f"O{r}", f"=DATE({y},{m},1)", size=9, fmt="DD/MM/YYYY", border=box)
for i, t in enumerate(TYPES):
    cell(f"P{ROW_PAR0+i}", t, size=9, border=box)

# ---------------------------------------------------------------- vols additionnels
ws.merge_cells(f"B{ROW_ADD_TITLE}:L{ROW_ADD_TITLE}")
cell(f"B{ROW_ADD_TITLE}", "2) VOLS ADDITIONNELS (en plus de la semaine type)",
     bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
ws.merge_cells(f"B{ROW_ADD_TITLE+1}:L{ROW_ADD_TITLE+1}")
cell(f"B{ROW_ADD_TITLE+1}", "Une ligne = un vol supplementaire date (charter, renfort saisonnier, fret...). Il s'ajoute a la semaine type de la "
                            "route, sur n'importe quel jour, y compris un jour deja opere ou un jour hors semaine type. Il apparait dans la "
                            "grille entre parentheses avec un + et une bordure rouge, et il est compte dans le recapitulatif du mois concerne.",
     size=8, italic=True, color=BROWN, halign="left", wrap=True)
ws.row_dimensions[ROW_ADD_TITLE+1].height = 24

add_hdr = ["Date du vol", "Route", "Type avion", "Nb de rotations", "Commentaire (libre)"]
for i, h in enumerate(add_hdr):
    cell(f"{get_column_letter(2+i)}{ROW_ADD_HDR}", h, bold=True, size=9, color="FFFFFF", fill=BROWN, wrap=True, border=box)
cell(f"J{ROW_ADD_HDR}", "Cle technique (auto)", bold=True, size=8, color="FFFFFF", fill="A6A6A6", wrap=True, border=box)

for i in range(NB_ADD_ROWS):
    r = ROW_ADD0 + i
    vals = VOLS_ADD[i] if i < len(VOLS_ADD) else None
    cell(f"B{r}", vals[0] if vals else None, size=10, color=BLUE_IN, fmt="DD/MM/YYYY", border=box)
    cell(f"C{r}", vals[1] if vals else None, size=10, color=BLUE_IN, border=box)
    cell(f"D{r}", vals[2] if vals else None, size=10, color=BLUE_IN, border=box)
    cell(f"E{r}", vals[3] if vals else None, size=10, color=BLUE_IN, fmt="0", border=box)
    cell(f"F{r}", vals[4] if vals else None, size=9, italic=True, color="595959", halign="left", border=box)
    cell(f"J{r}", f'=IF(OR($B{r}="",$C{r}=""),"",$B{r}&"|"&$C{r})', size=8, color="A6A6A6", border=box)
ws[f"B{ROW_ADD0}"].comment = Comment(
    "Saisir la date du vol (dans l'annee avril 2027 - mars 2028).\n"
    "Nb de rotations : 1 par defaut, 2 pour une double rotation le meme jour.\n"
    "Les 3 premieres lignes sont des exemples : les supprimer avant exploitation.", "Modele")

# ---------------------------------------------------------------- listes deroulantes
def add_dv(dv, rng):
    ws.add_data_validation(dv)
    dv.add(rng)

add_dv(DataValidation(type="list", formula1=f"=$N${ROW_PAR0}:$N${ROW_PAR_END}", allow_blank=False), "C4")
add_dv(DataValidation(type="list", formula1=f"=$B${ROW_PAR0}:$B${ROW_PAR_END}", allow_blank=False), "C5")
add_dv(DataValidation(type="list", formula1=f"=$P${ROW_PAR0}:$P${ROW_PAR0+len(TYPES)-1}", allow_blank=True,
                      error="Choisir un type avion dans la liste, ou laisser vide si la route n'opere pas ce jour-la.",
                      errorTitle="Type avion invalide"),
       f"C{ROW_PAR0}:I{ROW_PAR_END}")
add_dv(DataValidation(type="list", formula1=f"=$B${ROW_PAR0}:$B${ROW_PAR_END}", allow_blank=True),
       f"C{ROW_ADD0}:C{ROW_ADD_END}")
add_dv(DataValidation(type="list", formula1=f"=$P${ROW_PAR0}:$P${ROW_PAR0+len(TYPES)-1}", allow_blank=True),
       f"D{ROW_ADD0}:D{ROW_ADD_END}")
add_dv(DataValidation(type="whole", operator="between", formula1="1", formula2="20", allow_blank=True,
                      error="Nombre de rotations : entier entre 1 et 20.", errorTitle="Valeur invalide"),
       f"E{ROW_ADD0}:E{ROW_ADD_END}")
add_dv(DataValidation(type="date", operator="between", formula1="DATE(2027,4,1)", formula2="DATE(2028,3,31)",
                      allow_blank=True, errorStyle="warning", errorTitle="Date hors annee d'exploitation",
                      error="La date est hors de l'annee avril 2027 - mars 2028 : le vol ne sera visible dans aucun mois."),
       f"B{ROW_ADD0}:B{ROW_ADD_END}")

# ---------------------------------------------------------------- mise en forme conditionnelle
grid_ranges = " ".join(f"{get_column_letter(GRID_C0)}{r}:{get_column_letter(GRID_C0+6)}{r}" for r in type_rows)
c0, r0 = get_column_letter(GRID_C0), ROW_W1
# 1) marqueur "vol additionnel" : bordure rouge (prioritaire, n'interrompt pas les regles couleur)
ws.conditional_formatting.add(
    grid_ranges,
    FormulaRule(formula=[f'ISNUMBER(SEARCH("(+",{c0}{r0}))'],
                border=Border(left=Side(style="medium", color=RED), right=Side(style="medium", color=RED),
                              top=Side(style="medium", color=RED), bottom=Side(style="medium", color=RED)),
                stopIfTrue=False))
# 2) une couleur fixe par route (grille + recapitulatif + parametres)
for route, days, colr, white in ROUTES:
    fc = "FFFFFF" if white else "000000"
    ws.conditional_formatting.add(
        grid_ranges,
        FormulaRule(formula=[f'AND($C$5="{route}",{c0}{r0}<>"")'],
                    fill=PatternFill("solid", fgColor=colr), font=Font(name=F, size=11, bold=True, color=fc),
                    stopIfTrue=True))
    for rng, first in ((f"B{ROW_REC0}:B{ROW_REC_TOT-1}", ROW_REC0), (f"B{ROW_PAR0}:B{ROW_PAR_END}", ROW_PAR0)):
        ws.conditional_formatting.add(
            rng,
            FormulaRule(formula=[f'$B{first}="{route}"'],
                        fill=PatternFill("solid", fgColor=colr),
                        font=Font(name=F, size=10, bold=True, color=fc), stopIfTrue=True))

# ---------------------------------------------------------------- mise en page
widths = {"A": 2, "B": 30, "J": 14, "K": 20, "L": 12, "M": 3, "N": 22, "O": 14, "P": 12}
for j in range(7):
    widths[get_column_letter(GRID_C0+j)] = 15
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = "C8"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_area = f"B1:L{ROW_REC_TOT+2}"

wb.save(OUT)
print("ecrit ->", OUT)
