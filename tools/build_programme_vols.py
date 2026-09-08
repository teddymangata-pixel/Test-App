# -*- coding: utf-8 -*-
"""Genere le classeur 'Programme de vols mensuel' (feuille unique, pilotee par 2 filtres).

Saisie : une table de PROGRAMMES HEBDOMADAIRES. Une ligne = une route + un mois + le type avion
opere chaque jour J1-J7. Plusieurs lignes pour un meme couple (route, mois) = plusieurs vols le
meme jour, avec des types differents si besoin. Le mois "TOUS" sert de programme par defaut,
utilise pour les mois qui n'ont pas de ligne specifique.
Vols additionnels : table de vols dates qui s'ajoutent au programme du mois.
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
_ = None
TOUS = "TOUS"

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

# Programmes livres = EXEMPLES a ajuster. (route, mois, [type J1..J7])
# Mois "TOUS" = programme par defaut, applique aux mois sans ligne specifique.
PROGRAMMES = [
    ("CDGRUN", TOUS, ["77W","77W","77W","77W","77W","77W","77W"]),
    ("CDGRUN", TOUS, ["778",_,_,_,"778",_,_]),
    ("CDGRUN", TOUS, ["A320",_,_,_,_,_,_]),
    ("CDGDZA", TOUS, [_,"778",_,_,"778",_,"77W"]),
    ("BKKRUN", TOUS, [_,_,"778",_,_,"778",_]),
    ("DZARUN", TOUS, ["A320","A320","A320","A320","A320",_,_]),
    ("DZARUN", TOUS, [_,_,"778",_,_,_,_]),
    ("MRURUN", TOUS, ["A320",_,"A320",_,"A320",_,"A320"]),
    ("NOSRUN", TOUS, ["A320",_,_,"A320",_,"A320",_]),
    ("RUNTNR", TOUS, [_,"A320",_,"A320",_,"A320",_]),
    ("JNBRUN", TOUS, ["778",_,_,_,"778",_,_]),
    ("CPTRUN", TOUS, [_,_,_,"778",_,_,_]),
    ("DIERUN", TOUS, [_,"A320",_,_,_,"A320",_]),
    ("RRGRUN", TOUS, ["A320","A320","A320","A320","A320","A320",_]),
    ("RRGRUN", TOUS, [_,_,_,_,"A320",_,_]),
    ("RRGRUN", TOUS, [_,_,_,_,"A320",_,_]),
    ("TMMRUN", TOUS, [_,_,"A320",_,_,_,"A320"]),
    # --- exemples de programmes specifiques a un mois (saisonnalite) ---
    ("NOSRUN", "juillet 2027",  ["A320",_,"A320","A320",_,"A320",_]),
    ("NOSRUN", "decembre 2027", ["A320","A320",_,"A320",_,"A320","A320"]),
    ("CDGRUN", "aout 2027",     ["77W","77W","77W","77W","77W","77W","77W"]),
    ("CDGRUN", "aout 2027",     ["778","778","778","778","778","778","778"]),
]

# Vols additionnels : 3 lignes d'EXEMPLE montrant le format attendu (a supprimer).
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
ROW_PROG_TITLE, ROW_PROG_HDR, ROW_PROG0 = 41, 45, 46
NB_PROG_ROWS = 80
ROW_PROG_END = ROW_PROG0 + NB_PROG_ROWS - 1       # 125
ROW_ADD_TITLE, ROW_ADD_HDR, ROW_ADD0 = 128, 132, 133
NB_ADD_ROWS = 40
ROW_ADD_END = ROW_ADD0 + NB_ADD_ROWS - 1          # 172
# zone auto a droite (colonnes N..U)
R_REP0, R_HELP_TITLE, R_HELP_HDR, R_HELP_TXT, R_HELP_T0, R_HELP_NB = 4, 12, 13, 14, 15, 18
R_MAT_TITLE, R_MAT_HDR = 20, 21
MAT_BASE = {"77W": 23, "778": 37, "A320": 51}     # 1re ligne de chaque bloc (12 routes)
HELP_C0 = 15                                      # colonne O
REF0 = 4                                          # listes de reference : lignes 4..

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
NAVY, LIGHT, GREY, BROWN, RED, GREEN = "1F3864", "D9E2F3", "F2F2F2", "833C0C", "C00000", "E2EFDA"
BLUE_IN = "0000FF"

# plages de la table des programmes et des vols additionnels
PR_R = f"$B${ROW_PROG0}:$B${ROW_PROG_END}"
PR_M = f"$C${ROW_PROG0}:$C${ROW_PROG_END}"
PR_D = [f"${get_column_letter(4+j)}${ROW_PROG0}:${get_column_letter(4+j)}${ROW_PROG_END}" for j in range(7)]
ADD_D = f"$B${ROW_ADD0}:$B${ROW_ADD_END}"
ADD_R = f"$C${ROW_ADD0}:$C${ROW_ADD_END}"
ADD_T = f"$D${ROW_ADD0}:$D${ROW_ADD_END}"
ADD_N = f"$E${ROW_ADD0}:$E${ROW_ADD_END}"
ADD_K = f"$H${ROW_ADD0}:$H${ROW_ADD_END}"
OCC = f"$C${ROW_OCC}:$I${ROW_OCC}"
mat = lambda t, i: f"$O${MAT_BASE[t]+i}:$U${MAT_BASE[t]+i}"     # ligne du bloc auto
matc = lambda t, i, j: f"${get_column_letter(HELP_C0+j)}${MAT_BASE[t]+i}"

# ---------------------------------------------------------------- titre
ws.merge_cells("B1:I1")
cell("B1", "PROGRAMME DE VOLS MENSUEL - ANNEE D'EXPLOITATION AVRIL 2027 / MARS 2028",
     bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
ws.row_dimensions[1].height = 30
ws.merge_cells("B2:I2")
cell("B2", "Feuille unique : les deux filtres ci-dessous (Mois et Route) pilotent la grille, le recapitulatif ne suit que le filtre "
           "Mois. Le programme hebdomadaire se saisit MOIS PAR MOIS dans la table du bas : une ligne = une route + un mois + le type "
           "avion opere chaque jour. Plusieurs lignes pour un meme mois = plusieurs vols le meme jour. Tout est calcule par formule.",
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
cell("E4", "<-- liste deroulante : pilote la grille, le recapitulatif ET le programme applique", size=9, italic=True, halign="left", color="808080")
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
    ("Programme applique",   f'=IF(COUNTIFS({PR_R},$C$5,{PR_M},$C$4)>0,$C$4,"{TOUS}")', "General"),
    ("Type(s) avion",        f'=IFERROR(INDEX($C${ROW_REC0}:$C${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "General"),
    ("Rotations / semaine",  f'=IFERROR(INDEX($E${ROW_REC0}:$E${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "0"),
    ("Total rotations mois", f'=IFERROR(INDEX($G${ROW_REC0}:$G${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "0"),
]
for i, (lab, fml, fmt) in enumerate(reperes):
    r = R_REP0 + i
    cell(f"N{r}", lab, size=9, halign="left", fill=GREY, border=box)
    cell(f"O{r}", fml, size=9, bold=True, fmt=fmt, border=box)

# ---------------------------------------------------------------- bloc auto : programme de la route filtree
hc = lambda j: get_column_letter(HELP_C0+j)
ws.merge_cells(f"N{R_HELP_TITLE}:{hc(6)}{R_HELP_TITLE}")
cell(f"N{R_HELP_TITLE}", "PROGRAMME DE LA ROUTE FILTREE POUR LE MOIS FILTRE (auto - alimente la grille)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"N{R_HELP_HDR}", "Jour", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_HELP_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
labels = [(R_HELP_TXT, "Vols du jour (affichage grille)"), (R_HELP_T0, "dont 77W"),
          (R_HELP_T0+1, "dont 778"), (R_HELP_T0+2, "dont A320"), (R_HELP_NB, "Nb de vols")]
for r, lab in labels:
    cell(f"N{r}", lab, bold=(r in (R_HELP_TXT, R_HELP_NB)), size=8, halign="left", fill=GREY, border=box)
for j in range(7):
    c = hc(j)
    for k, t in enumerate(TYPES):
        cell(f"{c}{R_HELP_T0+k}", f'=COUNTIFS({PR_R},$C$5,{PR_M},$O$7,{PR_D[j]},"{t}")',
             size=9, fmt="0", border=box)
    a, b, d = f"{c}${R_HELP_T0}", f"{c}${R_HELP_T0+1}", f"{c}${R_HELP_T0+2}"
    cell(f"{c}{R_HELP_TXT}",
         f'=IF({a}=0,"",IF({a}=1,"77W",{a}&"x77W"))'
         f'&IF({b}=0,"",IF({a}>0,"/","")&IF({b}=1,"778",{b}&"x778"))'
         f'&IF({d}=0,"",IF({a}+{b}>0,"/","")&IF({d}=1,"A320",{d}&"xA320"))',
         bold=True, size=9, border=box)
    cell(f"{c}{R_HELP_NB}", f"=SUM({c}{R_HELP_T0}:{c}{R_HELP_T0+2})", bold=True, size=9, fmt="0", border=box)

# ---------------------------------------------------------------- bloc auto : nb de vols par route et par jour
ws.merge_cells(f"N{R_MAT_TITLE}:{hc(6)}{R_MAT_TITLE}")
cell(f"N{R_MAT_TITLE}", "NB DE VOLS PAR ROUTE ET PAR JOUR POUR LE MOIS FILTRE (auto - alimente le recapitulatif)",
     bold=True, size=9, color="FFFFFF", fill=NAVY, halign="left")
cell(f"N{R_MAT_HDR}", "Route", bold=True, size=8, fill=GREY, border=box)
for j in range(7):
    cell(f"{hc(j)}{R_MAT_HDR}", f"J{j+1}", bold=True, size=8, fill=GREY, border=box)
for t in TYPES:
    base = MAT_BASE[t]
    cell(f"N{base-1}", f"Vols en {t}", bold=True, size=8, halign="left", fill="E7E6E6", border=box)
    ws.merge_cells(f"N{base-1}:{hc(6)}{base-1}")
    for i in range(len(ROUTES)):
        rec = ROW_REC0 + i
        cell(f"N{base+i}", f"=$B{rec}", size=8, halign="left", border=box)
        for j in range(7):
            cell(f"{hc(j)}{base+i}", f'=COUNTIFS({PR_R},$B{rec},{PR_M},$H{rec},{PR_D[j]},"{t}")',
                 size=8, fmt="0", border=box)

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
    rt = ROW_W1 + 2*(k-1)
    rd = rt + 1
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
ws.merge_cells(f"B{ROW_REC_TITLE}:M{ROW_REC_TITLE}")
cell(f"B{ROW_REC_TITLE}", "RECAPITULATIF MENSUEL - TOUTES ROUTES  (depend uniquement du filtre Mois)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
rec_hdr = ["Route", "Type(s) avion", "Semaine type du mois (jours J1-J7)", "Rotations/semaine",
           "Nb de semaines/occurrences dans le mois", "Total rotations du mois",
           "Programme applique", "dont programme hebdo", "dont vols additionnels",
           "Total 77W", "Total 778", "Total A320"]
for i, h in enumerate(rec_hdr):
    cell(f"{get_column_letter(2+i)}{ROW_REC_HDR}", h, bold=True, size=9, color="FFFFFF",
         fill="2F5597", wrap=True, border=box)
ws.row_dimensions[ROW_REC_HDR].height = 44

for i in range(len(ROUTES)):
    r = ROW_REC0 + i
    add = lambda extra="": f'SUMIFS({ADD_N},{ADD_R},$B{r}{extra},{ADD_D},">="&$O$4,{ADD_D},"<="&$O$5)'
    sums = {t: f"SUM({mat(t,i)})" for t in TYPES}
    prods = {t: f"SUMPRODUCT({OCC},{mat(t,i)})" for t in TYPES}
    cell(f"B{r}", f"=$Z${REF0+i}", bold=True, size=10, border=box)
    tl = "&".join(f'IF({sums[t]}>0,"{t}("&{sums[t]}&") ","")' for t in TYPES)
    cell(f"C{r}", f'=IF(E{r}=0,"-",SUBSTITUTE(TRIM({tl})," ",", "))', size=9, wrap=True, border=box)
    jl = "&".join(f'IF({matc("77W",i,j)}+{matc("778",i,j)}+{matc("A320",i,j)}>0,"J{j+1} ","")' for j in range(7))
    cell(f"D{r}", f'=IF(E{r}=0,"aucune operation",SUBSTITUTE(TRIM({jl})," ",", "))', size=9, wrap=True, border=box)
    cell(f"E{r}", "=" + "+".join(sums[t] for t in TYPES), size=10, fmt="0", border=box)
    cell(f"F{r}", f"=$C${ROW_NBW}", size=10, fmt="0", border=box)
    cell(f"G{r}", f"=I{r}+J{r}", bold=True, size=11, fmt="0", border=box)
    cell(f"H{r}", f'=IF(COUNTIFS({PR_R},$B{r},{PR_M},$C$4)>0,$C$4,"{TOUS}")', size=8, italic=True, border=box)
    cell(f"I{r}", "=" + "+".join(prods[t] for t in TYPES), size=10, fmt="0", border=box)
    cell(f"J{r}", f"={add()}", size=10, fmt="0", border=box)
    for t, col in zip(TYPES, ("K", "L", "M")):
        cell(f"{col}{r}", f'={prods[t]}+{add(f",{ADD_T},"+chr(34)+t+chr(34))}', size=9, fmt="0", border=box)
    ws.row_dimensions[r].height = 24

cell(f"B{ROW_REC_TOT}", "TOTAL GENERAL", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
for c in ("C", "D", "H"):
    cell(f"{c}{ROW_REC_TOT}", "", fill=NAVY, border=box)
for c in ("E", "F", "G", "I", "J", "K", "L", "M"):
    f = f"=$C${ROW_NBW}" if c == "F" else f"=SUM({c}{ROW_REC0}:{c}{ROW_REC_TOT-1})"
    cell(f"{c}{ROW_REC_TOT}", f, bold=True, size=11 if c == "G" else 10,
         color="FFFFFF", fill=NAVY, fmt="0", border=box)

r = ROW_REC_TOT + 1
ws.merge_cells(f"B{r}:M{r+1}")
cell(f"B{r}", "Methode : le programme retenu pour chaque route est celui du mois filtre s'il existe une ligne pour ce mois, sinon "
              "le programme \"TOUS\" (colonne Programme applique). Pour chaque jour opere, le nombre d'occurrences de ce jour dans "
              "le mois est compte a partir des dates reelles de la grille (ligne \"Occurrences du jour dans le mois\", semaines "
              "partielles incluses), chaque vol de la journee etant compte separement. Total rotations du mois = programme hebdo "
              "+ vols additionnels dates du mois ; les colonnes Total 77W / 778 / A320 le ventilent par type avion.",
     size=8, italic=True, color="595959", halign="left", wrap=True)

# ---------------------------------------------------------------- 1) programmes hebdomadaires
ws.merge_cells(f"B{ROW_PROG_TITLE}:L{ROW_PROG_TITLE}")
cell(f"B{ROW_PROG_TITLE}", "1) PROGRAMMES HEBDOMADAIRES PAR ROUTE ET PAR MOIS (saisie principale)",
     bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
ws.merge_cells(f"B{ROW_PROG_TITLE+1}:L{ROW_PROG_TITLE+3}")
cell(f"B{ROW_PROG_TITLE+1}",
     "Cellules en BLEU = saisie utilisateur. UNE LIGNE = une route + un mois + le type avion opere chaque jour J1-J7 "
     "(cellule vide = pas de vol ce jour-la). La frequence hebdomadaire peut donc etre differente pour chaque mois : il suffit "
     "d'ajouter une ligne au mois concerne. AJOUTER PLUSIEURS LIGNES pour un meme couple route/mois = plusieurs vols le meme jour, "
     "avec le meme type ou des types differents.  >>> Mois = \"TOUS\" : programme par defaut, applique uniquement aux mois qui n'ont "
     "AUCUNE ligne specifique pour cette route ; des qu'une ligne existe pour un mois, elle remplace le programme \"TOUS\" pour ce "
     "mois. La colonne Statut indique en vert les lignes reellement appliquees au mois filtre. Astuce : utiliser les fleches de "
     "filtre de l'en-tete pour n'afficher qu'un mois ou qu'une route, et copier/coller une ligne pour creer le programme d'un "
     "nouveau mois. Lignes livrees = exemples a ajuster.",
     size=8, italic=True, color=BROWN, halign="left", wrap=True)
ws.row_dimensions[ROW_PROG_TITLE+1].height = 58

prog_hdr = ["Route", "Mois du programme"] + JOURS + ["Rot./sem. (auto)", "Statut (auto)"]
for i, h in enumerate(prog_hdr):
    cell(f"{get_column_letter(2+i)}{ROW_PROG_HDR}", h, bold=True, size=9, color="FFFFFF", fill=BROWN, wrap=True, border=box)
ws.row_dimensions[ROW_PROG_HDR].height = 28

for i in range(NB_PROG_ROWS):
    r = ROW_PROG0 + i
    p = PROGRAMMES[i] if i < len(PROGRAMMES) else None
    cell(f"B{r}", p[0] if p else None, bold=True, size=10, halign="left", color=BLUE_IN, border=box)
    cell(f"C{r}", p[1] if p else None, size=9, color=BLUE_IN, border=box)
    for j in range(7):
        cell(f"{get_column_letter(4+j)}{r}", (p[2][j] if p else None), size=10, color=BLUE_IN, border=box)
    cell(f"K{r}", f'=IF($B{r}="","",COUNTA(D{r}:J{r}))', size=10, bold=True, fmt="0", border=box)
    cell(f"L{r}", f'=IF($B{r}="","",IF(OR($C{r}=$C$4,AND($C{r}="{TOUS}",'
                  f'COUNTIFS({PR_R},$B{r},{PR_M},$C$4)=0)),"APPLIQUE",""))',
         size=8, bold=True, border=box)
ws[f"B{ROW_PROG0}"].comment = Comment(
    "Une ligne par route et par mois.\n"
    "Mois = TOUS -> programme par defaut des mois non renseignes.\n"
    "Plusieurs lignes pour un meme route/mois = plusieurs vols le meme jour.\n"
    "Les lignes livrees sont des exemples : les adapter.", "Modele")
ws.auto_filter.ref = f"B{ROW_PROG_HDR}:L{ROW_PROG_END}"

# ---------------------------------------------------------------- 2) vols additionnels
ws.merge_cells(f"B{ROW_ADD_TITLE}:L{ROW_ADD_TITLE}")
cell(f"B{ROW_ADD_TITLE}", "2) VOLS ADDITIONNELS (en plus du programme hebdomadaire)",
     bold=True, size=11, color="FFFFFF", fill=BROWN, halign="left")
ws.merge_cells(f"B{ROW_ADD_TITLE+1}:L{ROW_ADD_TITLE+2}")
cell(f"B{ROW_ADD_TITLE+1}", "Une ligne = un vol supplementaire date (charter, renfort ponctuel, fret...). Il s'ajoute au programme du mois, "
                            "sur n'importe quel jour, y compris un jour deja opere ou un jour hors programme. Il apparait dans la grille "
                            "entre parentheses avec un + et une bordure rouge, et il est compte dans le recapitulatif du mois concerne. "
                            "Pour une modification recurrente sur tout un mois, preferer une ligne de programme ci-dessus.",
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

# ---------------------------------------------------------------- listes de reference (W..AA)
for col, h in (("W", "Mois (filtre)"), ("X", "1er jour"), ("Y", "Types avion"), ("Z", "Routes"),
               ("AA", "Mois (saisie programme)")):
    cell(f"{col}{REF0-1}", h, bold=True, size=8, color="FFFFFF", fill="808080", wrap=True, border=box)
for i, (lab, y, m) in enumerate(MOIS):
    cell(f"W{REF0+i}", lab, size=8, halign="left", border=box)
    cell(f"X{REF0+i}", f"=DATE({y},{m},1)", size=8, fmt="DD/MM/YYYY", border=box)
    cell(f"AA{REF0+1+i}", lab, size=8, halign="left", border=box)
cell(f"AA{REF0}", TOUS, size=8, bold=True, halign="left", border=box)
for i, t in enumerate(TYPES):
    cell(f"Y{REF0+i}", t, size=8, border=box)
for i, (route, colr, white) in enumerate(ROUTES):
    cell(f"Z{REF0+i}", route, size=8, border=box)

# ---------------------------------------------------------------- listes deroulantes
def add_dv(dv, rng):
    ws.add_data_validation(dv); dv.add(rng)

L_MOIS, L_TYPE = f"=$W${REF0}:$W${REF0+11}", f"=$Y${REF0}:$Y${REF0+len(TYPES)-1}"
L_ROUTE, L_PMOIS = f"=$Z${REF0}:$Z${REF0+11}", f"=$AA${REF0}:$AA${REF0+12}"
add_dv(DataValidation(type="list", formula1=L_MOIS, allow_blank=False), "C4")
add_dv(DataValidation(type="list", formula1=L_ROUTE, allow_blank=False), "C5")
add_dv(DataValidation(type="list", formula1=L_ROUTE, allow_blank=True,
                      error="Choisir une route dans la liste.", errorTitle="Route inconnue"),
       f"B{ROW_PROG0}:B{ROW_PROG_END}")
add_dv(DataValidation(type="list", formula1=L_PMOIS, allow_blank=True,
                      error="Choisir un mois, ou TOUS pour le programme par defaut.", errorTitle="Mois invalide"),
       f"C{ROW_PROG0}:C{ROW_PROG_END}")
add_dv(DataValidation(type="list", formula1=L_TYPE, allow_blank=True,
                      error="Choisir un type avion, ou laisser vide s'il n'y a pas de vol ce jour-la.",
                      errorTitle="Type avion invalide"),
       f"D{ROW_PROG0}:J{ROW_PROG_END}")
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
for route, colr, white in ROUTES:
    fc = "FFFFFF" if white else "000000"
    ws.conditional_formatting.add(
        grid_ranges,
        FormulaRule(formula=[f'AND($C$5="{route}",{c0}{ROW_W1}<>"")'],
                    fill=PatternFill("solid", fgColor=colr), font=Font(name=F, size=10, bold=True, color=fc),
                    stopIfTrue=True))
    for rng, first in ((f"B{ROW_REC0}:B{ROW_REC_TOT-1}", ROW_REC0), (f"B{ROW_PROG0}:B{ROW_PROG_END}", ROW_PROG0)):
        ws.conditional_formatting.add(
            rng, FormulaRule(formula=[f'$B{first}="{route}"'],
                             fill=PatternFill("solid", fgColor=colr),
                             font=Font(name=F, size=10, bold=True, color=fc), stopIfTrue=True))
# lignes de programme appliquees au mois filtre : vertes ; les autres : grisees
ws.conditional_formatting.add(
    f"C{ROW_PROG0}:L{ROW_PROG_END}",
    FormulaRule(formula=[f'$L{ROW_PROG0}="APPLIQUE"'], fill=PatternFill("solid", fgColor=GREEN), stopIfTrue=True))
ws.conditional_formatting.add(
    f"C{ROW_PROG0}:L{ROW_PROG_END}",
    FormulaRule(formula=[f'AND($B{ROW_PROG0}<>"",$L{ROW_PROG0}<>"APPLIQUE")'],
                font=Font(name=F, size=10, color="A6A6A6"), stopIfTrue=True))

# ---------------------------------------------------------------- mise en page
widths = {"A": 2, "B": 30, "J": 18, "K": 18, "L": 14, "M": 12, "N": 26, "V": 3,
          "W": 18, "X": 12, "Y": 10, "Z": 10, "AA": 20}
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
ws.print_area = f"B1:M{ROW_REC_TOT+2}"

wb.save(OUT)
print("ecrit ->", OUT)
