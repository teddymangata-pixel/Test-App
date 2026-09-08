# -*- coding: utf-8 -*-
"""Genere le classeur 'Programme de vols mensuel' (feuille unique, pilotee par 2 filtres)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

OUT = "/home/user/Test-App/Programme_vols_mensuel_2027-2028.xlsx"
F = "Arial"

# ---------------------------------------------------------------- donnees de reference
# Semaines types : valeurs d'EXEMPLE a ajuster par l'utilisateur dans la zone de parametres.
# NOSRUN correspond a l'exemple fourni dans la demande : 3/7 en A320 sur J1, J4, J6.
ROUTES = [
    # (route, type avion, jours J1..J7, couleur, police blanche ?)
    ("CDGRUN", "77W",  [1,1,1,1,1,1,1], "9DC3E6", False),
    ("CDGDZA", "778",  [0,1,0,0,1,0,1], "2E75B6", True),
    ("BKKRUN", "778",  [0,0,1,0,0,1,0], "FFE699", False),
    ("DZARUN", "A320", [1,1,1,1,1,0,0], "F4B183", False),
    ("MRURUN", "A320", [1,0,1,0,1,0,1], "C6E0B4", False),
    ("NOSRUN", "A320", [1,0,0,1,0,1,0], "FFD966", False),
    ("RUNTNR", "A320", [0,1,0,1,0,1,0], "D9D2E9", False),
    ("JNBRUN", "778",  [1,0,0,0,1,0,0], "A9D08E", False),
    ("CPTRUN", "778",  [0,0,0,1,0,0,0], "8FAADC", False),
    ("DIERUN", "A320", [0,1,0,0,0,1,0], "FFC7CE", False),
    ("RRGRUN", "A320", [1,1,1,1,1,1,0], "B7DEE8", False),
    ("TMMRUN", "A320", [0,0,1,0,0,0,1], "D5A6BD", False),
]
TYPES = ["77W", "778", "A320"]
MOIS = [("avril 2027",2027,4),("mai 2027",2027,5),("juin 2027",2027,6),("juillet 2027",2027,7),
        ("aout 2027",2027,8),("septembre 2027",2027,9),("octobre 2027",2027,10),
        ("novembre 2027",2027,11),("decembre 2027",2027,12),("janvier 2028",2028,1),
        ("fevrier 2028",2028,2),("mars 2028",2028,3)]
JOURS = ["J1 - Lundi","J2 - Mardi","J3 - Mercredi","J4 - Jeudi","J5 - Vendredi","J6 - Samedi","J7 - Dimanche"]

NB_WEEKS = 6            # un mois peut s'etaler sur 6 semaines calendaires
GRID_C0 = 3             # colonne C = J1
ROW_HDR = 7
ROW_W1 = 8              # semaine 1 : ligne type = 8, ligne dates = 9
ROW_OCC = ROW_W1 + 2*NB_WEEKS          # 20 : occurrences de chaque jour dans le mois
ROW_NBW = ROW_OCC + 1                  # 21 : nb de semaines calendaires
ROW_REC_TITLE = 23
ROW_REC_HDR = 24
ROW_REC0 = 25                          # 25..36
ROW_REC_TOT = ROW_REC0 + len(ROUTES)   # 37
ROW_PAR_TITLE = 40
ROW_PAR_HDR = 42
ROW_PAR0 = 43                          # 43..54
ROW_PAR_END = ROW_PAR0 + len(ROUTES) - 1

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
NAVY, LIGHT, GREY = "1F3864", "D9E2F3", "F2F2F2"
BLUE_IN = "0000FF"   # convention : saisies utilisateur en bleu

# ---------------------------------------------------------------- titre
ws.merge_cells("B1:I1")
cell("B1", "PROGRAMME DE VOLS MENSUEL - ANNEE D'EXPLOITATION AVRIL 2027 / MARS 2028",
     bold=True, size=16, color="FFFFFF", fill=NAVY, halign="left")
ws.row_dimensions[1].height = 30
ws.merge_cells("B2:I2")
cell("B2", "Feuille unique : les deux filtres ci-dessous (Mois et Route) pilotent la grille. "
           "Le recapitulatif ne suit que le filtre Mois. Toutes les valeurs sont calculees par formule "
           "a partir des semaines types saisies dans la zone de parametres (bas de feuille).",
     size=9, italic=True, halign="left", wrap=True, color="404040")
ws.row_dimensions[2].height = 26

# ---------------------------------------------------------------- filtres
cell("B4", "FILTRE  MOIS", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
ws.merge_cells("C4:D4")
cell("C4", "avril 2027", bold=True, size=12, color=BLUE_IN, fill="FFF2CC", border=Border(left=med, right=med, top=med, bottom=med))
cell("B5", "FILTRE  ROUTE", bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
ws.merge_cells("C5:D5")
cell("C5", "NOSRUN", bold=True, size=12, color=BLUE_IN, fill="FFF2CC", border=Border(left=med, right=med, top=med, bottom=med))
cell("E4", "<-- listes deroulantes", size=9, italic=True, halign="left", color="808080")
cell("E5", "<-- listes deroulantes", size=9, italic=True, halign="left", color="808080")
ws["C4"].comment = Comment("Liste deroulante : avril 2027 -> mars 2028.\nPilote la grille ET le recapitulatif.", "Modele")
ws["C5"].comment = Comment("Liste deroulante : 12 routes.\nPilote uniquement la grille.", "Modele")

# reperes calcules (colonnes K/L)
PAR_R = f"$B${ROW_PAR0}:$B${ROW_PAR_END}"
PAR_T = f"$C${ROW_PAR0}:$C${ROW_PAR_END}"
PAR_F = f"$D${ROW_PAR0}:$J${ROW_PAR_END}"
cell("K3", "REPERES CALCULES", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left")
ws.merge_cells("K3:L3")
reperes = [
    ("1er jour du mois",                        f"=INDEX($O${ROW_PAR0}:$O${ROW_PAR_END},MATCH($C$4,$N${ROW_PAR0}:$N${ROW_PAR_END},0))", "DD/MM/YYYY"),
    ("Dernier jour du mois",                    "=EOMONTH($L$4,0)", "DD/MM/YYYY"),
    ("Lundi de la semaine 1",                   "=$L$4-WEEKDAY($L$4,3)", "DD/MM/YYYY"),
    ("Type avion (route filtree)",              f'=IFERROR(INDEX({PAR_T},MATCH($C$5,{PAR_R},0)),"")', "General"),
    ("Rotations / semaine (route filtree)",     f'=IFERROR(SUM(INDEX({PAR_F},MATCH($C$5,{PAR_R},0),0)),"")', "0"),
    ("Total rotations du mois (route filtree)", f'=IFERROR(INDEX($G${ROW_REC0}:$G${ROW_REC_TOT-1},MATCH($C$5,$B${ROW_REC0}:$B${ROW_REC_TOT-1},0)),"")', "0"),
]
for i, (lab, fml, fmt) in enumerate(reperes):
    r = 4 + i
    cell(f"K{r}", lab, size=9, halign="left", fill=GREY, border=box)
    cell(f"L{r}", fml, size=9, bold=True, fmt=fmt, border=box)

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
        cell(f"{col}{rd}",
             f'=IF(OR($L$6+{off}<$L$4,$L$6+{off}>$L$5),"",$L$6+{off})',
             size=8, italic=True, color="595959", fmt="DD/MM", border=box)
        cell(f"{col}{rt}",
             f'=IF({col}{rd}="","",IFERROR(IF(INDEX({PAR_F},MATCH($C$5,{PAR_R},0),{j+1})=1,'
             f'INDEX({PAR_T},MATCH($C$5,{PAR_R},0)),""),""))',
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
cell(f"D{ROW_NBW}", "Semaines partielles de debut / fin de mois incluses : seuls les jours reellement compris "
                    "dans le mois sont dates, donc comptes.", size=8, italic=True, color="595959", halign="left")

# ---------------------------------------------------------------- recapitulatif mensuel
ws.merge_cells(f"B{ROW_REC_TITLE}:G{ROW_REC_TITLE}")
cell(f"B{ROW_REC_TITLE}", "RECAPITULATIF MENSUEL - TOUTES ROUTES  (depend uniquement du filtre Mois)",
     bold=True, size=11, color="FFFFFF", fill=NAVY, halign="left")
rec_hdr = ["Route", "Type avion", "Semaine type (jours J1-J7)", "Rotations/semaine",
           "Nb de semaines/occurrences dans le mois", "Total rotations du mois"]
for i, h in enumerate(rec_hdr):
    cell(f"{get_column_letter(2+i)}{ROW_REC_HDR}", h, bold=True, size=10, color="FFFFFF",
         fill="2F5597", wrap=True, border=box)
ws.row_dimensions[ROW_REC_HDR].height = 32

for i in range(len(ROUTES)):
    r = ROW_REC0 + i
    p = ROW_PAR0 + i
    cell(f"B{r}", f"=$B${p}", bold=True, size=10, border=box)
    cell(f"C{r}", f"=$C${p}", size=10, border=box)
    jl = "&".join(f'IF($'+get_column_letter(4+j)+f'${p}=1,"J{j+1} ","")' for j in range(7))
    cell(f"D{r}", f'=IF(SUM($D${p}:$J${p})=0,"aucune operation",'
                  f'SUBSTITUTE(TRIM({jl})," ",", "))', size=10, border=box)
    cell(f"E{r}", f"=SUM($D${p}:$J${p})", size=10, fmt="0", border=box)
    cell(f"F{r}", f"=$C${ROW_NBW}", size=10, fmt="0", border=box)
    cell(f"G{r}", f"=SUMPRODUCT($C${ROW_OCC}:$I${ROW_OCC},$D${p}:$J${p})", bold=True, size=10, fmt="0", border=box)

cell(f"B{ROW_REC_TOT}", "TOTAL GENERAL", bold=True, size=10, color="FFFFFF", fill=NAVY, halign="left", border=box)
for c in ("C", "D"):
    cell(f"{c}{ROW_REC_TOT}", "", fill=NAVY, border=box)
cell(f"E{ROW_REC_TOT}", f"=SUM(E{ROW_REC0}:E{ROW_REC_TOT-1})", bold=True, size=10, color="FFFFFF", fill=NAVY, fmt="0", border=box)
cell(f"F{ROW_REC_TOT}", f"=$C${ROW_NBW}", bold=True, size=10, color="FFFFFF", fill=NAVY, fmt="0", border=box)
cell(f"G{ROW_REC_TOT}", f"=SUM(G{ROW_REC0}:G{ROW_REC_TOT-1})", bold=True, size=11, color="FFFFFF", fill=NAVY, fmt="0", border=box)

r = ROW_REC_TOT + 1
ws.merge_cells(f"B{r}:G{r}")
cell(f"B{r}", "Methode de calcul : pour chaque jour d'operation de la semaine type, le nombre d'occurrences de ce "
              "jour dans le mois est compte via les dates reelles de la grille (ligne \"Occurrences du jour dans le mois\", "
              "semaines partielles incluses). Total rotations du mois = somme de ces occurrences. La colonne "
              "\"Nb de semaines/occurrences\" indique le nombre de semaines calendaires couvertes par le mois.",
     size=8, italic=True, color="595959", halign="left", wrap=True)
ws.row_dimensions[r].height = 26

# ---------------------------------------------------------------- zone de parametres
ws.merge_cells(f"B{ROW_PAR_TITLE}:P{ROW_PAR_TITLE}")
cell(f"B{ROW_PAR_TITLE}", "ZONE DE PARAMETRES - SEMAINES TYPES PAR ROUTE (saisie unique, valable pour les 12 mois)",
     bold=True, size=11, color="FFFFFF", fill="833C0C", halign="left")
ws.merge_cells(f"B{ROW_PAR_TITLE+1}:P{ROW_PAR_TITLE+1}")
cell(f"B{ROW_PAR_TITLE+1}", "Cellules en BLEU = saisie utilisateur. Type avion : liste deroulante (77W / 778 / A320). "
                            "Jours J1-J7 : saisir 1 si la route opere ce jour-la, 0 sinon. Toute modification recalcule "
                            "immediatement la grille et le recapitulatif, pour les 12 mois. "
                            "Valeurs livrees = exemples a ajuster (NOSRUN = 3/7 A320 sur J1/J4/J6, conforme a l'exemple fourni).",
     size=8, italic=True, color="833C0C", halign="left", wrap=True)
ws.row_dimensions[ROW_PAR_TITLE+1].height = 24

par_hdr = ["Route", "Type avion", "J1", "J2", "J3", "J4", "J5", "J6", "J7", "Rot./sem. (auto)", "Couleur"]
for i, h in enumerate(par_hdr):
    cell(f"{get_column_letter(2+i)}{ROW_PAR_HDR}", h, bold=True, size=9, color="FFFFFF", fill="833C0C", wrap=True, border=box)
cell(f"N{ROW_PAR_HDR}", "Mois (liste)", bold=True, size=9, color="FFFFFF", fill="833C0C", border=box)
cell(f"O{ROW_PAR_HDR}", "1er jour", bold=True, size=9, color="FFFFFF", fill="833C0C", border=box)
cell(f"P{ROW_PAR_HDR}", "Types avion", bold=True, size=9, color="FFFFFF", fill="833C0C", border=box)

for i, (route, typ, days, colr, white) in enumerate(ROUTES):
    r = ROW_PAR0 + i
    cell(f"B{r}", route, bold=True, size=10, halign="left", border=box)
    cell(f"C{r}", typ, size=10, color=BLUE_IN, border=box)
    for j, d in enumerate(days):
        cell(f"{get_column_letter(4+j)}{r}", d, size=10, color=BLUE_IN, border=box)
    cell(f"K{r}", f"=SUM(D{r}:J{r})&\"/7\"", size=10, bold=True, border=box)
    cell(f"L{r}", "", fill=colr, border=box)

for i, (lab, y, m) in enumerate(MOIS):
    r = ROW_PAR0 + i
    cell(f"N{r}", lab, size=9, halign="left", border=box)
    cell(f"O{r}", f"=DATE({y},{m},1)", size=9, fmt="DD/MM/YYYY", border=box)
for i, t in enumerate(TYPES):
    cell(f"P{ROW_PAR0+i}", t, size=9, border=box)

# ---------------------------------------------------------------- listes deroulantes
def dv(formula, cells):
    d = DataValidation(type="list", formula1=formula, allow_blank=False, showDropDown=False)
    ws.add_data_validation(d)
    for c in cells:
        d.add(ws[c] if ":" not in c else c)
    return d

d1 = DataValidation(type="list", formula1=f"=$N${ROW_PAR0}:$N${ROW_PAR_END}", allow_blank=False)
ws.add_data_validation(d1); d1.add(ws["C4"])
d2 = DataValidation(type="list", formula1=f"=$B${ROW_PAR0}:$B${ROW_PAR_END}", allow_blank=False)
ws.add_data_validation(d2); d2.add(ws["C5"])
d3 = DataValidation(type="list", formula1=f"=$P${ROW_PAR0}:$P${ROW_PAR0+len(TYPES)-1}", allow_blank=False)
ws.add_data_validation(d3); d3.add(f"C{ROW_PAR0}:C{ROW_PAR_END}")
d4 = DataValidation(type="list", formula1='"0,1"', allow_blank=False,
                    error="Saisir 1 (jour opere) ou 0 (pas d'operation).", errorTitle="Valeur invalide")
ws.add_data_validation(d4); d4.add(f"D{ROW_PAR0}:J{ROW_PAR_END}")

# ---------------------------------------------------------------- mise en forme conditionnelle (1 couleur = 1 route)
grid_ranges = " ".join(f"{get_column_letter(GRID_C0)}{r}:{get_column_letter(GRID_C0+6)}{r}" for r in type_rows)
for route, typ, days, colr, white in ROUTES:
    fnt = Font(name=F, size=11, bold=True, color="FFFFFF" if white else "000000")
    ws.conditional_formatting.add(
        grid_ranges,
        FormulaRule(formula=[f'AND($C$5="{route}",{get_column_letter(GRID_C0)}{ROW_W1}<>"")'],
                    fill=PatternFill("solid", fgColor=colr), font=fnt, stopIfTrue=True))
    ws.conditional_formatting.add(
        f"B{ROW_REC0}:B{ROW_REC_TOT-1}",
        FormulaRule(formula=[f'$B{ROW_REC0}="{route}"'],
                    fill=PatternFill("solid", fgColor=colr),
                    font=Font(name=F, size=10, bold=True, color="FFFFFF" if white else "000000"),
                    stopIfTrue=True))
    ws.conditional_formatting.add(
        f"B{ROW_PAR0}:B{ROW_PAR_END}",
        FormulaRule(formula=[f'$B{ROW_PAR0}="{route}"'],
                    fill=PatternFill("solid", fgColor=colr),
                    font=Font(name=F, size=10, bold=True, color="FFFFFF" if white else "000000"),
                    stopIfTrue=True))

# ---------------------------------------------------------------- mise en page
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 30
for j in range(7):
    ws.column_dimensions[get_column_letter(GRID_C0+j)].width = 15
ws.column_dimensions["J"].width = 3
ws.column_dimensions["K"].width = 34
ws.column_dimensions["L"].width = 14
ws.column_dimensions["M"].width = 3
ws.column_dimensions["N"].width = 16
ws.column_dimensions["O"].width = 12
ws.column_dimensions["P"].width = 12
ws.freeze_panes = "C8"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_area = f"B1:L{ROW_REC_TOT+1}"

wb.save(OUT)
print("ecrit ->", OUT)
