"""
app.py — Application Streamlit pour l'analyse de facturation dentaire v16
Groupe Efficentres Horizon

Lancement : streamlit run app.py
"""

import io
import zipfile
import pandas as pd
import streamlit as st
from moteur_v16 import analyser_base

# ─────────────────────────────────────────────
# Configuration page
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Analyse Facturation v16 — Efficentres Horizon",
    page_icon="🦷",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# CSS global
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* Palette */
:root {
    --navy:      #1B2A4A;
    --dark-red:  #8B1A1A;
    --gold:      #C9A84C;
    --bg:        #F7F8FA;
    --card-bg:   #FFFFFF;
    --text:      #1C1C1E;
    --muted:     #6B7280;
    --border:    #E5E7EB;
    --critique:  #DC2626;
    --eleve:     #D97706;
    --ok:        #059669;
}
body { background: var(--bg); color: var(--text); }

/* En-tête */
.app-header {
    background: linear-gradient(135deg, var(--navy) 0%, #2C3E6B 100%);
    color: white;
    padding: 1.5rem 2rem;
    border-radius: 12px;
    margin-bottom: 2rem;
}
.app-header h1 { margin: 0; font-size: 1.6rem; }
.app-header p  { margin: 0.25rem 0 0; color: #B8C5E0; font-size: 0.9rem; }

/* Cartes métriques */
.metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.metric-card {
    background: var(--card-bg);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1rem 1.25rem;
    min-width: 160px;
    flex: 1;
}
.metric-card .label { font-size: 0.78rem; color: var(--muted); text-transform: uppercase;
                      letter-spacing: .04em; margin-bottom: .25rem; }
.metric-card .value { font-size: 1.6rem; font-weight: 700; }
.metric-card .value.rouge  { color: var(--critique); }
.metric-card .value.orange { color: var(--eleve); }
.metric-card .value.vert   { color: var(--ok); }

/* Badges signal */
.badge-critique {
    background: #FEF2F2; color: var(--critique);
    border: 1px solid #FECACA;
    padding: 2px 8px; border-radius: 4px; font-size: 0.78rem; font-weight: 600;
}
.badge-eleve {
    background: #FFFBEB; color: var(--eleve);
    border: 1px solid #FDE68A;
    padding: 2px 8px; border-radius: 4px; font-size: 0.78rem; font-weight: 600;
}

/* Section anomalies */
.section-title {
    font-size: 1rem; font-weight: 700; color: var(--navy);
    border-bottom: 2px solid var(--navy); padding-bottom: .4rem;
    margin: 1.5rem 0 1rem;
}

/* Tableau résultat (surcharge agGrid Streamlit) */
.stDataFrame { border: 1px solid var(--border) !important; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# En-tête
# ─────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <h1>🦷 Analyse Facturation v16</h1>
  <p>Groupe Efficentres Horizon · Détection des anomalies de facturation dentaire</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Zone d'upload
# ─────────────────────────────────────────────
st.subheader("1 · Charger la base de facturation")

uploaded_file = st.file_uploader(
    "Déposer la base de facturation (CSV, XLS, XLSX ou XLSB)",
    type=["csv", "xls", "xlsx", "xlsb"],
    help="Le fichier doit contenir les colonnes : Numéro facture, Date facture, "
         "Praticien, Honoraire acte, Code acte, Panier, Analytique, Numero_dossier …"
)

# Options CSV (affichées uniquement si le fichier est CSV)
_is_csv = (uploaded_file is not None and
           uploaded_file.name.lower().endswith(".csv"))
SEP_OPTIONS = {
    "Point-virgule  ;  (défaut France)": ";",
    "Virgule  ,": ",",
    "Tabulation  \\t": "\t",
    "Pipe  |": "|",
}
if _is_csv or uploaded_file is None:
    col1, col2 = st.columns([2, 1])
    with col1:
        sep_label = st.selectbox("Séparateur CSV", list(SEP_OPTIONS.keys()), index=0)
    with col2:
        encoding = st.selectbox("Encodage", ["utf-8", "latin-1", "cp1252"], index=0)
    sep = SEP_OPTIONS[sep_label]
else:
    sep = ";"
    encoding = "utf-8"

# Pour XLSX/XLS : choix de la feuille (affiché après upload)
_sheet_name = None
if uploaded_file is not None and not _is_csv:
    _ext = uploaded_file.name.lower().rsplit(".", 1)[-1]
    if _ext in ("xlsx", "xls"):
        try:
            import openpyxl
            _wb_tmp = openpyxl.load_workbook(
                io.BytesIO(uploaded_file.read()), read_only=True, data_only=True
            )
            _sheets = _wb_tmp.sheetnames
            uploaded_file.seek(0)
            if len(_sheets) > 1:
                _sheet_name = st.selectbox(
                    "Feuille à analyser", _sheets, index=0
                )
            else:
                _sheet_name = _sheets[0]
        except Exception:
            uploaded_file.seek(0)
            _sheet_name = 0
    elif _ext == "xlsb":
        try:
            from pyxlsb import open_workbook as open_xlsb
            _bytes_tmp = uploaded_file.read()
            with open_xlsb(io.BytesIO(_bytes_tmp)) as _wb_tmp:
                _sheets_xlsb = _wb_tmp.sheets
            uploaded_file.seek(0)
            if len(_sheets_xlsb) > 1:
                _sheet_name = st.selectbox(
                    "Feuille à analyser", _sheets_xlsb, index=0
                )
            else:
                _sheet_name = _sheets_xlsb[0]
        except Exception:
            uploaded_file.seek(0)
            _sheet_name = 0


# ─────────────────────────────────────────────
# Génération du fichier Excel de synthèse
# ─────────────────────────────────────────────

def generer_excel(df_enrichi: pd.DataFrame, df_anomalies: pd.DataFrame) -> bytes:
    """
    Génère un fichier Excel avec deux onglets :
      - « Base enrichie »  : toutes les lignes de la base enrichie (base intacte),
                             avec une colonne helper « Première occurrence » (1 pour
                             la 1re occurrence de chaque Numéro facture, 0 sinon)
      - « Synthèse »       : tableau StatutxDétail × Centre reproduisant exactement
                             la structure du fichier de référence v17
                             (couleurs, formules SUMIFS sur colonne helper, formats)
                             — équivalent de NBVAL(UNIQUE(FILTRE(…))) sans UNIQUE/FILTRE
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    # Supprimer la feuille par défaut créée par Workbook() — on ne veut
    # que la feuille "Synthèse" créée ci-dessous dans ce workbook.
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # ══════════════════════════════════════════════════════════════
    # Helpers de style
    # ══════════════════════════════════════════════════════════════
    def mk_fill(hex6):
        """PatternFill solid depuis un code hex 6 ou 8 caractères."""
        return PatternFill("solid", fgColor=hex6)

    def mk_font(bold=False, color="000000", size=10, italic=False, name="Calibri"):
        return Font(name=name, bold=bold, color=color, size=size, italic=italic)

    def mk_border(style="thin", color="BFBFBF"):
        s = Side(style=style, color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    def mk_align(h="center", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    # ── Couleurs extraites du xlsb de référence ──────────────────
    # Bleu nuit en-tête          : 002060  (fill index 2 : FF002060)
    # Rouge en-tête              : C00000  (fill index 3 : FFC00000)
    # Bleu accent centres        : 4472C4  (theme accent1)
    # Gris très léger sous-total : fond blanc cassé ≈ theme 0 tint -0.05 → D9D9D9 approx
    # Bleu clair TOTAL           : 5B9BD5 tint 0.8 → DEEAF1 approx (accent5 tint +0.8)
    # Blanc                      : FFFFFF
    C_BLUE_NUIT  = "002060"   # en-tête Statut×Détail
    C_RED_DARK   = "C00000"   # en-tête TOTAL # et TOTAL %
    C_ACCENT_BLU = "4472C4"   # en-têtes centres (accent1)
    C_SOUS_TOT   = "D9D9D9"   # fond gris clair sous-totaux (theme:0 tint -0.05)
    C_TOTAL_BG   = "DEEAF1"   # fond bleu très clair ligne TOTAL (accent5 tint 0.8)
    C_BORDEAUX   = "C00000"   # fond bordeaux/rouge sur cellules centres KO
    C_WHITE      = "FFFFFF"
    C_BLACK      = "000000"

    # Police blanche = couleur police sur fonds colorés (theme:1 = blanc)
    C_WHITE_FONT = "FFFFFF"
    # Police noire sur fonds clairs
    C_BLACK_FONT = "000000"

    # Format nombre de référence
    FMT_NB   = '#,##0;-#,##0;"-"'       # entiers
    FMT_PCT  = '0.0%;-0.0%;"-"'         # pourcentages (pas de parenthèses → pas de rouge auto)
    FMT_GEN  = "General"

    # ══════════════════════════════════════════════════════════════
    # Onglet 1 — Base enrichie  (préparation du DataFrame)
    # ══════════════════════════════════════════════════════════════

    # Base intacte (toutes les lignes)
    df_base = df_enrichi.copy()

    # Colonne helper : 1 pour la 1re occurrence de chaque Numéro facture
    if "Numéro facture" in df_base.columns:
        df_base["Première occurrence"] = (~df_base["Numéro facture"].duplicated(keep="first")).astype(int)
    else:
        df_base["Première occurrence"] = 1

    cols_base = list(df_base.columns)

    # ══════════════════════════════════════════════════════════════
    # Onglet 2 — Synthèse
    # Structure conforme au fichier v17 de référence :
    #   Col C = Statut × Détail (label)
    #   Col E = TOTAL #  (nb absolus)
    #   Col F = TOTAL %  (%)
    #   Col H:AB = ARC AUB BEA BEL BOB BON BOR COL DAV FIV LAM LUM LY9 LYO NEU REN ROE RSS SUR TOU VIL WAZ
    #   Lignes :
    #     3  = en-têtes
    #     4  = vide
    #     5  = OK OK
    #     6  = vide
    #     7..13 = KO Intraday détails (7 lignes)
    #     14 = SOUS-TOTAL KO Intraday
    #     15 = vide
    #     16..22 = KO Ultérieur détails (7 lignes)
    #     23 = SOUS-TOTAL KO Ultérieur
    #     24 = vide
    #     25 = KO Hors base AVOIR - Non refacturé
    #     26 = vide
    #     27 = TOTAL
    # ══════════════════════════════════════════════════════════════
    ws = wb.create_sheet("Synthèse")

    # Masquer le quadrillage
    try:
        ws.sheet_view.showGridLines = False
    except Exception:
        pass

    # ── Références à la Base enrichie ────────────────────────────
    BASE = "'Base enrichie'"

    # Trouver les colonnes utiles dans df_base
    def col_letter_for(col_name_search):
        try:
            idx = cols_base.index(col_name_search) + 1
            return get_column_letter(idx)
        except ValueError:
            return None

    col_sxd = col_letter_for("Statut x Détail") or col_letter_for("StatutxDétail") or "K"
    col_ana = col_letter_for("Analytique") or col_letter_for("Code centre") or col_letter_for("Centre") or "C"
    col_sta = col_letter_for("Statut") or "I"
    col_occ = col_letter_for("Première occurrence") or get_column_letter(len(cols_base))

    # ── CENTRES : valeurs uniques de la colonne Analytique, triées alphabétiquement ──
    # Dynamique — fonctionne pour Efficentres ET Vertuo, et tout autre périmètre
    _col_ana_name = None
    for _cn in ("Analytique", "Code centre", "Centre"):
        if _cn in df_base.columns:
            _col_ana_name = _cn
            break
    if _col_ana_name:
        CENTRES = sorted(df_base[_col_ana_name].dropna().unique().tolist())
    else:
        # Fallback : liste Efficentres historique
        CENTRES = [
            "ARC", "AUB", "BEA", "BEL", "BOB", "BON", "BOR", "COL",
            "DAV", "FIV", "LAM", "LUM", "LY9", "LYO", "NEU", "REN",
            "ROE", "RSS", "SUR", "TOU", "VIL", "WAZ",
        ]
    N_CENTRES = len(CENTRES)

    # ── Colonnes (indices 1-based) ────────────────────────────────
    COL_LABEL  = 3   # C
    COL_NB     = 5   # E  → TOTAL #
    COL_PCT    = 6   # F  → TOTAL %
    COL_FIRST  = 8   # H  → premier centre
    COL_LAST   = COL_FIRST + N_CENTRES - 1

    # Largeurs colonnes
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 2
    ws.column_dimensions["C"].width = 52
    ws.column_dimensions["D"].width = 2
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 2
    for k in range(N_CENTRES):
        ws.column_dimensions[get_column_letter(COL_FIRST + k)].width = 9

    # ── Formules ──────────────────────────────────────────────────
    # Les formules centres référencent {col}$3 (en-tête ligne 3) et non le nom en dur
    # → robuste si les centres changent entre deux fichiers

    def f_nb_sxd(sxd_label):
        return (f'=SUMIFS({BASE}!${col_occ}:${col_occ},'
                f'{BASE}!${col_sxd}:${col_sxd},"{sxd_label}")')

    def f_nb_statut(statut_label):
        return (f'=SUMIFS({BASE}!${col_occ}:${col_occ},'
                f'{BASE}!${col_sta}:${col_sta},"{statut_label}")')

    def f_pct_sxd(sxd_label, row_total):
        return (f'=IFERROR(SUMIFS({BASE}!${col_occ}:${col_occ},'
                f'{BASE}!${col_sxd}:${col_sxd},"{sxd_label}")/E${row_total},0)')

    def f_pct_statut(statut_label, row_total):
        return (f'=IFERROR(SUMIFS({BASE}!${col_occ}:${col_occ},'
                f'{BASE}!${col_sta}:${col_sta},"{statut_label}")/E${row_total},0)')

    # Centre % : SUMIFS filtre sur col_ana = valeur lue en {col_centre}$3 (ligne en-tête)
    # → si on renomme un centre dans la Base enrichie, la Synthèse suit automatiquement
    def f_pct_sxd_centre(sxd_label, row_total, col_centre):
        return (f'=IFERROR(SUMIFS({BASE}!${col_occ}:${col_occ},'
                f'{BASE}!${col_sxd}:${col_sxd},"{sxd_label}",'
                f'{BASE}!${col_ana}:${col_ana},{col_centre}${ROW_HDR})'
                f'/{col_centre}${row_total},0)')

    def f_pct_statut_centre(statut_label, row_total, col_centre):
        return (f'=IFERROR(SUMIFS({BASE}!${col_occ}:${col_occ},'
                f'{BASE}!${col_sta}:${col_sta},"{statut_label}",'
                f'{BASE}!${col_ana}:${col_ana},{col_centre}${ROW_HDR})'
                f'/{col_centre}${row_total},0)')

    def f_total_nb():
        return f'=SUM({BASE}!${col_occ}:${col_occ})'

    def f_total_centre(col_centre):
        # Total par centre : SUMIFS sur col_ana = valeur lue en {col}$3
        return (f'=SUMIFS({BASE}!${col_occ}:${col_occ},'
                f'{BASE}!${col_ana}:${col_ana},{col_centre}${ROW_HDR})')

    # ── Numéros de lignes dans la synthèse ───────────────────────
    # (conformes au fichier de référence)
    ROW_HDR    = 3
    ROW_OK     = 5
    ROW_KOI    = list(range(7, 14))   # 7 8 9 10 11 12 13
    ROW_STI    = 14
    ROW_KOU    = list(range(16, 23))  # 16 17 18 19 20 21 22
    ROW_STU    = 23
    ROW_HB     = 25
    ROW_TOTAL  = 27

    # Labels détail dans l'ordre exact du fichier source
    SXD_KOI = [
        "KO Intraday AVOIR - Refacturé",
        "KO Intraday REFACTURATION",
        "KO Intraday FACTURATION INITIALE - Double facturation",
        "KO Intraday DOUBLE FACTURATION",
        "KO Intraday AVOIR - Double facturation",
        "KO Intraday FACTURATION INITIALE - Annulée sans suite",
        "KO Intraday AVOIR - Non refacturé",
    ]
    SXD_KOU = [
        "KO Ultérieur AVOIR - Refacturé",
        "KO Ultérieur REFACTURATION",
        "KO Ultérieur FACTURATION INITIALE - Annulée sans suite",
        "KO Ultérieur AVOIR - Non refacturé",
        "KO Ultérieur FACTURATION INITIALE - Double facturation",
        "KO Ultérieur DOUBLE FACTURATION",
        "KO Ultérieur AVOIR - Double facturation",
    ]

    # ── Fonctions d'écriture de cellule ──────────────────────────
    def write_cell(row, col, value=None, fill=None, font=None,
                   alignment=None, border=None, number_format=None):
        c = ws.cell(row=row, column=col, value=value)
        if fill:        c.fill           = fill
        if font:        c.font           = font
        if alignment:   c.alignment      = alignment
        if border:      c.border         = border
        if number_format: c.number_format = number_format
        return c

    # ── Bordure fine standard ─────────────────────────────────────
    brd = mk_border("thin", "BFBFBF")

    # ══ LIGNE 3 : EN-TÊTES ═══════════════════════════════════════
    ws.row_dimensions[ROW_HDR].height = 22

    # C3 : "Statut x Détail"  — fond bleu nuit, blanc gras
    write_cell(ROW_HDR, COL_LABEL, "Statut x Détail",
               fill=mk_fill(C_BLUE_NUIT),
               font=mk_font(bold=True, color=C_WHITE_FONT, size=10),
               alignment=mk_align("left", "center"),
               border=brd)

    # E3 : "TOTAL #"  — fond rouge foncé, blanc gras
    write_cell(ROW_HDR, COL_NB, "TOTAL #",
               fill=mk_fill(C_RED_DARK),
               font=mk_font(bold=True, color=C_WHITE_FONT, size=10),
               alignment=mk_align("center", "center"),
               border=brd)

    # F3 : "TOTAL %"  — fond rouge foncé, blanc gras
    write_cell(ROW_HDR, COL_PCT, "TOTAL %",
               fill=mk_fill(C_RED_DARK),
               font=mk_font(bold=True, color=C_WHITE_FONT, size=10),
               alignment=mk_align("center", "center"),
               border=brd)

    # H3:AB3 : noms des centres — fond bleu accent, blanc gras
    for k, ctr in enumerate(CENTRES):
        write_cell(ROW_HDR, COL_FIRST + k, ctr,
                   fill=mk_fill(C_ACCENT_BLU),
                   font=mk_font(bold=True, color=C_WHITE_FONT, size=10),
                   alignment=mk_align("center", "center"),
                   border=brd)

    # ══ LIGNE 5 : OK OK ═════════════════════════════════════════
    # Style : fond gris clair, police noire gras
    ws.row_dimensions[ROW_OK].height = 17
    fill_ok  = mk_fill(C_SOUS_TOT)
    font_ok  = mk_font(bold=True, color=C_BLACK_FONT, size=10)

    write_cell(ROW_OK, COL_LABEL, "OK",
               fill=fill_ok, font=font_ok,
               alignment=mk_align("left", "center"), border=brd)
    write_cell(ROW_OK, COL_NB,
               value=f_nb_statut("OK"),
               fill=fill_ok, font=font_ok,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_NB)
    write_cell(ROW_OK, COL_PCT,
               value=f_pct_statut("OK", ROW_TOTAL),
               fill=fill_ok, font=font_ok,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_PCT)
    for k in range(N_CENTRES):
        col_c = get_column_letter(COL_FIRST + k)
        write_cell(ROW_OK, COL_FIRST + k,
                   value=f_pct_statut_centre("OK", ROW_TOTAL, col_c),
                   fill=fill_ok, font=font_ok,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_PCT)

    # ══ LIGNES 7–13 : KO Intraday (détails) ════════════════════
    # Toutes cellules : fond blanc, police noire
    fill_white = mk_fill(C_WHITE)
    font_det   = mk_font(bold=False, color=C_BLACK_FONT, size=10)
    for i, (row_i, sxd) in enumerate(zip(ROW_KOI, SXD_KOI)):
        ws.row_dimensions[row_i].height = 17
        write_cell(row_i, COL_LABEL, sxd,
                   fill=fill_white, font=font_det,
                   alignment=mk_align("left", "center"), border=brd)
        write_cell(row_i, COL_NB,
                   value=f_nb_sxd(sxd),
                   fill=fill_white, font=font_det,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_NB)
        write_cell(row_i, COL_PCT,
                   value=f_pct_sxd(sxd, ROW_TOTAL),
                   fill=fill_white, font=font_det,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_PCT)
        for k in range(N_CENTRES):
            col_c = get_column_letter(COL_FIRST + k)
            write_cell(row_i, COL_FIRST + k,
                       value=f_pct_sxd_centre(sxd, ROW_TOTAL, col_c),
                       fill=fill_white, font=font_det,
                       alignment=mk_align("center", "center"), border=brd,
                       number_format=FMT_PCT)

    # ══ LIGNE 14 : SOUS-TOTAL KO Intraday ══════════════════════
    # Style : fond gris clair, police noire gras
    ws.row_dimensions[ROW_STI].height = 17
    fill_st  = mk_fill(C_SOUS_TOT)
    font_st  = mk_font(bold=True, color=C_BLACK_FONT, size=10)

    write_cell(ROW_STI, COL_LABEL, "SOUS-TOTAL KO Intraday",
               fill=fill_st, font=font_st,
               alignment=mk_align("left", "center"), border=brd)

    r_first_koi = ROW_KOI[0]
    r_last_koi  = ROW_KOI[-1]
    nb_col_e   = get_column_letter(COL_NB)
    pct_col_f  = get_column_letter(COL_PCT)

    write_cell(ROW_STI, COL_NB,
               value=f'=SUM({nb_col_e}{r_first_koi}:{nb_col_e}{r_last_koi})',
               fill=fill_st, font=font_st,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_NB)
    write_cell(ROW_STI, COL_PCT,
               value=f'=SUM({pct_col_f}{r_first_koi}:{pct_col_f}{r_last_koi})',
               fill=fill_st, font=font_st,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_PCT)
    for k, ctr in enumerate(CENTRES):
        col_c = get_column_letter(COL_FIRST + k)
        write_cell(ROW_STI, COL_FIRST + k,
                   value=f'=SUM({col_c}{r_first_koi}:{col_c}{r_last_koi})',
                   fill=fill_st, font=font_st,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_PCT)

    # ══ LIGNES 16–22 : KO Ultérieur (détails) ══════════════════
    for i, (row_i, sxd) in enumerate(zip(ROW_KOU, SXD_KOU)):
        ws.row_dimensions[row_i].height = 17
        write_cell(row_i, COL_LABEL, sxd,
                   fill=fill_white, font=font_det,
                   alignment=mk_align("left", "center"), border=brd)
        write_cell(row_i, COL_NB,
                   value=f_nb_sxd(sxd),
                   fill=fill_white, font=font_det,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_NB)
        write_cell(row_i, COL_PCT,
                   value=f_pct_sxd(sxd, ROW_TOTAL),
                   fill=fill_white, font=font_det,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_PCT)
        for k in range(N_CENTRES):
            col_c = get_column_letter(COL_FIRST + k)
            write_cell(row_i, COL_FIRST + k,
                       value=f_pct_sxd_centre(sxd, ROW_TOTAL, col_c),
                       fill=fill_white, font=font_det,
                       alignment=mk_align("center", "center"), border=brd,
                       number_format=FMT_PCT)

    # ══ LIGNE 23 : SOUS-TOTAL KO Ultérieur ═════════════════════
    ws.row_dimensions[ROW_STU].height = 17
    r_first_kou = ROW_KOU[0]
    r_last_kou  = ROW_KOU[-1]

    write_cell(ROW_STU, COL_LABEL, "SOUS-TOTAL KO Ultérieur",
               fill=fill_st, font=font_st,
               alignment=mk_align("left", "center"), border=brd)
    write_cell(ROW_STU, COL_NB,
               value=f'=SUM({nb_col_e}{r_first_kou}:{nb_col_e}{r_last_kou})',
               fill=fill_st, font=font_st,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_NB)
    write_cell(ROW_STU, COL_PCT,
               value=f'=SUM({pct_col_f}{r_first_kou}:{pct_col_f}{r_last_kou})',
               fill=fill_st, font=font_st,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_PCT)
    for k, ctr in enumerate(CENTRES):
        col_c = get_column_letter(COL_FIRST + k)
        write_cell(ROW_STU, COL_FIRST + k,
                   value=f'=SUM({col_c}{r_first_kou}:{col_c}{r_last_kou})',
                   fill=fill_st, font=font_st,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_PCT)

    # ══ LIGNE 25 : KO Hors base ══════════════════════════════
    ws.row_dimensions[ROW_HB].height = 17
    sxd_hb = "KO Hors base AVOIR - Non refacturé"

    write_cell(ROW_HB, COL_LABEL, sxd_hb,
               fill=fill_st, font=font_st,
               alignment=mk_align("left", "center"), border=brd)
    write_cell(ROW_HB, COL_NB,
               value=f_nb_sxd(sxd_hb),
               fill=fill_st, font=font_st,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_NB)
    write_cell(ROW_HB, COL_PCT,
               value=f_pct_sxd(sxd_hb, ROW_TOTAL),
               fill=fill_st, font=font_st,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_PCT)
    for k in range(N_CENTRES):
        col_c = get_column_letter(COL_FIRST + k)
        write_cell(ROW_HB, COL_FIRST + k,
                   value=f_pct_sxd_centre(sxd_hb, ROW_TOTAL, col_c),
                   fill=fill_st, font=font_st,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_PCT)

    # ══ LIGNE 27 : TOTAL ════════════════════════════════════════
    # Style : fond bleu clair (DEEAF1), police noire gras
    ws.row_dimensions[ROW_TOTAL].height = 17
    fill_tot = mk_fill(C_TOTAL_BG)
    font_tot = mk_font(bold=True, color=C_BLACK_FONT, size=10)

    write_cell(ROW_TOTAL, COL_LABEL, "TOTAL",
               fill=fill_tot, font=font_tot,
               alignment=mk_align("left", "center"), border=brd)

    # E27 : somme des lignes de la synthèse (garantit 100% en F27)
    write_cell(ROW_TOTAL, COL_NB,
               value=f'=E{ROW_OK}+E{ROW_STI}+E{ROW_STU}+E{ROW_HB}',
               fill=fill_tot, font=font_tot,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_NB)

    # F27 : 100% (E27/E27)
    nb_col_e_total = get_column_letter(COL_NB)
    write_cell(ROW_TOTAL, COL_PCT,
               value=f'=IFERROR({nb_col_e_total}{ROW_TOTAL}/{nb_col_e_total}{ROW_TOTAL},0)',
               fill=fill_tot, font=font_tot,
               alignment=mk_align("center", "center"), border=brd,
               number_format=FMT_PCT)

    # H27:…27 : nombre absolu de factures uniques par centre (même logique que E27)
    # SUMIFS direct sur Base enrichie, filtre sur le centre lu en ligne 3
    for k in range(N_CENTRES):
        col_c = get_column_letter(COL_FIRST + k)
        write_cell(ROW_TOTAL, COL_FIRST + k,
                   value=(f'=SUMIFS({BASE}!${col_occ}:${col_occ},'
                          f'{BASE}!${col_ana}:${col_ana},{col_c}${ROW_HDR})'),
                   fill=fill_tot, font=font_tot,
                   alignment=mk_align("center", "center"), border=brd,
                   number_format=FMT_NB)

    # Figer panneaux : E4 (conforme au freeze du fichier source)
    ws.freeze_panes = "H4"

    # ══════════════════════════════════════════════════════════════
    # Sauvegarde finale — architecture deux fichiers + fusion ZIP
    # ─────────────────────────────────────────────────────────────
    # Fichier A (synth.xlsx)  : openpyxl, Synthèse seule (~30 lignes)
    # Fichier B (base.xlsx)   : xlsxwriter via pandas, Base enrichie
    #                           (streaming natif, ~10× plus rapide qu'openpyxl)
    # Fusion               : manipulation directe du ZIP .xlsx pour
    #                           injecter la feuille de B dans A sans
    #                           jamais relire les 500k lignes avec openpyxl
    # ══════════════════════════════════════════════════════════════
    import tempfile, os, zipfile, shutil, re as _re

    with tempfile.TemporaryDirectory() as _tmp:
        _p_synth = os.path.join(_tmp, "synth.xlsx")
        _p_base  = os.path.join(_tmp, "base.xlsx")
        _p_final = os.path.join(_tmp, "final.xlsx")

        # ── 1) Synthèse via openpyxl (instantané) ─────────────────
        wb.save(_p_synth)

        # ── 2) Base enrichie via xlsxwriter (streaming natif) ────────
        with pd.ExcelWriter(_p_base, engine="xlsxwriter",
                            engine_kwargs={"options": {"strings_to_urls": False}}) as _wr:
            df_base.to_excel(_wr, index=False, sheet_name="Base enrichie")
            _xwb = _wr.book
            _xws = _wr.sheets["Base enrichie"]
            _hfmt = _xwb.add_format({
                "bold": True, "font_name": "Calibri", "font_size": 10,
                "font_color": "#FFFFFF", "bg_color": "#1B2A4A",
                "align": "center", "valign": "vcenter",
            })
            for _ci, _cn in enumerate(df_base.columns):
                _xws.write(0, _ci, _cn, _hfmt)
            _xws.set_row(0, 20)
            _xws.freeze_panes(1, 0)

        # ── 3) Fusion ZIP ─────────────────────────────────────────
        # Lire synth.xlsx
        with zipfile.ZipFile(_p_synth, "r") as _zs:
            _wb_xml   = _zs.read("xl/workbook.xml").decode("utf-8")
            _rels_xml = _zs.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            _ct_xml   = _zs.read("[Content_Types].xml").decode("utf-8")
            _synth_files = {n: _zs.read(n) for n in _zs.namelist()}

        # Lire base.xlsx — feuille + sharedStrings
        with zipfile.ZipFile(_p_base, "r") as _zb:
            _base_names = _zb.namelist()
            _ws_name = next((n for n in _base_names
                             if n.startswith("xl/worksheets/sheet")), None)
            _base_sheet = _zb.read(_ws_name)
            _base_ss_xml = (_zb.read("xl/sharedStrings.xml").decode("utf-8")
                            if "xl/sharedStrings.xml" in _base_names else None)
            _base_wb_xml = _zb.read("xl/workbook.xml").decode("utf-8")
            _base_rels   = _zb.read("xl/_rels/workbook.xml.rels").decode("utf-8")

        # ── Fusionner sharedStrings ────────────────────────────────
        # Stratégie : on garde sharedStrings de synth (formules) et on
        # concatène les strings de base en décalant les index dans sheet.xml
        import xml.etree.ElementTree as _ET

        _NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        _ET.register_namespace("", _NS)

        _synth_ss_xml = _synth_files.get("xl/sharedStrings.xml", b"").decode("utf-8")

        def _parse_ss(xml_str):
            """Retourne liste de strings (texte brut) depuis sharedStrings.xml."""
            if not xml_str:
                return []
            root = _ET.fromstring(xml_str)
            result = []
            for si in root.findall(f"{{{_NS}}}si"):
                t = si.find(f"{{{_NS}}}t")
                if t is not None and t.text:
                    result.append(t.text)
                else:
                    # texte avec runs <r><t>
                    parts = [r.find(f"{{{_NS}}}t") for r in si.findall(f"{{{_NS}}}r")]
                    result.append("".join(p.text or "" for p in parts if p is not None))
            return result

        _synth_strings = _parse_ss(_synth_ss_xml)
        _base_strings  = _parse_ss(_base_ss_xml) if _base_ss_xml else []

        # Index de décalage : les strings de base commencent après ceux de synth
        _offset = len(_synth_strings)

        # Construire le sharedStrings fusionné
        _all_strings = _synth_strings + _base_strings
        _ss_lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<sst xmlns="{_NS}" count="{len(_all_strings)}" uniqueCount="{len(_all_strings)}">',
        ]
        for _s in _all_strings:
            _s_esc = (_s.replace("&", "&amp;").replace("<", "&lt;")
                        .replace(">", "&gt;").replace('"', "&quot;"))
            _ss_lines.append(f'<si><t xml:space="preserve">{_s_esc}</t></si>')
        _ss_lines.append("</sst>")
        _merged_ss = "\n".join(_ss_lines).encode("utf-8")

        # Décaler les index <v> dans la feuille Base enrichie
        # (les cellules de type s="…" référencent sharedStrings par index)
        if _offset > 0 and _base_ss_xml:
            import re as _re2
            def _shift_v(m):
                return f"<v>{int(m.group(1)) + _offset}</v>"
            # On ne décale que les cellules de type t="s" (shared string)
            # Pattern : trouver <c ... t="s" ...><v>N</v>
            _base_sheet_str = _base_sheet.decode("utf-8")
            # Décaler toutes les <v> dans les cellules t="s"
            _base_sheet_str = _re2.sub(
                r'(<c [^>]*t="s"[^>]*>(?:<f>[^<]*</f>)?)<v>(\d+)</v>',
                lambda m: m.group(1) + f"<v>{int(m.group(2)) + _offset}</v>",
                _base_sheet_str
            )
            _base_sheet = _base_sheet_str.encode("utf-8")

        # ── IDs pour la nouvelle feuille ──────────────────────────
        _max_rid = max((int(x) for x in _re.findall(r'Id="rId(\d+)"', _rels_xml)), default=0)
        _max_sid = max((int(x) for x in _re.findall(r'sheetId="(\d+)"', _wb_xml)), default=0)
        _new_rid = f"rId{_max_rid + 1}"
        _new_sid = _max_sid + 1

        # workbook.xml : Base enrichie EN DERNIER
        # openpyxl déclare xmlns:r sur chaque <sheet>, on fait pareil
        _sheet_tag = (f'<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                      f'name="Base enrichie" sheetId="{_new_sid}" state="visible" r:id="{_new_rid}"/>')
        _wb_xml = _wb_xml.replace("</sheets>", f"{_sheet_tag}</sheets>", 1)

        # Relations — chemin absolu pour cohérence avec openpyxl (/xl/worksheets/...)
        _rel_tag = (f'<Relationship Id="{_new_rid}" '
                    f'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    f'relationships/worksheet" Target="/xl/worksheets/sheetBase.xml"/>')
        _rels_xml = _rels_xml.replace("</Relationships>", f"{_rel_tag}</Relationships>")

        # Content_Types : nouvelle feuille
        _ct_xml = _ct_xml.replace(
            "</Types>",
            '<Override PartName="/xl/worksheets/sheetBase.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.worksheet+xml"/></Types>'
        )
        # Content_Types : sharedStrings (si absent de synth)
        if "sharedStrings" not in _ct_xml:
            _ct_xml = _ct_xml.replace(
                "</Types>",
                '<Override PartName="/xl/sharedStrings.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.sharedStrings+xml"/></Types>'
            )
            # Ajouter relation sharedStrings dans workbook.rels
            _ss_rel_rid = f"rId{_max_rid + 2}"
            _rels_xml = _rels_xml.replace(
                "</Relationships>",
                f'<Relationship Id="{_ss_rel_rid}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                f'relationships/sharedStrings" Target="/xl/sharedStrings.xml"/>'
                f"</Relationships>"
            )

        # ── Écrire le fichier final ───────────────────────────────
        with zipfile.ZipFile(_p_final, "w", compression=zipfile.ZIP_DEFLATED) as _zout:
            for _name, _data in _synth_files.items():
                if _name == "xl/workbook.xml":
                    _zout.writestr(_name, _wb_xml.encode("utf-8"))
                elif _name == "xl/_rels/workbook.xml.rels":
                    _zout.writestr(_name, _rels_xml.encode("utf-8"))
                elif _name == "[Content_Types].xml":
                    _zout.writestr(_name, _ct_xml.encode("utf-8"))
                elif _name == "xl/sharedStrings.xml":
                    _zout.writestr(_name, _merged_ss)  # version fusionnée
                else:
                    _zout.writestr(_name, _data)
            # Nouvelles entrées
            _zout.writestr("xl/worksheets/sheetBase.xml", _base_sheet)
            if _base_ss_xml and "xl/sharedStrings.xml" not in _synth_files:
                _zout.writestr("xl/sharedStrings.xml", _merged_ss)

        with open(_p_final, "rb") as _f:
            return _f.read()


# ─────────────────────────────────────────────
# Traitement
# ─────────────────────────────────────────────
if uploaded_file is not None:
    st.divider()
    st.subheader("2 · Analyse en cours")

    progress_bar  = st.progress(0, text="Lecture du fichier …")
    status_text   = st.empty()

    # ── Lecture du fichier (CSV / XLS / XLSX / XLSB) ────────────
    _ext_upload = uploaded_file.name.lower().rsplit(".", 1)[-1]
    try:
        if _ext_upload == "csv":
            status_text.info("📂 Lecture du fichier CSV …")
            df_raw = pd.read_csv(
                uploaded_file, sep=sep, encoding=encoding,
                encoding_errors="replace",
                low_memory=False, on_bad_lines="skip"
            )
        elif _ext_upload in ("xlsx", "xls"):
            status_text.info(f"📂 Lecture du fichier {_ext_upload.upper()} …")
            _sheet_arg = _sheet_name if _sheet_name is not None else 0
            df_raw = pd.read_excel(
                uploaded_file, sheet_name=_sheet_arg,
                engine=("openpyxl" if _ext_upload == "xlsx" else "xlrd")
            )
        elif _ext_upload == "xlsb":
            status_text.info("📂 Lecture du fichier XLSB …")
            from pyxlsb import open_workbook as open_xlsb
            _raw_bytes = uploaded_file.read()
            _sheet_arg = _sheet_name if _sheet_name is not None else 0
            with open_xlsb(io.BytesIO(_raw_bytes)) as _wb_xlsb:
                _target_sheet = (
                    _sheet_arg if isinstance(_sheet_arg, str)
                    else _wb_xlsb.sheets[_sheet_arg]
                )
                with _wb_xlsb.get_sheet(_target_sheet) as _ws_xlsb:
                    _rows_xlsb = []
                    for _row in _ws_xlsb.rows():
                        # Convertir chaque cellule en str (None → "")
                        _rows_xlsb.append([
                            str(c.v) if c.v is not None else ""
                            for c in _row
                        ])
            if _rows_xlsb:
                # Supprimer les lignes entièrement vides en tête
                while _rows_xlsb and all(v == "" for v in _rows_xlsb[0]):
                    _rows_xlsb.pop(0)
                if _rows_xlsb:
                    _headers = [
                        h if h != "" else f"Col{i}"
                        for i, h in enumerate(_rows_xlsb[0])
                    ]
                    _data_rows = [
                        r for r in _rows_xlsb[1:]
                        if any(v != "" for v in r)
                    ]
                    df_raw = pd.DataFrame(_data_rows, columns=_headers)
                else:
                    df_raw = pd.DataFrame()
            else:
                df_raw = pd.DataFrame()
        else:
            st.error(f"Format non supporté : .{_ext_upload}")
            st.stop()

        # Convertir les colonnes numériques clés (au cas où lues en str)
        for _num_col in ["Honoraire acte", "BSS", "AMO", "AMC", "AMC2", "AES", "PP", "Coeff", "# actes"]:
            if _num_col in df_raw.columns:
                df_raw[_num_col] = pd.to_numeric(
                    df_raw[_num_col].astype(str).str.replace(",", ".").str.replace(" ", ""),
                    errors="coerce"
                )
        # Convertir les dates
        for _dt_col in ["Date facture", "Date réalisation acte"]:
            if _dt_col in df_raw.columns:
                df_raw[_dt_col] = pd.to_datetime(df_raw[_dt_col], errors="coerce", dayfirst=True)

        # Nettoyer les noms de colonnes : restaurer les accents corrompus
        col_fixes = {
            "Num�ro facture":        "Numéro facture",
            "Num?ro facture":             "Numéro facture",
            "Numro facture":              "Numéro facture",
            "Num\x8ero facture":          "Numéro facture",
            "Pr�nom praticien":      "Prénom praticien",
            "Pr?nom praticien":           "Prénom praticien",
            "Prnom praticien":            "Prénom praticien",
            "Prnom Praticien":            "Prénom praticien",
            "Date r�alisation acte": "Date réalisation acte",
            "Date r?alisation acte":      "Date réalisation acte",
            "Date ralisation acte":       "Date réalisation acte",
            "Dents concern�ees":     "Dents concernées",
            "Lettre cl� NGAP":       "Lettre clé NGAP",
        }
        df_raw = df_raw.rename(columns=col_fixes)
        # Supprimer les colonnes entièrement vides
        df_raw = df_raw.dropna(axis=1, how="all")
        progress_bar.progress(10, text="Fichier lu.")
    except Exception as e:
        st.error(f"Impossible de lire le fichier : {e}")
        st.stop()

    n_lignes   = len(df_raw)
    n_factures = df_raw["Numéro facture"].nunique() if "Numéro facture" in df_raw.columns else 0

    st.success(f"✅ Fichier chargé : **{n_lignes:,}** lignes · **{n_factures:,}** factures uniques")

    # ── Moteur v16 ───────────────────────────────────────────────
    status_text.info("⚙️ Construction des cycles et attribution des statuts …")
    progress_bar.progress(25, text="Construction des cycles …")

    try:
        with st.spinner("Analyse v16 en cours (peut prendre 1–2 min selon la taille) …"):
            df_enrichi, df_anomalies = analyser_base(df_raw)
        progress_bar.progress(85, text="Calcul des z-scores …")
    except Exception as e:
        st.error(f"Erreur pendant l'analyse : {e}")
        st.markdown("**Colonnes détectées dans votre fichier :**")
        st.code(", ".join(str(c) for c in df_raw.columns.tolist()))
        st.markdown("Copiez cette liste et envoyez-la pour que le moteur soit adapté.")
        st.exception(e)
        st.stop()

    progress_bar.progress(100, text="Analyse terminée ✅")
    status_text.success("✅ Analyse v16 terminée.")


    # ─────────────────────────────────────────────────────────────
    # 3 · Métriques globales
    # ─────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("3 · Vue d'ensemble")

    fac_uniques = df_enrichi.drop_duplicates("Numéro facture")
    total_fac   = len(fac_uniques)
    n_ko_i      = (fac_uniques["Statut"] == "KO Intraday").sum()
    n_ko_u      = (fac_uniques["Statut"] == "KO Ultérieur").sum()
    n_ko_hb     = (fac_uniques["Statut"] == "KO Hors base").sum()
    n_ok        = (fac_uniques["Statut"] == "OK").sum()
    n_critique  = (df_anomalies["z_max"] >= 3.0).sum() if not df_anomalies.empty else 0
    n_eleve     = ((df_anomalies["z_max"] >= 2.0) & (df_anomalies["z_max"] < 3.0)).sum() if not df_anomalies.empty else 0

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card">
        <div class="label">Total factures</div>
        <div class="value">{total_fac:,}</div>
      </div>
      <div class="metric-card">
        <div class="label">KO Intraday</div>
        <div class="value rouge">{n_ko_i:,}</div>
      </div>
      <div class="metric-card">
        <div class="label">KO Ultérieur</div>
        <div class="value orange">{n_ko_u:,}</div>
      </div>
      <div class="metric-card">
        <div class="label">KO Hors base</div>
        <div class="value orange">{n_ko_hb:,}</div>
      </div>
      <div class="metric-card">
        <div class="label">OK</div>
        <div class="value vert">{n_ok:,}</div>
      </div>
      <div class="metric-card">
        <div class="label">Praticiens Critique</div>
        <div class="value rouge">{n_critique}</div>
      </div>
      <div class="metric-card">
        <div class="label">Praticiens Élevé</div>
        <div class="value orange">{n_eleve}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Distribution des statuts
    with st.expander("📊 Distribution des statuts x détail"):
        dist = (fac_uniques.groupby(["Statut", "Detail"])
                            .size()
                            .reset_index(name="Nb factures")
                            .sort_values("Nb factures", ascending=False))
        st.dataframe(dist, use_container_width=True, hide_index=True)


    # ─────────────────────────────────────────────────────────────
    # 4 · Table des anomalies
    # ─────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("4 · Praticiens anomaliques")

    AXES = ["I_total", "U_total", "I_double", "U_double",
            "I_annul",  "U_annul",  "I_nr",    "U_nr"]

    LABELS = {
        "I_total":  ("KO Global", "Intraday"),
        "U_total":  ("KO Global", "Ultérieur"),
        "I_double": ("Double Fact.", "Intraday"),
        "U_double": ("Double Fact.", "Ultérieur"),
        "I_annul":  ("Annul. Sèche", "Intraday"),
        "U_annul":  ("Annul. Sèche", "Ultérieur"),
        "I_nr":     ("Avoir Non Refact.", "Intraday"),
        "U_nr":     ("Avoir Non Refact.", "Ultérieur"),
    }

    if df_anomalies.empty:
        st.info("Aucun praticien ne présente de signal anomalique (z ≥ 2σ) dans cette base.")
    else:
        # ── Filtres ──────────────────────────────────────────────
        col_f1, col_f2 = st.columns([1, 2])
        with col_f1:
            filtre_signal = st.selectbox(
                "Filtrer par niveau de signal",
                ["Tous", "Critique (z≥3)", "Élevé (2≤z<3)"]
            )
        with col_f2:
            centres_dispo = sorted(df_anomalies["centre"].dropna().unique())
            centres_sel   = st.multiselect("Filtrer par centre", centres_dispo, default=centres_dispo)

        # ── Application des filtres ───────────────────────────────
        dfA = df_anomalies.copy()
        if filtre_signal == "Critique (z≥3)":
            dfA = dfA[dfA["z_max"] >= 3.0]
        elif filtre_signal == "Élevé (2≤z<3)":
            dfA = dfA[(dfA["z_max"] >= 2.0) & (dfA["z_max"] < 3.0)]
        if centres_sel:
            dfA = dfA[dfA["centre"].isin(centres_sel)]

        st.caption(f"**{len(dfA)}** praticien(s) affiché(s)")

        # ── Construction du tableau d'affichage ──────────────────
        def signal_badge(sig):
            if sig == "Critique":
                return "🔴 Critique"
            elif sig == "Élevé":
                return "🟡 Élevé"
            return ""

        rows = []
        for _, r in dfA.iterrows():
            row = {
                "Praticien": f"{r.get('nom','')} {r.get('prenom','')}".strip(),
                "Centre":    r.get("centre", ""),
                "Nb factures": int(r.get("nb_fac", 0)),
            }
            for a in AXES:
                cat, timing = LABELS[a]
                taux  = r.get(f"tx_{a}", 0) * 100
                z     = r.get(f"z_{a}", 0)
                sig   = r.get(f"sig_{a}", "")
                col_k = f"{cat} {timing} — Taux"
                col_s = f"{cat} {timing} — Signal"
                row[col_k] = f"{taux:.1f}%"
                row[col_s] = signal_badge(sig) if sig else "—"
            rows.append(row)

        df_table = pd.DataFrame(rows)

        st.dataframe(
            df_table,
            use_container_width=True,
            hide_index=True,
            height=min(600, 60 + 35 * len(df_table)),
        )

        # ── Export Excel anomalies ───────────────────────────────
        st.markdown("---")
        st.subheader("Export — tableau anomalies praticiens")
        buf_anom = io.BytesIO()
        df_table.to_excel(buf_anom, index=False, engine="openpyxl")
        buf_anom.seek(0)
        st.download_button(
            "⬇️ Télécharger le tableau anomalies (.xlsx)",
            data=buf_anom,
            file_name="anomalies_praticiens.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


    # ─────────────────────────────────────────────────────────────
    # 5 · Export Excel complet (Base enrichie + Synthèse)
    # ─────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("5 · Télécharger l'export complet")

    st.markdown("""
Le fichier Excel contient **deux onglets** :
- 📋 **Base enrichie** — toutes les lignes avec Statut, Détail et Statut × Détail
- 📊 **Synthèse** — tableau Statut×Détail × Centre avec formules, pourcentages et mise en forme conditionnelle
    """)

    with st.spinner(f"⏳ Génération du fichier Excel ({len(df_enrichi):,} lignes) — quelques secondes…"):
        try:
            excel_data = generer_excel(df_enrichi, df_anomalies)
        except Exception as _e:
            import traceback
            st.error(f"Erreur génération Excel : {_e}\n\n```\n{traceback.format_exc()}\n```")
            excel_data = None

    if excel_data:
        st.success("✅ Fichier prêt !")
    st.download_button(
        "📥 Télécharger Base enrichie + Synthèse (.xlsx)",
        data=excel_data or b"",
        disabled=excel_data is None,
        file_name="analyse_facturation_v16.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Export CSV séparé (optionnel)
    with st.expander("🗂️ Export CSV uniquement (base enrichie)"):
        buf_csv = io.BytesIO()
        df_enrichi.to_csv(buf_csv, index=False, sep=";", encoding="utf-8-sig")
        buf_csv.seek(0)
        st.download_button(
            "⬇️ Télécharger base enrichie (.csv)",
            data=buf_csv,
            file_name="base_enrichie_v16.csv",
            mime="text/csv",
        )

    # ─────────────────────────────────────────────────────────────
    # 6 · Notes méthodologiques
    # ─────────────────────────────────────────────────────────────
    with st.expander("ℹ️ Méthodologie v16"):
        st.markdown("""
**Classification des cycles (hiérarchie v16)**

| Priorité | Type | Critères |
|---|---|---|
| 1 | Panier exact | Même dossier, même montant, même panier |
| 2 | Reventilation | Même dossier, même montant, panier différent |
| 3 | Codes identiques ≤30j | ≥1 code commun, délai ≤ 30j, montant différent |
| 4 | Montant différent | ≥1 code commun, délai ≤ 30j |
| 5 | Double facturation | Deux factures positives similaires à ≤1j |
| 6 | Annulation sèche | Aucune refacturation trouvée |

**Statuts**
- **KO Intraday** : l'avoir et la facturation initiale ont la même date (délai = 0 j)
- **KO Ultérieur** : délai ≥ 1 jour entre facturation initiale et avoir
- **KO Hors base** : avoir sans facture initiale identifiable

**Détection des anomalies**
- Population de référence : praticiens ayant ≥ 50 factures
- 8 axes analysés : taux KO Intraday/Ultérieur × Global / Double facturation / Annulation sèche / Avoir non refacturé
- Seuils z-score : **Élevé** z ≥ 2σ · **Critique** z ≥ 3σ
- Le dénominateur inclut TOUTES les factures (y compris les avoirs à honoraires négatifs)

**Synthèse Excel**
- Les formules COUNTIFS de l'onglet Synthèse pointent sur l'onglet Base enrichie
- La colonne « Analytique » identifie le centre de chaque acte
- Les pourcentages sont calculés par rapport au total des lignes du centre (pas des factures uniques)
        """)

else:
    # ── État initial — pas de fichier ─────────────────────────
    st.info("👆 Charger un fichier CSV pour démarrer l'analyse.")

    with st.expander("Format attendu"):
        st.markdown("""
Le fichier CSV doit contenir au minimum les colonnes suivantes :

| Colonne | Type | Description |
|---|---|---|
| `Numéro facture` | entier | Identifiant unique de la facture |
| `Date facture` | date (JJ/MM/AAAA) | Date d'émission |
| `Praticien` | texte | Code praticien |
| `Nom praticien` | texte | Nom de famille |
| `Prénom praticien` | texte | Prénom |
| `Honoraire acte` | décimal | Montant de l'acte (négatif = avoir) |
| `Code acte` | texte | Code CCAM ou NGAP |
| `Panier` | texte | Famille de soins |
| `Analytique` | texte | Code centre |
| `Numero_dossier` | texte | Numéro du dossier patient |
        """)
