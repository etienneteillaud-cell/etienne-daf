"""
moteur_v16.py — Algorithme de classification des factures dentaires v16
Groupe Efficentres Horizon

Usage:
    from moteur_v16 import analyser_base
    df_enrichi, df_anomalies = analyser_base(df_input)

Entrée  : DataFrame pandas issu de la base brute (voir COLONNES_REQUISES)
Sorties : (1) df_enrichi — base complète avec colonnes Statut, Detail, Statut x Détail
          (2) df_anomalies — DataFrame des 30 praticiens avec signaux ≥ 2σ
"""

import pandas as pd
import numpy as np
from typing import Tuple

# ─────────────────────────────────────────────
# Colonnes attendues dans la base brute
# ─────────────────────────────────────────────
COLONNES_REQUISES = [
    "Numéro facture",
    "Date facture",
    "Praticien",
    "Honoraire acte",
    "Code acte",
    "Panier",
    "BSS",
    "AMO",
    "AMC",
    "Analytique",
    "Numero_dossier",        # ou "Numéro dossier" — normalisé à l'import
]

# ─────────────────────────────────────────────
# 1. NORMALISATION DE LA BASE BRUTE
# ─────────────────────────────────────────────

def normaliser(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Normalise les noms de colonnes et les types."""
    df = df_raw.copy()

    # ── Normalisation robuste de TOUTES les colonnes clés ────────
    # Construit un mapping depuis les noms réels vers les noms canoniques
    # en ignorant accents, casse, espaces, caractères spéciaux
    import unicodedata, re

    def simplifier(s: str) -> str:
        """Supprime accents, met en minuscule, garde lettres/chiffres."""
        s = unicodedata.normalize("NFKD", str(s))
        s = "".join(c for c in s if not unicodedata.combining(c))
        return re.sub(r"[^a-z0-9]", "", s.lower())

    CANONIQUE = {
        # Colonne canonique → variantes simplifiées acceptées
        # Inclut les versions SANS accents (cas où l'encodage les a perdus)
        "Numéro facture":      ["numerofacture", "numrofacture", "numfacture",
                                "nfacture", "facture", "numerofacture"],
        "Date facture":        ["datefacture"],
        "Praticien":           ["praticien", "codepraticien"],
        "Nom praticien":       ["nompraticien"],
        "Prénom praticien":    ["prenompraticien", "prenompr", "prenomraticien",
                                "prenompatient", "prenom"],
        "Honoraire acte":      ["honoraireacte", "honoraire"],
        "Code acte":           ["codeacte"],
        "Panier":              ["panier"],
        "BSS":                 ["bss"],
        "AMO":                 ["amo"],
        "AMC":                 ["amc"],
        "Analytique":          ["analytique"],
        "Numero_dossier":      ["numerodossier", "numdossier", "dossier", "ndossier"],
        "Date réalisation acte": ["daterealisationacte", "datereali",
                                  "daterealisactionacte", "dateralisation",
                                  "dateralisation", "daterealisacte",
                                  "dateralisationacte"],
        "Nom praticien":       ["nompraticien"],
    }

    # Construire mapping nom_réel → nom_canonique
    rename_map = {}
    for col_reel in df.columns:
        simp = simplifier(col_reel)
        for canon, variantes in CANONIQUE.items():
            if simp in variantes or simp == simplifier(canon):
                if col_reel != canon:
                    rename_map[col_reel] = canon
                break

    df = df.rename(columns=rename_map)

    # Dates
    for col in ["Date facture", "Date réalisation acte"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

    # Numériques
    for col in ["Honoraire acte", "BSS", "AMO", "AMC", "AMC2"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ─────────────────────────────────────────────
# 2. CONSTRUCTION DES CYCLES v16
# ─────────────────────────────────────────────

def construire_cycles(df: pd.DataFrame) -> pd.DataFrame:
    """
    Identifie les cycles avoir↔facturation initiale↔refacturation.

    Hiérarchie v16 (ordre de priorité décroissant) :
      1. Panier exact    : même dossier + même praticien + même montant exact + même panier
      2. Reventilation   : même dossier + même praticien + montant identique + codes différents (changement de panier)
      3. Codes identiques: même dossier + ≥1 code commun + délai ≤ 30j + montant différent
      4. Montant différent: avoir le jour J ou soin antérieur + ≥1 code commun + refactu ≤ 30j
      5. Double facturation: deux factures positives même dossier/praticien/montant, délai ≤ 1j
      6. Annulation sèche: avoir sans aucune refacturation trouvée

    Retourne un DataFrame avec une ligne par avoir (Fac_avoir), contenant les
    références des factures liées et des indicateurs booléens.
    """
    # ── Agrégations par facture ──────────────────────────────────
    hon  = df.groupby("Numéro facture")["Honoraire acte"].sum()
    bss  = df.groupby("Numéro facture")["BSS"].sum()
    fdate = df.groupby("Numéro facture")["Date facture"].first()
    fprat = df.groupby("Numéro facture")["Praticien"].first()
    fdoss = df.groupby("Numéro facture")["Numero_dossier"].first()
    fpan  = df.groupby("Numéro facture")["Panier"].first() if "Panier" in df.columns else None
    fcodes = df.groupby("Numéro facture")["Code acte"].apply(frozenset)

    # ── Séparer avoirs et factures positives ─────────────────────
    avoirs_idx   = hon[hon < 0].index.tolist()
    positifs_idx = hon[hon > 0].index.tolist()

    # Index rapide par (praticien, dossier)
    from collections import defaultdict
    pos_by_prat_doss = defaultdict(list)
    for f in positifs_idx:
        key = (fprat.get(f), fdoss.get(f))
        pos_by_prat_doss[key].append(f)

    # ── Résultat ─────────────────────────────────────────────────
    rows = []

    # Ensemble des factures positives déjà utilisées comme init
    # (une init ne peut servir qu'une fois)
    init_used = set()
    # Ensemble des factures positives déjà utilisées comme refactu
    refactu_used = set()

    for fa in avoirs_idx:
        prat  = fprat.get(fa)
        doss  = fdoss.get(fa)
        hon_a = abs(hon[fa])
        date_a = fdate.get(fa)
        codes_a = fcodes.get(fa, frozenset())
        pan_a = fpan.get(fa) if fpan is not None else None

        candidats = [f for f in pos_by_prat_doss[(prat, doss)]
                     if f not in init_used and f not in refactu_used]

        row = {
            "Fac_avoir": fa,
            "Fac_init": None,
            "Fac_refactu": None,
            "Fac_refactu_codes": None,
            "Fac_refactu_montant": None,
            "Fac_refactu_rev": None,
            "Double": False, "Rev": False,
            "Annul": False, "Refactu_codes": False, "Refactu_montant": False,
        }

        # ── Priorité 1 : panier exact ────────────────────────────
        # Init = même montant exact + même panier
        init_exact = [
            f for f in candidats
            if abs(hon[f] - hon_a) < 0.01
            and (fpan is None or fpan.get(f) == pan_a)
            and f not in init_used
        ]
        # Prendre l'init la plus proche temporellement (avant l'avoir)
        init_exact_avant = [f for f in init_exact
                            if fdate.get(f) is not None
                            and fdate[f] <= date_a]
        if not init_exact_avant:
            init_exact_avant = init_exact  # fallback sans contrainte date

        if init_exact_avant:
            fi = min(init_exact_avant, key=lambda f: abs((fdate[f] - date_a).days))

            # Chercher une refacturation (montant identique, codes identiques, après avoir)
            refactu_exact = [
                f for f in candidats
                if f != fi
                and f not in refactu_used
                and abs(hon[f] - hon_a) < 0.01
                and fcodes.get(f, frozenset()) == fcodes.get(fi, frozenset())
                and fdate.get(f) is not None
                and fdate[f] >= date_a
            ]
            if refactu_exact:
                fr = min(refactu_exact, key=lambda f: fdate[f])

                # Vérifier si c'est une reventilation (panier différent)
                if fpan is not None and fpan.get(fr) != fpan.get(fi):
                    row.update({"Fac_init": fi, "Fac_refactu_rev": [fr], "Rev": True})
                    init_used.add(fi); refactu_used.add(fr)
                else:
                    row.update({"Fac_init": fi, "Fac_refactu": fr})
                    init_used.add(fi); refactu_used.add(fr)
                rows.append(row)
                continue

            # Refactu montant différent (≥1 code commun, ≤30j)
            refactu_diff = [
                f for f in candidats
                if f != fi
                and f not in refactu_used
                and fcodes.get(f, frozenset()) & fcodes.get(fi, frozenset())
                and fdate.get(f) is not None
                and 0 <= (fdate[f] - date_a).days <= 30
            ]
            if refactu_diff:
                fr = min(refactu_diff, key=lambda f: fdate[f])
                row.update({"Fac_init": fi, "Fac_refactu_montant": fr, "Refactu_montant": True})
                init_used.add(fi); refactu_used.add(fr)
                rows.append(row)
                continue

            # Codes identiques (≥1 commun, ≤30j, montant différent)
            refactu_codes = [
                f for f in candidats
                if f != fi
                and f not in refactu_used
                and fcodes.get(f, frozenset()) & fcodes.get(fi, frozenset())
                and fdate.get(f) is not None
                and abs((fdate[f] - date_a).days) <= 30
                and abs(hon.get(f, 0) - hon_a) > 0.01
            ]
            if refactu_codes:
                fr = min(refactu_codes, key=lambda f: abs((fdate[f] - date_a).days))
                row.update({"Fac_init": fi, "Fac_refactu_codes": fr, "Refactu_codes": True})
                init_used.add(fi); refactu_used.add(fr)
                rows.append(row)
                continue

            # Init trouvée mais pas de refactu → annulation sèche
            row.update({"Fac_init": fi, "Annul": True})
            init_used.add(fi)
            rows.append(row)
            continue

        # ── Priorité 5 : double facturation ─────────────────────
        # Deux factures positives même (prat, doss, montant) à ≤1j
        doubles = [
            f for f in candidats
            if f not in init_used
            and abs(hon.get(f, 0) - hon_a) < 0.01
            and fdate.get(f) is not None
            and abs((fdate[f] - date_a).days) <= 1
        ]
        if doubles:
            fi = min(doubles, key=lambda f: abs((fdate[f] - date_a).days))
            row.update({"Fac_init": fi, "Double": True})
            init_used.add(fi)
            rows.append(row)
            continue

        # ── Priorité 6 : annulation sèche sans init connue ───────
        row["Annul"] = True
        rows.append(row)

    return pd.DataFrame(rows)


# ─────────────────────────────────────────────
# 3. ATTRIBUTION DES STATUTS PAR FACTURE
# ─────────────────────────────────────────────

def attribuer_statuts(df: pd.DataFrame, cycles: pd.DataFrame) -> pd.DataFrame:
    """
    Attribue Statut et Detail à chaque ligne du DataFrame.

    Statuts : OK | KO Intraday | KO Ultérieur | KO Hors base
    Details  : OK | AVOIR - Refacturé | AVOIR - Non refacturé |
               AVOIR - Double facturation | REFACTURATION |
               FACTURATION INITIALE - Annulée sans suite |
               FACTURATION INITIALE - Double facturation |
               DOUBLE FACTURATION

    Hiérarchie Option A (un seul statut par facture) :
      avoir > refacturation > facturation initiale > doublon
    """
    hon   = df.groupby("Numéro facture")["Honoraire acte"].sum()
    fdate = df.groupby("Numéro facture")["Date facture"].first()

    def timing(delai_jours: int) -> str:
        return "Intraday" if delai_jours == 0 else "Ultérieur"

    def delai(fa, fi) -> int:
        if fi is None or pd.isna(fi):
            return 0
        fi = int(fi)
        if fa in fdate.index and fi in fdate.index:
            return abs((fdate[fi] - fdate[fa]).days)
        return 0

    # ── Construire les sets de factures et leur statut/detail ────
    fac_statut: dict = {}   # numéro facture → (Statut, Detail)

    def set_statut(fac_num, statut, detail):
        if fac_num not in fac_statut:
            fac_statut[fac_num] = (statut, detail)

    for _, row in cycles.iterrows():
        fa  = int(row["Fac_avoir"])
        fi  = int(row["Fac_init"]) if pd.notna(row.get("Fac_init")) else None

        if row.get("Double"):
            # Double facturation
            d = delai(fa, fi)
            t = timing(d)
            set_statut(fa, f"KO {t}", "AVOIR - Double facturation")
            if fi:
                set_statut(fi, f"KO {t}", "FACTURATION INITIALE - Double facturation")

        elif pd.notna(row.get("Fac_refactu")):
            # Panier exact
            fr = int(row["Fac_refactu"])
            d  = delai(fa, fi)
            t  = timing(d)
            set_statut(fa, f"KO {t}", "AVOIR - Refacturé")
            if fi:
                set_statut(fi, f"KO {t}", "REFACTURATION")
            set_statut(fr, "OK", "REFACTURATION")   # la nouvelle facture est OK

        elif row.get("Rev") and row.get("Fac_refactu_rev") is not None:
            # Reventilation
            fr_list = row["Fac_refactu_rev"]
            if isinstance(fr_list, list):
                d = delai(fa, fi)
                t = timing(d)
                set_statut(fa, f"KO {t}", "AVOIR - Refacturé")
                if fi:
                    set_statut(fi, f"KO {t}", "REFACTURATION")
                for fr in fr_list:
                    set_statut(int(fr), "OK", "REFACTURATION")

        elif pd.notna(row.get("Fac_refactu_codes")):
            # Codes identiques ≤30j
            fr = int(row["Fac_refactu_codes"])
            d  = delai(fa, fi)
            t  = timing(d)
            set_statut(fa, f"KO {t}", "AVOIR - Refacturé")
            if fi:
                set_statut(fi, f"KO {t}", "REFACTURATION")
            set_statut(fr, "OK", "REFACTURATION")

        elif pd.notna(row.get("Fac_refactu_montant")):
            # Montant différent
            fr = int(row["Fac_refactu_montant"])
            d  = delai(fa, fi)
            t  = timing(d)
            set_statut(fa, f"KO {t}", "AVOIR - Refacturé")
            if fi:
                set_statut(fi, f"KO {t}", "REFACTURATION")
            set_statut(fr, "OK", "REFACTURATION")

        elif row.get("Annul"):
            # Annulation sèche
            d = delai(fa, fi)
            t = timing(d)
            set_statut(fa, f"KO {t}", "AVOIR - Non refacturé")
            if fi:
                set_statut(fi, f"KO {t}", "FACTURATION INITIALE - Annulée sans suite")

    # ── Statut hors base : avoirs sans correspondance dans cycles ─
    avoirs_hors_base = set(hon[hon < 0].index) - set(cycles["Fac_avoir"].unique())
    for fa in avoirs_hors_base:
        set_statut(int(fa), "KO Hors base", "AVOIR - Non refacturé")

    # ── Appliquer à chaque ligne ──────────────────────────────────
    def get_statut(fac):
        return fac_statut.get(fac, ("OK", "OK"))

    df = df.copy()
    statuts = df["Numéro facture"].map(lambda f: get_statut(f))
    df["Statut"] = statuts.apply(lambda x: x[0])
    df["Detail"] = statuts.apply(lambda x: x[1])
    df["Statut x Détail"] = df["Statut"] + " " + df["Detail"]

    return df


# ─────────────────────────────────────────────
# 4. DÉTECTION DES ANOMALIES PAR PRATICIEN
# ─────────────────────────────────────────────

DETAILS_DOUBLE = frozenset([
    "DOUBLE FACTURATION",
    "FACTURATION INITIALE - Double facturation",
    "AVOIR - Double facturation",
])
DETAIL_ANNUL = "FACTURATION INITIALE - Annulée sans suite"
DETAIL_NR    = "AVOIR - Non refacturé"
NB_FAC_MIN   = 50    # seuil praticien inclus dans l'analyse
Z_SEUIL_ELEVE    = 2.0
Z_SEUIL_CRITIQUE = 3.0


def detecter_anomalies(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcule, pour chaque praticien ≥ NB_FAC_MIN factures,
    les taux et z-scores sur 8 axes intraday/ultérieur.

    Retourne un DataFrame ordonné par z-score maximal décroissant,
    avec seulement les praticiens ayant au moins un signal ≥ Z_SEUIL_ELEVE.
    """
    # 1 ligne par facture unique — sélection défensive (colonnes optionnelles)
    cols_voulues = ["Praticien", "Nom praticien", "Prénom praticien",
                    "Analytique", "Numéro facture", "Statut", "Detail"]
    cols_presentes = [c for c in cols_voulues if c in df.columns]
    fac = df.drop_duplicates("Numéro facture")[cols_presentes].copy()
    # Colonnes optionnelles absentes → chaîne vide
    for c in cols_voulues:
        if c not in fac.columns:
            fac[c] = ""

    # Praticiens avec suffisamment de factures
    nb = fac.groupby("Praticien").size()
    prats_ok = nb[nb >= NB_FAC_MIN].index
    fac = fac[fac["Praticien"].isin(prats_ok)]

    # Métadonnées praticien
    meta = (fac.groupby("Praticien")
               .agg(nom=("Nom praticien", "first"),
                    prenom=("Prénom praticien", "first"),
                    centre=("Analytique", lambda x: x.value_counts().index[0]))
               .join(nb.rename("nb_fac")))

    # Stats par praticien
    def stats(g):
        total   = len(g)
        ko_I    = g[g["Statut"] == "KO Intraday"]
        ko_U    = g[g["Statut"] == "KO Ultérieur"]
        return pd.Series({
            "total":    total,
            "I_total":  len(ko_I),
            "U_total":  len(ko_U),
            "I_double": len(ko_I[ko_I["Detail"].isin(DETAILS_DOUBLE)]),
            "U_double": len(ko_U[ko_U["Detail"].isin(DETAILS_DOUBLE)]),
            "I_annul":  len(ko_I[ko_I["Detail"] == DETAIL_ANNUL]),
            "U_annul":  len(ko_U[ko_U["Detail"] == DETAIL_ANNUL]),
            "I_nr":     len(ko_I[ko_I["Detail"] == DETAIL_NR]),
            "U_nr":     len(ko_U[ko_U["Detail"] == DETAIL_NR]),
        })

    st = fac.groupby("Praticien").apply(stats).join(meta)

    AXES = ["I_total", "U_total", "I_double", "U_double",
            "I_annul", "U_annul", "I_nr", "U_nr"]

    # Taux et z-scores
    for a in AXES:
        st[f"tx_{a}"] = st[a] / st["total"]

    means = {a: st[f"tx_{a}"].mean() for a in AXES}
    sigmas = {a: st[f"tx_{a}"].std()  for a in AXES}

    for a in AXES:
        s = sigmas[a]
        st[f"z_{a}"] = (st[f"tx_{a}"] - means[a]) / s if s > 0 else 0.0

    # Filtrer praticiens anomaliques
    z_cols = [f"z_{a}" for a in AXES]
    any_signal = (st[z_cols] >= Z_SEUIL_ELEVE).any(axis=1)
    anomalies = st[any_signal].copy()

    # Signal maximal pour trier
    anomalies["z_max"] = anomalies[z_cols].max(axis=1)
    anomalies = anomalies.sort_values("z_max", ascending=False)

    # Libellé signal
    def libelle_signal(z):
        if z >= Z_SEUIL_CRITIQUE: return "Critique"
        if z >= Z_SEUIL_ELEVE:    return "Élevé"
        return ""

    for a in AXES:
        anomalies[f"sig_{a}"] = anomalies[f"z_{a}"].apply(libelle_signal)

    return anomalies.reset_index()


# ─────────────────────────────────────────────
# 5. POINT D'ENTRÉE PRINCIPAL
# ─────────────────────────────────────────────

def analyser_base(df_raw: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Pipeline complet v16.

    Paramètre
    ---------
    df_raw : DataFrame brut issu du CSV export (colonnes COLONNES_REQUISES)

    Retour
    ------
    df_enrichi   : DataFrame complet avec Statut, Detail, Statut x Détail
    df_anomalies : DataFrame des praticiens anomaliques (signaux ≥ 2σ)
    """
    df = normaliser(df_raw)
    cycles = construire_cycles(df)
    df_enrichi = attribuer_statuts(df, cycles)
    df_anomalies = detecter_anomalies(df_enrichi)
    return df_enrichi, df_anomalies


# ─────────────────────────────────────────────
# Auto-test minimal
# ─────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python moteur_v16.py <base.csv>")
        sys.exit(1)

    print(f"Chargement de {sys.argv[1]} …")
    df_raw = pd.read_csv(sys.argv[1], sep=";", low_memory=False)
    print(f"  → {len(df_raw):,} lignes, {df_raw['Numéro facture'].nunique():,} factures uniques")

    print("Analyse v16 en cours …")
    df_enrichi, df_anomalies = analyser_base(df_raw)

    print("\n─── Statuts ───")
    print(df_enrichi["Statut"].value_counts())
    print("\n─── Praticiens anomaliques ───")
    cols = ["Praticien", "centre", "nb_fac", "z_max"] + [f"sig_{a}" for a in
            ["I_total","U_total","I_double","U_double","I_annul","U_annul","I_nr","U_nr"]]
    print(df_anomalies[[c for c in cols if c in df_anomalies.columns]].head(10).to_string())
