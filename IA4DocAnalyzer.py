
import re
import streamlit as st
import pandas as pd
import numpy as np
from pandas import DataFrame as _DataFrame
import os
import io
import zipfile


st.set_page_config(page_title='Analyse Fiches IA4Doc', layout='wide')


# =====================================================================
# STYLES / COULEURS
# =====================================================================

def color_verdict(val):
    if not isinstance(val, str):
        return ""
    v = val.strip().lower()
    if v == "bon":
        return "background-color: #c6efce; color: #006100; font-weight: bold;"
    if "partiellement" in v:
        return "background-color: #ffe699; color: #7f6000; font-weight: bold;"
    if "mauvais" in v:
        return "background-color: #fd4a4a; color: #9c0006; font-weight: bold;"
    return ""


def extract_pct(cell):
    if not isinstance(cell, str):
        return None
    cell = cell.strip()
    if not cell or cell in ("tbd", "NA"):
        return None
    if "(" not in cell or "%" not in cell:
        return None
    try:
        inside = cell.split("(")[1].split(")")[0]   # ex: '0.7%' ou '25.0%'
        inside = inside.replace("%", "").replace(",", ".")
        return float(inside)
    except Exception:
        return None


def color_tab3(val):
    pct = extract_pct(val)
    if pct is None:
        return ""
    # pct est en pourcentage (0–100)
    if pct < 25:
        return "background-color: #fd4a4a;"               # rouge
    elif pct < 50:
        return "background-color: #f4b183;"               # orange
    elif pct < 75:
        return "background-color: #fff2cc;"               # jaune
    elif pct < 90:
        return "background-color: #c6efce;"               # vert clair
    else:
        return "background-color: #00b050; color: white;" # vert foncé


def build_tab2bis(data, ref_base):
    # FS réellement testés
    fs_testes = data["fs_id"].astype(str).str.strip().unique()
    tab2bis = ref_base[ref_base["N° fs"].isin(fs_testes)].copy()

    # moyenne taux_justes par FS
    taux_fs = (
        data.groupby("fs_id")["taux_justes"]
        .mean()
        .rename(lambda x: str(x).strip())
        .to_dict()
    )

    fo_cols = ["fo1", "fo2", "fo3", "fo4", "fo5", "fo6", "fo7", "fo8", "fo9"]

    def fill_cell(row, fo):
        val = row[fo]
        if val != "x":
            return ""
        fs = row["N° fs"]
        actual = taux_fs.get(fs)
        if actual is None or pd.isna(actual):
            return ""
        return f"{actual:.1f}%"

    for fo in fo_cols:
        tab2bis[fo] = tab2bis.apply(lambda r: fill_cell(r, fo), axis=1)

    return tab2bis, fo_cols


def style_tab2bis(df, fo_cols):
    styles = _DataFrame("", index=df.index, columns=df.columns)

    for i in df.index:
        for fo in fo_cols:
            val = df.at[i, fo]
            if not isinstance(val, str) or not val:
                continue
            if val in ("tbd", "NA"):
                continue
            try:
                pct = float(val.replace("%", "").replace(",", "."))
            except Exception:
                continue

            if pct < 25:
                styles.at[i, fo] = "background-color: #fd4a4a;"
            elif pct < 50:
                styles.at[i, fo] = "background-color: #f4b183;"
            elif pct < 75:
                styles.at[i, fo] = "background-color: #fff2cc;"
            elif pct < 90:
                styles.at[i, fo] = "background-color: #c6efce;"
            else:
                styles.at[i, fo] = "background-color: #00b050; color: white;"

    return styles


def norm_verdict(x: str) -> str:
    if not isinstance(x, str):
        return "none"
    x = x.strip().lower()
    if x in ("", "nan", "none"):
        return "none"
    if x == "bon":
        return "bon"
    if "partiel" in x:
        return "partiel"
    if "mauvais" in x:
        return "mauvais"
    return "none"



def extract_modification(ref):
    if not isinstance(ref, str):
        return ""
    ref = ref.strip()
    match = re.search(r"(m\d{2})$", ref, re.IGNORECASE)
    return match.group(1) if match else ""



# =====================================================================
# PARSE FICHE
# =====================================================================

def parse_fiche(file):
    filename = file.name

    # Ex: "fs60-IA-v00-TUF-FFP_PV_...-Fiche-v01-CM.xlsx" → fs_id = "fs60"
    fs_id = filename.split("-")[0]

    pattern = r"^fs\d+-IA-v\d+-TUF-(.+)-Fiche-v\d+-[A-Za-z]+\.xlsx$"
    m = re.match(pattern, filename)
    if not m:
        raise ValueError(
            "Nom de fichier invalide. Format attendu : "
            "fsXX-IA-vXX-TUF-<ref>-Fiche-vXX-<initiales>.xlsx"
        )
    ref_from_filename = m.group(1)

    df = pd.read_excel(file, sheet_name="Template Fiche de Test", header=None)

    # -----------------------------------------------------------------
    # Détection du décalage : ligne 18 contient-elle "*Searchable:" ?
    # -----------------------------------------------------------------
    has_searchable_row = False
    try:
        row18 = df.iloc[17]  # ligne 18 (index 17)
        for val in row18:
            if isinstance(val, str) and "Searchable" in val:
                has_searchable_row = True
                break
    except Exception:
        pass

    offset = 0 if has_searchable_row else -1

    def r(base_row):
        # applique le décalage seulement après la ligne 18
        if base_row > 17:
            return base_row + offset
        return base_row

    # ===========================
    # Champs "avant 18"
    # ===========================
    type_test = df.iloc[4, 1] if pd.notna(df.iloc[4, 1]) else None
    type_doc = df.iloc[14, 1] if pd.notna(df.iloc[14, 1]) else None
    ref_coedm = ref_from_filename

    # Label fonctionnalité (ligne 24 → index 23, après 18 → offset)
    label_fonctionnalite = None
    try:
        val_lbl = df.iloc[r(23), 1]
        if pd.notna(val_lbl):
            label_fonctionnalite = val_lbl
    except Exception:
        pass

    if label_fonctionnalite is None or (
        isinstance(label_fonctionnalite, float) and pd.isna(label_fonctionnalite)
    ):
        label_fonctionnalite = fs_id

    # Date (B10 → ligne 10, index 9) = avant 18
    date_test = None
    try:
        val_date = df.iloc[9, 1]
        if not pd.isna(val_date):
            date_test = pd.to_datetime(val_date, errors="coerce")
    except Exception:
        pass

    # Site (B9 → ligne 9, index 8, colonne 1)
    site = None
    try:
        val_site = df.iloc[8, 1]  # B9
        if not pd.isna(val_site):
            site = str(val_site).strip()
            if site == "LTA":
                site = "STMA"
    except Exception:
        pass


    # Classe documentaire (D15 → ligne 15, index 14) = avant 18
    classe_documentaire = None
    try:
        val_cls = df.iloc[14, 3]
        if pd.isna(val_cls):
            classe_documentaire = None
        elif isinstance(val_cls, str):
            v = val_cls.strip()
            if not v:
                classe_documentaire = None
            elif v.lower() == "non":
                classe_documentaire = "Non Searchable"
            elif v.lower() == "oui":
                classe_documentaire = "Searchable"
            else:
                classe_documentaire = v
        else:
            classe_documentaire = val_cls
    except Exception:
        pass

    # Nombre de pages total (ligne 17 → index 16)
    try:
        nb_pages_total = int(df.iloc[16, 1])
    except Exception:
        nb_pages_total = None

    # Commentaire additionnel (ligne 46 → index 45, après 18 → offset)
    commentaire_add = None
    try:
        comment_row = df.iloc[r(45)]
        for val in comment_row:
            if isinstance(val, str) and val.strip():
                commentaire_add = val.strip()
                break
    except Exception:
        pass

    # Verdict doc : lignes 6 à 12 (index 5..11)
    verdict_doc = None
    try:
        for rr in range(5, 12):
            val = df.iloc[rr, 2]
            if isinstance(val, str) and val.strip():
                verdict_doc = val.strip()
                break
    except Exception:
        pass

    # Nom testeur : lignes 6 et 7 (index 5 & 6)
    nom_testeur = None
    try:
        val7 = df.iloc[6, 1]
        val6 = df.iloc[5, 1]
        out = []
        if isinstance(val7, str) and val7.strip():
            out.append(val7.strip())
        if isinstance(val6, str) and val6.strip():
            out.append(val6.strip())
        if out:
            nom_testeur = " ".join(out)
    except Exception:
        pass

    # Fonctionnalité : 25A / 26A avec offset
    fonctionnalite = None
    try:
        cell_26A = df.iloc[r(25), 0]  # ligne 26 (index 25)
        cell_25A = df.iloc[r(24), 0]  # ligne 25 (index 24)
        sentinel = "[NR] Nombre de test de répétabilité requis"

        if isinstance(cell_26A, str) and cell_26A.strip() and cell_26A.strip() != sentinel:
            fonctionnalite = cell_26A.strip()
        elif isinstance(cell_25A, str) and cell_25A.strip():
            fonctionnalite = cell_25A.strip()
    except Exception:
        pass

    # Noms des tests (ligne 31 → index 30, après 18 → offset)
    try:
        tests = list(df.iloc[r(30), 2:6].dropna())
    except Exception:
        tests = []

    records = []
    for k in range(len(tests)):
        col = 2 + k

        # lignes 32..37 → index 31..36 → après 18 → offset
        row_h = r(31)
        row_m = r(32)
        row_j = r(33)
        row_fn = r(34)
        row_fp = r(35)
        row_inc = r(36)

        raw_vals = [
            df.iloc[row_h, col],   # humain
            df.iloc[row_m, col],   # machine
            df.iloc[row_j, col],   # justes
            df.iloc[row_fn, col],  # fn
            df.iloc[row_fp, col],  # fp
            df.iloc[row_inc, col], # incertaines
        ]

        # Test totalement vide → ignoré
        if all(pd.isna(v) for v in raw_vals):
            continue

        def to_float_or_none(v):
            try:
                return float(v) if not pd.isna(v) else None
            except Exception:
                return None

        humain = to_float_or_none(raw_vals[0])
        machine = to_float_or_none(raw_vals[1])
        justes = to_float_or_none(raw_vals[2])
        fn = to_float_or_none(raw_vals[3])
        fp = to_float_or_none(raw_vals[4])
        inc = to_float_or_none(raw_vals[5])

        rec = {
            "fs_id": fs_id,
            "ref_coedm": ref_coedm,
            "date_test": date_test,
            "type_test": type_test,
            "type_document": type_doc,
            "classe_documentaire": classe_documentaire,
            "site": site,
            "label_fonctionnalite": label_fonctionnalite,
            "nb_pages_total": nb_pages_total,
            "commentaire_additionnel": commentaire_add,
            "verdict_doc": verdict_doc,
            "nom_testeur": nom_testeur,
            "fonctionnalite": fonctionnalite,
            "test_label": f"Test {k+1}",
            "temps_humain_s": humain,
            "temps_machine_s": machine,
            "nb_pages_justes": justes,
            "fn": fn,
            "fp": fp,
            "incertaines": inc,
        }

        if nb_pages_total and nb_pages_total != 0:
            rec["taux_justes"] = (justes / nb_pages_total * 100) if justes is not None else None
            rec["taux_fn"] = (fn / nb_pages_total * 100) if fn is not None else None
            rec["taux_fp"] = (fp / nb_pages_total * 100) if fp is not None else None
            rec["taux_incertaines"] = (inc / nb_pages_total * 100) if inc is not None else None
        else:
            rec["taux_justes"] = None
            rec["taux_fn"] = None
            rec["taux_fp"] = None
            rec["taux_incertaines"] = None


        if humain is not None and machine is not None:
            rec["gain_temps_s"] = humain - machine
        else:
            rec["gain_temps_s"] = None

        records.append(rec)

    return pd.DataFrame(records)


# =====================================================================
# INTERFACE STREAMLIT
# =====================================================================

st.title("Analyse automatique des fiches IA4Doc")

uploaded_files = st.file_uploader(
    "Fichiers Excel ou ZIP",
    type=["xlsx", "zip"],
    accept_multiple_files=True
)

def iter_excel_files(uploaded_files):
    """
    Génère des objets fichier Excel (avec un attribut .name)
    à partir de ce que l'utilisateur a uploadé :
    - .xlsx directs
    - .zip contenant des .xlsx
    """
    for up in uploaded_files:
        fname = up.name.lower()

        # 1) Cas ZIP
        if fname.endswith(".zip"):
            try:
                with zipfile.ZipFile(up) as zf:
                    for member in zf.namelist():
                        if not member.lower().endswith(".xlsx"):
                            continue
                        data = zf.read(member)
                        bio = io.BytesIO(data)
                        # On donne un "name" pour que parse_fiche puisse l'utiliser
                        bio.name = os.path.basename(member)
                        yield bio
            except Exception as e:
                st.error(f"Erreur en lisant le zip {up.name} : {e}")

        # 2) Cas Excel direct
        elif fname.endswith(".xlsx"):
            yield up


if uploaded_files:

    # -------------------------------
    # PARSE DES FICHES + anti-doublon
    # -------------------------------
    dfs = []
    seen_files = set()

    for file in iter_excel_files(uploaded_files):
        # file.name existe (xlsx direct ou fichier du zip)
        if file.name in seen_files:
            st.error(f"Fiche déjà chargée : {file.name} (ignorée)")
            continue
        seen_files.add(file.name)

        try:
            df = parse_fiche(file)
            if not df.empty:
                dfs.append(df)
        except Exception as e:
            st.error(f"{file.name} : {e}")

    if not dfs:
        st.error("Aucune fiche valide.")
        st.stop()

    data = pd.concat(dfs, ignore_index=True)
    data["fs_id"] = data["fs_id"].astype(str).str.strip()

    if not dfs:
        st.error("Aucune fiche valide.")
        st.stop()

    data = pd.concat(dfs, ignore_index=True)
    data["fs_id"] = data["fs_id"].astype(str).str.strip()

    # ==================================================================
    # CHARGEMENT RÉFÉRENTIEL (Feuil1 / Feuil2)
    # ==================================================================
    ref_xls = pd.ExcelFile("pourScript-tableauxJeremie.xlsx")

    ref1 = ref_xls.parse("Feuil1")
    ref2 = ref_xls.parse("Feuil2")

    ref1 = ref1.loc[:, ~ref1.columns.str.contains("Unnamed")]
    ref2 = ref2.loc[:, ~ref2.columns.str.contains("Unnamed")]

    comp1 = ref2["complexité.1"] if "complexité.1" in ref2.columns else None
    comp0 = ref2["complexité"] if "complexité" in ref2.columns else None

    if comp1 is not None and comp0 is not None:
        ref2["complexité_unifiee"] = comp1.where(comp1.notna(), comp0)
    elif comp1 is not None:
        ref2["complexité_unifiee"] = comp1
    else:
        ref2["complexité_unifiee"] = comp0

    ref2["complexité_unifiee"] = ref2["complexité_unifiee"].fillna("tbd")

    ref2 = ref2.drop(columns=[c for c in ["complexité", "complexité.1"] if c in ref2.columns])
    ref2 = ref2.rename(columns={"complexité_unifiee": "complexité"})

    ref_full = ref1.merge(ref2, on="N° fs", how="left")
    ref_full["N° fs"] = ref_full["N° fs"].astype(str).str.strip()

    base_cols = ["N° fs", "fo1", "fo2", "fo3", "fo4", "fo5", "fo6", "fo7", "fo8", "fo9", "complexité"]
    ref_base = ref_full[base_cols].copy()

    fo_cols = ["fo1", "fo2", "fo3", "fo4", "fo5", "fo6", "fo7", "fo8", "fo9"]

    def ensure_at_least_one_fo(row: pd.Series) -> pd.Series:

        if not any(str(row[fo]).strip().lower() == "x" for fo in fo_cols):
            row["fo1"] = "x"
        return row

    ref_base = ref_base.apply(ensure_at_least_one_fo, axis=1)


    crit = {
        "pc":  {"nb": 764, "critere": "99%", "echelle": "moins de 1% de FN soit 99% de TN"},
        "lc":  {"nb": 764, "critere": "99%", "echelle": "moins de 1% de FN soit 99% de TN"},
        "c":   {"nb": 252, "critere": "97%", "echelle": "moins de 3% de FN soit 97% de TN"},
        "tbd": {"nb": "tbd", "critere": "tbd", "echelle": ""},
        "NA":  {"nb": "NA", "critere": "NA", "echelle": ""},
    }

    # =====================================================================
    # TABLEAU 1 — Référentiel brut
    # =====================================================================
    st.subheader("Tableau 1 : Référentiel brut")
    show_tab1 = st.toggle("Afficher / masquer tableau 1")
    if show_tab1:
        st.dataframe(ref_base)

    # =====================================================================
    # TABLEAU 2 — Critères par fonctionnalité
    # =====================================================================
    st.subheader("Tableau 2 : Critères de réussite par fonctionnalité")
    show_tab2 = st.toggle("Afficher / masquer tableau 2")
    if show_tab2:
        tab2 = ref_base.copy()

        def convert_cell_percent(row, fo):
            val = row[fo]
            cx = row["complexité"]
            if val != "x":
                return ""
            if pd.isna(cx) or cx in ("tbd", "NA"):
                return "tbd"
            info = crit.get(cx)
            if not info:
                return "tbd"
            return info["critere"]

        for fo in ["fo1", "fo2", "fo3", "fo4", "fo5", "fo6", "fo7", "fo8", "fo9"]:
            tab2[fo] = tab2.apply(lambda row: convert_cell_percent(row, fo), axis=1)

        tab2 = tab2.drop(columns=["complexité"])
        st.dataframe(tab2)

    # =====================================================================
    # TABLEAU 2bis — Taux justes moyen (FS testées)
    # =====================================================================
    st.subheader("Tableau 2bis : Taux de justesse moyen (FS testées)")
    show_tab2bis = st.toggle("Afficher / masquer tableau 2bis")
    if show_tab2bis:
        tab2bis, fo_cols_2bis = build_tab2bis(data, ref_base)

        tab2bis_display = tab2bis[["N° fs"] + fo_cols_2bis]

        styled_2bis = tab2bis_display.style.apply(
            lambda df_: style_tab2bis(df_, fo_cols_2bis),
            axis=None,
            subset=fo_cols_2bis,
        )

        st.dataframe(styled_2bis)

    # =====================================================================
    # TABLEAU 3 — Progression des tests (FS testés uniquement)
    # =====================================================================
    st.subheader("Tableau 3 : Progression des tests (FS testés uniquement)")
    show_tab3 = st.toggle("Afficher / masquer tableau 3")
    if show_tab3:

        fs_testes = data["fs_id"].unique()
        tab3 = ref_base[ref_base["N° fs"].isin(fs_testes)].copy()

        tests_counts = data.groupby("fs_id").size().to_dict()

        tmp = data.copy()
        tmp["vcat"] = tmp["verdict_doc"].apply(norm_verdict)

        tver = tmp.groupby(["fs_id", "vcat"]).size().unstack(fill_value=0)
        tests_vdict = {fs: row.to_dict() for fs, row in tver.iterrows()}

        def convert_cell_progress(row, fo):
            val = row[fo]
            cx = row["complexité"]
            fs = row["N° fs"]
            if val != "x":
                return ""
            done = tests_counts.get(fs, 0)
            info = crit.get(cx, {})
            total = info.get("nb", "tbd")

            if isinstance(total, str) and total == "tbd":
                return f"{done}/tbd"

            if isinstance(total, (int, float)) and total > 0:
                pct = done / total
                return f"{done}/{int(total)} ({pct:.1%})"

            return str(done)

        fo_cols = ["fo1", "fo2", "fo3", "fo4", "fo5", "fo6", "fo7", "fo8", "fo9"]
        for fo in fo_cols:
            tab3[fo] = tab3.apply(lambda row: convert_cell_progress(row, fo), axis=1)

        def ratio_cat(fs, cat):
            fs = str(fs)
            total = tests_counts.get(fs, 0)
            if total == 0:
                return ""
            n = tests_vdict.get(fs, {}).get(cat, 0)
            return f"{int(n)}/{int(total)}"

        tab3["Bon"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "bon"))
        tab3["Partiellement bon"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "partiel"))
        tab3["Mauvais"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "mauvais"))
        tab3["None"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "none"))

        tab3_display = tab3.drop(columns=["complexité"])

        st.dataframe(tab3_display.style.applymap(color_tab3, subset=fo_cols))

    # ------------------------------------------------------
    # Données extraites (nettoyées)
    # ------------------------------------------------------
    st.write("### Données extraites")

    clean_cols = [
        "fs_id",
        "ref_coedm",
        "date_test",
        "nom_testeur",
        "type_test",
        "type_document",
        "classe_documentaire",
        "site",
        "fonctionnalite",
        "nb_pages_total",
        "test_label",
        "temps_humain_s",
        "temps_machine_s",
        "nb_pages_justes",
        "fn",
        "fp",
        "incertaines",
        "taux_justes",
        "taux_fn",
        "taux_fp",
        "taux_incertaines",
        "gain_temps_s",
        "verdict_doc",
    ]

    df_clean = data[clean_cols].copy()

    if "date_test" in df_clean.columns:
        df_clean["date_test"] = pd.to_datetime(
            df_clean["date_test"], errors="coerce"
        ).dt.date

    num_cols = df_clean.select_dtypes(include=["float", "int"]).columns
    df_clean[num_cols] = df_clean[num_cols].round(2)

    st.dataframe(df_clean.style.applymap(color_verdict, subset=["verdict_doc"]))

    # ------------------------------------------------------
    # Commentaires additionnels — menu déroulant
    # ------------------------------------------------------
    st.write("## Commentaires additionnels détectés")

    comment_rows = (
        data[["ref_coedm", "verdict_doc", "commentaire_additionnel"]]
        .dropna(subset=["commentaire_additionnel"])
        .drop_duplicates()
    )

    if comment_rows.empty:
        st.info("Aucun commentaire dans les fiches.")
    else:
        options = []
        for idx, row in comment_rows.iterrows():
            ref = row["ref_coedm"]
            verdict = row["verdict_doc"] or "Non renseigné"
            label = f"{ref} — Verdict : {verdict}"
            options.append((label, idx))

        labels = [o[0] for o in options]
        choice_label = st.selectbox(
            "Sélectionner un document pour voir le commentaire",
            labels,
        )

        chosen_idx = dict(options)[choice_label]
        chosen_row = comment_rows.loc[chosen_idx]

        st.info(
            f"📄 {chosen_row['ref_coedm']} — Verdict : {chosen_row['verdict_doc'] or 'Non renseigné'}\n\n"
            f"📝 Commentaire : {chosen_row['commentaire_additionnel']}"
        )

    # ------------------------------------------------------
    # KPI globaux
    # ------------------------------------------------------
    st.write("## KPI globaux")

    nb_tests_total = len(data)
    nb_fiches = data["ref_coedm"].nunique()
    nb_testeurs = data["nom_testeur"].nunique()

    col1, col2, col3 = st.columns(3)
    col1.metric("Nombre total de tests", nb_tests_total)
    col2.metric("Nombre total de fiches", nb_fiches)
    col3.metric("Nombre de testeurs", nb_testeurs)

    # ------------------------------------------------------
    # KPI par fonctionnalité
    # ------------------------------------------------------
    st.write("## KPI par fonctionnalité")

    kpi_fct = (
        data.groupby("fs_id")
        .agg(
            nb_tests=("test_label", "count"),
            temps_humain_moy=("temps_humain_s", "mean"),
            temps_machine_moy=("temps_machine_s", "mean"),
            taux_justes_moy=("taux_justes", "mean"),
            nb_fiches=("ref_coedm", lambda s: s.nunique()),
        )
        .reset_index()
    )

    tmp_kpi = data.copy()
    tmp_kpi["vcat"] = tmp_kpi["verdict_doc"].apply(norm_verdict)

    counts = (
        tmp_kpi.groupby(["fs_id", "vcat"])
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    for c in ["bon", "partiel", "mauvais", "none"]:
        if c not in counts.columns:
            counts[c] = 0

    counts["total_tests"] = counts[["bon", "partiel", "mauvais", "none"]].sum(axis=1)

    def ratio_str(n, total):
        if total == 0:
            return ""
        return f"{int(n)}/{int(total)}"

    counts["bon_ratio"] = counts.apply(lambda r: ratio_str(r["bon"], r["total_tests"]), axis=1)
    counts["partiel_ratio"] = counts.apply(lambda r: ratio_str(r["partiel"], r["total_tests"]), axis=1)
    counts["mauvais_ratio"] = counts.apply(lambda r: ratio_str(r["mauvais"], r["total_tests"]), axis=1)

    kpi_fct = kpi_fct.merge(
        counts[["fs_id", "bon_ratio", "partiel_ratio", "mauvais_ratio"]],
        on="fs_id",
        how="left",
    )

    kpi_fct = kpi_fct.rename(columns={
        "bon_ratio": "bon",
        "partiel_ratio": "partiellement bon",
        "mauvais_ratio": "mauvais",
    })

    st.dataframe(kpi_fct)

    # ------------------------------------------------------
    # Réussite par classe documentaire et fonctionnalité
    # ------------------------------------------------------
    st.write("## Réussite par classe documentaire et fonctionnalité")

    if "classe_documentaire" in data.columns:
        tmp_cd = data.copy()
        tmp_cd["vcat"] = tmp_cd["verdict_doc"].apply(norm_verdict)

        tot_cd = (
            tmp_cd.groupby(["fs_id", "classe_documentaire"])
            .size()
            .rename("total_tests")
        )

        bon_cd = (
            tmp_cd[tmp_cd["vcat"] == "bon"]
            .groupby(["fs_id", "classe_documentaire"])
            .size()
            .rename("bon_tests")
        )

        cd_df = pd.concat([tot_cd, bon_cd], axis=1).fillna(0)
        cd_df["bon_tests"] = cd_df["bon_tests"].astype(int)
        cd_df["total_tests"] = cd_df["total_tests"].astype(int)

        def fmt_ratio(row):
            if row["total_tests"] == 0:
                return ""
            return f"{row['bon_tests']}/{row['total_tests']}"

        cd_df["ratio"] = cd_df.apply(fmt_ratio, axis=1)

        pivot_cd = cd_df.reset_index().pivot(
            index="fs_id",
            columns="classe_documentaire",
            values="ratio"
        )

        st.dataframe(pivot_cd)
    else:
        st.info("Aucune classe documentaire trouvée dans les fiches.")


    # ------------------------------------------------------
    # Résumé par type de document
    # ------------------------------------------------------
    st.write("## Résumé par type de document")

    summary = (
        data.groupby("type_document")
        .agg(
            nb_tests=("test_label", "count"),
            temps_humain_moy=("temps_humain_s", "mean"),
            temps_machine_moy=("temps_machine_s", "mean"),
            taux_justes_moy=("taux_justes", "mean"),
            taux_fn_moy=("taux_fn", "mean"),
            taux_fp_moy=("taux_fp", "mean"),
        )
        .reset_index()
    )

    st.dataframe(summary)

    # ------------------------------------------------------
    # Taux de justesse par classe documentaire
    # ------------------------------------------------------
    st.write("## Taux de justesse par classe documentaire")

    if "classe_documentaire" in data.columns:
        classe_summary = (
            data.groupby("classe_documentaire")
            .agg(
                nb_tests=("test_label", "count"),
                taux_justes_moy=("taux_justes", "mean"),
            )
            .reset_index()
        )

        classe_summary["taux_justes_moy"] = classe_summary["taux_justes_moy"].round(2)

        st.dataframe(classe_summary)
    else:
        st.info("Aucune classe documentaire trouvée dans les fiches.")

    # ------------------------------------------------------
    # Export suivi recettage (Excel)
    # ------------------------------------------------------
    st.write("## Export suivi recettage")

    def verdict_to_score(v):
        v_norm = norm_verdict(v)
        if v_norm == "bon":
            return 1.0
        if v_norm == "partiel":
            return 0.5
        if v_norm == "mauvais":
            return 0.0
        return np.nan

    recettage_df = data[
        ["fs_id", "verdict_doc", "classe_documentaire", "site", "ref_coedm"]
    ].copy()

    # Score
    recettage_df["score"] = recettage_df["verdict_doc"].apply(verdict_to_score)

    #Modification
    recettage_df["Modification"] = recettage_df["ref_coedm"].apply(extract_modification)

    # Réordonner les colonnes
    recettage_df = recettage_df[
        ["fs_id", "verdict_doc", "score", "classe_documentaire", "site", "Modification"]
    ]

    from io import BytesIO
    output_recettage = BytesIO()
    with pd.ExcelWriter(output_recettage, engine="xlsxwriter") as writer:
        recettage_df.to_excel(writer, sheet_name="Suivi_recettage", index=False)

    st.download_button(
        "Export suivi recettage (Excel)",
        data=output_recettage.getvalue(),
        file_name="suivi_recettage.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ------------------------------------------------------
    # Téléchargement des résultats
    # ------------------------------------------------------
    from io import BytesIO

    # 1) CSV brut des tests
    st.download_button(
        "Exporter les données détaillées (CSV)",
        data=data.to_csv(index=False).encode("utf-8"),
        file_name="resultats_ia_detail.csv",
        mime="text/csv"
    )

    # 2) Excel multi-feuilles avec un peu de mise en forme

    # --- Reconstruire Tableau 2bis (données, sans style Streamlit) ---
    tab2bis_raw, fo_cols_2bis = build_tab2bis(data, ref_base)
    tab2bis_xlsx = tab2bis_raw[["N° fs"] + fo_cols_2bis]

    # --- Reconstruire Tableau 3 (valeurs texte "X/Y (%)") ---
    fs_testes = data["fs_id"].unique()
    tab3 = ref_base[ref_base["N° fs"].isin(fs_testes)].copy()

    tests_counts = data.groupby("fs_id").size().to_dict()

    tmp_export = data.copy()
    tmp_export["vcat"] = tmp_export["verdict_doc"].apply(norm_verdict)

    tver = tmp_export.groupby(["fs_id", "vcat"]).size().unstack(fill_value=0)
    tests_vdict = {fs: row.to_dict() for fs, row in tver.iterrows()}

    def convert_cell_progress_export(row, fo):
        val = row[fo]
        cx = row["complexité"]
        fs = row["N° fs"]
        if val != "x":
            return ""
        done = tests_counts.get(fs, 0)
        info = crit.get(cx, {})
        total = info.get("nb", "tbd")

        if isinstance(total, str) and total == "tbd":
            return f"{done}/tbd"

        if isinstance(total, (int, float)) and total > 0:
            pct = done / total
            return f"{done}/{int(total)} ({pct:.1%})"

        return str(done)

    fo_cols = ["fo1", "fo2", "fo3", "fo4", "fo5", "fo6", "fo7", "fo8", "fo9"]
    for fo in fo_cols:
        tab3[fo] = tab3.apply(lambda row: convert_cell_progress_export(row, fo), axis=1)

    def ratio_cat(fs, cat):
        fs = str(fs)
        total = tests_counts.get(fs, 0)
        if total == 0:
            return ""
        n = tests_vdict.get(fs, {}).get(cat, 0)
        return f"{int(n)}/{int(total)}"

    tab3["Bon"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "bon"))
    tab3["Partiellement bon"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "partiel"))
    tab3["Mauvais"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "mauvais"))
    tab3["None"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "none"))

    tab3_display = tab3.drop(columns=["complexité"])

    # --- Réussite par classe doc & fonctionnalité (pivot_cd) ---
    if "classe_documentaire" in data.columns:
        tmp_cd = data.copy()
        tmp_cd["vcat"] = tmp_cd["verdict_doc"].apply(norm_verdict)

        tot_cd = (
            tmp_cd.groupby(["fs_id", "classe_documentaire"])
            .size()
            .rename("total_tests")
        )

        bon_cd = (
            tmp_cd[tmp_cd["vcat"] == "bon"]
            .groupby(["fs_id", "classe_documentaire"])
            .size()
            .rename("bon_tests")
        )

        cd_df = pd.concat([tot_cd, bon_cd], axis=1).fillna(0)
        cd_df["bon_tests"] = cd_df["bon_tests"].astype(int)
        cd_df["total_tests"] = cd_df["total_tests"].astype(int)

        def fmt_ratio(row):
            if row["total_tests"] == 0:
                return ""
            return f"{row['bon_tests']}/{row['total_tests']}"

        cd_df["ratio"] = cd_df.apply(fmt_ratio, axis=1)

        pivot_cd = cd_df.reset_index().pivot(
            index="fs_id",
            columns="classe_documentaire",
            values="ratio"
        )
    else:
        pivot_cd = pd.DataFrame()

    # df_clean, kpi_fct, summary, classe_summary existent déjà plus haut

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # --- Écrire toutes les feuilles sans style ---
        tab2bis_xlsx.to_excel(writer, sheet_name="Tableau_2bis", index=False)
        tab3_display.to_excel(writer, sheet_name="Tableau_3", index=False)
        df_clean.to_excel(writer, sheet_name="Donnees_extraites", index=False)
        kpi_fct.to_excel(writer, sheet_name="KPI_par_fonctionnalite", index=False)
        pivot_cd.to_excel(writer, sheet_name="Reussite_cls_doc_fct")
        summary.to_excel(writer, sheet_name="Resume_type_document", index=False)
        classe_summary.to_excel(writer, sheet_name="Taux_par_classe_doc", index=False)

        # --- Formats Excel (couleurs) ---
        workbook = writer.book

        fmt_red = workbook.add_format({"bg_color": "#fb3f35"})
        fmt_orange = workbook.add_format({"bg_color": "#F4B183"})
        fmt_yellow = workbook.add_format({"bg_color": "#FFF2CC"})
        fmt_light_green = workbook.add_format({"bg_color": "#C6EFCE"})
        fmt_dark_green = workbook.add_format({"bg_color": "#00B050", "font_color": "#FFFFFF"})

        fmt_verdict_bon = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "bold": True})
        fmt_verdict_partiel = workbook.add_format({"bg_color": "#FFE699", "font_color": "#7F6000", "bold": True})
        fmt_verdict_mauvais = workbook.add_format({"bg_color": "#fb3f35", "font_color": "#9C0006", "bold": True})

        # --------------------------
        # Couleurs dans Tableau_2bis
        # --------------------------
        ws2 = writer.sheets["Tableau_2bis"]
        n_rows_2bis = len(tab2bis_xlsx)

        for r in range(n_rows_2bis):
            for fo in fo_cols_2bis:
                c_idx = tab2bis_xlsx.columns.get_loc(fo)
                val = tab2bis_xlsx.iloc[r, c_idx]
                if not isinstance(val, str) or not val:
                    continue
                if val in ("tbd", "NA"):
                    continue
                try:
                    pct = float(val.replace("%", "").replace(",", "."))
                except Exception:
                    continue

                if pct < 25:
                    fmt = fmt_red
                elif pct < 50:
                    fmt = fmt_orange
                elif pct < 75:
                    fmt = fmt_yellow
                elif pct < 90:
                    fmt = fmt_light_green
                else:
                    fmt = fmt_dark_green

                ws2.write(r + 1, c_idx, val, fmt)  # +1 car ligne 0 = en-têtes

        # --------------------------
        # Couleurs dans Tableau_3
        # --------------------------
        ws3 = writer.sheets["Tableau_3"]
        n_rows_3 = len(tab3_display)

        for r in range(n_rows_3):
            for fo in fo_cols:
                c_idx = tab3_display.columns.get_loc(fo)
                val = tab3_display.iloc[r, c_idx]
                pct = extract_pct(val)
                if pct is None:
                    continue
                if pct < 25:
                    fmt = fmt_red
                elif pct < 50:
                    fmt = fmt_orange
                elif pct < 75:
                    fmt = fmt_yellow
                elif pct < 90:
                    fmt = fmt_light_green
                else:
                    fmt = fmt_dark_green
                ws3.write(r + 1, c_idx, val, fmt)

        # --------------------------
        # Couleurs verdict dans Donnees_extraites
        # --------------------------
        wsD = writer.sheets["Donnees_extraites"]
        n_rows_D = len(df_clean)
        verdict_col_idx = df_clean.columns.get_loc("verdict_doc")

        for r in range(n_rows_D):
            val = df_clean.iloc[r, verdict_col_idx]
            if not isinstance(val, str):
                continue
            v = val.strip().lower()
            if v == "bon":
                fmt = fmt_verdict_bon
            elif "partiellement" in v:
                fmt = fmt_verdict_partiel
            elif "mauvais" in v:
                fmt = fmt_verdict_mauvais
            else:
                continue
            wsD.write(r + 1, verdict_col_idx, val, fmt)

    st.download_button(
        "Exporter tous les tableaux (Excel)",
        data=output.getvalue(),
        file_name="IA4Doc_tableaux_complets.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )







