# -*- coding: utf-8 -*-
"""
Created on Wed Dec  3 20:34:46 2025

@author: ravis
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO


# ============================================
# Helpers: read Excel and common saldo logic
# ============================================

def load_excel_range(file, start_row, end_row):
    """
    Read Excel range B[start_row]:E[end_row] with no header,
    and return a DataFrame with columns:
    ['Namen', 'Datum', 'Debet', 'Credit'].
    """
    nrows = max(end_row - start_row + 1, 0)
    if nrows <= 0:
        return pd.DataFrame(columns=["Namen", "Datum", "Debet", "Credit"])

    df = pd.read_excel(
        file,
        sheet_name=0,
        header=None,
        usecols="B:E",
        skiprows=start_row - 1,
        nrows=nrows,
        dtype=str
    )
    df.columns = ["Namen", "Datum", "Debet", "Credit"]

    # Normalize decimals and convert to numeric
    for col in ["Debet", "Credit"]:
        df[col] = df[col].str.replace(",", ".", regex=False)
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Drop rows without any numeric content
    df = df.dropna(subset=["Debet", "Credit"], how="all")

    return df


def compute_saldo_tables(df):
    """
    For each unique 'Namen', compute a detail table with a 'Saldo' row and
    a final 'Total' row, and a summary table 'total_saldo' with
    per-name debet/credit saldo + overall total.
    Returns:
        tables: dict{name -> detail DataFrame}
        total_saldo: DataFrame with columns ['Naam','Debet','Credit']
    """
    tables = {}
    total_rows = []

    # Ensure required cols
    for c in ["Namen", "Datum", "Debet", "Credit"]:
        if c not in df.columns:
            raise ValueError(f"Missing column: {c}")

    # Group by name
    for name, sub in df.groupby("Namen"):
        sub = sub.copy()
        # Keep only relevant columns
        sub2 = sub[["Datum", "Debet", "Credit"]].copy()

        total_debet = sub2["Debet"].sum(skipna=True)
        total_credit = sub2["Credit"].sum(skipna=True)
        diff = (total_credit or 0) - (total_debet or 0)

        debet_saldo = max(diff, 0)
        credit_saldo = max(-diff, 0)

        # Add saldo row
        saldo_row = {"Datum": "Saldo", "Debet": debet_saldo, "Credit": credit_saldo}
        sub2 = pd.concat([sub2, pd.DataFrame([saldo_row])], ignore_index=True)

        # Add total row (like janitor::adorn_totals)
        total_row = {
            "Datum": "Total",
            "Debet": sub2["Debet"].sum(skipna=True),
            "Credit": sub2["Credit"].sum(skipna=True),
        }
        sub2 = pd.concat([sub2, pd.DataFrame([total_row])], ignore_index=True)

        tables[name] = sub2

        total_rows.append({"Naam": name, "Debet": debet_saldo, "Credit": credit_saldo})

    total_saldo = pd.DataFrame(total_rows)

    # Remove zero-zero saldo rows
    non_zero = (total_saldo["Debet"] != 0) | (total_saldo["Credit"] != 0)
    total_saldo_nz = total_saldo[non_zero].copy()

    # Total line at the bottom
    total_row = {
        "Naam": "Total",
        "Debet": total_saldo_nz["Debet"].sum(skipna=True),
        "Credit": total_saldo_nz["Credit"].sum(skipna=True),
    }
    total_saldo_final = pd.concat(
        [total_saldo_nz, pd.DataFrame([total_row])], ignore_index=True
    )

    return tables, total_saldo_final


def write_tables_to_excel_kruis(
    buffer,
    tables_kpkn,
    total_saldo1,
    tables_other,
    total_saldo2,
):
    """
    Build the Excel workbook for choice == 'Kruisposten'.
    """
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        sheet_name = "Data"
        row = 0

        # Section 1: KP/KN accounts
        for name in sorted(tables_kpkn.keys()):
            df = tables_kpkn[name].copy()
            # Add a header row with the name
            header_df = pd.DataFrame([[f"{name}"]], columns=["Rekening"])
            header_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=1,
                index=False,
                header=False,
            )
            row += 1

            df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=0,
                index=False,
            )
            row += len(df) + 2

        # Total saldo for KP/KN
        if not total_saldo1.empty:
            header_df = pd.DataFrame([["Total all saldo KP/KN"]], columns=[""])
            header_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=1,
                index=False,
                header=False,
            )
            row += 1
            total_saldo1.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=0,
                index=False,
            )
            row += len(total_saldo1) + 3

        # Section 2: Data zonder KP of KN
        header_df = pd.DataFrame([["Data zonder KP of KN"]], columns=[""])
        header_df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=row,
            startcol=1,
            index=False,
            header=False,
        )
        row += 2

        for name in sorted(tables_other.keys()):
            df = tables_other[name].copy()
            sub_header = pd.DataFrame([[f"{name}"]], columns=["Rekening"])
            sub_header.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=1,
                index=False,
                header=False,
            )
            row += 1
            df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=0,
                index=False,
            )
            row += len(df) + 2

        # Total saldo for non-KP/KN
        if not total_saldo2.empty:
            header_df = pd.DataFrame([["Total all saldo (zonder KP/KN)"]], columns=[""])
            header_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=1,
                index=False,
                header=False,
            )
            row += 1
            total_saldo2.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=0,
                index=False,
            )
            row += len(total_saldo2) + 3

        # Combined open saldo list
        # D1: all open balances
        if not total_saldo1.empty or not total_saldo2.empty:
            combined = pd.concat(
                [
                    total_saldo1[total_saldo1["Naam"] != "Total"],
                    total_saldo2[total_saldo2["Naam"] != "Total"],
                ],
                ignore_index=True,
            )

            if not combined.empty:
                total_row = {
                    "Naam": "Total",
                    "Debet": combined["Debet"].sum(skipna=True),
                    "Credit": combined["Credit"].sum(skipna=True),
                }
                combined_all = pd.concat(
                    [combined, pd.DataFrame([total_row])], ignore_index=True
                )

                header_df = pd.DataFrame(
                    [["Lijst met alle openstaande saldi"]], columns=[""]
                )
                header_df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=row,
                    startcol=1,
                    index=False,
                    header=False,
                )
                row += 2

                combined_all.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=row,
                    startcol=0,
                    index=False,
                )
                row += len(combined_all) + 3

                # D2: only > 5 SRD
                mask = (combined["Debet"] > 5) | (combined["Credit"] > 5)
                filtered = combined[mask].copy()
                if not filtered.empty:
                    total_row2 = {
                        "Naam": "Total",
                        "Debet": filtered["Debet"].sum(skipna=True),
                        "Credit": filtered["Credit"].sum(skipna=True),
                    }
                    filtered_all = pd.concat(
                        [filtered, pd.DataFrame([total_row2])], ignore_index=True
                    )

                    header_df = pd.DataFrame(
                        [["Lijst met alle openstaande saldi van meer dan 5 SRD"]],
                        columns=[""],
                    )
                    header_df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        startrow=row,
                        startcol=1,
                        index=False,
                        header=False,
                    )
                    row += 2
                    filtered_all.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        startrow=row,
                        startcol=0,
                        index=False,
                    )

        writer.close()
    buffer.seek(0)
    return buffer


def write_tables_to_excel_simple(buffer, tables, total_saldo, title_prefix):
    """
    Build the Excel workbook for Crediteuren / Debiteuren / Voorschotten:
    - detail tables per name
    - total_saldo
    - list of open balances > 5 SRD
    """
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        sheet_name = "Data"
        row = 0

        for name in sorted(tables.keys()):
            df = tables[name].copy()
            header_df = pd.DataFrame([[f"{name}"]], columns=["Rekening"])
            header_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=1,
                index=False,
                header=False,
            )
            row += 1
            df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=0,
                index=False,
            )
            row += len(df) + 2

        # Total saldo for all names
        if not total_saldo.empty:
            header_df = pd.DataFrame([[f"{title_prefix} - Total all saldo"]], columns=[""])
            header_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=1,
                index=False,
                header=False,
            )
            row += 1
            total_saldo.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=row,
                startcol=0,
                index=False,
            )
            row += len(total_saldo) + 3

            # > 5 SRD list
            non_total = total_saldo[total_saldo["Naam"] != "Total"].copy()
            mask = (non_total["Debet"] > 5) | (non_total["Credit"] > 5)
            filtered = non_total[mask].copy()
            if not filtered.empty:
                total_row = {
                    "Naam": "Total",
                    "Debet": filtered["Debet"].sum(skipna=True),
                    "Credit": filtered["Credit"].sum(skipna=True),
                }
                filtered_all = pd.concat(
                    [filtered, pd.DataFrame([total_row])], ignore_index=True
                )
                header_df = pd.DataFrame(
                    [["Lijst met alle openstaande saldi van meer dan 5 SRD"]],
                    columns=[""],
                )
                header_df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=row,
                    startcol=1,
                    index=False,
                    header=False,
                )
                row += 2
                filtered_all.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=row,
                    startcol=0,
                    index=False,
                )

        writer.close()
    buffer.seek(0)
    return buffer


# ============================================
# Cleaning functions per choice
# ============================================

def preprocess_kruis(df):
    """
    Implements the KP/KN logic from the R 'Kruisposten' branch.
    Returns: used_file_kpkn, used_file_other
    """
    df = df.copy()

    # Detect KP and KN patterns
    kp_mask = df["Namen"].str.contains(r"( KP|\.KP)", na=False)
    kn_mask = df["Namen"].str.contains(r"( KN|\.KN)", na=False)

    # Extract parts after KP or KN
    kp_part = df["Namen"].where(kp_mask).str.replace(r".* KP", "", regex=True)
    kp_part = kp_part.where(kp_mask).str.replace(r".*\.KP", "", regex=True)

    kn_part = df["Namen"].where(kn_mask).str.replace(r".* KN", "", regex=True)
    kn_part = kn_part.where(kn_mask).str.replace(r".*\.KN", "", regex=True)

    # Normalize KP/KN part
    def normalize_part(s, prefix):
        # If None or NaN → skip
        if s is None or (isinstance(s, float) and np.isnan(s)):
            return np.nan

        # Convert safely to string
        s = str(s)

        # Replace '=' with '-'
        s = s.replace("=", "-")

        # Extract the first sequence of digits
        m = re.search(r"([0-9]+)", s)
        if not m:
            return np.nan

        num = m.group(1)

        # Ensure leading "-" as in your R logic
        if "-" not in num:
            num = "-" + num

        num = num.replace(" ", "")
        return prefix + num

    # Apply robustly
    kp_code = kp_part.apply(lambda x: normalize_part(x, "KP"))
    kn_code = kn_part.apply(lambda x: normalize_part(x, "KN"))

    df["variables2"] = kp_code
    df["variable5"] = kn_code

    # Combine KP/KN
    df["variables"] = df["variables2"]
    df.loc[df["variable5"].notna(), "variables"] = df.loc[df["variable5"].notna(), "variable5"]

    # Rows WITHOUT KP/KN codes
    used_file2 = df.loc[df["variables"].isna(), ["Namen", "Datum", "Debet", "Credit"]].copy()
    used_file2 = used_file2.dropna(subset=["Namen"])
    used_file2 = used_file2.sort_values("Namen")

    # Rows WITH KP/KN codes
    used_file = df.loc[df["variables"].notna(), ["variables", "Datum", "Debet", "Credit"]].copy()
    used_file = used_file.rename(columns={"variables": "Namen"})
    used_file = used_file.sort_values("Namen", ascending=False)

    return used_file, used_file2




def preprocess_cred(df):
    df = df.copy()
    s = df["Namen"].fillna("").astype(str)
    
    # 1) inhoud opschonen (woorden/tokens verwijderen)
    patterns = [
    r"\bAANK\b",
    r"\bRIJST\b",
    r"\bCOMM\w*\b",
    r"\bCARGO\w*\b",
    r"\bHN\w*\b",
    r"\bPARB\w*\b",
    r"\bBREUK\w*\b",
    r"\bKEUR\w*\b",
    r"\bBESTR\w*\b",
    r"\bMT\b\.?",          # MT of MT.
    r"\bI\b",              # losse letter I
    r"\d+",                # cijfers
    r"[\$\/\-,\.]",        # symbols/punct (komma/punt ook hier)
    ]
    
    for pat in patterns:
        s = s.str.replace(pat, " ", regex=True)
    
    # 2) whitespace normaliseren
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    
    # 3) ABaboelal -> A Baboelal (CamelCase)
    s = s.str.replace(r"^([A-Z]{1,2})([A-Z][a-z].+)$", r"\1 \2", regex=True)
    
    # 4) "Baboelal A" / "Baboelal A." -> "A Baboelal"
    parts = s.str.split()
    
    last = parts.str[-1].fillna("")
    last_clean = last.str.replace(".", "", regex=False)
    
    mask_swap = (
    parts.str.len().ge(2) &
    last_clean.str.fullmatch(r"[A-Za-z]{1,2}")
    )
    
    s_swap = parts[mask_swap].apply(lambda x: " ".join([x[-1].replace(".", "")] + x[:-1]))
    s.loc[mask_swap] = s_swap
    
    # 5) final tidy + leeg -> Geen Naam
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    s = s.mask(s.eq(""), "Geen Naam")
    
    df["Namen"] = s
    
    # sorteren zonder display te slopen
    df["Namen_sort"] = df["Namen"].str.replace(r"\s+", "", regex=True).str.upper()
    df = df.sort_values("Namen_sort").drop(columns=["Namen_sort"])
    return df[["Namen", "Datum", "Debet", "Credit"]]


def preprocess_deb(df):
    df = df.copy()

    patterns = [
        r" HN.*",
        r" AFL.*",
        r" EXP.R.*",
        r" MT.*",
        r"RIJST",
        r" F.*",
        r" KG",
        r" ZILVERVL\.",
        r" VERK.*",
        r"BREUKF",
        r" BREUK",
        r" RECYLL.BREUK",
        r"CARGOBREUK",
        r"COMP.BREUK",
        r"CARGOBR\.",
        r",",
        r"\$",
        r"(?<=\d)\.(?=\d)",
        r"[0-9]+",
        r" H -",
        r" EXP.*",
        r"BREUK",
        r"\.BREUK",
        r" BON.*",
        r" ZKN..*",
    ]

    for pat in patterns:
        df["Namen"] = df["Namen"].str.replace(pat, " ", regex=True)

    df["Namen"] = df["Namen"].str.strip()
    df["Namen"] = df["Namen"].str.replace(r" W\. +.*", "", regex=True)
    df["Namen"] = df["Namen"].str.replace(r" W\.+.*", "", regex=True)
    df["Namen"] = df["Namen"].str.replace(
        r"W.+ZILVERVL.+PARB", "", regex=True
    )
    df["Namen"] = df["Namen"].str.replace(r" STORTING.*", "", regex=True)
    df["Namen"] = df["Namen"].str.replace(r" DSB.*", "", regex=True)

    # Split + check first word length (difference == 2 in R)
    split_series = df["Namen"].fillna("").str.split()
    df["Namen2"] = split_series.str[0].fillna("")
    df["Namen4"] = df["Namen"].str.replace(" ", "", regex=False)

    len_namen4 = df["Namen4"].str.len()
    len_namen2 = df["Namen2"].str.len()
    mask_len = (len_namen4 - len_namen2) == 2

    df["Namen3"] = df["Namen"]
    df.loc[mask_len, "Namen3"] = df.loc[mask_len, "Namen2"]

    df["Namen"] = df["Namen3"].str.replace(" ", "", regex=False)
    df["Namen"] = df["Namen"].replace("", "Geen Naam")

    df = df.sort_values("Namen")
    return df[["Namen", "Datum", "Debet", "Credit"]]


def preprocess_voors(df):
    df = df.copy()

    patterns = [
        r"AFL.*",
        r"ALF.*",
        r"HN.*",
        r"VSH.*",
        r"FSRD.*",
        r"VOORSCHOT",
        r" I .*",
        r"KN.*",
        r"VOORSCH.*",
        r"KP.*",
    ]

    for pat in patterns:
        df["Namen"] = df["Namen"].str.replace(pat, " ", regex=True)

    # Remove 'H$...' pattern
    df["Namen"] = df["Namen"].str.replace(r"H\$.*", "", regex=True)

    # First word only
    df["Namen"] = df["Namen"].str.split().str[0].fillna("")
    df["Namen"] = df["Namen"].str.replace(" ", ".", regex=False)
    df["Namen"] = df["Namen"].str.strip()
    df["Namen"] = df["Namen"].replace("", "Geen Naam")

    df = df.sort_values("Namen")
    return df[["Namen", "Datum", "Debet", "Credit"]]


# ============================================
# Streamlit app
# ============================================

st.title("Paloma grootboek Saldo Excel Tool")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
choice = st.selectbox(
    "Select processing type",
    options=["Kruisposten", "Crediteuren", "Debiteuren", "Voorschotten"],
    index=0,
)
col1, col2 = st.columns(2)
with col1:
    range_start = st.number_input("Start row (e.g. 9160)", min_value=1, value=1)
with col2:
    range_end = st.number_input("End row (e.g. 10359)", min_value=1, value=100)

output_name = st.text_input("Output filename (without .xlsx)", value="output")

if st.button("Process file"):
    if uploaded_file is None:
        st.error("Please upload an Excel file first.")
    elif range_end < range_start:
        st.error("End row must be ≥ start row.")
    else:
        with st.spinner("Processing..."):
            df_raw = load_excel_range(uploaded_file, range_start, range_end)

            if df_raw.empty:
                st.warning("No data found in the selected range.")
            else:
                if choice == "Kruisposten":
                    used_kpkn, used_other = preprocess_kruis(df_raw)
                    tables_kpkn, total_saldo1 = compute_saldo_tables(used_kpkn)
                    tables_other, total_saldo2 = compute_saldo_tables(used_other)

                    buffer = BytesIO()
                    buffer = write_tables_to_excel_kruis(
                        buffer,
                        tables_kpkn,
                        total_saldo1,
                        tables_other,
                        total_saldo2,
                    )

                elif choice == "Crediteuren":
                    used_file = preprocess_cred(df_raw)
                    tables, total_saldo = compute_saldo_tables(used_file)
                    buffer = BytesIO()
                    buffer = write_tables_to_excel_simple(
                        buffer, tables, total_saldo, title_prefix="Crediteuren"
                    )

                elif choice == "Debiteuren":
                    used_file = preprocess_deb(df_raw)
                    tables, total_saldo = compute_saldo_tables(used_file)
                    buffer = BytesIO()
                    buffer = write_tables_to_excel_simple(
                        buffer, tables, total_saldo, title_prefix="Debiteuren"
                    )

                elif choice == "Voorschotten":
                    used_file = preprocess_voors(df_raw)
                    tables, total_saldo = compute_saldo_tables(used_file)
                    buffer = BytesIO()
                    buffer = write_tables_to_excel_simple(
                        buffer, tables, total_saldo, title_prefix="Voorschotten"
                    )

                st.success("Excel file generated.")

                st.download_button(
                    label="Download Excel",
                    data=buffer,
                    file_name=f"{output_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )









