import pandas as pd

# ── CONFIG ─────────────────────────────────────────────────────────────────────
excel_path       = "WOS_Filtered.xlsx"            # your input Excel file
tabdelim_path    = "TabDelimited_Filtered.txt"    # output #1: tab-delimited file
plaintext_path   = "PlainText_Filtered.txt"       # output #2: plain-text WOS tagged format
last_column_name = "UT (Unique WOS ID)"  # drop columns after this
# ────────────────────────────────────────────────────────────────────────────────

# 1) Map Excel headers → WOS short codes
col_map = {
    "Publication Type":    "PT",
    "Authors":             "AU",
    "Book Authors":        "BA",
    "Book Editors":        "BE",
    "Book Group Authors":  "GP",
    "Author Full Names":   "AF",
    "Book Author Full Names":"BF",
    "Group Authors":       "CA",
    "Article Title":       "TI",
    "Source Title":        "SO",
    "Book Series Title":   "SE",
    "Book Series Subtitle":"BS",
    "Language":            "LA",
    "Document Type":       "DT",
    "Conference Title":    "CT",
    "Conference Date":     "CY",
    "Conference Location": "CL",
    "Conference Sponsor":  "SP",
    "Conference Host":     "HO",
    "Author Keywords":     "DE",
    "Keywords Plus":       "ID",
    "Abstract":            "AB",
    "Addresses":           "C1",
    "Affiliations":        "C3",
    "Reprint Addresses":   "RP",
    "Email Addresses":     "EM",
    "Researcher Ids":      "RI",
    "ORCIDs":              "OI",
    "Funding Orgs":        "FG",
    "Funding Name Preferred":"FP",
    "Funding Text":        "FX",
    "Cited References":    "CR",
    "Cited Reference Count":"NR",
    "Times Cited, WoS Core":"TC",
    "Times Cited, All Databases":"Z9",
    "180 Day Usage Count": "U1",
    "Since 2013 Usage Count":"U2",
    "Publisher":           "PU",
    "Publisher City":      "PI",
    "Publisher Address":   "PA",
    "ISSN":                "SN",
    "eISSN":               "EI",
    "ISBN":                "BN",
    "Journal Abbreviation":"J9",
    "Journal ISO Abbreviation":"JI",
    "Publication Date":    "PD",
    "Publication Year":    "PY",
    "Volume":              "VL",
    "Issue":               "IS",
    "Part Number":         "PN",
    "Supplement":          "SU",
    "Special Issue":       "SI",
    "Meeting Abstract":    "MA",
    "Start Page":          "BP",
    "End Page":            "EP",
    "Article Number":      "AR",
    "DOI":                 "DI",
    "DOI Link":            "DL",
    "Book DOI":            "D2",
    "Early Access Date":   "EA",
    "Number of Pages":     "PG",
    "WoS Categories":      "WC",
    "Web of Science Index":"WE",
    "Research Areas":      "SC",
    "IDS Number":          "GA",
    "Pubmed Id":           "PM",
    "Open Access Designations":"OA",
    "Highly Cited Status": "HC",
    "Hot Paper Status":    "HP",
    "Date of Export":      "DA",
    "UT (Unique WOS ID)":  "UT"
}

# Load Excel file
df = pd.read_excel(excel_path, dtype=str)

# Map Excel headers to WoS short codes
df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})
print("Columns after renaming:", df.columns.tolist())

# Check for critical columns
critical_cols = ["TI", "SO", "AU"]
missing_critical = [col for col in critical_cols if col not in df.columns]
if missing_critical:
    raise ValueError(f"Missing critical columns: {missing_critical}. Update col_map.")

# Generate placeholder UT if missing
if "UT" not in df.columns:
    print("UT column not found. Generating placeholders for all records.")
    df["UT"] = [f"SCOPUS:ID{idx:06d}" for idx in range(len(df))]
else:
    def generate_ut(row, idx):
        if pd.notna(row["UT"]) and str(row["UT"]).strip():
            return row["UT"]
        if pd.notna(row.get("DI")) and str(row["DI"]).strip():
            return f"SCOPUS:{row['DI']}"
        return f"SCOPUS:ID{idx:06d}"
    df["UT"] = [generate_ut(row, idx) for idx, row in df.iterrows()]

# Deduplicate based on critical columns
print("Rows before deduplication:", len(df))
df = df.drop_duplicates(subset=critical_cols, keep="first")
print("Rows after deduplication:", len(df))

# Normalize field separators
def normalize_field(s):
    if pd.isna(s) or not str(s).strip():
        return ""
    return str(s).replace(",", ";").strip()

for col in ["AU", "AF", "C1", "CR", "DE", "ID"]:
    if col in df.columns:
        df[col] = df[col].apply(normalize_field)

# Output 1: Tab-delimited file
df.to_csv(tabdelim_path, sep="\t", index=False, encoding="utf-8")
print(f"Tab-delimited file written to {tabdelim_path}")

# Output 2: Plaintext file in WoS format
def format_list_field(s, tag):
    if pd.isna(s) or not str(s).strip():
        return []
    parts = [p.strip() for p in str(s).split(";") if p.strip()]
    if not parts:
        return []
    lines = [f"{tag} {parts[0]}"]
    lines += [f"   {p}" for p in parts[1:]]
    return lines

tag_order = [
    "PT", "AU", "AF", "TI", "SO", "LA", "DT", "DE", "ID", "AB", "C1", "C3", "RP", "EM",
    "RI", "OI", "CR", "NR", "TC", "Z9", "U1", "U2", "PU", "PI", "PA", "SN", "EI", "J9", "JI",
    "PD", "PY", "DI", "EA", "PG", "WC", "WE", "SC", "GA", "UT", "OA", "DA"
]

lines = ["FN Clarivate Analytics Web of Science", "VR 1.0"]
skipped_records = 0
for idx, row in df.iterrows():
    record = []
    has_critical = all(pd.notna(row[col]) and str(row[col]).strip() for col in critical_cols)
    for tag in tag_order:
        val = row.get(tag)
        if pd.isna(val) or not str(val).strip():
            continue
        if tag in {"AU", "AF", "C1", "CR"}:
            formatted = format_list_field(val, tag)
            if formatted:
                record += formatted
        else:
            record.append(f"{tag} {val}")
    if record and has_critical:
        lines.extend(record)
        lines.append("ER")
        lines.append("")
    else:
        skipped_records += 1
        print(f"Skipped record {idx}: Missing critical fields")

with open(plaintext_path, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

# Diagnostics
print(f"Plaintext file written to {plaintext_path}")
print(f"Records processed: {len(df) - skipped_records}")
print(f"Records skipped: {skipped_records}")
with open(plaintext_path, "r", encoding="utf-8") as f:
    er_count = f.read().count("ER")

