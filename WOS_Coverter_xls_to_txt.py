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

# 2) Load Excel
df = pd.read_excel(excel_path, dtype=str)

# 3) Drop all columns after "UT (Unique WOS ID)"
if last_column_name in df.columns:
    cutoff_index = df.columns.get_loc(last_column_name) + 1
    df = df.iloc[:, :cutoff_index]
else:
    raise ValueError(f"'{last_column_name}' column not found in input file.")

# 4) Rename to WOS short tags
df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

# 5) Output #1: TabDelimited.txt
df.to_csv(tabdelim_path, sep="\t", index=False, encoding="utf-8")
print(f"→ Tab‑delimited file written to {tabdelim_path}")

# 6) Output #2: PlainText.txt (Web of Science format)
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
    "PT","AU","AF","TI","SO","LA","DT","DE","ID","AB","C1","C3","RP","EM",
    "RI","OI","CR","NR","TC","Z9","U1","U2","PU","PI","PA","SN","EI","J9","JI",
    "PD","PY","DI","EA","PG","WC","WE","SC","GA","UT","OA","DA"
]

lines = ["FN Clarivate Analytics Web of Science", "VR 1.0"]

for _, row in df.iterrows():
    record = []
    for tag in tag_order:
        val = row.get(tag)
        if pd.isna(val) or not val:
            continue
        if tag in {"AU","AF","C1","CR"}:
            record += format_list_field(val, tag)
        else:
            record.append(f"{tag} {val}")
    if record:
        lines.extend(record)
        lines.append("ER")
        lines.append("")  # blank line between records

with open(plaintext_path, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"→ Plain‑text tagged file written to {plaintext_path}")
