#!/usr/bin/env python3
"""
contact_lot_cleaner.py
=======================
Cleans a contact lot (Excel/CSV) following an 8-step workflow:

  1. Insert two leftmost columns: "Upload" (= "Yes") and "Result Code" (= "New").
     Only filled for non-empty rows.
  2. Insert "First Name" and "Last Name" before "Full Name", split from Full Name.
  3. Clean contact numbers by target country. Non-target numbers are dropped
     UNLESS a row's only number is non-target -> tagged "Number outside specified country".
  4. Normalise numbers for an auto-dialer: strip spaces, +, (, ), -  ->
     digits only WITH country code. Max 3 phone columns (Phone, Phone 2, Phone3).
  5. Add "Region" column, auto-filled from Country.
  6. Screen Job Title against a user-supplied EXCLUDED-titles .txt list.
     Matches are tagged "DNC - Not Target Job Title".
  7. Sort A->Z by: Phone, Phone 2, Phone3, Company Name, Country.
  8. Re-order columns (right after Result Code):
     Country, Region, Phone, Phone 2, Phone3, LinkedIn URL, Company Name,
     First Name, Last Name, Full Name, Email, Addl Email, Job Title,
     Employee Size, Turnover, Industry

Usage:
    pip install pandas openpyxl phonenumbers      # phonenumbers optional but recommended
    python contact_lot_cleaner.py

The script prompts interactively for:
    - input file (or auto-detects the first .xlsx/.csv in this folder)
    - target country/countries
    - path to the excluded-titles .txt file (one title/keyword per line)
"""

import os
import re
import sys
import glob

import pandas as pd

# phonenumbers is optional. If installed, country detection is far more reliable.
try:
    import phonenumbers
    HAVE_PHONENUMBERS = True
except ImportError:
    HAVE_PHONENUMBERS = False


# ----------------------------------------------------------------------------
# Step 5 data: Country -> Region map (provided by user, kept verbatim)
# ----------------------------------------------------------------------------
COUNTRY_REGION = {
    "Afghanistan": "MEA", "Algeria": "Africa", "Angola": "Africa", "Argentina": "SA",
    "Australia": "ANZ", "Austria": "DACH", "Azerbaijan": "FSE", "Bahrain": "MEA",
    "Bangladesh": "South Asia", "Belgium": "BAN", "Benin": "Africa",
    "Bosnia and Herzegovina": "FSE", "Botswana": "Africa", "Brazil": "SA",
    "Bulgaria": "FSE", "Burkina Faso": "Africa", "Cambodia": "SEA", "Cameroon": "Africa",
    "Canada": "NA", "Cape Verde": "Africa", "Chad": "Africa", "Chile": "SA",
    "China": "East Asia", "Colombia": "SA", "Congo": "Africa", "Côte d'Ivoire": "Africa",
    "Cuba": "SA", "Czech Republic": "FSE", "Democratic Republic of the Congo": "Africa",
    "Denmark": "BAN", "Dominican Republic": "SA", "Egypt": "MEA", "Estonia": "FSE",
    "Ethiopia": "Africa", "Faroe Islands": "BAN", "Finland": "BAN", "France": "FSE",
    "Gabon": "Africa", "Germany": "DACH", "Ghana": "Africa", "Greece": "FSE",
    "Guinea": "Africa", "Hong Kong": "East Asia", "Hungary": "FSE", "Iceland": "UKI",
    "India": "South Asia", "Indonesia": "SEA", "Iran": "MEA", "Iraq": "MEA",
    "Ireland": "UKI", "Isle of Man": "UKI", "Israel": "MEA", "Italy": "FSE",
    "Japan": "East Asia", "Jordan": "MEA", "Kenya": "Africa", "Kuwait": "MEA",
    "Lebanon": "MEA", "Lesotho": "Africa", "Liberia": "Africa", "Libya": "Africa",
    "Lithuania": "FSE", "Luxembourg": "BAN", "Macedonia": "FSE", "Madagascar": "Africa",
    "Malawi": "Africa", "Malaysia": "SEA", "Maldives": "South Asia", "Mali": "Africa",
    "Mauritius": "Africa", "Mexico": "SA", "Mongolia": "East Asia", "Morocco": "Africa",
    "Mozambique": "Morocco", "Myanmar": "South Asia", "Namibia": "Africa",
    "Netherlands": "BAN", "New Zealand": "ANZ", "Nigeria": "Africa", "Norway": "BAN",
    "Oman": "MEA", "Pakistan": "West Asia", "Palestine": "MEA", "Philippines": "SEA",
    "Poland": "FSE", "Portugal": "FSE", "Qatar": "MEA", "Réunion": "Africa",
    "Romania": "FSE", "Russia": "North Asia", "Rwanda": "Africa", "Saudi Arabia": "MEA",
    "Senegal": "Africa", "Serbia": "FSE", "Sierra Leone": "Africa", "Singapore": "SEA",
    "Slovakia": "FSE", "Slovenia": "FSE", "South Africa": "Africa",
    "South Korea": "East Asia", "South Sudan": "Africa", "Spain": "FSE",
    "Sri Lanka": "South Asia", "Sudan": "Africa", "Sweden": "BAN",
    "Switzerland": "DACH", "Syria": "MEA", "Taiwan": "East Asia", "Tanzania": "Africa",
    "Thailand": "SEA", "Tunisia": "Africa", "Turkey": "MEA", "UAE": "MEA",
    "Uganda": "Africa", "Ukraine": "FSE", "United Arab Emirates": "MEA",
    "United Kingdom": "UKI", "United States": "NA", "USA": "NA", "Uzbekistan": "Central Asia",
    "Venezuela": "SA", "Vietnam": "SEA", "Yemen": "Africa", "Zambia": "Africa",
    "Zimbabwe": "Africa", "Monaco": "FSE", "Albania": "FSE", "Malta": "FSE",
    "Bhutan": "Asia", "Nepal": "Asia", "Saudia Arabia": "MEA", "Niger": "Africa",
    "Cyprus": "MEA", "Latvia": "FSE", "Burundi": "Africa", "Cote d’Ivoire": "Africa",
    "Equatorial Guinea": "Africa", "Seychelles": "Africa", "Croatia": "FSE",
    "Somalia": "Africa", "Montenegro": "FSE",
}

# Common name aliases -> canonical name used in COUNTRY_REGION
COUNTRY_ALIASES = {
    "u.s.": "United States", "u.s.a.": "United States", "us": "United States",
    "usa": "United States", "united states of america": "United States",
    "uk": "United Kingdom", "u.k.": "United Kingdom", "great britain": "United Kingdom",
    "england": "United Kingdom", "uae": "United Arab Emirates",
    "korea": "South Korea", "republic of korea": "South Korea",
    "russian federation": "Russia", "czechia": "Czech Republic",
    "ivory coast": "Côte d'Ivoire", "cote d'ivoire": "Côte d'Ivoire",
    "viet nam": "Vietnam", "the netherlands": "Netherlands", "holland": "Netherlands",
}

# Country name (lowercase) -> ISO 3166-1 alpha-2 region code, used for phone matching.
NAME_TO_ISO2 = {
    "afghanistan": "AF", "algeria": "DZ", "angola": "AO", "argentina": "AR",
    "australia": "AU", "austria": "AT", "azerbaijan": "AZ", "bahrain": "BH",
    "bangladesh": "BD", "belgium": "BE", "benin": "BJ", "bosnia and herzegovina": "BA",
    "botswana": "BW", "brazil": "BR", "bulgaria": "BG", "burkina faso": "BF",
    "cambodia": "KH", "cameroon": "CM", "canada": "CA", "cape verde": "CV",
    "chad": "TD", "chile": "CL", "china": "CN", "colombia": "CO", "congo": "CG",
    "côte d'ivoire": "CI", "cote d'ivoire": "CI", "cote d’ivoire": "CI",
    "cuba": "CU", "czech republic": "CZ", "czechia": "CZ",
    "democratic republic of the congo": "CD", "denmark": "DK",
    "dominican republic": "DO", "egypt": "EG", "estonia": "EE", "ethiopia": "ET",
    "faroe islands": "FO", "finland": "FI", "france": "FR", "gabon": "GA",
    "germany": "DE", "ghana": "GH", "greece": "GR", "guinea": "GN", "hong kong": "HK",
    "hungary": "HU", "iceland": "IS", "india": "IN", "indonesia": "ID", "iran": "IR",
    "iraq": "IQ", "ireland": "IE", "isle of man": "IM", "israel": "IL", "italy": "IT",
    "japan": "JP", "jordan": "JO", "kenya": "KE", "kuwait": "KW", "lebanon": "LB",
    "lesotho": "LS", "liberia": "LR", "libya": "LY", "lithuania": "LT",
    "luxembourg": "LU", "macedonia": "MK", "madagascar": "MG", "malawi": "MW",
    "malaysia": "MY", "maldives": "MV", "mali": "ML", "mauritius": "MU", "mexico": "MX",
    "mongolia": "MN", "morocco": "MA", "mozambique": "MZ", "myanmar": "MM",
    "namibia": "NA", "netherlands": "NL", "new zealand": "NZ", "nigeria": "NG",
    "norway": "NO", "oman": "OM", "pakistan": "PK", "palestine": "PS",
    "philippines": "PH", "poland": "PL", "portugal": "PT", "qatar": "QA",
    "réunion": "RE", "reunion": "RE", "romania": "RO", "russia": "RU",
    "russian federation": "RU", "rwanda": "RW", "saudi arabia": "SA",
    "saudia arabia": "SA", "senegal": "SN", "serbia": "RS", "sierra leone": "SL",
    "singapore": "SG", "slovakia": "SK", "slovenia": "SI", "south africa": "ZA",
    "south korea": "KR", "korea": "KR", "south sudan": "SS", "spain": "ES",
    "sri lanka": "LK", "sudan": "SD", "sweden": "SE", "switzerland": "CH",
    "syria": "SY", "taiwan": "TW", "tanzania": "TZ", "thailand": "TH", "tunisia": "TN",
    "turkey": "TR", "uae": "AE", "united arab emirates": "AE", "uganda": "UG",
    "ukraine": "UA", "united kingdom": "GB", "uk": "GB", "england": "GB",
    "great britain": "GB", "united states": "US", "usa": "US", "us": "US",
    "uzbekistan": "UZ", "venezuela": "VE", "vietnam": "VN", "viet nam": "VN",
    "yemen": "YE", "zambia": "ZM", "zimbabwe": "ZW", "monaco": "MC", "albania": "AL",
    "malta": "MT", "bhutan": "BT", "nepal": "NP", "niger": "NE", "cyprus": "CY",
    "latvia": "LV", "burundi": "BI", "equatorial guinea": "GQ", "seychelles": "SC",
    "croatia": "HR", "somalia": "SO", "montenegro": "ME",
}

# ISO2 -> international calling code (fallback when phonenumbers is unavailable).
ISO2_TO_CC = {
    "AF": "93", "DZ": "213", "AO": "244", "AR": "54", "AU": "61", "AT": "43",
    "AZ": "994", "BH": "973", "BD": "880", "BE": "32", "BJ": "229", "BA": "387",
    "BW": "267", "BR": "55", "BG": "359", "BF": "226", "KH": "855", "CM": "237",
    "CA": "1", "CV": "238", "TD": "235", "CL": "56", "CN": "86", "CO": "57",
    "CG": "242", "CI": "225", "CU": "53", "CZ": "420", "CD": "243", "DK": "45",
    "DO": "1", "EG": "20", "EE": "372", "ET": "251", "FO": "298", "FI": "358",
    "FR": "33", "GA": "241", "DE": "49", "GH": "233", "GR": "30", "GN": "224",
    "HK": "852", "HU": "36", "IS": "354", "IN": "91", "ID": "62", "IR": "98",
    "IQ": "964", "IE": "353", "IM": "44", "IL": "972", "IT": "39", "JP": "81",
    "JO": "962", "KE": "254", "KW": "965", "LB": "961", "LS": "266", "LR": "231",
    "LY": "218", "LT": "370", "LU": "352", "MK": "389", "MG": "261", "MW": "265",
    "MY": "60", "MV": "960", "ML": "223", "MU": "230", "MX": "52", "MN": "976",
    "MA": "212", "MZ": "258", "MM": "95", "NA": "264", "NL": "31", "NZ": "64",
    "NG": "234", "NO": "47", "OM": "968", "PK": "92", "PS": "970", "PH": "63",
    "PL": "48", "PT": "351", "QA": "974", "RE": "262", "RO": "40", "RU": "7",
    "RW": "250", "SA": "966", "SN": "221", "RS": "381", "SL": "232", "SG": "65",
    "SK": "421", "SI": "386", "ZA": "27", "KR": "82", "SS": "211", "ES": "34",
    "LK": "94", "SD": "249", "SE": "46", "CH": "41", "SY": "963", "TW": "886",
    "TZ": "255", "TH": "66", "TN": "216", "TR": "90", "AE": "971", "UG": "256",
    "UA": "380", "GB": "44", "US": "1", "UZ": "998", "VE": "58", "VN": "84",
    "YE": "967", "ZM": "260", "ZW": "263", "MC": "377", "AL": "355", "MT": "356",
    "BT": "975", "NP": "977", "NE": "227", "CY": "357", "LV": "371", "BI": "257",
    "GQ": "240", "SC": "248", "HR": "385", "SO": "252", "ME": "382",
}

# reverse map: ISO2 -> a canonical country name present in COUNTRY_REGION (first-seen wins)
ISO2_TO_COUNTRY = {}
for _c in COUNTRY_REGION:
    _i = NAME_TO_ISO2.get(re.sub(r"\s+", " ", _c.strip()).lower())
    if _i and _i not in ISO2_TO_COUNTRY:
        ISO2_TO_COUNTRY[_i] = _c

# Final canonical column order (after the two leftmost Upload / Result Code columns).
FINAL_ORDER = [
    "Upload", "Result Code", "Country", "Region", "Phone", "Phone 2", "Phone3",
    "LinkedIn URL", "Company Name", "First Name", "Last Name", "Full Name",
    "Email", "Addl Email", "Job Title", "Employee Size", "Turnover", "Industry",
]

# Result-code tags
TAG_OUTSIDE = "Number outside specified country"
TAG_DNC_TITLE = "DNC - Not Target Job Title"


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def norm(s):
    """Lower-case, collapse whitespace."""
    return re.sub(r"\s+", " ", str(s).strip()).lower()


def is_blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "" or str(v).strip().lower() == "nan"


def canonical_country(raw):
    """Return canonical country name found in COUNTRY_REGION, or None."""
    if is_blank(raw):
        return None
    n = norm(raw)
    if n in COUNTRY_ALIASES:
        return COUNTRY_ALIASES[n]
    for key in COUNTRY_REGION:
        if norm(key) == n:
            return key
    return None


def country_to_iso2(raw):
    if is_blank(raw):
        return None
    n = norm(raw)
    if n in NAME_TO_ISO2:
        return NAME_TO_ISO2[n]
    canon = canonical_country(raw)
    if canon and norm(canon) in NAME_TO_ISO2:
        return NAME_TO_ISO2[norm(canon)]
    return None


def region_for_country(raw):
    canon = canonical_country(raw)
    if canon:
        return COUNTRY_REGION.get(canon, "")
    return ""


def find_col(columns, *keywords):
    """Find first column whose normalised name contains any keyword."""
    lc = {c: norm(c) for c in columns}
    for kw in keywords:
        for c, n in lc.items():
            if kw in n:
                return c
    return None


# ----------------------------------------------------------------------------
# Phone handling (steps 3 & 4)
# ----------------------------------------------------------------------------
PHONE_COL_HINTS = ("phone", "mobile", "tel", "contact", "cell", "direct", "number", "hq", "fax")


def detect_phone_columns(columns):
    cols = []
    for c in columns:
        n = norm(c)
        if any(h in n for h in PHONE_COL_HINTS) and "email" not in n:
            cols.append(c)
    return cols


def digits_only(raw):
    """Strip spaces, +, (, ), -, dots and any non-digit. Returns digit string."""
    return re.sub(r"\D", "", str(raw))


def detect_number_iso2(raw, row_country_iso2, target_iso2):
    """Return (cleaned_digits_with_cc, iso2_or_None)."""
    raw = str(raw).strip()
    if not raw:
        return None, None
    has_plus = "+" in raw
    digits = digits_only(raw)
    if not digits:
        return None, None

    if HAVE_PHONENUMBERS:
        candidates = []
        if has_plus:
            candidates.append("+" + digits)
        else:
            candidates.append("+" + digits)
            regions = [row_country_iso2] + list(target_iso2)
            for reg in regions:
                if reg:
                    candidates.append((digits, reg))
        for cand in candidates:
            try:
                if isinstance(cand, tuple):
                    pn = phonenumbers.parse(cand[0], cand[1])
                else:
                    pn = phonenumbers.parse(cand, None)
                if phonenumbers.is_valid_number(pn):
                    iso2 = phonenumbers.region_code_for_number(pn)
                    e164 = phonenumbers.format_number(
                        pn, phonenumbers.PhoneNumberFormat.E164)
                    return e164.lstrip("+"), iso2
            except Exception:
                continue
        return digits, None

    # ---- Fallback: prefix matching against calling codes ----
    cc_map = {ISO2_TO_CC[i]: i for i in target_iso2 if i in ISO2_TO_CC}
    if row_country_iso2 in ISO2_TO_CC:
        cc_map.setdefault(ISO2_TO_CC[row_country_iso2], row_country_iso2)
    for cc in sorted(cc_map, key=len, reverse=True):
        if digits.startswith(cc):
            return digits, cc_map[cc]
    if row_country_iso2 in ISO2_TO_CC:
        return ISO2_TO_CC[row_country_iso2] + digits, row_country_iso2
    return digits, None


def clean_row_phones(raw_numbers, row_country, target_iso2):
    """Return (list_of_up_to_3_cleaned_numbers, outside_flag, outside_iso2)."""
    row_iso2 = country_to_iso2(row_country)
    parsed = []
    seen = set()
    for raw in raw_numbers:
        if is_blank(raw):
            continue
        for piece in re.split(r"[;/]| or ", str(raw)):
            if not piece.strip():
                continue
            cleaned, iso2 = detect_number_iso2(piece, row_iso2, target_iso2)
            if cleaned and cleaned not in seen:
                seen.add(cleaned)
                parsed.append((cleaned, iso2))

    if not parsed:
        return [], False, None

    target = [p for p in parsed if p[1] in target_iso2]
    if target:
        kept = [c for c, _ in target][:3]
        return kept, False, None
    return [parsed[0][0]], True, parsed[0][1]


# ----------------------------------------------------------------------------
# Main cleaning pipeline
# ----------------------------------------------------------------------------
def clean(df, target_countries, excluded_titles):
    target_iso2 = {country_to_iso2(c) for c in target_countries}
    target_iso2.discard(None)

    cols = list(df.columns)
    full_name_col = find_col(cols, "full name", "name")
    company_col = find_col(cols, "company")
    country_col = find_col(cols, "country")
    job_col = find_col(cols, "job title", "title", "designation")
    phone_cols = detect_phone_columns(cols)
    linkedin_col = find_col(cols, "linkedin")
    email_col = find_col(cols, "email")
    addl_email_col = find_col(cols, "addl email", "additional email", "alt email", "secondary email", "email 2")
    size_col = find_col(cols, "employee", "company size", "headcount", "size")
    turnover_col = find_col(cols, "turnover", "revenue")
    industry_col = find_col(cols, "industry", "sector")

    # Any input column not mapped to a standard field is "extra" — keep it at the far right
    consumed = {c for c in [full_name_col, company_col, country_col, job_col,
                            linkedin_col, email_col, addl_email_col, size_col,
                            turnover_col, industry_col] if c} | set(phone_cols)
    extra_cols = [c for c in cols if c not in consumed and str(c).strip() != ""]
    extra_out = [(c + " (original)") if c in FINAL_ORDER else c for c in extra_cols]

    out_rows = []
    for _, row in df.iterrows():
        row_blank = all(is_blank(row[c]) for c in cols)
        rec = {}

        rec["Country"] = _get(row, country_col)
        rec["Region"] = region_for_country(rec["Country"]) if rec["Country"] else ""
        rec["LinkedIn URL"] = _get(row, linkedin_col)
        rec["Company Name"] = _get(row, company_col)
        rec["Full Name"] = _get(row, full_name_col)
        rec["Email"] = _get(row, email_col)
        rec["Addl Email"] = _get(row, addl_email_col)
        rec["Job Title"] = _get(row, job_col)
        rec["Employee Size"] = _get(row, size_col)
        rec["Turnover"] = _get(row, turnover_col)
        rec["Industry"] = _get(row, industry_col)

        # Step 2: split Full Name
        first, last = split_name(rec["Full Name"])
        rec["First Name"] = first
        rec["Last Name"] = last

        # Steps 3 & 4: phones
        raw_numbers = [row[c] for c in phone_cols]
        outside_iso2 = None
        if target_iso2:
            kept, outside_flag, outside_iso2 = clean_row_phones(raw_numbers, rec["Country"], target_iso2)
        else:
            kept = [digits_only(x) for x in raw_numbers if not is_blank(x)][:3]
            outside_flag = False
        rec["Phone"] = kept[0] if len(kept) > 0 else ""
        rec["Phone 2"] = kept[1] if len(kept) > 1 else ""
        rec["Phone3"] = kept[2] if len(kept) > 2 else ""

        # When the only number is outside the target country, re-derive Country/Region from its country code
        if outside_flag and outside_iso2 and outside_iso2 in ISO2_TO_COUNTRY:
            rec["Country"] = ISO2_TO_COUNTRY[outside_iso2]
            rec["Region"] = region_for_country(rec["Country"])

        # Step 1: Upload / Result Code
        rec["Upload"] = "" if row_blank else "Yes"
        result_codes = [] if row_blank else ["New"]

        # Step 6: job title screen against excluded list
        title_dnc = title_is_excluded(rec["Job Title"], excluded_titles)
        if not row_blank and title_dnc:
            result_codes = [TAG_DNC_TITLE]
        elif not row_blank and outside_flag:
            result_codes = [TAG_OUTSIDE]
        rec["Result Code"] = "; ".join(result_codes)

        # carry leftover original columns to the right end
        for c, oc in zip(extra_cols, extra_out):
            rec[oc] = _get(row, c)

        out_rows.append(rec)

    out = pd.DataFrame(out_rows, columns=FINAL_ORDER + extra_out)

    # Step 7: sort A-Z by Phone, Phone 2, Phone3, Company Name, Country
    sort_keys = ["Phone", "Phone 2", "Phone3", "Company Name", "Country"]
    out["_blank"] = out["Upload"].eq("")
    out = out.sort_values(
        by=["_blank"] + sort_keys,
        key=lambda s: s if s.name == "_blank" else s.astype(str).str.lower(),
        kind="stable",
    ).drop(columns="_blank").reset_index(drop=True)

    return out


def _get(row, col):
    if col is None or col not in row or is_blank(row[col]):
        return ""
    return str(row[col]).strip()


def split_name(full):
    full = str(full).strip()
    if not full:
        return "", ""
    parts = full.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def title_is_excluded(title, excluded):
    if is_blank(title) or not excluded:
        return False
    t = norm(title)
    return any(ex in t for ex in excluded)


# ----------------------------------------------------------------------------
# I/O & interactive entry point
# ----------------------------------------------------------------------------
def load_excluded_titles(path):
    titles = []
    if path and os.path.isfile(path):
        with open(path, encoding="utf-8-sig") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    titles.append(norm(line))
    return titles


def read_any(path):
    if path.lower().endswith((".xlsx", ".xls", ".xlsm")):
        return pd.read_excel(path, dtype=str)
    return pd.read_csv(path, dtype=str)


def main():
    here = os.path.dirname(os.path.abspath(__file__))

    in_path = input("Path to contact lot file (blank = auto-detect in this folder): ").strip().strip('"')
    if not in_path:
        candidates = [p for p in glob.glob(os.path.join(here, "*"))
                      if p.lower().endswith((".xlsx", ".xls", ".csv"))
                      and "cleaned" not in os.path.basename(p).lower()]
        if not candidates:
            print("No .xlsx/.csv found in this folder. Aborting.")
            sys.exit(1)
        in_path = candidates[0]
        print(f"  -> using {os.path.basename(in_path)}")

    df = read_any(in_path)
    print(f"Loaded {len(df)} rows, {len(df.columns)} columns.")

    raw = input("Target country/countries (comma-separated): ").strip()
    target_countries = [c.strip() for c in raw.split(",") if c.strip()]
    unknown = [c for c in target_countries if country_to_iso2(c) is None]
    if unknown:
        print(f"  WARNING: could not map these to a country code: {unknown}")
    if not HAVE_PHONENUMBERS:
        print("  NOTE: 'phonenumbers' not installed - using calling-code prefix matching "
              "(less precise). Install with: pip install phonenumbers")

    tpath = input("Path to excluded-titles .txt (blank = excluded_titles.txt here): ").strip().strip('"')
    if not tpath:
        tpath = os.path.join(here, "excluded_titles.txt")
    excluded = load_excluded_titles(tpath)
    print(f"Loaded {len(excluded)} excluded title keyword(s).")

    out = clean(df, target_countries, excluded)

    base = os.path.splitext(os.path.basename(in_path))[0]
    out_path = os.path.join(here, f"{base}_cleaned.xlsx")
    out.to_excel(out_path, index=False)
    print(f"\nDone. Wrote {len(out)} rows -> {out_path}")
    vc = out["Result Code"].value_counts(dropna=False)
    print("Result Code breakdown:")
    for k, v in vc.items():
        print(f"  {k or '(blank)'}: {v}")


if __name__ == "__main__":
    main()
