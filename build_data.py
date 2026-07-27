#!/usr/bin/env python3
"""
build_data.py  -  MOVE Fund Coalition Evaluation dashboard preprocessor.

Reads the survey files listed in data/manifest.csv, scores the collective-impact
constructs (Likert -> 1..5), parses program-impact fields, and writes
dashboard_data.js  ->  `const DASHBOARD_DATA = {...}` consumed by dashboard.html.

Two file formats are supported:
  * Qualtrics CSV (one coalition per file): rows 2-3 are question text / importId.
    Likert answers are TEXT ("Strongly agree"). Manifest gives the coalition name.
  * Pre-coded analysis XLSX (many coalitions per file): a "Coalition" column labels
    each row, Likert answers are NUMERIC 1-5, multi-selects are exploded binary columns.
    Leave the manifest 'coalition' cell BLANK to split the file by its Coalition column.

Usage:   python3 build_data.py
Re-run whenever you add/update a file in data/. (Reading .xlsx needs `openpyxl`.)

To add data:
  1. Drop the export into the right data/ subfolder.
  2. Add a row to data/manifest.csv:  file,coalition,timepoint
     - file: path relative to data/ (quote it if it contains commas)
     - coalition: the name, OR blank to split a multi-coalition file by its Coalition column
     - timepoint: T0 / T1 / T2 ...
  3. Re-run this script.
"""

import csv
import json
import math
import os
import re

try:
    import openpyxl
except ImportError:
    openpyxl = None

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
MANIFEST = os.path.join(DATA_DIR, "manifest.csv")
OUT_JS = os.path.join(HERE, "dashboard_data.js")
# The dashboard page (index.html is now the landing page; the dashboard lives in dashboard.html).
INDEX_HTML = os.path.join(HERE, "dashboard.html")
STANDALONE_HTML = os.path.join(HERE, "dashboard_standalone.html")

# --------------------------------------------------------------------------------------
# Scale maps : survey text  ->  numeric value
# --------------------------------------------------------------------------------------
AGREEMENT = {
    "strongly disagree": 1,
    "somewhat disagree": 2,
    "neither agree nor disagree": 3,
    "somewhat agree": 4,
    "strongly agree": 5,
}
FREQUENCY = {
    "none at all": 1,
    "a little": 2,
    "a moderate amount": 3,
    "a lot": 4,
    "a great deal": 5,
}
POSITIVITY = {
    "extremely negative": 1,
    "somewhat negative": 2,
    "neither positive nor negative": 3,
    "somewhat positive": 4,
    "extremely positive": 5,
}
SCALE_MAPS = {"agreement": AGREEMENT, "frequency": FREQUENCY, "positivity": POSITIVITY}
SCALE_MAX = 5

# --------------------------------------------------------------------------------------
# Construct definitions.  prefix -> {name, items, scale}
# `items` are the exact variable-name columns (row 1 of the Qualtrics export).
# Add `reverse: [list of item keys]` to a construct to reverse-code those items (6 - v).
# --------------------------------------------------------------------------------------
CONSTRUCTS = [
    {"key": "TACT",      "name": "Motivation: Functional",          "items": ["TACT_1", "TACT_2", "TACT_3"], "scale": "agreement"},
    {"key": "TRAN",      "name": "Motivation: Transformational",    "items": ["TRAN_1", "TRAN_2", "TRAN_3"], "scale": "agreement"},
    {"key": "GC",        "name": "Goal Consensus",                  "items": ["GC_1", "GC_2", "GC_3"],       "scale": "agreement"},
    {"key": "COM",       "name": "Communication",                   "items": ["COM_1", "COM_2", "COM_3"],    "scale": "agreement"},
    {"key": "COH",       "name": "Cohesion",                        "items": ["COH_1", "COH_2", "COH_3"],    "scale": "agreement"},
    {"key": "TF",        "name": "Task Focus",                      "items": ["TF_1", "TF_2", "TF_3"],       "scale": "agreement", "reverse": ["TF_2"]},
    {"key": "DM",        "name": "Decision-Making",                 "items": ["DM_1", "DM_2", "DM_3"],       "scale": "frequency"},
    {"key": "LEAD",      "name": "Leadership",                      "items": ["LEAD_1", "LEAD_2", "LEAD_3"], "scale": "agreement"},
    {"key": "STAFF",     "name": "Staff",                           "items": ["STAFF_1", "STAFF_2", "STAFF_3"], "scale": "agreement"},
    {"key": "TRUST",     "name": "Trust",                           "items": ["TRUST_1", "TRUST_2", "TRUST_3"], "scale": "agreement"},
    {"key": "MemEng_Sat","name": "Member Engagement",              "items": ["MemEng_Sat_1", "MemEng_Sat_2", "MemEng_Sat_3"], "scale": "agreement", "reverse": ["MemEng_Sat_1", "MemEng_Sat_2", "MemEng_Sat_3"]},
    {"key": "PROD",      "name": "Productivity",                    "items": ["PROD_1", "PROD_2", "PROD_3"], "scale": "agreement", "reverse": ["PROD_1", "PROD_2", "PROD_3"]},
    {"key": "LEGIT",     "name": "External Legitimacy",             "items": ["LEGIT_1", "LEGIT_2", "LEGIT_3"], "scale": "agreement"},
    {"key": "AGENDA",    "name": "Common Agenda",                   "items": ["AGENDA_1", "AGENDA_2", "AGENDA_3"], "scale": "agreement"},
    {"key": "MA",        "name": "Mutually Reinforcing Activities", "items": ["MA_1", "MA_2", "MA_3"],       "scale": "agreement"},
    {"key": "SM",        "name": "Shared Measurement",              "items": ["SM_1", "SM_2", "SM_3"],       "scale": "agreement"},
    {"key": "EQUITY",    "name": "Equity",                          "items": ["EQUITY_1", "EQUITY_2", "EQUITY_3"], "scale": "agreement"},
    {"key": "OC",        "name": "Organizational Capacity",         "items": ["OC_I_1", "OC_G_2", "OC_O_3", "OC_C_4"], "scale": "positivity"},
    {"key": "SC",        "name": "Systems Change",                  "items": ["SC_1", "SC_2", "SC_3", "SC_4", "SC_5", "SC_6"], "scale": "agreement"},
    {"key": "O",         "name": "Social Outcomes",                 "items": ["O_1", "O_2"],                 "scale": "agreement"},
]

# Higher-order domains (from the codebook) that group the constructs. Ordered for the menu.
DOMAINS = [
    {"key": "processes", "name": "Collaborative Processes", "constructs": ["TACT", "TRAN", "GC", "COM", "COH", "TF", "DM"]},
    {"key": "backbone",  "name": "Backbone Support",        "constructs": ["LEAD", "STAFF", "TRUST", "EQUITY"]},
    {"key": "impact",    "name": "Collective Impact",       "constructs": ["AGENDA", "MA", "SM"]},
    {"key": "outputs",   "name": "Collaborative Outputs",   "constructs": ["MemEng_Sat", "PROD", "LEGIT"]},
    {"key": "outcomes",  "name": "Social Outcomes",         "constructs": ["OC", "SC", "O"]},
]
DOMAIN_OF = {ck: d["key"] for d in DOMAINS for ck in d["constructs"]}

# Items that are reverse-scored (6 - value), collected from the construct definitions.
REVERSE_ITEMS = {it for c in CONSTRUCTS for it in c.get("reverse", [])}

# Map each item variable-name to its code in the codebook (data columns differ from codebook).
CODEBOOK_CODE = {it: it for c in CONSTRUCTS for it in c["items"]}
CODEBOOK_CODE.update({"TACT_1": "FUNC_1", "TACT_2": "FUNC_2", "TACT_3": "FUNC_3", "TF_2": "TF_2_re"})
CODEBOOK_FILE = os.path.join(DATA_DIR, "Survey Codebook - for Will.xlsx")

# Program-impact (free-text numeric) fields. Columns are resolved by QUESTION TEXT
# (a lowercase substring of row 1), not by variable name -- the variable names (Q81 etc.)
# are reused for different questions across survey versions (e.g. Delaware's Q81 is a Likert
# item), so name-matching would mis-read them.
PROGRAM_FIELDS = {
    "youth_served":    {"match": "how many youth has your organization served", "label": "Youth served (last 12 mo.)",         "agg": "sum"},
    "pct_aid":         {"match": "what percentage of youth",                    "label": "% of youth receiving financial aid", "agg": "wmean"},
    "aid_dollars":     {"match": "how many dollars have you disbursed",          "label": "Financial aid disbursed ($)",        "agg": "sum"},
    "coaches_worked":  {"match": "how many coaches have worked",                 "label": "Coaches worked",                     "agg": "sum"},
    "coaches_trained": {"match": "how many coaches have been trained",           "label": "Coaches trained",                    "agg": "sum"},
}

# --------------------------------------------------------------------------------------
# Respondent / organization profile questions (each chart is per-coalition, no comparison).
# Percentages use the coalition's respondent count (N) as the denominator.
#   mode "present"  -> show only options this coalition selected, by count desc
#   mode "order"    -> show a fixed ordered list of categories (0% bars kept)
#   mode "universe" -> show every option seen across ALL files (0% bars kept), by global freq
# --------------------------------------------------------------------------------------
FREQ_ORDER = ["None at all", "A little", "A moderate amount", "A lot", "A great deal"]
AGE_ORDER = ["0-5", "6-12", "13-17", "18-24"]
# Membership-duration codes (MemEng_Dur; codebook). Present in the coded Time 0 file only.
DUR_ORDER = ["Less than 1 year", "1-3 years", "4-6 years", "6-8 years", "More than 9 years"]
# Fixed master list of populations from the survey question (always shown, in this order, 0% kept).
POPS_ORDER = [
    "Girls",
    "Youth in poverty",
    "BIPOC Youth",
    "Immigrant and newcomer youth",
    "LGBTQIA+ youth",
    "Non-binary, gender non-conforming, and/or trans youth",
    "Youth involved in the criminal justice system",
    "Youth in foster care",
    "Youth with physical disabilities",
    "Youth with intellectual disabilities",
]

PROFILE_QUESTIONS = {
    "sector":     {"col": "Sector",      "label": "Organization sector",                "type": "single", "mode": "present"},
    "field":      {"col": "Field",       "label": "Focus areas",                        "type": "multi",  "mode": "present"},
    "engagement": {"col": "MemEng_Part", "label": "Overall engagement in the coalition","type": "single", "mode": "order", "order": FREQ_ORDER},
    "duration":   {"col": "MemEng_Dur",  "label": "Time as a coalition member",         "type": "single", "mode": "order", "order": DUR_ORDER},
    "age":        {"col": "Age",         "label": "Ages of youth served",               "type": "multi",  "mode": "order", "order": AGE_ORDER},
    "pops":       {"col": "AtRiskPops",  "label": "Populations intentionally served",   "type": "multi",  "mode": "order", "order": POPS_ORDER},
}
ENGAGE_ITEMS = [f"EngageFeedback_{i}" for i in range(1, 7)]

# Multi-select option labels that themselves contain commas (so a naive comma-split breaks).
KNOWN_COMMA_OPTIONS = [
    "Parks, recreation, and leisure services",
    "Non-binary, gender non-conforming, and/or trans youth",
]
# Display fixes for obvious instrument typos (counting is unaffected).
LABEL_FIXES = {"Youth develpoment": "Youth development"}


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------
def is_blank(v):
    return v is None or str(v).strip() in ("", "None")


def to_numeric(raw, scale):
    """Likert value as 1..SCALE_MAX. Handles NUMERIC-coded cells (xlsx) and TEXT cells (csv)."""
    if is_blank(raw):
        return None
    if isinstance(raw, (int, float)):
        v = float(raw)
        return v if 1 <= v <= SCALE_MAX else None
    t = str(raw).strip()
    try:                                # numeric stored as text, e.g. "4"
        v = float(t)
        return v if 1 <= v <= SCALE_MAX else None
    except ValueError:
        return SCALE_MAPS[scale].get(t.lower())


def map_likert(text, scale):           # text-only helper used by profile parsing
    if text is None:
        return None
    t = text.strip().lower()
    if not t:
        return None
    return SCALE_MAPS[scale].get(t)


def norm_label(s):
    return LABEL_FIXES.get(s.strip(), s.strip())


def parse_multi(cell):
    """Split a Qualtrics multi-select cell into option labels, handling option labels
    that contain commas and dropping 'Click to write Choice N' placeholders."""
    s = cell or ""
    found = []
    for opt in KNOWN_COMMA_OPTIONS:        # pull comma-containing options out first
        if opt in s:
            found.append(opt)
            s = s.replace(opt, "")
    for tok in s.split(","):
        t = tok.strip()
        if not t or t.lower().startswith("click to write"):
            continue
        found.append(t)
    return [norm_label(x) for x in found]


def engage_label(qtext):
    """The matrix item label is the text after the ' - ' in the question text."""
    return qtext.split(" - ", 1)[1].strip() if " - " in qtext else qtext.strip()


def parse_number(text):
    """Best-effort numeric extraction from messy free text.
    Returns (value or None, confidence 'ok'|'approx'|'none')."""
    if text is None:
        return None, "none"
    raw = text.strip()
    if not raw:
        return None, "none"
    low = raw.lower()
    if low in ("n/a", "na", "none", "no", "unknown", "-"):
        return None, "none"

    # Pull all number tokens (allow commas + decimals).
    nums = [float(n.replace(",", "")) for n in re.findall(r"\d[\d,]*\.?\d*", raw)]
    if not nums:
        return None, "none"

    approx = False
    # Range like "500-700" or "500 - 700" -> midpoint.
    rng = re.search(r"(\d[\d,]*\.?\d*)\s*[-–to]+\s*(\d[\d,]*\.?\d*)", raw)
    if rng:
        a = float(rng.group(1).replace(",", ""))
        b = float(rng.group(2).replace(",", ""))
        return (a + b) / 2.0, "approx"

    val = nums[0]                       # leading number is the headline figure
    if len(nums) > 1:                   # extra numbers ("1500 Gymnastics, 60 Preschool") -> uncertain
        approx = True
    if re.search(r"\b(approx|around|about|over|more than|\+|~)\b", low) or "+" in raw:
        approx = True
    return val, ("approx" if approx else "ok")


def mean(vals):
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else None


def stdev(vals):
    vals = [v for v in vals if v is not None]
    if len(vals) < 2:
        return 0.0 if vals else None
    m = sum(vals) / len(vals)
    return (sum((v - m) ** 2 for v in vals) / (len(vals) - 1)) ** 0.5


def variance(vals):
    if len(vals) < 2:
        return 0.0
    m = sum(vals) / len(vals)
    return sum((v - m) ** 2 for v in vals) / (len(vals) - 1)


# ---- Welch's two-sample t-test (pure Python; p-value via the incomplete beta function) ----
def _betacf(a, b, x):
    MAXIT, EPS, FPMIN = 200, 3e-7, 1e-30
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < FPMIN:
        d = FPMIN
    d = 1.0 / d
    h = d
    for m in range(1, MAXIT + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        de = d * c
        h *= de
        if abs(de - 1.0) < EPS:
            break
    return h


def _betai(a, b, x):
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    lb = (math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
          + a * math.log(x) + b * math.log(1.0 - x))
    bt = math.exp(lb)
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * _betacf(a, b, x) / a
    return 1.0 - bt * _betacf(b, a, 1.0 - x) / b


def welch_p(m1, v1, n1, m2, v2, n2):
    """Two-tailed p-value for the difference of two means (Welch). 1.0 if not computable."""
    if n1 < 2 or n2 < 2:
        return 1.0
    se2 = v1 / n1 + v2 / n2
    if se2 <= 0:
        return 1.0
    t = (m1 - m2) / math.sqrt(se2)
    df = se2 ** 2 / ((v1 / n1) ** 2 / (n1 - 1) + (v2 / n2) ** 2 / (n2 - 1))
    return _betai(df / 2.0, 0.5, df / (df + t * t))


def load_codebook_text():
    """Map each item's codebook code -> clean question wording (blank if unavailable)."""
    text = {}
    if openpyxl is None or not os.path.exists(CODEBOOK_FILE):
        return text
    wb = openpyxl.load_workbook(CODEBOOK_FILE, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        code = str(row[0]).strip() if row and row[0] else ""
        qt = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        if code and qt and code != "Code":
            text[code] = qt
    return text


def clean_item_text(txt):
    """Generalize the codebook's KCPEC-specific wording to any coalition."""
    txt = txt.replace("***", "").strip()
    txt = re.sub(r"\bthe\s+K?CPEC\b", "the coalition", txt, flags=re.I)
    txt = re.sub(r"\bK?CPEC\b", "the coalition", txt, flags=re.I)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt[0].upper() + txt[1:] if txt else txt


def build_items(codebook_text):
    """Ordered item metadata: code, domain, construct, clean question text, reverse flag."""
    items = []
    for c in CONSTRUCTS:
        for it in c["items"]:
            txt = clean_item_text(codebook_text.get(CODEBOOK_CODE.get(it, it), ""))
            rev = it in REVERSE_ITEMS
            if rev and txt:
                txt += " (reverse-scored)"
            items.append({"code": it, "domain": DOMAIN_OF.get(c["key"]), "construct": c["key"],
                          "constructName": c["name"], "text": txt or it, "reverse": rev})
    return items


def _read_raw_rows(path):
    """Return the raw grid of cells for a .csv or .xlsx file."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        if openpyxl is None:
            raise RuntimeError("openpyxl is required to read .xlsx files: pip3 install openpyxl")
        import warnings
        warnings.filterwarnings("ignore")               # "no default style" on non-Excel exports
        wb = openpyxl.load_workbook(path, data_only=True)   # full load: dims are reliable
        return list(wb.active.iter_rows(values_only=True))
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def read_file(path):
    """Return (header, question_text, data_rows, coded), detecting the format by content
    (not extension) so Qualtrics exports work as either .csv or .xlsx.

    * Qualtrics export  -> header[0] == 'StartDate'; row 1 is question text; an optional
      importId row (row 2) is skipped when present; Likert answers are text. coded=False.
    * Pre-coded workbook -> anything else (header[0] == 'ID'); no question-text row;
      Likert answers are numeric; multi-selects are binary columns. coded=True.
    """
    rows = _read_raw_rows(path)
    header = [str(h) if h is not None else "" for h in rows[0]]
    if header and header[0] == "StartDate":
        qtext = [str(x) if x is not None else "" for x in rows[1]]
        row2 = rows[2] if len(rows) > 2 else []
        has_importid = any("importid" in str(c).lower() for c in row2)
        data = rows[3:] if has_importid else rows[2:]
        return header, qtext, data, False
    return header, header, rows[1:], True


# EngageFeedback matrix item labels (the xlsx has no question text; csv supplies these).
ENGAGE_LABELS = [
    "Increased stipend", "Strong facilittaion", "More networking and breakouts",
    "Clear agenda and goals", "In person meetings", "More varied and interesting content",
]
# In the coded xlsx, ages are exploded into binary columns in this order.
AGE_BINARY_COLS = [("Age_1", "0-5"), ("Age_2", "6-12"), ("Age_3", "13-17"), ("Age_4", "18-24")]

# Codebook value labels for the coded (Time 0) main-survey member-characteristic questions.
# (Age and AtRiskPops are NOT in the standardized codebook -- they vary by coalition -- so
# only Sector and Field are decoded for the coded format.)
SECTOR_CODES = {
    1: "Nonprofit or not-for-profit organization", 2: "For-profit organization",
    3: "Public organization or agency", 4: "Informal organization or association", 5: "Other",
}
FIELD_CODES = {
    1: "Business", 2: "Community-based organization", 3: "Criminal justice/safety",
    4: "Education", 5: "Faith", 6: "Outdoor recreation",
    7: "Parks, recreation, and leisure services", 8: "Politics",
    9: "Public health or healthcare", 10: "Research", 11: "Sports",
    12: "Social/human services", 13: "Youth development", 14: "Other",
}


def int_code(v):
    """Parse an integer option code from a coded cell (1, 1.0, '1', ' ' -> None)."""
    if is_blank(v):
        return None
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return None


# Social-connections network: the "list up to 10 organizations you're connected with" write-in
# slots. Same names in both formats; the name "Network_1" also appears in the Yes/No block, so
# we take the FIRST occurrence of each (the write-in block precedes the Yes/No block).
NET_WRITEIN_NAMES = ["Network_1", "Network_20", "Network_29", "Network_30", "Network_31",
                     "Network_32", "Network_33", "Network_34", "Network_35", "Network_36"]
NET_TOP_NODES = 30            # cap the map to the most-named organizations
ALIAS_FILE = os.path.join(DATA_DIR, "org_aliases.csv")


def norm_org(s):
    """Normalize a free-text organization name for matching (lowercase, strip punctuation and
    a few noise words). Two names that normalize equal are treated as the same node."""
    t = str(s).strip().lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\b(the|inc|llc|of|a)\b", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def load_aliases():
    """Optional editable data/org_aliases.csv (columns: variant,canonical). Maps a raw/variant
    org name onto a canonical display name so variants merge into one node."""
    amap = {}
    if os.path.exists(ALIAS_FILE):
        with open(ALIAS_FILE, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                variant = (r.get("variant") or "").strip()
                canonical = (r.get("canonical") or "").strip()
                if variant and canonical:
                    amap[norm_org(variant)] = canonical
    return amap


def extract_network(header, rows, alias_map):
    """Build a mention + co-mention org network for one coalition cell.
    nodes: organizations sized by how many members named them (where ties concentrate).
    edges: two orgs named by the same member (co-mention), weighted by that count."""
    net_idx, seen = [], set()
    for i, h in enumerate(header):                       # first occurrence of each write-in slot
        if h in NET_WRITEIN_NAMES and h not in seen:
            net_idx.append(i)
            seen.add(h)
    if not net_idx:
        return None

    def resolve(raw):
        nk = norm_org(raw)
        if not nk:
            return None, None
        if nk in alias_map:
            return norm_org(alias_map[nk]), alias_map[nk]
        return nk, raw.strip()

    mentions, disp, comention, n_resp = {}, {}, {}, 0
    for row in rows:
        orgs = {}
        for i in net_idx:
            raw = str(cell(row, i)).strip()
            if not raw or raw.lower() in ("n/a", "na", "none", "."):
                continue
            key, label = resolve(raw)
            if key:
                orgs[key] = label
        if not orgs:
            continue
        n_resp += 1
        for key, label in orgs.items():
            mentions[key] = mentions.get(key, 0) + 1
            d = disp.setdefault(key, {})
            d[label] = d.get(label, 0) + 1
        keys = sorted(orgs)
        for a in range(len(keys)):
            for b in range(a + 1, len(keys)):
                pair = (keys[a], keys[b])
                comention[pair] = comention.get(pair, 0) + 1
    if not mentions:
        return None

    top = sorted(mentions, key=lambda k: (-mentions[k], k))[:NET_TOP_NODES]
    topset = set(top)
    nodes = [{"id": k, "label": max(disp[k], key=disp[k].get), "mentions": mentions[k]} for k in top]
    edges = sorted(
        ({"source": a, "target": b, "weight": w} for (a, b), w in comention.items()
         if a in topset and b in topset),
        key=lambda e: -e["weight"])
    return {"nodes": nodes, "edges": edges, "respondents": n_resp, "totalOrgs": len(mentions)}


def extract_zip_youth(qtext, rows):
    """Cleveland-style 'youth served annually in each zip code' -> total youth per zip."""
    cols = []
    for i, qt in enumerate(qtext):
        s = str(qt)
        if "how many youth do you serve annually in each zip" in s.lower() and s.lower().rstrip().endswith("text"):
            parts = s.split(" - ")
            cols.append((i, parts[-2].strip() if len(parts) >= 2 else f"col{i}"))
    if not cols:
        return None
    items = []
    for i, zlabel in cols:
        total, n = 0.0, 0
        for row in rows:
            v, _ = parse_number(str(cell(row, i)))
            if v is not None:
                total += v
                n += 1
        if n:
            items.append({"label": zlabel, "value": round(total, 1), "n": n})
    return {"items": items} if items else None


def cell(raw_row, i):
    """Safe indexed access into a csv list-row or xlsx tuple-row."""
    if i is None or i >= len(raw_row):
        return ""
    v = raw_row[i]
    return "" if v is None else v


def process_cell(coalition, timepoint, header, idx, qtext, rows, coded,
                 cells, program, question_text_map,
                 prof_n, prof_counts, prof_answered, engage, engage_labels, universe,
                 all_item_cols, alias_map, network, zipyouth, item_stats, cseq_acc):
    # ---- Construct scoring (numeric-aware; reverse-coding applied to both formats) ----
    out = {}
    item_vals_map = {}      # item -> list of reverse-coded values for this cell
    for c in CONSTRUCTS:
        reverse = set(c.get("reverse", []))
        present = [it for it in c["items"] if it in idx]
        for it in present:
            if header[idx[it]] not in question_text_map:
                question_text_map[header[idx[it]]] = qtext[idx[it]] if idx[it] < len(qtext) else ""
        resp_means = []
        for row in rows:
            vals = []
            for it in present:
                v = to_numeric(cell(row, idx[it]), c["scale"])
                if v is not None and it in reverse:
                    v = (SCALE_MAX + 1) - v
                if v is not None:
                    vals.append(v)
                    item_vals_map.setdefault(it, []).append(v)
            if vals:
                resp_means.append(sum(vals) / len(vals))
        out[c["key"]] = {
            "mean": round(mean(resp_means), 3) if resp_means else None,
            "sd": round(stdev(resp_means), 3) if resp_means else None,
            "n": len(resp_means),
        }
    cells[(coalition, timepoint)] = out

    # Per-item stats for this cell, and pool Time 0 values into the CSEq baseline.
    istats = {}
    for it, vs in item_vals_map.items():
        istats[it] = {"mean": round(mean(vs), 3), "sd": round(stdev(vs), 3) if len(vs) > 1 else 0.0,
                      "var": variance(vs), "n": len(vs)}
        if timepoint == "T0":
            acc = cseq_acc.setdefault(it, {"n": 0, "sum": 0.0, "sumsq": 0.0})
            acc["n"] += len(vs)
            acc["sum"] += sum(vs)
            acc["sumsq"] += sum(v * v for v in vs)
    item_stats[(coalition, timepoint)] = istats

    # ---- Program impact (free-text fields; resolved by question text, not variable name) ----
    prog = {}
    for fkey, meta in PROGRAM_FIELDS.items():
        # Find the column whose question text contains the field's match string.
        ci = next((i for i, qt in enumerate(qtext) if meta["match"] in str(qt).lower()), None)
        parsed, raws = [], []
        if ci is not None:
            for row in rows:
                txt = str(cell(row, ci))
                val, conf = parse_number(txt)
                if txt.strip():
                    raws.append({"raw": txt.strip(), "value": val, "conf": conf})
                if val is not None:
                    parsed.append(val)
        agg_val = (round(sum(parsed), 2) if meta["agg"] == "sum" else round(mean(parsed), 2)) if parsed else None
        prog[fkey] = {"value": agg_val, "n": len(parsed), "raw": raws}
    program[(coalition, timepoint)] = prog

    # ---- Respondent count N ----
    if coded:
        n_resp = sum(1 for row in rows if any(not is_blank(cell(row, idx[it]))
                                              for it in all_item_cols if it in idx))
    else:
        ri = idx.get("Respondent")
        n_resp = sum(1 for row in rows if not is_blank(cell(row, ri)))
    prof_n[(coalition, timepoint)] = n_resp

    # ---- Profile distributions (format-aware) ----
    counts = {qk: {} for qk in PROFILE_QUESTIONS}
    answered = {qk: 0 for qk in PROFILE_QUESTIONS}

    def add(qk, label):
        counts[qk][label] = counts[qk].get(label, 0) + 1
        universe[qk][label] = universe[qk].get(label, 0) + 1

    if coded:
        # Decoded via the survey codebook: engagement (numeric), ages (binary),
        # sector (numeric code) and field (binary columns). AtRiskPops is a coalition-specific
        # question absent from the codebook, so populations stay hidden for the coded format.
        mi = idx.get("MemEng_Part")
        if mi is not None:
            for row in rows:
                v = to_numeric(cell(row, mi), "frequency")
                if v is not None:
                    answered["engagement"] += 1
                    add("engagement", FREQ_ORDER[int(round(v)) - 1])
        if any(col in idx for col, _ in AGE_BINARY_COLS):
            for row in rows:
                picked = [band for col, band in AGE_BINARY_COLS
                          if col in idx and str(cell(row, idx[col])).strip() in ("1", "1.0")]
                if picked:
                    answered["age"] += 1
                    for band in picked:
                        add("age", band)
        di = idx.get("MemEng_Dur")
        if di is not None:
            for row in rows:
                v = int_code(cell(row, di))
                if v and 1 <= v <= 5:
                    answered["duration"] += 1
                    add("duration", DUR_ORDER[v - 1])
        si = idx.get("Sector")
        if si is not None:
            for row in rows:
                label = SECTOR_CODES.get(int_code(cell(row, si)))
                if label:
                    answered["sector"] += 1
                    add("sector", label)
        field_cols = [(idx[f"Field_{j}"], FIELD_CODES[j]) for j in range(1, 15) if f"Field_{j}" in idx]
        if field_cols:
            for row in rows:
                picked = [label for ci, label in field_cols
                          if str(cell(row, ci)).strip() in ("1", "1.0")]
                if picked:
                    answered["field"] += 1
                    for label in picked:
                        add("field", label)
    else:
        for qk, meta in PROFILE_QUESTIONS.items():
            col = meta["col"]
            if col not in idx:
                continue
            for row in rows:
                raw = str(cell(row, idx[col])).strip()
                if not raw:
                    continue
                answered[qk] += 1
                opts = parse_multi(raw) if meta["type"] == "multi" else [norm_label(raw)]
                for o in opts:
                    add(qk, o)
    prof_counts[(coalition, timepoint)] = counts
    prof_answered[(coalition, timepoint)] = answered

    # ---- EngageFeedback matrix: mean degree (1-5) per item ----
    eng = {}
    for n, item in enumerate(ENGAGE_ITEMS):
        if item not in idx:
            continue
        i = idx[item]
        if coded:
            label = ENGAGE_LABELS[n] if n < len(ENGAGE_LABELS) else item
        else:
            label = engage_label(qtext[i] if i < len(qtext) else item)
        if label not in engage_labels:
            engage_labels.append(label)
        vals = [to_numeric(cell(row, i), "frequency") for row in rows]
        vals = [v for v in vals if v is not None]
        eng[label] = {"sum": sum(vals), "n": len(vals)}
    engage[(coalition, timepoint)] = eng

    # ---- Social-connections network + coalition-specific extras ----
    net = extract_network(header, rows, alias_map)
    if net:
        network[(coalition, timepoint)] = net
    zy = extract_zip_youth(qtext, rows)
    if zy:
        zipyouth[(coalition, timepoint)] = zy


# --------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------
def main():
    with open(MANIFEST, newline="", encoding="utf-8-sig") as f:
        manifest = [r for r in csv.DictReader(f) if r.get("file", "").strip()]

    coalitions, timepoints = [], []
    # cells[(coalition, timepoint)] = {construct_key: {"mean", "sd", "n"}, ...}
    cells = {}
    program = {}        # program[(coalition, timepoint)] = {field: {value, n, raw:[...]}}
    question_text_map = {}

    # Respondent / org profile accumulators.
    prof_n = {}                       # prof_n[(co, tp)] = respondent count
    prof_counts = {}                  # prof_counts[(co, tp)][qkey] = {label: count}
    prof_answered = {}                # prof_answered[(co, tp)][qkey] = # who answered
    engage = {}                       # engage[(co, tp)] = {label: {"sum": x, "n": k}}
    engage_labels = []                # ordered EngageFeedback item labels
    universe = {qk: {} for qk in PROFILE_QUESTIONS}   # qkey -> {label: global count}
    network = {}                      # network[(co, tp)] = {nodes, edges, ...}
    zipyouth = {}                     # zipyouth[(co, tp)] = {items:[{label,value,n}]}
    item_stats = {}                   # item_stats[(co, tp)][item] = {mean, sd, var, n}
    cseq_acc = {}                     # cseq_acc[item] = {n, sum, sumsq}  (Time 0 pool)
    alias_map = load_aliases()
    codebook_text = load_codebook_text()
    items_meta = build_items(codebook_text)

    print("MOVE Fund dashboard build")
    print("=" * 70)

    all_item_cols = [it for c in CONSTRUCTS for it in c["items"]]

    for entry in manifest:
        fname = entry["file"].strip()
        man_coalition = entry["coalition"].strip()
        timepoint = entry["timepoint"].strip()
        path = os.path.join(DATA_DIR, fname)
        if not os.path.exists(path):
            print(f"  !! missing file, skipping: {fname}")
            continue
        if timepoint not in timepoints:
            timepoints.append(timepoint)

        header, qtext, data, coded = read_file(path)
        idx = {h: i for i, h in enumerate(header)}

        # Group rows into one cell per coalition (split a multi-coalition file by its column).
        if man_coalition:
            groups = [(man_coalition, data)]
        else:
            ci = idx.get("Coalition")
            buckets = {}
            for row in data:
                co = (str(row[ci]).strip() if ci is not None and row[ci] is not None else "")
                if co:
                    buckets.setdefault(co, []).append(row)
            groups = sorted(buckets.items())

        print(f"\n{fname}  [{timepoint}, {'coded' if coded else 'qualtrics'}]  "
              f"-> {len(groups)} coalition(s)")

        for coalition, rows in groups:
            if coalition not in coalitions:
                coalitions.append(coalition)
            print(f"    [{coalition}]  {len(rows)} rows")
            process_cell(coalition, timepoint, header, idx, qtext, rows, coded,
                         cells, program, question_text_map,
                         prof_n, prof_counts, prof_answered,
                         engage, engage_labels, universe, all_item_cols,
                         alias_map, network, zipyouth, item_stats, cseq_acc)

    timepoints = sorted(set(timepoints))      # T0, T1, T2 ...
    coalitions = sorted(set(coalitions))       # alphabetical for the dropdown

    # ---- Build per-cell profile output (ordered items with percentages) ----
    def pct(count, n):
        return round(100.0 * count / n, 1) if n else 0.0

    profiles = {}
    for (co, tp) in cells:
        n = prof_n.get((co, tp), 0)
        counts = prof_counts.get((co, tp), {})
        cellprof = {"n": n}
        for qk, meta in PROFILE_QUESTIONS.items():
            qc = counts.get(qk, {})
            if meta["mode"] == "order":
                labels = meta["order"]
            elif meta["mode"] == "universe":
                labels = sorted(universe[qk], key=lambda L: (-universe[qk][L], L))
            else:  # present
                labels = sorted(qc, key=lambda L: (-qc[L], L))
            items = [{"label": L, "count": qc.get(L, 0), "pct": pct(qc.get(L, 0), n)} for L in labels]
            cellprof[qk] = {"answered": prof_answered.get((co, tp), {}).get(qk, 0), "items": items}
        # EngageFeedback means in item order
        eng = engage.get((co, tp), {})
        cellprof["engageFeedback"] = {"items": [
            {"label": L, "mean": round(eng[L]["sum"] / eng[L]["n"], 2) if eng.get(L, {}).get("n") else None,
             "n": eng.get(L, {}).get("n", 0)}
            for L in engage_labels
        ]}
        profiles[(co, tp)] = cellprof

    # ---- Grand means: average of coalition means per construct, per timepoint ----
    grand = {}
    for tp in timepoints:
        grand[tp] = {}
        for c in CONSTRUCTS:
            cmeans = [cells[(co, tp)][c["key"]]["mean"]
                      for co in coalitions
                      if (co, tp) in cells and cells[(co, tp)][c["key"]]["mean"] is not None]
            grand[tp][c["key"]] = round(mean(cmeans), 3) if cmeans else None

    # ---- CSEq baseline (pooled Time 0 respondents) per item, and per-item significance ----
    cseq_item = {}
    for it, acc in cseq_acc.items():
        n = acc["n"]
        if n >= 1:
            m = acc["sum"] / n
            var = (acc["sumsq"] - acc["sum"] ** 2 / n) / (n - 1) if n > 1 else 0.0
            cseq_item[it] = {"mean": round(m, 3), "var": var, "n": n}

    item_scores = {}
    for (co, tp), istats in item_stats.items():
        cellout = {}
        for it, st in istats.items():
            base = cseq_item.get(it)
            sig, direction = False, None
            if base:
                p = welch_p(st["mean"], st["var"], st["n"], base["mean"], base["var"], base["n"])
                sig = p < 0.05
                direction = "higher" if st["mean"] > base["mean"] else "lower"
            cellout[it] = {"mean": st["mean"], "n": st["n"],
                           "cseqMean": base["mean"] if base else None,
                           "sig": sig, "dir": direction}
        item_scores[(co, tp)] = cellout

    # ---- Domain scores per cell (mean of the domain's construct means) + CSEq domain avg ----
    def domain_mean(source, dkey):
        vals = [source[ck]["mean"] if isinstance(source[ck], dict) else source[ck]
                for d in DOMAINS if d["key"] == dkey for ck in d["constructs"]
                if ck in source and (source[ck]["mean"] if isinstance(source[ck], dict) else source[ck]) is not None]
        return round(mean(vals), 3) if vals else None

    domain_scores = {}
    for (co, tp) in cells:
        domain_scores[(co, tp)] = {d["key"]: domain_mean(cells[(co, tp)], d["key"]) for d in DOMAINS}
    cseq_domain = {d["key"]: (round(mean([grand["T0"][ck] for ck in d["constructs"]
                                          if grand.get("T0", {}).get(ck) is not None]), 3)
                              if grand.get("T0") else None)
                   for d in DOMAINS}

    # ---- Serialize ----
    out = {
        "generatedFrom": [e["file"] for e in manifest],
        "coalitions": coalitions,
        "timepoints": timepoints,
        "scaleMax": SCALE_MAX,
        "constructs": [{"key": c["key"], "name": c["name"], "scale": c["scale"],
                        "items": c["items"]} for c in CONSTRUCTS],
        "questionText": question_text_map,
        "programFields": [{"key": k, "label": v["label"], "agg": v["agg"]}
                          for k, v in PROGRAM_FIELDS.items()],
        "cells": {f"{co}||{tp}": cells[(co, tp)] for (co, tp) in cells},
        "grandMeans": grand,
        "program": {f"{co}||{tp}": program[(co, tp)] for (co, tp) in program},
        "profileQuestions": [{"key": k, "label": v["label"], "type": v["type"], "mode": v["mode"]}
                             for k, v in PROFILE_QUESTIONS.items()],
        "engageFeedbackPrompt": "To what degree would any of the following help you engage "
                                "more regularly in action team or committee meetings?",
        "profiles": {f"{co}||{tp}": profiles[(co, tp)] for (co, tp) in profiles},
        "network": {f"{co}||{tp}": network[(co, tp)] for (co, tp) in network},
        "zipYouth": {f"{co}||{tp}": zipyouth[(co, tp)] for (co, tp) in zipyouth},
        "domains": DOMAINS,
        "items": items_meta,
        "itemScores": {f"{co}||{tp}": item_scores[(co, tp)] for (co, tp) in item_scores},
        "cseqItems": cseq_item,
        "domainScores": {f"{co}||{tp}": domain_scores[(co, tp)] for (co, tp) in domain_scores},
        "cseqDomain": cseq_domain,
    }

    data_js = "// Auto-generated by build_data.py - do not edit by hand.\nconst DASHBOARD_DATA = " \
        + json.dumps(out, indent=2, ensure_ascii=False) + ";\n"
    with open(OUT_JS, "w", encoding="utf-8") as f:
        f.write(data_js)

    # Also emit a single self-contained HTML with the data inlined -- easiest to host
    # (one file, no separate dashboard_data.js path to get wrong on a website).
    wrote_standalone = write_standalone(data_js)

    print("\n" + "=" * 70)
    print(f"Wrote {OUT_JS}")
    if wrote_standalone:
        print(f"Wrote {STANDALONE_HTML}  (single self-contained file for hosting)")
    print(f"  coalitions: {coalitions}")
    print(f"  timepoints: {timepoints}")
    print(f"  constructs: {len(CONSTRUCTS)}")


def write_standalone(data_js):
    """Inline dashboard_data.js into index.html -> dashboard_standalone.html (one file)."""
    if not os.path.exists(INDEX_HTML):
        return False
    with open(INDEX_HTML, encoding="utf-8") as f:
        html = f.read()
    tag = '<script src="dashboard_data.js"></script>'
    if tag not in html:
        return False
    # Guard against an accidental </script> inside the data breaking the inline tag.
    safe = data_js.replace("</script>", "<\\/script>")
    html = html.replace(tag, "<script>\n" + safe + "\n</script>")
    with open(STANDALONE_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    return True


if __name__ == "__main__":
    main()
