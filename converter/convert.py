"""Supplier file -> TyresCart format converter (Lookin supplier, v1: rules only).

Usage:  python3 convert.py "<mapping guide xlsx>" [out.xlsx]
Implements Mapping Guide steps 1-11. Pattern/OEM/RunFlat/EV are parsed with rules
plus the TyresCart pattern vocabulary. A Claude fallback can be added later for
rows flagged low-confidence.
"""
import re
import sys
import difflib
import warnings
from collections import Counter

import pandas as pd

warnings.filterwarnings("ignore")

from brands import BrandResolver  # noqa: E402

# ---------- legacy Lookin config (kept for reference; brand now comes from brands.py) ---------- Add one dict per new supplier ----------
LOOKIN_BRAND_PREFIX = {  # description prefix -> TyresCart mgs_brand
    "BR": "Bridgestone", "CL": "Continental", "D": "Dunlop", "GY": "Goodyear",
    "P": "Pirelli", "YOKO": "Yokohama", ".YOKO": "Yokohama", ".TOYO": "Toyo",
    "HB": "Habilead", "HK": "Hankook", "IMP": "Imperial", "MIC": "Michelin",
    "TRACAMX": "Tracmax", "MATARX": "Matrax", "TRAZNO": "Zelda", "TRZ": "Zelda",
    "NEOLIN-EV": "Neolin", "WINDFORCE": "Windforce",
    "MASIMO": "Masimon", "NIPON": "Masimon", "NANKING": "Nankang",
    "HONOUR": "Honour", "ANNITE": "Annaite",
}
# supplier Brand column spelling -> TyresCart brand (None = unknown)
BRAND_ALIAS = {"WIND FORCE": "Windforce", "DEESTONE": "Deestone", "TRAZANO": "Zelda",
               "CHINA": None, "ACCCELERA": "Accelera"}
NOISE = {"RFID", "TL", "TT", "LOAD", "EXTRA", "(OE)", "OE", "S", "THA", "THAI", "TH", "JPN", "JAP", "KOR", "XL", "FR", "NO", "OE", "THI"}

SIZE_RE = re.compile(r"(\d{3})/(\d{2})\s*Z?R\s*(\d{2})C?")
LOAD_RE = re.compile(r"\b(\d{2,3}(?:/\d{2,3})?[A-Z])\b")
YEAR_RE = re.compile(r"(?<!\d)(20[12]\d)(?!\d)")
SECOND_SIZE_RE = re.compile(r"\b\d{3}/\d{2}R\d{2}\b")


def norm(s):
    return re.sub(r"[^A-Z0-9]", "", str(s).upper()) if pd.notna(s) and s != "" else ""


def clean_text(s):
    """Step 1: trim and collapse repeated spaces."""
    return re.sub(r"\s+", " ", str(s)).strip() if pd.notna(s) else s


def build_vocab(T):
    pats = {b: sorted(set(g.pattern.dropna()), key=lambda x: (-len(x), x))
            for b, g in T.groupby("mgs_brand")}
    oems = sorted(set(T.oem_tyres.dropna()), key=lambda x: (-len(x), x))
    return pats, oems


def resolve_brand(desc, brand_col, brands_lc):
    first = desc.split()[0].upper() if desc else ""
    if first in LOOKIN_BRAND_PREFIX:
        return LOOKIN_BRAND_PREFIX[first]
    b = str(brand_col).strip()
    if b.upper() in BRAND_ALIAS:
        return BRAND_ALIAS[b.upper()]
    return brands_lc.get(b.lower())


def parse_description(desc, brand, pats, oems, pool=None, extra_pats=None):
    """Step 2: size/load/year by regex; oem/runflat/ev/pattern by rules + TyresCart vocab."""
    out = dict(size=None, width=None, height=None, rim=None, load=None, year=None,
               pattern=None, oem=None, runflat=None, ev=None, conf="high")
    desc, rfid_year = strip_rfid(desc.upper())
    d = " " + desc + " "
    m = SIZE_RE.search(d)
    if m:
        out.update(width=int(m[1]), height=int(m[2]), rim=int(m[3]),
                   size=f"{m[1]}/{m[2]} R{m[3]}")
        d = d.replace(m[0], " ", 1)
    d = SECOND_SIZE_RE.sub(" ", d)  # staggered second size
    m = LOAD_RE.search(d)
    if m:
        out["load"] = m[1]
        d = d.replace(m[0], " ", 1)
    m = YEAR_RE.search(d)
    if m:
        out["year"] = int(m[1])
        d = d.replace(m[0], " ", 1)
    elif rfid_year:
        out["year"] = rfid_year
    m = re.search(r"-\s*(2\d)\b(?!\s*[/\dX])", d)       # short year: "-25", "- 25", "RFT-25"
    if m:
        if out["year"] is None:
            out["year"] = 2000 + int(m[1])
        d = d.replace(m[0], " ", 1)
    d = re.sub(r"-\d{2}\b", " ", d)
    d = " " + " ".join(d.split()[1:]) + " "  # drop brand prefix word
    d = d.replace("-", " ")

    if re.search(r"\b(RFT|RUNFLAT|SSR|ZP|ROF)\b|\bRUN FLAT\b", d):
        out["runflat"] = "RunFlat"
        d = re.sub(r"\b(RFT|RUNFLAT|SSR|ZP|ROF)\b|\bRUN FLAT\b", " ", d)
    if re.search(r"\bEV\b", d):
        out["ev"] = "EV"
        d = re.sub(r"\bEV\b", " ", d)
    found = []
    for o in oems:
        pat = r"(?<![A-Z0-9])" + re.escape(o.upper()) + r"(?![A-Z0-9])"
        if o.strip() and re.search(pat, d):
            found.append(o)
            d = re.sub(pat, " ", d, count=1)
            if len(found) == 2:
                break
    out["oem"] = " ".join(found) if found else None

    toks = [t for t in d.split() if t not in NOISE]
    left = " ".join(toks)
    nl = norm(left)

    def score(p):
        """How well supplier leftover text explains a TyresCart pattern."""
        np_ = norm(p)
        if not np_ or not nl:
            return 0
        if np_ in nl:
            return 0.6 + 0.4 * len(np_) / len(nl)
        # abbreviation: every supplier token appears inside the pattern
        ptoks = [norm(t) for t in toks if norm(t)]
        hit = sum(len(t) for t in ptoks if t in np_)
        cover = hit / max(sum(len(t) for t in ptoks), 1)
        return max(difflib.SequenceMatcher(None, nl, np_).ratio(), 0.55 + 0.45 * cover if cover > 0.6 else 0)

    best, score_ = None, 0
    # narrow pool (same brand+size+load+year) first, then brand-wide vocab
    stages = [(sorted({p for p in (pool or [])}), 0.62), (pats.get(brand, []), 0.8)]
    if extra_pats:  # last resort: pattern spellings seen on other websites (e.g. Pitstop Arabia)
        stages.append((extra_pats.get(brand, []), 0.8))
    for source, thr in stages:
        for p in source:
            sc = score(p)
            if sc > score_:
                best, score_ = p, sc
        if best and score_ >= thr:
            break
        best, score_ = None, 0
    if best:
        out["pattern"] = best
    else:
        out["pattern"] = left.title() if left else None
        out["conf"] = "low"
    return out


RFID_RE = re.compile(r"-?\s*(?:RFID\s*-?\s*(\d{2})|(\d{2})\s*R\s?FID)\b|-?\s*RFID\b", re.I)


def strip_rfid(desc):
    """'... - 25RFID', '- RFID 24', '- 23R FID' -> (clean description, 2-digit year or None)."""
    yr = None
    def sub(m):
        nonlocal yr
        v = m.group(1) or m.group(2)
        if v and yr is None:
            yr = 2000 + int(v)
        return " "
    return RFID_RE.sub(sub, desc), yr


def quick_year(desc):
    """Year from the description only: 2026 / 26 after a dash (e.g. '-25', '- 25 RFID')."""
    desc, ry = strip_rfid(desc.upper())
    if ry:
        return ry
    d = SIZE_RE.sub(" ", desc, 1)
    m = YEAR_RE.search(d)
    if m:
        return int(m[1])
    m = re.search(r"-\s*(2\d)\b(?!\s*[/\dX])", d)
    return 2000 + int(m[1]) if m else None


def key1(b, s, l, y): return "|".join(map(norm, (b, s, l, y)))
def key2(b, s, y): return "|".join(map(norm, (b, s, y)))
def key3(p, o, e, r): return "|".join(map(norm, (p, o, e, r)))


def convert(L, T, extra_vocab=None):
    """Convert a supplier DataFrame (Lookin layout) to TyresCart format.
    Returns (tyrescart_format_df, working_df, stats_dict)."""
    L = L.copy()
    T = T.copy()

    # Step 1
    for c in L.select_dtypes("object"):
        L[c] = L[c].map(clean_text)

    pats, oems = build_vocab(T)
    extra_pats = {}
    if extra_vocab:  # e.g. Pitstop Arabia: more OEM markings + pattern spellings
        extra_pats = extra_vocab.get("patterns", {})
    brands_lc = {b.lower(): b for b in T.mgs_brand.unique()}

    # Step 2
    T["_k1"] = [key1(*a) for a in zip(T.mgs_brand, T.tyre_size, T.load_index, T.year)]
    T["_k2"] = [key2(*a) for a in zip(T.mgs_brand, T.tyre_size, T.year)]
    pool1 = T.groupby("_k1").pattern.apply(list).to_dict()
    pool2 = T.groupby("_k2").pattern.apply(list).to_dict()
    resolver = BrandResolver(T.mgs_brand.unique())
    rows = []
    sup_year = L["Year"].tolist() if "Year" in L.columns else [None] * len(L)
    sup_load = L["Load"].tolist() if "Load" in L.columns else [None] * len(L)
    for desc0, sy, sl in zip(L.Description, sup_year, sup_load):
        desc = desc0
        # the description lacks a year / load-speed but the sheet has them in their own columns -> use them
        if quick_year(desc) is None and pd.notna(sy) and re.fullmatch(r"20[12]\d", str(sy).split(".")[0]):
            desc += f" {str(sy).split('.')[0]}"
        if not LOAD_RE.search(SIZE_RE.sub(" ", desc.upper(), 1)) and pd.notna(sl) and sl:
            desc += f" {sl}"
        brand = resolver(desc)
        m = SIZE_RE.search(desc.upper())
        sz = f"{m[1]}/{m[2]} R{m[3]}" if m else None
        ld = LOAD_RE.search(SIZE_RE.sub(" ", desc.upper(), 1))
        y = quick_year(desc)
        pool = pool1.get(key1(brand, sz, ld[1] if ld else None, y)) \
            or pool2.get(key2(brand, sz, y))
        p = parse_description(desc, brand, pats, oems, pool, extra_pats)
        p["brand"] = brand
        rows.append(p)
    P = pd.DataFrame(rows)
    P["year"] = P["year"].astype("Int64")

    # Steps 3, 4, 8
    L["mgs_brand"] = P.brand
    L["Key1"] = [key1(*a) for a in zip(P.brand, P["size"], P.load, P.year)]
    L["Key2"] = [key2(*a) for a in zip(P.brand, P["size"], P.year)]
    L["Key3"] = [key3(*a) for a in zip(P.pattern, P.oem, P.ev, P.runflat)]
    # Steps 5, 6, 9
    T["Key1"] = [key1(*a) for a in zip(T.mgs_brand, T.tyre_size, T.load_index, T.year)]
    T["Key2"] = [key2(*a) for a in zip(T.mgs_brand, T.tyre_size, T.year)]
    T["Key3"] = [key3(*a) for a in zip(T.pattern, T.oem_tyres, T.ev_tyre, T.runflat)]
    # Step 7
    c1, c2 = Counter(T.Key1), Counter(T.Key2)
    L["Key1 Count"] = L.Key1.map(c1).fillna(0).astype(int)
    L["Key2 Count"] = L.Key2.map(c2).fillna(0).astype(int)
    # Steps 10, 11
    by1 = T.groupby(["Key1", "Key3"]).sku.apply(list).to_dict()
    by2 = T.groupby(["Key2", "Key3"]).sku.apply(list).to_dict()
    L["TyresCart SKU"] = None
    L["TyresCart SKU 2"] = None
    L["Review note"] = ""
    for i, r in L.iterrows():
        if r["Key1 Count"] > 0:
            m = by1.get((r.Key1, r.Key3))
            if m:
                L.at[i, "TyresCart SKU"] = m[0]
                if len(m) > 1:
                    L.at[i, "Review note"] = f"{len(m)} SKUs matched Key1+Key3"
                continue
        if r["Key2 Count"] > 0:
            m = by2.get((r.Key2, r.Key3))
            if m:
                L.at[i, "TyresCart SKU 2"] = m[0]
                if len(m) > 1:
                    L.at[i, "Review note"] = f"{len(m)} SKUs matched Key2+Key3"
    L.loc[(P.conf == "low").values, "Review note"] += " pattern not in TyresCart vocab"
    L.loc[P.brand.isna().values, "Review note"] += " brand unknown"
    L["Review note"] = L["Review note"].str.strip()

    # ---- TyresCart-format output ----
    cat = T.groupby("mgs_brand").tyres_category.agg(lambda s: s.mode().iat[0]).to_dict()
    war = T.groupby("mgs_brand").warranty_period.agg(lambda s: s.mode().iat[0]).to_dict()
    out = pd.DataFrame(index=L.index, columns=list(T.columns.drop(["Key1", "Key2", "Key3", "_k1", "_k2"])))
    out["sku"] = L["TyresCart SKU"].fillna(L["TyresCart SKU 2"])
    out["mgs_brand"] = P.brand
    out["pattern"] = P.pattern
    out["runflat"] = P.runflat
    out["oem_tyres"] = P.oem
    out["ev_tyre"] = P.ev
    out["display_name"] = [" ".join(x for x in (a, b) if x)
                           for a, b in zip(P.pattern.fillna(""), P.oem.fillna(""))]
    out["tyre_size"], out["width"], out["height"], out["rim"] = P["size"], P.width, P.height, P.rim
    out["load_index"], out["year"] = P.load, P.year
    out["construction"] = "R"
    out["country"] = L["Origin"].str.title() if "Origin" in L.columns else None
    out["tyres_category"] = P.brand.map(cat)
    out["warranty_period"] = P.brand.map(war)
    out["name"] = [" ".join(str(x) for x in (b, s, l, d, y) if pd.notna(x) and x != "")
                   for b, s, l, d, y in zip(P.brand, P["size"], P.load, out.display_name, P.year)]
    if "Price" in L.columns:
        out["price"] = L["Price"]
    for col in ["ItemCode", "Description", "Stock", "Credit"]:
        out[f"supplier_{col.lower()}"] = L[col] if col in L.columns else None
    for col in [c for c in L.columns if c.startswith("supplier_")]:
        out[col] = L[col]
    out["match_status"] = ["MATCHED (Size1)" if pd.notna(a) else "MATCHED (Size2)" if pd.notna(b)
                           else "NO MATCH" for a, b in zip(L["TyresCart SKU"], L["TyresCart SKU 2"])]
    out["review_note"] = L["Review note"]
    # TyresCart side of the matched record, next to the supplier_* columns
    tc = T.drop_duplicates("sku").set_index("sku")
    for col in [c for c in T.columns if not c.startswith(("Key", "_k"))]:  # all TyresCart columns
        out["tyrescart_" + col] = out["sku"].map(tc[col]) if col != "sku" else out["sku"]
    sup_price = pd.to_numeric(out["supplier_credit"], errors="coerce")
    out["price_difference"] = sup_price - pd.to_numeric(out["tyrescart_price"], errors="coerce")

    stats = dict(rows=len(L), size1=int(L["TyresCart SKU"].notna().sum()),
                 size2=int(L["TyresCart SKU 2"].notna().sum()),
                 unmatched=int(out.match_status.eq("NO MATCH").sum()),
                 low_conf=int((P.conf == "low").sum()), unknown_brand=int(P.brand.isna().sum()),
                 no_size=int(P["size"].isna().sum()), no_load=int(P.load.isna().sum()))
    return out, L, stats


def write_excel(out, L, path):
    with pd.ExcelWriter(path) as w:
        out.to_excel(w, sheet_name="TyresCart_Format", index=False)
        L.to_excel(w, sheet_name="Lookin_Working", index=False)
        out[out.review_note.ne("") | out.match_status.eq("NO MATCH")].to_excel(
            w, sheet_name="Review", index=False)


def build_stock_view(out, T, n_supplier_rows):
    """Every TyresCart row + the supplier's matching item/stock/price (blank when the supplier lacks it).
    Returns (summary_df, full_stock_df, supplier_not_in_tyrescart_df)."""
    m = out[out.sku.notna()].copy()
    m["_st"] = pd.to_numeric(m.supplier_stock, errors="coerce")
    m = m.sort_values("_st", ascending=False)
    dup = m.groupby("sku").size()
    first = m.drop_duplicates("sku").set_index("sku")
    res = T.copy()
    res["in_supplier"] = res.sku.map(lambda s: "YES" if s in first.index else "NO")
    res["match_type"] = res.sku.map(first.match_status)
    res["supplier_itemcode"] = res.sku.map(first.supplier_itemcode)
    res["supplier_description"] = res.sku.map(first.supplier_description)
    res["supplier_stock"] = res.sku.map(first.supplier_stock)
    res["supplier_price"] = pd.to_numeric(res.sku.map(first.supplier_credit), errors="coerce")
    res["supplier_rows_for_sku"] = res.sku.map(dup)
    res["price_diff"] = res.supplier_price - pd.to_numeric(res.price, errors="coerce")
    un = out[out.match_status == "NO MATCH"][["supplier_itemcode", "supplier_description", "supplier_stock",
         "supplier_credit", "mgs_brand", "pattern", "tyre_size", "load_index", "year", "review_note"]]
    summary = pd.DataFrame({
        "Item": ["TyresCart tyres (total)", "Found in supplier file", "Not in supplier file", "Supplier rows",
                 "Supplier rows matched", "Supplier rows NOT in TyresCart", "SKUs hit by more than 1 supplier row"],
        "Count": [len(T), int((res.in_supplier == "YES").sum()), int((res.in_supplier == "NO").sum()),
                  n_supplier_rows, len(m), len(un), int((dup > 1).sum())]})
    return summary, res, un


def write_stock_excel(summary, res, un, path):
    with pd.ExcelWriter(path) as w:
        summary.to_excel(w, sheet_name="Summary", index=False)
        res.to_excel(w, sheet_name="TyresCart_Full_Stock", index=False)
        un.to_excel(w, sheet_name="Supplier_Not_In_TyresCart", index=False)


def main(path, out_path="Lookin_converted.xlsx"):
    out, L, stats = convert(pd.read_excel(path, "Lookin"), pd.read_excel(path, "TyresCart"))
    write_excel(out, L, out_path)
    print(stats)
    print("saved", out_path)


if __name__ == "__main__":
    main(sys.argv[1], *sys.argv[2:3])
