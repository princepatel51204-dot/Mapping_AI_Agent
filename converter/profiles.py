"""Find the DESCRIPTION column in any supplier sheet. Everything else (brand, size, load/speed,
year, pattern, OEM, RunFlat, EV) is split out of that one column by convert.py.
Optional pass-through columns (item code, stock, price, origin) are only carried along, never parsed."""
import re
import pandas as pd

DESC_NAMES = ["description", "itdesc", "item description", "item_description", "itemdescription",
              "product description", "full name", "fullname", "full_name", "desc", "product name", "item name", "product", "details", "name", "title"]
PASS = {"ItemCode": ["itemcode", "itcode", "item code", "sku", "code"],
        "Stock": ["stock", "qty", "quantity"],
        "Origin": ["origin", "originname", "country"],
        "Credit": ["credit", "price", "price_per_tyre_aed"],
        "Year": ["year"]}
SIZE = re.compile(r"\d{3}/\d{2}\s*Z?R\s*\d{2}", re.I)


def _s(x):
    return re.sub(r"\s+", " ", str(x)).strip() if pd.notna(x) else x


def _find_desc(df):
    cols = {str(c).strip().lower(): c for c in df.columns}
    for n in DESC_NAMES:
        if n in cols:
            return cols[n]
    best, score = None, 0.0  # fall back to the column that looks most like full tyre descriptions
    for c in df.columns:
        col = df[c].dropna().astype(str).head(300)
        if not len(col):
            continue
        has_size = col.str.contains(SIZE).mean()
        # a real description has a size AND other words (brand/pattern/year), not just the size itself
        extra = (col.str.replace(SIZE, " ", regex=True).str.split().str.len() >= 3).mean()
        sc = has_size * extra * min(col.str.len().mean() / 30, 1)
        if sc > score:
            best, score = c, sc
    return best if score >= 0.3 else None


def normalise(df):
    """Returns (df with a clean 'Description' column + pass-through columns, note) or (df, None)."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    c = _find_desc(df)
    if c is None:
        return df, None
    out = pd.DataFrame()
    desc = df[c].map(_s)
    out["Description"] = desc
    low = {x.lower(): x for x in df.columns}
    for std, names in PASS.items():
        for n in names:
            if n in low and low[n] != c:
                out[std] = df[low[n]].map(_s) if df[low[n]].dtype == object else df[low[n]]
                break
    # separate load-index + speed-rating columns (e.g. load_index=95, speed_rating=W) -> "95W"
    if "load_index" in low and "speed_rating" in low:
        li, sr = df[low["load_index"]], df[low["speed_rating"]]
        out["Load"] = [f"{int(float(a))}{b}" if pd.notna(a) and pd.notna(b) and str(a).strip() != "" else None
                       for a, b in zip(li, sr)]
    for extra in df.columns:  # keep EVERY other supplier column, labelled supplier_<name>
        if extra == c:
            continue
        name = "supplier_" + re.sub(r"\W+", "_", str(extra).lower()).strip("_")
        if name not in out.columns:
            out[name] = df[extra]
    return out, f"description column '{c}'"
