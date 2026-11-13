"""Brand detection from the DESCRIPTION text only (first word = supplier brand code/name)."""
import re

# description prefix -> brand (TyresCart spelling where the brand exists there)
CODE_MAP = {
    "BR": "Bridgestone", "M": "Michelin", "MIC": "Michelin", "P": "Pirelli", "D": "Dunlop",
    "ACL": "Accelera", "LS": "Landspider", "K": "Kumho", "NX": "Nexen", "Y": "Yokohama",
    "YOKO": "Yokohama", "HAB": "Habilead", "HB": "Habilead", "DE": "Deestone", "CNL": "Continental",
    "CL": "Continental", "G": "Goodyear", "GY": "Goodyear", "LAUFENN": "Laufenn", "LAUFEN": "Laufenn",
    "SUM": "Sumitomo", "HK": "Hankook", "AT": "Atturo", "WIN": "Winrun", "WI": "Windforce",
    "WINDFORCE": "Windforce", "FN": "Falken", "T": "Toyo", "BF": "Bfgoodrich", "F": "Firestone",
    "SE": "Seam", "GM": "Gripmax", "COP": "CooperTires", "AMERICAN": "American Star",
    "TBB": "TBB Tires", "MTRX": "Matrax", "MATARX": "Matrax", "SAI": "Sailun", "NAN": "Nankang",
    "NANKING": "Nankang", "MXS": "Maxxis", "PR": "Pearly", "MC": "Mastercraft", "VI": "Vitour",
    "ATL": "Atlas", "TRZ": "Trazano", "TRAZNO": "Trazano", "Q": "Quattro", "FOE": "Fortune",
    "ZET": "Zetum", "IMP": "Imperial", "TRACAMX": "Tracmax", "NEOLIN-EV": "Neolin",
    "MASIMO": "Masimon", "NIPON": "Masimon", "HONOUR": "Honour", "ANNITE": "Annaite",
    "BLACKARROW": "Black Arrow", "HEADWAY": "Head Way", "KINGBOSS": "Kingboss", "HILO": "Hilo",
    "WHEEL": None, "TUBE": None,
}


def _n(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


class BrandResolver:
    def __init__(self, tyrescart_brands):
        self.tc = {_n(b): b for b in tyrescart_brands}

    def __call__(self, description):
        words = str(description).strip().split()
        if not words:
            return None
        first = words[0].upper().lstrip(".")
        if first in CODE_MAP:
            return CODE_MAP[first]
        n1 = _n(first)
        if n1 in self.tc:
            return self.tc[n1]
        if len(words) > 1 and _n(words[0] + words[1]) in self.tc:  # e.g. "ROAD KING"
            return self.tc[_n(words[0] + words[1])]
        if len(n1) >= 4:  # abbreviation like MICH -> Michelin (only if unique)
            hits = [b for k, b in self.tc.items() if k.startswith(n1)]
            if len(hits) == 1:
                return hits[0]
        return first.title() if len(n1) >= 4 and not re.search(r"\d", n1) else None
