"""Small web UI: upload a supplier sheet -> see TyresCart-format table -> download xlsx.

Run:  .venv/bin/python webapp/app.py   then open http://127.0.0.1:5000
"""
import io
import sys
import uuid
from pathlib import Path

import pandas as pd
from flask import Flask, render_template_string, request, send_file, abort, redirect

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "converter"))
from convert import convert, write_excel, build_stock_view, write_stock_excel  # noqa: E402
from profiles import normalise  # noqa: E402

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
CATALOGUE = pd.read_pickle(ROOT / "data" / "tyrescart.pkl")
_hay = None
import json
VOCAB = json.load(open(ROOT / "data" / "pitstop_vocab.json"))  # pattern spellings from Pitstop Arabia
RESULTS = {}  # id -> (out_df, working_df, stats, filename)
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
REQUIRED = ["Description"]
OPTIONAL = ["ItemCode", "Stock", "Credit"]
SHOW_HEAD = ["match_status", "sku"]
SHOW_TAIL = ["price_difference", "mgs_brand", "pattern", "oem_tyres", "runflat", "ev_tyre", "tyre_size", "load_index",
             "year", "country", "review_note"]


def show_cols(out):
    """match status + SKU, then ALL supplier_* columns, then ALL tyrescart_* columns, then the parsed fields."""
    cols = list(out.columns)
    sup = [c for c in cols if c.startswith("supplier_")]
    tcc = [c for c in cols if c.startswith("tyrescart_")]
    return SHOW_HEAD + sup + tcc + SHOW_TAIL

PAGE = """<!doctype html><html lang=en><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1"><title>Tyre Mapper</title>
<style>
:root{color-scheme:light;--navy:#0b2239;--navy2:#12385c;--ac:#f59e0b;--ac2:#d97706;--bg:#f4f6f9;--card:#fff;--tx:#17212b;--mut:#64748b;--bd:#e2e8f0;
--ok:#15803d;--okbg:#dcfce7;--warn:#b45309;--warnbg:#fef3c7;--no:#b91c1c;--nobg:#fee2e2;--grid:#e5e7eb;--hd:#f1f5f9}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--tx);font:15px/1.55 "Segoe UI",Inter,system-ui,-apple-system,sans-serif}
.wrap{max-width:1380px;margin:0 auto;padding:0 20px}
/* top bar */
.nav{background:var(--navy);color:#fff}.nav .wrap{display:flex;align-items:center;gap:12px;height:60px}
.logo{display:flex;align-items:center;gap:10px;font-weight:700;font-size:18px;letter-spacing:.2px}.nav small{color:#9db4cc;margin-left:auto}
/* hero */
.hero{background:linear-gradient(135deg,var(--navy) 0%,var(--navy2) 60%,#1b5a8c 100%);color:#fff;overflow:hidden;position:relative}
.hero .wrap{display:grid;grid-template-columns:1.25fr .75fr;gap:24px;align-items:center;padding-top:44px;padding-bottom:52px}
.hero h1{font-size:40px;line-height:1.12;margin:0 0 14px;font-weight:800}.hero h1 span{color:var(--ac)}
.hero p{color:#cfe0f1;font-size:17px;margin:0 0 22px;max-width:620px}
.chips{display:flex;gap:10px;flex-wrap:wrap}.chip{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.18);padding:6px 12px;border-radius:99px;font-size:13px}
.art{display:flex;justify-content:center}.art svg{width:min(340px,80%);height:auto;filter:drop-shadow(0 20px 30px rgba(0,0,0,.35));animation:spin 40s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
@media(max-width:860px){.hero .wrap{grid-template-columns:1fr}.art{display:none}.hero h1{font-size:30px}}
.hero.slim .wrap{grid-template-columns:1fr;padding-top:22px;padding-bottom:50px}.hero.slim h1{font-size:26px;margin:0}.hero.slim p,.hero.slim .chips,.hero.slim .art{display:none}
/* upload card (overlaps hero) */
.up{margin-top:-34px;position:relative}
.card{background:var(--card);border:1px solid var(--bd);border-radius:14px;box-shadow:0 8px 24px rgba(15,34,57,.07);padding:20px;margin-bottom:18px}
.drop{display:flex;align-items:center;gap:18px;flex-wrap:wrap}
.drop .ico{width:56px;height:56px;border-radius:14px;background:#fff7e6;display:grid;place-items:center;flex:none}
.drop .txt{flex:1;min-width:220px}.drop b{font-size:17px}.mut{color:var(--mut)}
input[type=file]{font:inherit;padding:9px;border:1.5px dashed #cbd5e1;border-radius:10px;background:#f8fafc;max-width:100%}
button,.btn{background:var(--ac);color:#1b1305;border:0;border-radius:10px;padding:10px 18px;font:inherit;font-weight:700;cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;gap:8px;transition:.15s}
button:hover,.btn:hover{background:var(--ac2);color:#fff;transform:translateY(-1px)}
.btn.dk{background:var(--navy);color:#fff}.btn.dk:hover{background:var(--navy2)}.btn.gh{background:#fff;color:var(--navy);border:1px solid var(--bd)}.btn.gh:hover{background:var(--hd);color:var(--navy)}
.btn.rd{background:#fff;color:var(--no);border:1px solid #fecaca}.btn.rd:hover{background:var(--nobg);color:var(--no)}
.err{background:var(--nobg);color:var(--no);border:1px solid #fecaca;padding:10px 14px;border-radius:10px;margin-top:14px}
/* steps */
.steps{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:26px 0}
.step{background:var(--card);border:1px solid var(--bd);border-radius:14px;padding:16px}
.step i{display:grid;place-items:center;width:38px;height:38px;border-radius:10px;background:var(--navy);color:#fff;font-style:normal;font-weight:700;margin-bottom:10px}
.step b{display:block}.step span{color:var(--mut);font-size:13.5px}
@media(max-width:860px){.steps{grid-template-columns:1fr 1fr}}
/* result */
.file{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:14px}.file .name{font-weight:700;font-size:17px}
.pill{background:var(--hd);border:1px solid var(--bd);border-radius:99px;padding:3px 11px;font-size:12.5px;color:var(--mut)}
.acts{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px}
.stats{display:grid;grid-template-columns:repeat(6,1fr);gap:12px}
@media(max-width:1000px){.stats{grid-template-columns:repeat(3,1fr)}}@media(max-width:560px){.stats{grid-template-columns:1fr 1fr}}
.stat{border:1px solid var(--bd);border-radius:12px;padding:12px 14px;background:#fff;border-top:4px solid var(--navy)}
.stat span{color:var(--mut);font-size:13px}.stat b{display:block;font-size:28px;line-height:1.2}
.stat.g{border-top-color:#16a34a}.stat.g b{color:var(--ok)}.stat.o{border-top-color:#f59e0b}.stat.o b{color:var(--warn)}.stat.r{border-top-color:#dc2626}.stat.r b{color:var(--no)}
.bar2{height:10px;border-radius:99px;background:var(--hd);overflow:hidden;display:flex;margin-top:14px}
.bar2 i{display:block;height:100%}
/* tabs */
.tabs{display:inline-flex;gap:4px;background:var(--hd);border:1px solid var(--bd);padding:4px;border-radius:12px;margin-bottom:14px;flex-wrap:wrap}
.tabs button{background:transparent;color:var(--mut);padding:8px 14px;border-radius:9px;font-weight:600}
.tabs button:hover{background:#e2e8f0;color:var(--navy);transform:none}.tabs button.on{background:#fff;color:var(--navy);box-shadow:0 1px 3px rgba(0,0,0,.12)}
.bar{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:12px}
select,input[type=search]{padding:9px 12px;border:1px solid var(--bd);border-radius:10px;background:#fff;color:var(--tx);font:inherit}
input[type=search]{min-width:260px}
.tw{overflow:auto;max-height:66vh;border:1px solid var(--bd);border-radius:12px;background:#fff}
table{border-collapse:separate;border-spacing:0;font-size:13px;width:100%}
th,td{padding:7px 12px;border-bottom:1px solid var(--grid);border-right:1px solid #f1f3f5;white-space:nowrap;text-align:left}
th{position:sticky;top:0;background:var(--hd);color:#334155;font-weight:700;z-index:2;font-size:12px;text-transform:uppercase;letter-spacing:.3px}
td.rn,th.rn{position:sticky;left:0;background:var(--hd);color:var(--mut);text-align:center;min-width:46px;z-index:1}th.rn{z-index:3}
tbody tr:nth-child(even) td:not(.rn){background:#fafbfc}tbody tr:hover td:not(.rn){background:#fff7e6}
.badge{display:inline-block;padding:2px 10px;border-radius:99px;font-size:12px;font-weight:700}
.b1{background:var(--okbg);color:var(--ok)}.b2{background:var(--warnbg);color:var(--warn)}.b0{background:var(--nobg);color:var(--no)}
.dlg{display:grid;grid-template-columns:repeat(2,1fr);gap:16px}@media(max-width:860px){.dlg{grid-template-columns:1fr}}
.dl{border:1px solid var(--bd);border-radius:14px;padding:18px;background:#fff;display:flex;flex-direction:column;gap:8px}
.dl h3{margin:0;font-size:17px}.dl p{margin:0 0 8px;color:var(--mut);flex:1}
footer{color:var(--mut);text-align:center;padding:26px 0 40px;font-size:13px}
</style></head><body>
<div class=nav><div class=wrap><div class=logo>
<svg width=30 height=30 viewBox="0 0 100 100"><circle cx=50 cy=50 r=46 fill="#f59e0b"/><circle cx=50 cy=50 r=32 fill="#0b2239"/><circle cx=50 cy=50 r=14 fill="#cbd5e1"/><circle cx=50 cy=50 r=5 fill="#0b2239"/></svg>
Tyre Mapper</div><small>Supplier sheet &rarr; TyresCart format</small></div></div>

<section class="hero{% if stats %} slim{% endif %}"><div class=wrap><div>
<h1>Turn any supplier price list into <span>TyresCart-ready data</span></h1>
<p>Upload a supplier sheet. We read the <b>description</b> column, split it into brand, size, load/speed, year, pattern, OEM marking, run-flat and EV, then match every tyre to your TyresCart SKU.</p>
<div class=chips><span class=chip>&#10003; Reads only the Description column</span><span class=chip>&#10003; Works with any supplier layout</span><span class=chip>&#10003; Exact-match SKU mapping</span><span class=chip>&#10003; Excel download</span></div></div>
<div class=art><svg viewBox="0 0 300 300" aria-hidden=true>
<defs><radialGradient id=rg cx=50% cy=50% r=50%><stop offset=0 stop-color="#e2e8f0"/><stop offset=1 stop-color="#94a3b8"/></radialGradient></defs>
<circle cx=150 cy=150 r=140 fill="#0a0f14"/>
<g stroke="#1f2933" stroke-width=10 stroke-dasharray="14 9" fill=none><circle cx=150 cy=150 r=132 /></g>
<circle cx=150 cy=150 r=112 fill="#161d25"/><circle cx=150 cy=150 r=96 fill="none" stroke="#f59e0b" stroke-width=4 />
<circle cx=150 cy=150 r=84 fill="url(#rg)"/>
<g fill="#475569"><circle cx=150 cy=150 r=26 /></g>
<g stroke="#64748b" stroke-width=14 stroke-linecap=round><path d="M150 150 L150 82"/><path d="M150 150 L208 116"/><path d="M150 150 L208 184"/><path d="M150 150 L150 218"/><path d="M150 150 L92 184"/><path d="M150 150 L92 116"/></g>
<circle cx=150 cy=150 r=18 fill="#f59e0b"/><circle cx=150 cy=150 r=7 fill="#0b2239"/></svg></div></div></section>

<div class=wrap>
<div class="card up"><form method=post action="/" enctype=multipart/form-data class=drop>
<div class=ico><svg width=30 height=30 viewBox="0 0 24 24" fill=none stroke="#d97706" stroke-width=2 stroke-linecap=round stroke-linejoin=round><path d="M12 16V4M7 9l5-5 5 5"/><path d="M4 16v3a2 2 0 002 2h12a2 2 0 002-2v-3"/></svg></div>
<div class=txt><b>Upload supplier sheet</b><div class=mut>.xlsx, .xls or .csv. The sheet needs a tyre description column such as <i>Description</i> or <i>ITDESC</i>. Stock, price and origin columns are carried along if present.</div></div>
<input type=file name=file accept=".xlsx,.xls,.csv" required><button><svg width=16 height=16 viewBox="0 0 24 24" fill=none stroke=currentColor stroke-width=3><path d="M5 12h14M13 6l6 6-6 6"/></svg> Upload &amp; convert</button></form>
{% if error %}<div class=err>{{error}}</div>{% endif %}</div>

{% if not stats %}
<div class=steps>
<div class=step><i>1</i><b>Upload</b><span>Drop in any supplier price list or stock sheet.</span></div>
<div class=step><i>2</i><b>Split the description</b><span>Brand, size, load/speed, year, pattern, OEM, run-flat, EV.</span></div>
<div class=step><i>3</i><b>Match to TyresCart</b><span>Size 1 key first, then Size 2, confirmed by the full display name.</span></div>
<div class=step><i>4</i><b>Download</b><span>Get the TyresCart-format sheet, plus a review list.</span></div></div>
{% endif %}

{% if stats %}
<div class=card>
<div class=file><span class=name>{{filename}}</span><span class=pill>{{stats.sheet}}</span><span class=pill>{{stats.rows}} rows</span><span class=pill>Pitstop Arabia vocabulary on</span></div>
<div class=acts><a class=btn href="/download/{{rid}}/tyrescart">&#8681; Download TyresCart sheet</a>
<a class="btn dk" href="/download/{{rid}}">&#8681; Full result (3 sheets)</a>
<a class="btn gh" href="/download/{{rid}}/original">&#8681; Uploaded sheet</a>
<form method=post action="/remove/{{rid}}" style="margin:0" onsubmit="return confirm('Remove this uploaded file and its results?')"><button class="btn rd" style="background:#fff">&#10005; Remove file</button></form></div>
<div class=stats>
<div class=stat><span>Rows</span><b>{{stats.rows}}</b></div>
<div class="stat g"><span>Matched Size 1</span><b>{{stats.size1}}</b></div>
<div class="stat o"><span>Matched Size 2</span><b>{{stats.size2}}</b></div>
<div class="stat r"><span>Unmatched</span><b>{{stats.unmatched}}</b></div>
<div class=stat><span>Pattern unsure</span><b>{{stats.low_conf}}</b></div>
<div class=stat><span>Unknown brand</span><b>{{stats.unknown_brand}}</b></div></div>
<div class=bar2 title="matched vs unmatched"><i style="width:{{pct1}}%;background:#16a34a"></i><i style="width:{{pct2}}%;background:#f59e0b"></i><i style="width:{{pct0}}%;background:#fca5a5"></i></div>
<div class=mut style="font-size:13px;margin-top:6px">{{pct1+pct2}}% of rows matched to a TyresCart SKU</div></div>

<div class=card>
<div class=tabs><button id=tb1 type=button class=on>TyresCart format (converted)</button><button id=tb2 type=button>Uploaded sheet (original)</button><button id=tb3 type=button>Reference (TyresCart catalogue)</button><button id=tb4 type=button>Download</button></div>
<div id=v1><div class=bar>
<select id=st><option value="">All rows</option><option>MATCHED (Size1)</option><option>MATCHED (Size2)</option><option>NO MATCH</option></select>
<input type=search id=q placeholder="Search any text..."><span class=mut id=cnt></span>{% if stats.shown < stats.rows %}<span class=mut>(first {{stats.shown}} of {{stats.rows}} rows shown; downloads have all rows)</span>{% endif %}</div>
<div class=tw><table id=t><thead><tr>{% for c in cols %}<th>{{c}}</th>{% endfor %}</tr></thead><tbody>
{% for r in rows %}<tr data-s="{{r[0]}}">{% for v in r %}<td{% if loop.index==1 %}><span class="badge {{ {'MATCHED (Size1)':'b1','MATCHED (Size2)':'b2','NO MATCH':'b0'}[v] }}">{{v}}</span>{% else %}>{{v}}{% endif %}</td>{% endfor %}</tr>{% endfor %}
</tbody></table></div></div>
<div id=v3 style="display:none"><div class=bar><input type=search id=rq placeholder="Search catalogue (brand, pattern, size, SKU...)"><span class=mut id=rcnt></span></div>
<div class=tw><table id=rt><thead></thead><tbody></tbody></table></div></div>
<div id=v4 style="display:none"><div class=dlg>
<div class=dl><h3>TyresCart sheet</h3><p>The updated file in TyresCart format: the 24 TyresCart columns with the SKU filled where matched. {{stats.rows}} rows.</p><a class=btn href="/download/{{rid}}/tyrescart">&#8681; Download .xlsx</a></div>
<div class=dl><h3>Full result (3 sheets)</h3><p>TyresCart_Format with match status, Lookin_Working with every key and count, and Review with the rows to check.</p><a class="btn dk" href="/download/{{rid}}">&#8681; Download .xlsx</a></div>
<div class=dl><h3>Full TyresCart stock view</h3><p>Every TyresCart tyre compared with this supplier: in stock or not, supplier item, stock and price, plus supplier tyres missing from TyresCart.</p><a class="btn dk" href="/download/{{rid}}/stock">&#8681; Download .xlsx</a></div>
<div class=dl><h3>Uploaded sheet</h3><p>Your supplier sheet exactly as it was read from {{filename}}.</p><a class="btn gh" href="/download/{{rid}}/original">&#8681; Download .xlsx</a></div></div></div>
<div id=v2 style="display:none"><div class=tw><table><thead><tr>{% for c in ocols %}<th>{{c}}</th>{% endfor %}</tr></thead><tbody>
{% for r in orows %}<tr>{% for v in r %}<td>{{v}}</td>{% endfor %}</tr>{% endfor %}</tbody></table></div></div></div>
<script>
const V=['v1','v2','v3','v4'].map(i=>document.getElementById(i)),B=['tb1','tb2','tb3','tb4'].map(i=>document.getElementById(i));
function show(i){V.forEach((v,j)=>{v.style.display=i===j?'':'none';B[j].classList.toggle('on',i===j)});if(i===2&&!rt.tHead.rows.length)loadRef()}
B.forEach((b,i)=>b.onclick=()=>show(i));
const rt=document.getElementById('rt'),rq=document.getElementById('rq'),rcnt=document.getElementById('rcnt');let tm;
async function loadRef(){const d=await (await fetch('/reference?q='+encodeURIComponent(rq.value))).json();
 const esc=x=>x.replace(/&/g,'&amp;').replace(/</g,'&lt;');
 rt.tHead.innerHTML='<tr><th class=rn></th>'+d.cols.map(c=>'<th>'+c+'</th>').join('')+'</tr>';
 rt.tBodies[0].innerHTML=d.rows.map((r,i)=>'<tr><td class=rn>'+(i+1)+'</td>'+r.map(v=>'<td>'+esc(v)+'</td>').join('')+'</tr>').join('');
 rcnt.textContent='showing '+d.rows.length+' of '+d.total+' catalogue rows'}
rq.oninput=()=>{clearTimeout(tm);tm=setTimeout(loadRef,250)};
const st=document.getElementById('st'),q=document.getElementById('q'),rows=[...document.querySelectorAll('#t tbody tr')],cnt=document.getElementById('cnt');
function f(){const s=st.value,t=q.value.toLowerCase();let n=0;rows.forEach(r=>{const ok=(!s||r.dataset.s===s)&&(!t||r.textContent.toLowerCase().includes(t));r.style.display=ok?'':'none';if(ok)n++});cnt.textContent=n+' rows shown'}
document.querySelectorAll('#t, #v2 table').forEach(t=>{
 const hr=t.tHead.rows[0];hr.insertBefore(Object.assign(document.createElement('th'),{className:'rn',textContent:'#'}),hr.firstChild);
 [...t.tBodies[0].rows].forEach((r,i)=>r.insertBefore(Object.assign(document.createElement('td'),{className:'rn',textContent:i+1}),r.firstChild));
});
st.onchange=f;q.oninput=f;f();
</script>
{% endif %}
<footer>Tyre Mapper &middot; reads the description column only &middot; matches by Size 1 / Size 2 key + full display name</footer>
</div></body></html>"""


def _fix_cols(df):
    lc = {c.lower(): c for c in REQUIRED + OPTIONAL}
    df.columns = [lc.get(str(c).strip().lower(), str(c).strip()) for c in df.columns]
    return df


def load_supplier(file):
    """Return (DataFrame, sheet_name). Finds the sheet/header row containing the required columns."""
    name = file.filename.lower()
    if name.endswith(".csv"):
        raw = file.read()
        text = None
        for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):  # Excel CSVs are often not UTF-8
            try:
                text = raw.decode(enc)
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        for sep in (None, ",", ";", "\t", "|"):  # None = auto-detect the separator
            for hdr in range(0, 8):  # header may not be on row 1
                try:
                    df = pd.read_csv(io.StringIO(text), sep=sep, engine="python", header=hdr,
                                     skipinitialspace=True, on_bad_lines="skip")
                except Exception:
                    break
                df, prof = normalise(df)
                if all(c in df.columns for c in REQUIRED):
                    return df, "csv" + (f" - {prof}" if prof else "")
        raise ValueError("could not find a tyre description column in the CSV (e.g. 'Description', 'ITDESC' or 'Full Name')")
    xl = pd.ExcelFile(file)
    order = sorted(xl.sheet_names, key=lambda n: n.lower() != "lookin")  # prefer "Lookin"
    for sh in order:
        for hdr in range(0, 8):  # header may not be on row 1
            try:
                df = xl.parse(sh, header=hdr)
            except Exception:
                break
            df, prof = normalise(df)  # finds the description column in any layout
            if all(c in df.columns for c in REQUIRED):
                return df, sh + (f" - {prof}" if prof else "")
    raise ValueError("could not find a tyre description column (e.g. 'Description' or 'ITDESC'). Sheets found: "
                     + ", ".join(xl.sheet_names))


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template_string(PAGE)
    file = request.files.get("file")
    try:
        L, sheet = load_supplier(file)
        for c in OPTIONAL:
            if c not in L.columns:
                L[c] = None
        L = L.dropna(subset=["Description"]).reset_index(drop=True)
        out, work, stats = convert(L, CATALOGUE, VOCAB)
    except Exception as e:  # show a readable error instead of a 500 page
        return render_template_string(PAGE, error=f"Could not process file: {e}")
    rid = uuid.uuid4().hex[:12]
    RESULTS[rid] = (out, work, stats, file.filename, L)
    stats["sheet"] = sheet
    CAP = 3000  # keep the browser fast; downloads always contain every row
    fmt = lambda d: d.head(CAP).astype(object).where(d.head(CAP).notna(), "").astype(str).values.tolist()
    stats["shown"] = min(CAP, len(out))
    n = max(stats["rows"], 1)
    pct1, pct2 = round(100 * stats["size1"] / n, 1), round(100 * stats["size2"] / n, 1)
    return render_template_string(
        PAGE, stats=stats, rid=rid, filename=file.filename, pct1=pct1, pct2=pct2, pct0=round(100 - pct1 - pct2, 1),
        cols=show_cols(out), rows=fmt(out[show_cols(out)]),
        ocols=list(L.columns), orows=fmt(L))


@app.route("/download/<rid>")
def download(rid):
    if rid not in RESULTS:
        abort(404)
    out, work, _, fn, _L = RESULTS[rid]
    buf = io.BytesIO()
    write_excel(out, work, buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"{Path(fn).stem}_tyrescart.xlsx",
                     mimetype=XLSX)


@app.route("/download/<rid>/stock")
def download_stock(rid):
    """Full TyresCart list compared with the uploaded supplier sheet."""
    if rid not in RESULTS:
        abort(404)
    out, _w, _s, fn, L = RESULTS[rid]
    summary, res, un = build_stock_view(out, CATALOGUE, len(L))
    buf = io.BytesIO()
    write_stock_excel(summary, res, un, buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"TyresCart_vs_{Path(fn).stem}.xlsx", mimetype=XLSX)


@app.route("/download/<rid>/<kind>")
def download_kind(rid, kind):
    """kind = tyrescart (clean TyresCart-format sheet only) | original (uploaded sheet)"""
    if rid not in RESULTS:
        abort(404)
    out, _w, _s, fn, L = RESULTS[rid]
    df = out.drop(columns=[c for c in out.columns if c.startswith(("supplier_", "tyrescart_")) or c in
                           ("match_status", "review_note", "price_difference")]) if kind == "tyrescart" else L
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"{Path(fn).stem}_{kind}.xlsx", mimetype=XLSX)


@app.route("/remove/<rid>", methods=["POST"])
def remove(rid):
    RESULTS.pop(rid, None)  # forget the old uploaded file and its results
    return redirect("/")


@app.route("/reference")
def reference():
    """TyresCart catalogue rows for the Reference tab (searchable, capped for speed)."""
    q = request.args.get("q", "").strip().lower()
    cols = ["sku", "mgs_brand", "pattern", "oem_tyres", "runflat", "ev_tyre", "tyre_size", "load_index",
            "year", "country", "tyres_category", "price", "name"]
    df = CATALOGUE[cols]
    if q:
        global _hay
        if _hay is None:
            _hay = CATALOGUE[cols].fillna("").astype(str).apply(" ".join, axis=1).str.lower()
        df = df[_hay.loc[df.index].str.contains(q, regex=False)]
    total = len(df)
    return {"cols": cols, "rows": df.head(300).fillna("").astype(str).values.tolist(), "total": total}


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
