"""Survey: which page titles, column headers and labels the trustee reports use, over time.

Standalone (not part of WIT). Reads only the top of each page; keeps no values.

    python survey.py --deals active_deals.csv --root "E:/CDOA/Reports"
    python survey.py --deals active_deals.csv --root "E:/CDOA/Reports" --from 2024-01 --to 2026-09

active_deals.csv: one row per active deal, with columns
    Deal      the deal's folder name under --root
    Trustee   the trustee
(other column names: --deal-col / --trustee-col)

Folder layout read: <root>/<Deal>/<report type>/<YYYY_MM>/*.pdf  (a PDF named by its month
directly in the report-type folder works too).

Output: survey.xlsx
    titles    trustee x page title: in how many reports, deals, first/last month, pages per report
    headers   trustee x page title x column headers (in order): how often, first/last month
    labels    trustee x page title x label of a 'label  number' line (pages chosen by --label-pages)
    reports   one row per PDF read (deal, month, pages, problem if any)
Deal names and dates are blanked out of titles/headers/labels; digits become '#'.
The 'example' columns hold a file path so you can open an odd one: delete them before sharing.
"""
import argparse
import re
import statistics as st
import sys
import time
from collections import defaultdict
from pathlib import Path

import pandas as pd
import pymupdf

MONTH = re.compile(r"(?<!\d)(\d{4})[ _.-]?(\d{1,2})(?!\d)")
DATE = re.compile(r"\b\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}\b|\b\d{1,2}[- ][A-Za-z]{3,9}[- ,]+\d{2,4}\b|"
                  r"\b[A-Za-z]{3,9}\.? \d{1,2},? \d{4}\b|\b\d{4}-\d{2}-\d{2}\b")
NOISE = re.compile(r"^(as of|as at|next payment|payment date|determination date|page \S+ of|"
                   r"report date|distribution date)\b", re.I)
HAS_DIGIT = re.compile(r"\d")


# ---------------------------------------------------------------------------------------------
# finding the PDFs
# ---------------------------------------------------------------------------------------------
def month_of(name):
    found = None
    for m in MONTH.finditer(name):
        y, mo = int(m.group(1)), int(m.group(2))
        if 1990 < y < 2100 and 1 <= mo <= 12:
            found = (y, mo)
    return found


def reports_of(deal_dir):
    """{(year, month): pdf} for one deal: one PDF per month (the first, by report-type folder
    order then name)."""
    out = {}
    if not deal_dir.is_dir():
        return out
    for type_dir in sorted(p for p in deal_dir.iterdir() if p.is_dir()):
        for p in sorted(type_dir.iterdir()):
            if p.is_dir() and month_of(p.name):
                pdfs = sorted(f for f in p.rglob("*") if f.suffix.lower() == ".pdf")
                if pdfs:
                    out.setdefault(month_of(p.name), pdfs[0])
            elif p.suffix.lower() == ".pdf" and month_of(p.stem):
                out.setdefault(month_of(p.stem), p)
    return out


# ---------------------------------------------------------------------------------------------
# reading the top of a page
# ---------------------------------------------------------------------------------------------
def words_of(page):
    m = page.rotation_matrix
    out = []
    for x0, y0, x1, y1, t, *_ in page.get_text("words"):
        r = pymupdf.Rect(x0, y0, x1, y1) * m
        out.append(dict(x0=r.x0, y0=r.y0, x1=r.x1, y1=r.y1, t=t))
    return out


def lines_of(ws):
    if not ws:
        return []
    tol = st.median(w["y1"] - w["y0"] for w in ws) * 0.5
    rows = []
    for w in sorted(ws, key=lambda w: (w["y0"] + w["y1"]) / 2):
        c = (w["y0"] + w["y1"]) / 2
        if rows and abs(c - rows[-1]["c"]) <= tol:
            rows[-1]["ws"].append(w)
            rows[-1]["c"] = st.mean((v["y0"] + v["y1"]) / 2 for v in rows[-1]["ws"])
        else:
            rows.append({"c": c, "ws": [w]})
    for r in rows:
        r["ws"].sort(key=lambda w: w["x0"])
        r["text"] = " ".join(w["t"] for w in r["ws"])
    return rows


def phrases(ws, gap):
    out, cur = [], []
    for w in ws:
        if cur and w["x0"] - cur[-1]["x1"] > gap:
            out.append(cur)
            cur = []
        cur.append(w)
    return out + [cur] if cur else out


def clean(text, deal_words):
    """Dates out, deal name words out, digits -> '#', spaces normalised."""
    t = DATE.sub("", text)
    if deal_words:
        toks = [x for x in t.split() if x.lower().strip(".,()") not in deal_words]
        t = " ".join(toks)
    t = HAS_DIGIT.sub("#", t)
    t = re.sub(r"#+", "#", t)
    return " ".join(t.split()).strip(" :-|")


def is_deal_line(text, deal_words):
    toks = [x.lower().strip(".,()") for x in text.split()]
    toks = [x for x in toks if len(x) > 1]
    return bool(toks) and sum(x in deal_words for x in toks) / len(toks) >= 0.5


def is_data(ws):
    return any(HAS_DIGIT.search(w["t"]) and not DATE.fullmatch(w["t"]) for w in ws)


def rules_y(page):
    """y of horizontal rules: thin drawings or images at least 30% of the page wide."""
    m = page.rotation_matrix
    rects = [d["rect"] for d in page.get_drawings()] + [pymupdf.Rect(i["bbox"]) for i in page.get_image_info()]
    W = page.rect.width
    return sorted(((r * m).y0 + (r * m).y1) / 2 for r in rects
                  if (r * m).height <= 10 and (r * m).width >= 0.3 * W)


def page_top(page, title_band):
    """-> (top lines: [(text, single phrase?)], the page's lines, lh). Top lines = lines in the
    top `title_band` of the page, in order."""
    words = words_of(page)
    if not words:
        return [], [], 0
    H = page.rect.height
    lh = st.median(w["y1"] - w["y0"] for w in words)
    lines = lines_of(words)
    top = [(r["text"], len(phrases(r["ws"], lh * 1.5)) == 1, r) for r in lines if r["c"] <= H * title_band]
    return top, lines, lh


def page_survey(page, top, lines, lh, repeated, deal_words, header_band):
    """-> (title, [header columns], [labels]). `repeated` = top lines printed on most pages of the
    report (deal name, 'As of', logo text): never part of a title."""
    if not lines:
        return "(no text)", [], []
    H = page.rect.height
    # title: single-phrase top lines that aren't repeated / dates / 'As of' ...; the first
    # multi-phrase line (column headers) ends the title
    title_lines, start = [], 0
    for k, (text, single, r) in enumerate(top):
        if text in repeated or NOISE.match(DATE.sub("", text).strip()) or is_deal_line(text, deal_words):
            start = k + 1
            continue
        if not single:
            break
        t = clean(text, deal_words)
        if t and len(t) > 2:
            title_lines.append(t)
        start = k + 1
    title = " | ".join(title_lines) or "(no title)"
    # header: the lines after the title, before the first data line (max 6 lines)
    first = lines.index(top[start - 1][2]) + 1 if start else 0
    hdr = []
    rules = rules_y(page)
    for r in lines[first:]:
        if r["text"] in repeated:
            continue
        if r["c"] > H * header_band or is_data(r["ws"]) or len(hdr) >= 6:
            break
        if hdr and (r["c"] - hdr[-1]["c"] > 2.5 * lh                       # a clear gap, or
                    or any(hdr[-1]["c"] < y < r["c"] for y in rules)):      # the rule under the header
            break
        hdr.append(r)
    columns = []
    if hdr:
        ps = [(p[0]["x0"], p[-1]["x1"], r["c"], " ".join(w["t"] for w in p))
              for r in hdr for p in phrases(r["ws"], lh * 0.8)]
        ps.sort(key=lambda p: p[0])
        groups = []                                   # phrases overlapping in x = one column
        for x0, x1, c, t in ps:
            if groups and x0 < groups[-1]["x1"] - 1:
                g = groups[-1]
                g["x1"] = max(g["x1"], x1)
                g["parts"].append((c, t))
            else:
                groups.append({"x0": x0, "x1": x1, "parts": [(c, t)]})
        for g in groups:
            text = clean(" ".join(t for _, t in sorted(g["parts"])), deal_words)
            if text:
                columns.append(text)
    # labels: 'label  number' lines (first phrase text, a later phrase a number)
    labels = []
    number = re.compile(r"[(\-]?[\d,]+(\.\d+)?%?\)?")
    for r in lines[first:]:
        ps = [" ".join(w["t"] for w in p) for p in phrases(r["ws"], lh * 0.9)]
        for text, nxt in zip(ps, ps[1:]):             # every text directly followed by a number
            if not HAS_DIGIT.search(text) and len(text) >= 3 and number.fullmatch(nxt.replace(" ", "")):
                labels.append(clean(text, deal_words))
    return title, columns, labels


# ---------------------------------------------------------------------------------------------
# the survey
# ---------------------------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--deals", required=True, help="CSV of active deals")
    ap.add_argument("--root", required=True, help="reports root folder")
    ap.add_argument("--deal-col", default="Deal")
    ap.add_argument("--trustee-col", default="Trustee")
    ap.add_argument("--from", dest="first", help="first month, YYYY-MM")
    ap.add_argument("--to", dest="last", help="last month, YYYY-MM")
    ap.add_argument("--label-pages", default="summary|executive|portfolio summary",
                    help="regex: page titles whose 'label  number' lines are listed")
    ap.add_argument("--title-band", type=float, default=0.2, help="top share of the page holding the title")
    ap.add_argument("--header-band", type=float, default=0.40, help="headers must start above this share")
    ap.add_argument("--out", default="survey.xlsx")
    a = ap.parse_args(argv)

    deals = pd.read_csv(a.deals, dtype=str).fillna("")
    for col in (a.deal_col, a.trustee_col):
        if col not in deals.columns:
            sys.exit(f"column '{col}' not in {a.deals} (columns: {', '.join(deals.columns)})")
    first = month_of(a.first) if a.first else (1990, 1)
    last = month_of(a.last) if a.last else (2100, 12)
    label_rx = re.compile(a.label_pages, re.I)

    titles = defaultdict(lambda: dict(reports=set(), deals=set(), months=set(), pages=[], example=""))
    headers = defaultdict(lambda: dict(reports=set(), months=set(), example=""))
    labels = defaultdict(lambda: dict(reports=set(), months=set()))
    reports = []
    jobs = []
    for _, d in deals.iterrows():
        deal, trustee = d[a.deal_col].strip(), d[a.trustee_col].strip()
        for ym, pdf in sorted(reports_of(Path(a.root) / deal).items()):
            if first <= ym <= last:
                jobs.append((deal, trustee, ym, pdf))
    print(f"{len(deals)} deals, {len(jobs)} reports to read")
    t0 = time.time()
    for n, (deal, trustee, ym, pdf) in enumerate(jobs, 1):
        month = f"{ym[0]}-{ym[1]:02d}"
        deal_words = {x.lower().strip(".,()") for x in re.split(r"[\s_]+", deal) if len(x) > 1}
        key = (deal, month)
        try:
            doc = pymupdf.open(pdf)
        except Exception as e:                       # noqa: BLE001
            reports.append(dict(trustee=trustee, deal=deal, month=month, pages=0, problem=f"cannot open: {e}",
                                file=str(pdf)))
            continue
        per_title = defaultdict(int)
        with doc:
            tops = [page_top(page, a.title_band) for page in doc]
            # top lines printed on more than half of the pages: deal name, 'As of', logo text
            seen = defaultdict(int)
            for top, _, _ in tops:
                for text in {t for t, _, _ in top}:
                    seen[text] += 1
            repeated = {t for t, c in seen.items() if c > 0.5 * len(doc) and len(doc) > 3}
            for page, (top, lines, lh) in zip(doc, tops):
                try:
                    title, cols, labs = page_survey(page, top, lines, lh, repeated, deal_words, a.header_band)
                except Exception as e:               # noqa: BLE001
                    title, cols, labs = f"(error: {type(e).__name__})", [], []
                per_title[title] += 1
                t = titles[(trustee, title)]
                t["reports"].add(key)
                t["deals"].add(deal)
                t["months"].add(month)
                t["example"] = t["example"] or f"{pdf} p{page.number + 1}"
                h = headers[(trustee, title, " | ".join(cols) or "(no header)")]
                h["reports"].add(key)
                h["months"].add(month)
                h["example"] = h["example"] or f"{pdf} p{page.number + 1}"
                if label_rx.search(title):
                    for lab in dict.fromkeys(labs):
                        lb = labels[(trustee, title, lab)]
                        lb["reports"].add(key)
                        lb["months"].add(month)
            for title, cnt in per_title.items():
                titles[(trustee, title)]["pages"].append(cnt)
            reports.append(dict(trustee=trustee, deal=deal, month=month, pages=len(doc), problem="",
                                file=str(pdf)))
        if n % 25 == 0 or n == len(jobs):
            print(f"  {n}/{len(jobs)} reports, {time.time() - t0:.0f}s")

    n_rep = defaultdict(set)
    for r in reports:
        n_rep[r["trustee"]].add((r["deal"], r["month"]))

    def span(ms):
        return (min(ms), max(ms)) if ms else ("", "")

    t_rows = [dict(trustee=tr, title=ti, reports=len(v["reports"]), of_reports=len(n_rep[tr]),
                   deals=len(v["deals"]), first=span(v["months"])[0], last=span(v["months"])[1],
                   pages_per_report=f"{min(v['pages'])}-{max(v['pages'])}" if v["pages"] else "",
                   example=v["example"]) for (tr, ti), v in titles.items()]
    h_rows = [dict(trustee=tr, title=ti, headers=hd, reports=len(v["reports"]),
                   first=span(v["months"])[0], last=span(v["months"])[1], example=v["example"])
              for (tr, ti, hd), v in headers.items()]
    l_rows = [dict(trustee=tr, title=ti, label=lb, reports=len(v["reports"]),
                   first=span(v["months"])[0], last=span(v["months"])[1])
              for (tr, ti, lb), v in labels.items()]
    with pd.ExcelWriter(a.out) as xw:
        pd.DataFrame(t_rows).sort_values(["trustee", "reports"], ascending=[True, False]) \
            .to_excel(xw, sheet_name="titles", index=False)
        pd.DataFrame(h_rows).sort_values(["trustee", "title", "reports"], ascending=[True, True, False]) \
            .to_excel(xw, sheet_name="headers", index=False)
        pd.DataFrame(l_rows, columns=["trustee", "title", "label", "reports", "first", "last"]) \
            .sort_values(["trustee", "title", "reports"], ascending=[True, True, False]) \
            .to_excel(xw, sheet_name="labels", index=False)
        pd.DataFrame(reports).to_excel(xw, sheet_name="reports", index=False)
    print(f"done: {a.out}  ({len(reports)} reports, {time.time() - t0:.0f}s)")


if __name__ == "__main__":
    main()
