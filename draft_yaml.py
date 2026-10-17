"""Draft one trustee's YAML from survey.xlsx.

    python draft_yaml.py --survey survey.xlsx --trustee "US Bank"
    -> us_bank.draft.yaml   (copy to us_bank.yaml and edit; re-running never touches your copy)

What the draft holds
    balance / performance / collateral   sources guessed from GUESS below (edit GUESS and re-run,
                                         or move blocks by hand)
    unused                               every other title, same block format, most frequent first
Each source block:
    titles     the page title(s); spellings that differ only by case, punctuation or a plural 's'
               are grouped into one block
    shape      values | tests | grid | grouped   (a guess; see SHAPES)
    columns    standard_name: [header as printed, ...]  - every header seen in any layout.
               Rename the standard names, put spellings of the same thing under one name,
               delete what you don't need.
    labels     the same for 'label  number' lines (summary pages)
    # layout   comments: each distinct header line, how many reports, first and last month
Nothing here is data: only titles, headers and labels as the survey wrote them.
"""
import argparse
import json
import re
from collections import OrderedDict

import pandas as pd

# Which titles go where in the draft (regex, case-insensitive). Report wording only.
GUESS = {
    "balance": [r"executive summary"],
    "performance": [r"executive summary", r"concentration limitation", r"portfolio profile test",
                    r"coverage test", r"quality test", r"par value test", r"interest coverage"],
    "collateral": [r"detail of assets", r"rating detail", r"minimum diversity", r"rating change"],
}
SHAPES = "values = label -> number | tests = test rows under a header | grid = table | " \
         "grouped = table whose first column is printed once and copied down"
TEST_WORDS = r"\b(test|threshold|trigger|required|result|current|actual|min|max)\b"
EMPTY = "no asset records"


def q(s):
    """A YAML-safe quoted string."""
    return json.dumps(str(s), ensure_ascii=False)


def group_key(title):
    words = re.sub(r"[^a-z0-9 ]", " ", title.lower()).split()
    return " ".join(w[:-1] if len(w) > 3 and w.endswith("s") else w for w in words)


def std_name(header, taken):
    s = header.lower().replace("'", "").replace("&", " and ").replace("%", " pct ")
    s = re.sub(r"[^a-z0-9]+", "_", s).strip("_") or "col"
    if s[0].isdigit():
        s = "c_" + s
    name, i = s, 2
    while name in taken:
        name, i = f"{s}_{i}", i + 1
    taken.add(name)
    return name


def is_text(h):
    """A header line that is really a sentence (notice, disclaimer, 'no asset records')."""
    return h in ("(no header)", "", "nan") or EMPTY in h.lower() or \
        any(len(c.split()) > 8 for c in h.split(" | "))


def block(titles, t_rows, h_rows, l_rows, shape_hint=None, indent="    "):
    i, j = indent, indent + "  "
    reports = int(t_rows["reports"].max())
    of = int(t_rows["of_reports"].max())
    first, last = t_rows["first"].min(), t_rows["last"].max()
    pages = "/".join(sorted(set(t_rows["pages_per_report"].astype(str))))
    out = [f"{i}- titles: [{', '.join(q(t) for t in titles)}]"
           f"   # in {reports} of {of} reports, {first}..{last}, pages per report {pages}"]

    layouts = h_rows.sort_values("reports", ascending=False)
    heads = [h for h in layouts["headers"].astype(str) if not is_text(h)]
    empty = any(EMPTY in h.lower() for h in layouts["headers"].astype(str))
    has_labels = len(l_rows) > 0

    if shape_hint:
        shape = shape_hint
    elif heads and re.search(TEST_WORDS, heads[0], re.I):
        shape = "tests"
    elif has_labels and not heads:
        shape = "values"
    else:
        shape = "grid"
    out.append(f"{j}shape: {shape}")
    if empty:
        out.append(f"{j}empty_ok: true            # some reports print 'No asset records ...'")

    if shape != "values":
        cols, taken = OrderedDict(), set()
        for h in heads:
            for c in h.split(" | "):
                c = c.strip()
                if c and c not in cols:
                    cols[c] = std_name(c, taken)
        if cols:
            out.append(f"{j}columns:")
            for printed, name in cols.items():
                out.append(f"{j}  {name}: [{q(printed)}]")
    if has_labels and shape in ("values", "tests") or (has_labels and not heads):
        out.append(f"{j}labels:")
        taken = set()
        for _, r in l_rows.sort_values("reports", ascending=False).iterrows():
            if str(r["label"]) in ("nan", ""):
                continue
            out.append(f"{j}  {std_name(str(r['label']), taken)}: [{q(r['label'])}]"
                       f"   # {r['reports']} reports, {r['first']}..{r['last']}")
    for n, (_, r) in enumerate(layouts.iterrows()):
        h = str(r["headers"])
        out.append(f"{j}# layout {chr(65 + n) if n < 26 else n}: {r['reports']} reports, "
                   f"{r['first']}..{r['last']}: {h[:200]}")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--survey", default="survey.xlsx")
    ap.add_argument("--trustee", required=True, help="trustee as written in the CSV (part of it is enough)")
    ap.add_argument("--out", help="default: <trustee>.draft.yaml")
    a = ap.parse_args()

    x = pd.read_excel(a.survey, sheet_name=None)
    pick = lambda df: df[df["trustee"].astype(str).str.contains(a.trustee, case=False, regex=False)]
    t, h, l = pick(x["titles"]), pick(x["headers"]), pick(x["labels"])
    if t.empty:
        raise SystemExit(f"no trustee matching {a.trustee!r}; in the survey: "
                         + ", ".join(sorted(x["titles"]["trustee"].astype(str).unique())))
    names = sorted(t["trustee"].astype(str).unique())
    t = t[t["title"].astype(str) != "(no title)"]

    # group near-identical titles, most frequent first
    groups = OrderedDict()
    for _, r in t.sort_values("reports", ascending=False).iterrows():
        groups.setdefault(group_key(str(r["title"])), []).append(str(r["title"]))

    def parts(titles):
        return (t[t["title"].isin(titles)], h[h["title"].isin(titles)], l[l["title"].isin(titles)])

    placed = {k: [] for k in GUESS}
    used = set()
    for key, titles in groups.items():
        for out_name, pats in GUESS.items():
            if any(re.search(p, titles[0], re.I) for p in pats):
                placed[out_name].append(key)
                used.add(key)

    slug = re.sub(r"[^a-z0-9]+", "_", a.trustee.lower()).strip("_")
    path = a.out or f"{slug}.draft.yaml"
    lines = [f"# Draft from {a.survey} for trustee {', '.join(names)}.",
             "# Copy to <trustee>.yaml before editing; re-running overwrites this draft only.",
             f"# shapes: {SHAPES}",
             "",
             f"trustee: {q(a.trustee)}",
             f"names: [{', '.join(q(n) for n in names)}]   # as the deal list spells it",
             "title_zone: {top: 0.20}                     # page titles are looked for in the top 20%",
             ""]
    for out_name in ("balance", "performance", "collateral"):
        lines.append(f"{out_name}:")
        if out_name == "collateral":
            lines.append("  key: [issuer_name, facility_name]   # one row per issuer + issue; "
                         "use the standard names below")
        if out_name == "performance":
            lines.append("  # test words (Test / Threshold / Current / Result ...) are common to "
                         "all trustees, not listed here")
        if not placed[out_name]:
            lines.append("  sources: []")
        else:
            lines.append("  sources:")
            for key in placed[out_name]:
                if out_name == "collateral":
                    first = str(parts(groups[key])[1]["headers"].iloc[0]).split(" | ")[0] \
                        if len(parts(groups[key])[1]) else ""
                    hint = "grouped" if re.search(r"industry", first, re.I) else "grid"
                else:
                    hint = "values" if out_name == "balance" else "tests"
                lines += block(groups[key], *parts(groups[key]), shape_hint=hint)
        lines.append("")
    lines.append("unused:   # every other title; move a block up into an output to use it")
    for key, titles in groups.items():
        if key not in used:
            lines += block(titles, *parts(titles), indent="  ")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"done: {path}  ({len(groups)} titles, {len(used)} placed in outputs)")


if __name__ == "__main__":
    main()
