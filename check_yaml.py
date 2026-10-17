"""Check an edited trustee YAML against survey.xlsx: does every layout have the columns you kept?

    python check_yaml.py --survey survey.xlsx --yaml us_bank.yaml

For every source in balance / performance / collateral it prints each layout (distinct header line)
of its titles, with the standard names that layout does NOT have. A missing collateral key or a
missing column you need means: add that layout's spelling to the column's list.
Titles of a source the survey never saw are listed too (typo?).
"""
import argparse

import pandas as pd
import yaml


def norm(s):
    return " ".join(str(s).lower().split())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--survey", default="survey.xlsx")
    ap.add_argument("--yaml", required=True)
    a = ap.parse_args()

    cfg = yaml.safe_load(open(a.yaml, encoding="utf-8"))
    x = pd.read_excel(a.survey, sheet_name=None)
    names = [norm(n) for n in cfg.get("names", [cfg["trustee"]])]
    mine = lambda df: df[df["trustee"].map(norm).isin(names)]
    t, h, l = mine(x["titles"]), mine(x["headers"]), mine(x["labels"])
    total = int(t["of_reports"].max()) if len(t) else 0
    problems = 0

    for out in ("balance", "performance", "collateral"):
        sec = cfg.get(out) or {}
        key = sec.get("key") or []
        for src in sec.get("sources") or []:
            titles = src.get("titles") or []
            print(f"\n[{out}] {' / '.join(titles)}  ({src.get('shape')})")
            seen = t[t["title"].isin(titles)]
            for ti in titles:
                if ti not in set(seen["title"]):
                    print(f"   ! title never seen in the survey: {ti!r}")
                    problems += 1
            n = int(seen["reports"].sum()) if len(seen) else 0
            print(f"   in {n} of {total} reports")

            cols = src.get("columns") or {}
            if cols:
                want = {k: {norm(v) for v in (vs or [])} for k, vs in cols.items()}
                for col in key:
                    if col not in want:
                        print(f"   ! key column {col!r} is not among this source's columns")
                        problems += 1
                for _, r in h[h["title"].isin(titles)].sort_values("reports", ascending=False).iterrows():
                    cells = {norm(c) for c in str(r["headers"]).split(" | ")}
                    if "no asset records" in str(r["headers"]).lower():
                        print(f"   - {r['reports']:>4} reports  empty page" +
                              ("" if src.get("empty_ok") else "   ! add empty_ok: true"))
                        continue
                    missing = [k for k, vs in want.items() if not (vs & cells)]
                    flag = " !" if missing else "  "
                    problems += bool(missing)
                    print(f"  {flag} {r['reports']:>4} reports {r['first']}..{r['last']}  "
                          + ("all columns found" if not missing else "missing: " + ", ".join(missing)))
                    if missing:
                        print(f"        header: {r['headers']}")

            labels = src.get("labels") or {}
            if labels:
                ll = l[l["title"].isin(titles)]
                for k, vs in labels.items():
                    hit = ll[ll["label"].map(norm).isin({norm(v) for v in (vs or [])})]
                    got = int(hit["reports"].sum())
                    flag = " !" if got < n else "  "
                    problems += got < n
                    print(f"  {flag} label {k}: in {got} of {n} reports")

    print(f"\n{problems} problem(s)" if problems else "\nno problems: every layout has every column you kept")


if __name__ == "__main__":
    main()
