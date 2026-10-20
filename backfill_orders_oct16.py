#!/usr/bin/env python3
"""
backfill_orders_oct16.py - fill the missing HKTV Staff parent-order count (and everything that hangs off it)
for a day the 03:00 job skipped, using the Tableau MTD figures already stored in history.json.

Method (same Tableau basis end to end, so the cost-report-sourced days never get mixed in):
  HKTV orders(T) = productivityMtdBasis[N] - productivityMtdBasis[P] - sum(HKTV orders of every other day in (P, N])
      T = target day, P = latest day before T that has an MTD basis, N = first day after T that has one.
  ODS orders(T)  = unique Parent Orders of ODS/VAN users in OIX_Record_YYYYMMDD (pipeline's own order_count_for_group).
      Already in the log for 2026-10-16 - the script only CHECKS it unless --apply-ods is given.

Writes (only with --write; a .bak copy is made first):
  dailyProductivityLog[T].hktvStaff      orderCount + productivity, per district and total
  hktvStaff[T]                           daily productivity (7-day rolling forecast input)
  productivityMtdBasis[T]                basis[P] + HKTV orders(T)
  productivityMtdLog[T]                  MTD productivity = basis[T] / cumulative manpower to T
Afterwards the next `kpi_pipeline.py --section tableau` run (or the 08:15 cron) re-trims productivity_history.json /
data.json / index.html from the patched history.json.

Usage (in the folder with kpi_pipeline.py and history.json):
  python3 backfill_orders_oct16.py --oix "C:/.../Digimobi Report/OIX_Record_20261016.csv"          # dry run
  python3 backfill_orders_oct16.py --oix ... --write
"""
import argparse, json, os, shutil, sys

DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX"]


def prod(orders, manpower):
    return round(orders / manpower, 2) if manpower and orders is not None else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--history", default="./history.json")
    ap.add_argument("--date", default="2026-10-16")
    ap.add_argument("--oix", help="OIX_Record_YYYYMMDD.csv for the target day (ODS check / --apply-ods)")
    ap.add_argument("--apply-ods", action="store_true", help="also overwrite the ODS order count from the OIX file")
    ap.add_argument("--write", action="store_true", help="save changes (default is a dry run)")
    a = ap.parse_args()

    with open(a.history, encoding="utf-8") as f:
        h = json.load(f)
    log, basis, mtd = h["dailyProductivityLog"], h["productivityMtdBasis"], h.setdefault("productivityMtdLog", {})
    T = a.date
    month = T[:7]

    P = max((d for d in basis if d < T and d.startswith(month)), default=None)
    N = min((d for d in basis if d > T and d.startswith(month)), default=None)
    if not P or not N:
        sys.exit(f"Need an MTD basis before and after {T}; found before={P}, after={N}.")
    others = [d for d in sorted(log) if P < d <= N and d != T]
    gaps = [d for d in others if log[d]["hktvStaff"]["total"].get("orderCount") is None]
    if gaps:
        sys.exit(f"These days between {P} and {N} also have no HKTV order count, fix them first: {gaps}")

    def hk(d, x=None):
        g = log[d]["hktvStaff"]
        return (g["total"] if x is None else g["districts"][x])["orderCount"]

    def bs(d, x=None):
        return basis[d]["overall"] if x is None else basis[d]["districts"][x]

    new_d = {x: int(round(bs(N, x) - bs(P, x) - sum(hk(d, x) for d in others))) for x in DISTRICTS}
    new_total = int(round(bs(N) - bs(P) - sum(hk(d) for d in others)))
    print(f"Target {T}: MTD basis {P} -> {N}, other days subtracted: {others}")
    print(f"HKTV parent orders {T}: total {new_total:,} (sum of districts {sum(new_d.values()):,})")
    print("  " + ", ".join(f"{x} {v:,}" for x, v in new_d.items()))
    if abs(new_total - sum(new_d.values())) > 5:
        sys.exit("District sum and overall differ by more than rounding - stopping, nothing written.")

    # ---- ODS: check against the OIX file with the pipeline's own functions ----
    ods = log[T]["odsRatio"]
    if a.oix:
        os.environ.setdefault("TABLEAU_USER", "x")
        os.environ.setdefault("TABLEAU_PASS", "x")
        sys.path.insert(0, os.getcwd())
        import kpi_pipeline as kp
        oix = kp.order_count_for_group(kp.process_oix(kp.load_oix(a.oix)), ("ODS", "VAN"))
        same = all(ods["districts"][x].get("orderCount") == oix["districts"][x] for x in DISTRICTS)
        print(f"ODS/VAN parent orders from OIX: {oix['overall']:,}; in log: {ods['total'].get('orderCount')}; "
              f"districts match: {same}")
        if a.apply_ods:
            for x in DISTRICTS:
                g = ods["districts"][x]
                g["orderCount"] = oix["districts"][x]
                g["productivity"] = prod(g["orderCount"], g.get("manpower"))
            ods["total"]["orderCount"] = oix["overall"]
            ods["total"]["productivity"] = prod(oix["overall"], ods["total"].get("manpower"))
            print("  ODS order count / productivity overwritten from OIX.")
        print(f"Check: HKTV {new_total:,} + ODS {oix['overall']:,} = network {new_total + oix['overall']:,}")

    # ---- apply HKTV ----
    g = log[T]["hktvStaff"]
    for x in DISTRICTS:
        g["districts"][x]["orderCount"] = float(new_d[x])
        g["districts"][x]["productivity"] = prod(new_d[x], g["districts"][x].get("manpower"))
    g["total"]["orderCount"] = float(new_total)
    g["total"]["productivity"] = prod(new_total, g["total"].get("manpower"))

    h.setdefault("hktvStaff", {})[T] = {
        "overall": g["total"]["productivity"],
        "districts": {x: g["districts"][x]["productivity"] for x in DISTRICTS},
    }

    basis[T] = {"overall": float(bs(P) + new_total),
                "districts": {x: float(bs(P, x) + new_d[x]) for x in DISTRICTS}}

    def mp(x=None):
        s = 0
        for d, e in log.items():
            if d.startswith(month) and d <= T:
                gg = e["hktvStaff"]
                s += ((gg["total"] if x is None else gg["districts"][x]).get("manpower") or 0)
        return s

    mtd[T] = {"overall": prod(basis[T]["overall"], mp()),
              "districts": {x: prod(basis[T]["districts"][x], mp(x)) for x in DISTRICTS}}
    print(f"HKTV productivity {T}: {g['total']['productivity']}  | MTD productivity {T}: {mtd[T]['overall']}")
    print("  basis chain check: basis[T] + HKTV(N) == basis[N] ->",
          basis[T]["overall"] + hk(N) == basis[N]["overall"])

    if not a.write:
        print("\nDRY RUN - nothing written. Add --write to save.")
        return
    shutil.copy2(a.history, a.history + ".bak_orders_" + T.replace("-", ""))
    with open(a.history, "w", encoding="utf-8") as f:
        json.dump(h, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {a.history} (backup: {a.history}.bak_orders_{T.replace('-', '')})")


if __name__ == "__main__":
    main()
