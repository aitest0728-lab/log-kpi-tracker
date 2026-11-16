import csv, shutil, sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HK = ZoneInfo("Asia/Hong_Kong")
src = Path(r"C:\Users\chipanl\Downloads\AI Fetching for Order Date = T-1_data.csv.part.bad")
stable = Path(r"C:\Users\chipanl\Downloads\AI Fetching for Order Date = T-1_data.utf16.tsv")
out = Path(r"C:\Users\chipanl\Downloads\AI Fetching for Order Date = T-1_data.converted.csv")
LAYOUT_A_PARENT = "Order (Unlimited & Add-on Group to Parent)"

if not stable.exists() or stable.stat().st_size != src.stat().st_size:
    shutil.copy2(src, stable)
print("stable size", stable.stat().st_size)

raw = stable.read_bytes()
text = raw.decode("utf-16")
rows = list(csv.reader(text.splitlines(), delimiter="\t"))
header = [h.strip().lstrip("\ufeff") for h in rows[0]]
while header and header[-1] == "":
    header.pop()
is_a = LAYOUT_A_PARENT in header and any(h.lower() in ("order_number", "order number") for h in header)
print("layout_A", is_a, "n_data_rows", len(rows) - 1, "header_ok", True)
if not is_a:
    raise SystemExit("not layout A")

width = len(header)
body = []
for r in rows[1:]:
    if len(r) < width:
        r = r + [""] * (width - len(r))
    else:
        r = r[:width]
    body.append(r)

od_i = next((i for i, h in enumerate(header) if "order date" in h.lower()), None)
dd_i = next((i for i, h in enumerate(header) if "delivery" in h.lower() and "date" in h.lower()), None)
on_i = next((i for i, h in enumerate(header) if h.lower() in ("order_number", "order number")), 1)
ods, dds = [], []
orders, parents = set(), set()
sample_od, sample_dd = [], []
for r in body:
    if od_i is not None and r[od_i]:
        ods.append(r[od_i])
        if len(sample_od) < 3:
            sample_od.append(r[od_i])
    if dd_i is not None and r[dd_i]:
        dds.append(r[dd_i])
        if len(sample_dd) < 3:
            sample_dd.append(r[dd_i])
    if r[on_i]:
        orders.add(r[on_i])
    if r[0]:
        parents.add(r[0])

od_min, od_max = (min(ods) if ods else None), (max(ods) if ods else None)
dd_min, dd_max = (min(dds) if dds else None), (max(dds) if dds else None)
print("sample_od", sample_od)
print("sample_dd", sample_dd)
print("n_rows", len(body), "unique_orders", len(orders), "unique_parents", len(parents))
print("od_span", od_min, "..", od_max)
print("dd_span", dd_min, "..", dd_max)

tmp = out.with_suffix(".csv.tmp")
with tmp.open("w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f)
    w.writerow(header)
    w.writerows(body)

now = datetime.now(HK)
meta_lines = [
    f"exported_at_hkt={now:%Y-%m-%d %H:%M:%S}",
    "suggested_name=AI Fetching for Order Date = T-1.csv",
    f"path={out}",
    f"size={tmp.stat().st_size}",
    "layout=A",
    f"n_rows={len(body)}",
    f"unique_orders={len(orders)}",
    f"unique_parents={len(parents)}",
    f"order_date_span={od_min} .. {od_max}",
    f"delivery_date_span={dd_min} .. {dd_max}",
    "view_url=https://inhouse-analytics.hktv.com.hk/#/views/AIFetchingforOrderDateT-1/AIFetchingforOrderDateT-1_1",
    "start_date=2026-11-01",
    "encoding=utf-8-sig",
    "source_export=utf-16-le TSV converted to utf-8-sig CSV (detector tab/comma tie)",
    "refreshed_this_run=1",
    f"source_raw={stable}",
    f"source_raw_size={stable.stat().st_size}",
]
stamp = Path(r"C:\Users\chipanl\Downloads\AI Fetching for Order Date = T-1_data.converted.export_meta.txt")
stamp.write_text("\n".join(meta_lines) + "\n", encoding="utf-8")

if out.exists():
    same_bytes = out.read_bytes() == tmp.read_bytes()
    print("previous size", out.stat().st_size, "new size", tmp.stat().st_size, "identical", same_bytes)
    if not same_bytes:
        bak = out.with_name(out.name + f".bak_{now:%Y%m%d_%H%M}")
        out.replace(bak)
        print("backed up previous")
        tmp.replace(out)
    else:
        tmp.unlink()
        print("identical to previous converted; kept previous file")
else:
    tmp.replace(out)
print("final", out, "size", out.stat().st_size)
