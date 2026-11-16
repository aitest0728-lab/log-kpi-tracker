from pathlib import Path
p = Path(r"C:\Users\chipanl\Downloads\AI Fetching for Order Date = T-1_data.csv.part.bad")
print("exists", p.exists(), "size", p.stat().st_size if p.exists() else None)
if p.exists():
    raw = p.read_bytes()
    print("bom hex", raw[:16].hex())
    text = raw.decode("utf-16")
    first = text.splitlines()[0]
    print("header repr:", repr(first[:500]))
    print("tab count", first.count("\t"), "comma count", first.count(","))
    print("ncols tab", len(first.split("\t")), "ncols comma", len(first.split(",")))
c = Path(r"C:\Users\chipanl\Downloads\AI Fetching for Order Date = T-1_data.converted.csv")
if c.exists():
    print("converted size", c.stat().st_size)
    chead = c.read_text(encoding="utf-8-sig").splitlines()[0]
    print("converted header:", repr(chead[:500]))
