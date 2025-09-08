from pathlib import Path

BASE = Path(r"C:\kb\output")

def group_key(p: Path) -> str:
    rel = p.relative_to(BASE)
    return rel.parts[0] if len(rel.parts) > 1 else "root"

# gather .txt files but ignore previously created bundles
txt_files = [p for p in BASE.rglob("*.txt") if not p.name.startswith("bundle__")]

if not txt_files:
    print(f"No .txt files found under {BASE}. Did you run the converter?")
    raise SystemExit(0)

groups = {}
for p in txt_files:
    groups.setdefault(group_key(p), []).append(p)

for grp, files in sorted(groups.items()):
    out = BASE / f"bundle__{grp}.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with out.open("w", encoding="utf-8", errors="ignore") as w:
        for f in sorted(files):
            w.write(f"\n\n===== FILE: {f.relative_to(BASE)} =====\n\n")
            try:
                w.write(f.read_text(encoding="utf-8", errors="ignore"))
            except Exception as e:
                w.write(f"[READ ERROR] {e}")
            count += 1
    print(f"Wrote {out} ({count} files)")

# OPTIONAL: also bundle all Mermaid flowcharts into one file so they count as 1 upload
mmd_files = list(BASE.rglob("*.mermaid.mmd"))
if mmd_files:
    mmd_out = BASE / "bundle__mermaid.mmd"
    with mmd_out.open("w", encoding="utf-8", errors="ignore") as w:
        for f in sorted(mmd_files):
            w.write(f"%% FILE: {f.relative_to(BASE)}\n")
            w.write(f.read_text(encoding="utf-8", errors="ignore"))
            w.write("\n\n")
    print(f"Wrote {mmd_out} ({len(mmd_files)} mermaid files)")
else:
    print("No .mermaid.mmd files found (only made text bundles).")
