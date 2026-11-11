"""demo/manifest.json dosyasını üretir: tarayıcı demosunun yüklediği Python dosyaları,
paketli wheel'ler ve sentetik örnek veriler. Örnek veri veya modül eklenince çalıştırın:
    python tools/demo_manifest.py
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PY = ROOT / "python"
SKIP = {"dialog.py", "run_tests.py"}

files = sorted(p.name for p in PY.glob("*.py") if p.name not in SKIP)
files += sorted(p.relative_to(PY).as_posix() for p in (PY / "wheels").glob("*.whl"))
files += sorted(
    p.relative_to(PY).as_posix()
    for p in (PY / "ornek").rglob("*")
    if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
)
(ROOT / "demo" / "manifest.json").write_text(
    json.dumps({"files": files}, ensure_ascii=False, indent=1) + "\n", encoding="utf-8"
)
print(len(files), "dosya")
