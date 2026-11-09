"""VBA kaynak araçları: vba_elle kopyalarını üretir ve VBA modüllerini statik olarak denetler.

Kullanım (depo kökünden):
    python tools/vba_araclari.py uret      # makro/vba (UTF-8) -> makro/vba_elle (Windows-1254, CRLF)
    python tools/vba_araclari.py kontrol   # kopyalar güncel mi + statik denetim (CI bu komutu çalıştırır)

Tek kaynak makro/vba klasörüdür; vba_elle elle düzenlenmez.
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "makro" / "vba"
TARGET = ROOT / "makro" / "vba_elle"
INSTALLER = ROOT / "makro" / "Kurulum.vbs"

# VBA'da prosedür/değişken adı olamayan deyimler (ör. eski "Sub Put" derleme hatası).
RESERVED = {
    "put", "get", "open", "close", "print", "input", "line", "seek", "lock", "unlock", "write", "name",
    "kill", "reset", "width", "let", "set", "dim", "redim", "call", "goto", "gosub", "return", "resume",
    "stop", "end", "exit", "with", "select", "case", "loop", "next", "wend", "erase", "error", "event",
    "friend", "implements", "new", "nothing", "type", "enum", "property", "date", "mid", "load", "unload",
}
DECLARATION = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.IGNORECASE,
)


def encoded(text):
    return text.replace("\r\n", "\n").replace("\n", "\r\n").encode("cp1254")


def lint(path, text):
    problems = []
    lines = text.splitlines()
    if not lines or lines[0] != f'Attribute VB_Name = "{path.stem}"':
        problems.append(f"{path.name}: ilk satır Attribute VB_Name = \"{path.stem}\" olmalı")
    if "Option Explicit" not in lines[:3]:
        problems.append(f"{path.name}: Option Explicit eksik")
    for number, line in enumerate(lines, 1):
        match = DECLARATION.match(line)
        if match and match.group(2).lower() in RESERVED:
            problems.append(f"{path.name}:{number}: '{match.group(2)}' VBA'da ayrılmış sözcük; prosedür adı olamaz")
        try:
            line.encode("cp1254")
        except UnicodeEncodeError as error:
            problems.append(
                f"{path.name}:{number}: Windows-1254'te olmayan karakter {line[error.start]!r}; ChrW() kullanın"
            )
    return problems


def project_checks():
    """Modüller arası denetim: aynı adlı genel prosedür (derleme belirsizliği) ve düğmelerin çağırdığı makrolar."""
    problems, public, everything = [], {}, set()
    texts = {path.name: path.read_text(encoding="utf-8") for path in modules()}
    for name, text in texts.items():
        for line in text.splitlines():
            match = DECLARATION.match(line)
            if not match:
                continue
            everything.add(match.group(2).lower())
            if not line.lstrip().lower().startswith("private"):
                key = match.group(2).lower()
                if key in public and public[key] != name:
                    problems.append(f"{match.group(2)}: hem {public[key]} hem {name} içinde genel prosedür")
                public.setdefault(key, name)
    actions = set()
    for text in texts.values():
        actions |= set(re.findall(r'AddWorkButton\s+\w+,\s*"[^"]*",\s*"(\w+)"', text))
        for group in re.findall(r"actions\s*=\s*Array\(([^)]*)\)", text):
            actions |= set(re.findall(r'"(\w+)"', group))
    for action in sorted(actions):
        if action.lower() not in public:
            problems.append(f"Düğme makrosu bulunamadı veya Private: {action}")
    return problems


def modules():
    return sorted(SOURCE.glob("*.bas"))


def generate():
    TARGET.mkdir(exist_ok=True)
    for path in modules():
        (TARGET / path.name).write_bytes(encoded(path.read_text(encoding="utf-8")))
    for stale in set(p.name for p in TARGET.glob("*.bas")) - set(p.name for p in modules()):
        (TARGET / stale).unlink()
    print(f"{len(modules())} modül vba_elle klasörüne yazıldı.")


def check():
    problems = []
    names = {p.stem for p in modules()}
    for path in modules():
        text = path.read_text(encoding="utf-8")
        problems += lint(path, text)
        copy = TARGET / path.name
        try:
            if copy.read_bytes() != encoded(text):
                problems.append(f"vba_elle/{path.name} güncel değil: python tools/vba_araclari.py uret")
        except FileNotFoundError:
            problems.append(f"vba_elle/{path.name} yok: python tools/vba_araclari.py uret")
        except UnicodeEncodeError:
            pass  # lint zaten bildirdi
    problems += project_checks()
    extra = {p.stem for p in TARGET.glob("*.bas")} - names
    problems += [f"vba_elle/{n}.bas kaynakta yok" for n in sorted(extra)]
    listed = set(re.findall(r'"(Makro\w+)"', INSTALLER.read_text(encoding="utf-8")))
    if listed != names:
        problems.append(f"Kurulum.vbs modül listesi farklı: eksik {sorted(names - listed)}, fazla {sorted(listed - names)}")
    for problem in problems:
        print("HATA:", problem)
    print(f"{len(modules())} modül denetlendi; {len(problems)} sorun.")
    return 1 if problems else 0


if __name__ == "__main__":
    command = sys.argv[1] if len(sys.argv) > 1 else "kontrol"
    if command == "uret":
        generate()
    elif command == "kontrol":
        sys.exit(check())
    else:
        sys.exit(__doc__)
