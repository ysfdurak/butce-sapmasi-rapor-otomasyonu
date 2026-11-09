"""Offline launcher: bundled pure-Python wheels, no pip or TLS required."""

from pathlib import Path
import sys

BASE = Path(__file__).resolve().parent


def dependencies():
    if sys.version_info < (3, 9):
        raise RuntimeError("Python 3.9 veya üzeri gerekli. 3.13 / 3.14 kullanılabilir.")
    # Prefer bundled tested versions, independent of globally installed packages.
    for wheel in sorted((BASE / "wheels").glob("*.whl")):
        sys.path.insert(0, str(wheel))
    import openpyxl

    if openpyxl.__version__ != "3.1.5":
        raise RuntimeError("Paketin wheels klasörü eksik veya değişmiş.")


if __name__ == "__main__":
    try:
        dependencies()
        from app import main

        main()
    except Exception:
        import traceback

        error = traceback.format_exc()
        (BASE / "Baslatma_Hatasi.txt").write_text(error, encoding="utf-8")
        print(error)
        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Bütçe Analiz", error)
            root.destroy()
        except Exception:
            pass
        sys.exit(1)
