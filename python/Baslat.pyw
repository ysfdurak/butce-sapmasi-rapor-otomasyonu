try:
    from bootstrap import dependencies
    dependencies()
    from app import main
    main()
except Exception:
    from pathlib import Path
    import traceback,tkinter as tk
    from tkinter import messagebox
    error=traceback.format_exc()
    Path(__file__).with_name('Baslatma_Hatasi.txt').write_text(error,encoding='utf-8')
    root=tk.Tk();root.withdraw();messagebox.showerror('Bütçe Dashboard',error);root.destroy()
