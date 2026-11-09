"""Native file/folder dialogs, run in their own interpreter UI thread."""
import sys,json
import tkinter as tk
from tkinter import filedialog
root=tk.Tk();root.withdraw();root.attributes('-topmost',True)
try:
    mode=sys.argv[1]
    if mode=='folder':value=filedialog.askdirectory(title='Çıktı klasörünü seçin')
    elif mode=='files':value=list(filedialog.askopenfilenames(title='Açıklamalı raporları seçin',filetypes=[('Excel','*.xlsx')]))
    else:value=filedialog.askopenfilename(title='Excel dosyasını seçin',filetypes=[('Excel','*.xlsx')])
    print(json.dumps(value,ensure_ascii=True))
finally:root.destroy()
