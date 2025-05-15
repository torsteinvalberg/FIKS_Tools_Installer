import tkinter as tk

def center_window(win, parent=None):
    
    win.update_idletasks()

    # 2) Les av win‐størrelse
    w = win.winfo_width()
    h = win.winfo_height()

    # 3) Beregn senterkoordinater
    if parent:
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        x = px + (pw - w)//2
        y = py + (ph - h)//2
    else:
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = (sw - w)//2
        y = (sh - h)//2

    # 4) Sett posisjon + la det ligge øverst
    win.geometry(f"{w}x{h}+{x}+{y}")
    win.lift()
    win.focus_force()
    win.attributes('-topmost', True)
