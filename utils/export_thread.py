# utils/export_thread.py
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import TclError
print("[DEBUG] Laster export_thread.py som brukes nå...")

def show_loading_popup_with_progress(parent, total, message="Laster og genererer strekkoder…"):
    popup = tk.Toplevel()
    popup.title("Vennligst vent")
    popup.configure(bg='#bfdfff')
    popup.resizable(False, False)

    # Innhold
    tk.Label(popup, text=message, font=("Arial", 12, "bold"), bg='#bfdfff').pack(pady=(15, 5))
    bar = ttk.Progressbar(popup, mode='determinate', length=250, maximum=total)
    bar.pack(pady=(0, 5))
    percent_label = tk.Label(popup, text="0%", font=("Arial", 10), bg='#bfdfff')
    percent_label.pack()

    popup.update_idletasks()

    screen_w = popup.winfo_screenwidth()
    screen_h = popup.winfo_screenheight()
    w, h = 320, 100
    x = (screen_w - w) // 2
    y = (screen_h - h) // 2
    popup.geometry(f"{w}x{h}+{x}+{y}")

    popup.lift()
    popup.attributes('-topmost', True)
    popup.grab_set()
    popup.focus_force()

    return popup, bar, percent_label


def threaded_export(file, result, with_barcodes, export_func):
    popup, bar, percent_label = show_loading_popup_with_progress(
        parent=None,
        total=100,
        message="Eksporterer til Excel..."
    )

    def update_progress(pct):
        try:
            bar["value"] = pct
            percent_label.config(text=f"{pct}%")
        except TclError:
            pass

    def do_export():
        try:
            export_func(
                file, result,
                with_barcodes=with_barcodes,
                parent_window=None,
                progress_func=update_progress
            )
        except Exception as e:
            popup.after(0, lambda err=str(e): messagebox.showerror("Feil ved eksport", err))
        finally:
            popup.after(0, popup.destroy)

    threading.Thread(target=do_export, daemon=True).start()
