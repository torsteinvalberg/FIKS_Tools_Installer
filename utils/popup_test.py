import tkinter as tk

def main():
    root = tk.Tk()
    root.title("Testvindu")
    tk.Label(root, text="Hvis du ser dette – Tkinter funker!").pack(pady=30)
    root.mainloop()

if __name__ == "__main__":
    main()
