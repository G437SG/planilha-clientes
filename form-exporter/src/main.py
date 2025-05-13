# File: /form-exporter/form-exporter/src/main.py

from tkinter import Tk
from logic.app_logic import AppLogic
from ui import UI

def main():
    root = Tk()
    root.title("Formulário de Projeto Arquitetônico")
    root.geometry("1000x700")
    root.minsize(800, 600)

    app_logic = AppLogic()
    app_logic.set_root(root)
    ui = UI(root, app_logic)

    root.mainloop()

if __name__ == "__main__":
    main()