import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from fpdf import FPDF
from openpyxl import Workbook
from PIL import ImageGrab
import os
import sys

sys.path.append("path_to_directory_containing_app_logic")


class UI:
    def __init__(self, root, app_logic):
        self.root = root
        self.app_logic = app_logic  # Instância da lógica do aplicativo
        self.fields = app_logic.fields
        self.checks = app_logic.checks
        self.radio_vars = app_logic.radio_vars
        self.demandas_entries = app_logic.demandas_entries

        self.setup_styles()
        self.setup_ui()

    def setup_styles(self):
        """Configura os estilos visuais da aplicação"""
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('.', background='#f0f0f0', font=('Helvetica', 10))
        style.configure('TLabel', padding=5)
        style.configure('TEntry', padding=5)
        style.configure('TButton', padding=5, font=('Helvetica', 10, 'bold'))

        style.configure('Header.TLabel',
                        background='#4b6cb7',
                        foreground='white',
                        font=('Helvetica', 12, 'bold'),
                        padding=10)

        style.configure('Section.TLabel',
                        background='#6c757d',
                        foreground='white',
                        font=('Helvetica', 11, 'bold'),
                        padding=8)

        style.map('TButton',
                  foreground=[('active', 'white'), ('!active', 'black')],
                  background=[('active', '#4b6cb7'), ('!active', '#f0f0f0')])

    def setup_ui(self):
        """Configura a interface do usuário"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Adiciona a logo e o título
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=10)

        # Carrega a logo
        logo_path = r"c:\Users\PC\Desktop\PROMPTS ARCHICAD\PLANILHA CLIENTES\logo_empresa.png"
        self.logo_image = tk.PhotoImage(file=logo_path)  # Armazena a imagem como atributo para evitar garbage collection
        logo_label = ttk.Label(header_frame, image=self.logo_image)
        logo_label.pack(side=tk.LEFT, padx=10)

        # Adiciona o título
        title_label = ttk.Label(header_frame, text="FORMULÁRIO DE PROJETO", font=("Arial", 16, "bold"))
        title_label.pack(side=tk.LEFT, padx=10)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.Frame(notebook)
        notebook.add(info_frame, text="Informações Básicas")

        self.setup_info_tab(info_frame)

        self.status_var = tk.StringVar()
        self.status_var.set("Pronto")
        status_bar = ttk.Label(self.root, textvariable=self.status_var,
                               relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_info_tab(self, parent_frame):
        """Configura a aba de informações"""
        canvas = tk.Canvas(parent_frame)
        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        # Atualiza o scrollregion dinamicamente
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        # Permite redimensionamento horizontal e vertical
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfig(
                "frame", width=e.width
            )
        )

        # Adiciona o frame rolável ao canvas
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", tags="frame")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Vincula o scroll do mouse ao canvas
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Adiciona as seções ao frame rolável
        self.app_logic.add_client_section(scrollable_frame)
        self.app_logic.add_property_section(scrollable_frame)
        self.app_logic.add_scope_section(scrollable_frame)
        self.app_logic.add_deadlines_section(scrollable_frame)
        self.app_logic.add_demands_section(scrollable_frame)
        self.app_logic.add_buttons_section(scrollable_frame)


if __name__ == "__main__":
    from main import AppLogic

    root = tk.Tk()
    root.title("Formulário de Projeto Arquitetônico")
    root.geometry("1000x700")
    root.minsize(800, 600)

    app_logic = AppLogic()
    app_logic.set_root(root)

    ui = UI(root, app_logic)

    root.mainloop()
