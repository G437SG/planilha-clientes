from tkinter import ttk, messagebox, Tk

class UI:
    def __init__(self, root, app_logic):
        self.root = root
        self.app_logic = app_logic
        self.create_widgets()

    def create_widgets(self):
        self.create_header()
        self.create_client_section()
        self.create_property_section()
        self.create_scope_section()
        self.create_demands_section()
        self.create_deadlines_section()
        self.create_buttons_section()

    def create_header(self):
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, pady=10)
        ttk.Label(header_frame, text="Formulário de Projeto Arquitetônico", font=("Arial", 16)).pack()

    def create_client_section(self):
        client_frame = ttk.LabelFrame(self.root, text="Informações do Cliente")
        client_frame.pack(fill=tk.X, padx=10, pady=10)
        self.app_logic.add_client_section(client_frame)

    def create_property_section(self):
        property_frame = ttk.LabelFrame(self.root, text="Informações do Imóvel")
        property_frame.pack(fill=tk.X, padx=10, pady=10)
        self.app_logic.add_property_section(property_frame)

    def create_scope_section(self):
        scope_frame = ttk.LabelFrame(self.root, text="Escopo")
        scope_frame.pack(fill=tk.X, padx=10, pady=10)
        self.app_logic.add_scope_section(scope_frame)

    def create_demands_section(self):
        demands_frame = ttk.LabelFrame(self.root, text="Demandas do Projeto")
        demands_frame.pack(fill=tk.X, padx=10, pady=10)
        self.app_logic.add_demands_section(demands_frame)

    def create_deadlines_section(self):
        deadlines_frame = ttk.LabelFrame(self.root, text="Prazos do Projeto")
        deadlines_frame.pack(fill=tk.X, padx=10, pady=10)
        self.app_logic.add_deadlines_section(deadlines_frame)

    def create_buttons_section(self):
        buttons_frame = ttk.Frame(self.root)
        buttons_frame.pack(fill=tk.X, pady=10)
        ttk.Button(buttons_frame, text="Salvar", command=self.app_logic.save_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Limpar", command=self.app_logic.clear_form).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Exportar para PDF", command=self.app_logic.export_to_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Exportar para Excel", command=self.app_logic.export_to_excel).pack(side=tk.LEFT, padx=5)