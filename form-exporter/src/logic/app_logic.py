class AppLogic:
    def __init__(self):
        self.fields = {}
        self.checks = {}
        self.radio_vars = {}
        self.demandas_entries = []
        self.export_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self.root = None

    def set_root(self, root):
        self.root = root

    def add_client_section(self, frame):
        section_frame = ttk.Frame(frame)
        section_frame.pack(fill=tk.X, pady=10)
        ttk.Label(section_frame, text="INFORMAÇÕES DO CLIENTE", style="Header.TLabel").pack(fill=tk.X)

        self.add_labeled_entry(section_frame, "Nome completo:", "nome", validate="nome", width=50)
        self.add_labeled_entry(section_frame, "Telefone:", "telefone", validate="telefone", width=50)
        self.add_labeled_entry(section_frame, "Email:", "email", width=50)
        self.add_labeled_entry(section_frame, "CNPJ/CPF:", "cnpj", width=50)
        self.add_labeled_entry(section_frame, "Responsável legal:", "responsavel", width=50)
        self.add_labeled_entry(section_frame, "Telefone do Responsável:", "telefone_responsavel", validate="telefone", width=50)
        self.add_labeled_entry(section_frame, "Endereço:", "endereco", width=50)
        self.add_labeled_entry(section_frame, "CEP:", "cep", validate="cep", width=50)

    def export_to_excel(self):
        if not self.export_path.get():
            messagebox.showerror("Erro", "Selecione uma pasta de exportação antes de continuar.")
            return

        nome_cliente = self.fields.get("nome", ttk.Entry()).get().strip() or "Cliente"
        data_atual = datetime.now().strftime("%d-%m-%Y")
        excel_filename = f"{nome_cliente} - RELATÓRIO - {data_atual}.xlsx"
        excel_path = os.path.join(self.export_path.get(), excel_filename)

        wb = Workbook()
        ws = wb.active
        ws.title = "Relatório"

        self.write_client_info(ws)
        self.write_property_info(ws)
        self.write_deadlines(ws)
        self.write_demands(ws)
        self.write_architecture_info(ws)
        self.write_complementary_projects(ws)

        wb.save(excel_path)
        messagebox.showinfo("Exportar para Excel", f"Dados exportados para {excel_path} com sucesso.")

    def write_client_info(self, ws):
        ws.append(["INFORMAÇÕES DO CLIENTE"])
        client_info = {
            "Nome completo": self.fields.get("nome", ttk.Entry()).get(),
            "Telefone": self.fields.get("telefone", ttk.Entry()).get(),
            "Email": self.fields.get("email", ttk.Entry()).get(),
            "CNPJ/CPF": self.fields.get("cnpj", ttk.Entry()).get(),
            "Responsável legal": self.fields.get("responsavel", ttk.Entry()).get(),
            "Telefone do Responsável": self.fields.get("telefone_responsavel", ttk.Entry()).get(),
            "Endereço": self.fields.get("endereco", ttk.Entry()).get(),
            "CEP": self.fields.get("cep", ttk.Entry()).get()
        }
        for label, value in client_info.items():
            ws.append([label, value])

    def write_property_info(self, ws):
        ws.append([])
        ws.append(["INFORMAÇÕES DO IMÓVEL"])
        property_info = {
            "Tipo de Imóvel": self.radio_vars.get("tipo_imovel", tk.StringVar()).get(),
            "Tipo de Construção": self.radio_vars.get("tipo_construcao", tk.StringVar()).get(),
            "Metragem Quadrada": f"{self.fields.get('metragem', ttk.Entry()).get()} m²"
        }
        for label, value in property_info.items():
            ws.append([label, value])

    def write_deadlines(self, ws):
        ws.append([])
        ws.append(["PRAZOS DO PROJETO"])
        deadlines = {
            "Levantamento": "levantamento",
            "Layout": "layout",
            "Modelagem 3D": "modelagem_3d",
            "Projeto Executivo": "projeto_executivo",
            "Complementares": "complementares"
        }
        for label, key in deadlines.items():
            days_value = self.fields.get(key, ttk.Entry()).get()
            ws.append([label, f"{days_value} DIAS" if days_value else "Não informado"])

    def write_demands(self, ws):
        ws.append([])
        ws.append(["DEMANDAS DO PROJETO"])
        for nome_entry, descricao_entry in self.demandas_entries:
            nome = nome_entry.get().strip()
            descricao = descricao_entry.get().strip()
            if nome or descricao:
                ws.append([nome, descricao])

    def write_architecture_info(self, ws):
        ws.append([])
        ws.append(["PROJETO DE ARQUITETURA"])
        architecture_options = [
            ("Layout", "layout"),
            ("3D", "3d"),
            ("Detalhamento", "detalhamento")
        ]
        for text, key in architecture_options:
            if self.checks.get(key, tk.BooleanVar()).get():
                ws.append([text])

    def write_complementary_projects(self, ws):
        ws.append([])
        ws.append(["PROJETOS COMPLEMENTARES"])
        complementary_options = [
            ("Ar Condicionado", "ar_condicionado"),
            ("Elétrica", "eletrica"),
            ("Dados e Voz", "dados_voz"),
            ("Hidráulica", "hidraulica"),
            ("CFTV", "cftv"),
            ("Alarme", "alarme"),
            ("Incêndio", "incendio")
        ]
        for text, key in complementary_options:
            if self.checks.get(key, tk.BooleanVar()).get():
                ws.append([text])