from openpyxl import Workbook
from datetime import datetime

def export_to_excel(fields, radio_vars, demandas_entries, export_path):
    if not export_path:
        raise ValueError("Export path must be provided.")

    nome_cliente = fields.get("nome", "").strip() or "Cliente"
    data_atual = datetime.now().strftime("%d-%m-%Y")
    excel_filename = f"{nome_cliente} - RELATÓRIO - {data_atual}.xlsx"
    excel_path = f"{export_path}/{excel_filename}"

    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"

    # INFORMAÇÕES DO CLIENTE
    ws.append(["INFORMAÇÕES DO CLIENTE"])
    client_info = {
        "Nome completo": fields.get("nome", ""),
        "Telefone": fields.get("telefone", ""),
        "Email": fields.get("email", ""),
        "CNPJ/CPF": fields.get("cnpj", ""),
        "Responsável legal": fields.get("responsavel", ""),
        "Telefone do Responsável": fields.get("telefone_responsavel", ""),
        "Endereço": fields.get("endereco", ""),
        "CEP": fields.get("cep", "")
    }
    for label, value in client_info.items():
        ws.append([label, value])

    # INFORMAÇÕES DO IMÓVEL
    ws.append([])
    ws.append(["INFORMAÇÕES DO IMÓVEL"])
    property_info = {
        "Tipo de Imóvel": radio_vars.get("tipo_imovel", ""),
        "Tipo de Construção": radio_vars.get("tipo_construcao", ""),
        "Metragem Quadrada": f"{fields.get('metragem', '')} m²"
    }
    for label, value in property_info.items():
        ws.append([label, value])

    # PRAZOS DO PROJETO
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
        days_value = fields.get(key, "")
        ws.append([label, f"{days_value} DIAS" if days_value else "Não informado"])

    # DEMANDAS DO PROJETO
    ws.append([])
    ws.append(["DEMANDAS DO PROJETO"])
    for nome_entry, descricao_entry in demandas_entries:
        nome = nome_entry.get().strip()
        descricao = descricao_entry.get().strip()
        if nome or descricao:
            ws.append([nome, descricao])

    # PROJETO DE ARQUITETURA
    ws.append([])
    ws.append(["PROJETO DE ARQUITETURA"])
    architecture_options = [
        ("Layout", "layout"),
        ("3D", "3d"),
        ("Detalhamento", "detalhamento")
    ]
    for text, key in architecture_options:
        if fields.get(key, False):
            ws.append([text])

    # PROJETOS COMPLEMENTARES
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
        if fields.get(key, False):
            ws.append([text])

    wb.save(excel_path)