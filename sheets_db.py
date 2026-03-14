"""
Módulo de persistência com Google Sheets
Azevedo Contabilidade
"""

import json
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

HEADERS_PROPOSTAS = [
    "id", "data", "cliente", "tratamento", "telefone", "email",
    "vendedor", "servicos", "valor", "status", "obs"
]

HEADERS_CONFIG = ["chave", "valor"]


def get_client():
    """Retorna cliente gspread autenticado via st.secrets"""
    if not GSPREAD_AVAILABLE:
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception:
        return None


def get_spreadsheet(client):
    """Abre a planilha pelo ID em st.secrets"""
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        if sheet_id:
            return client.open_by_key(sheet_id)
        return client.open("Propostas_Azevedo")
    except Exception:
        return None


def init_sheets(spreadsheet):
    """Garante que as abas existam com os headers corretos"""
    # Aba Propostas
    try:
        ws = spreadsheet.worksheet("Propostas")
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet("Propostas", rows=1000, cols=len(HEADERS_PROPOSTAS))
        ws.append_row(HEADERS_PROPOSTAS)

    # Aba Config
    try:
        ws_cfg = spreadsheet.worksheet("Config")
    except gspread.WorksheetNotFound:
        ws_cfg = spreadsheet.add_worksheet("Config", rows=50, cols=2)
        ws_cfg.append_row(HEADERS_CONFIG)
        ws_cfg.append_row(["meta_mensal", "12000"])
        ws_cfg.append_row(["vendedores", json.dumps(["Allan"])])

    return True


# ===== PROPOSTAS =====

def load_propostas():
    """Carrega todas as propostas do Google Sheets"""
    client = get_client()
    if not client:
        return []

    sp = get_spreadsheet(client)
    if not sp:
        return []

    init_sheets(sp)

    try:
        ws = sp.worksheet("Propostas")
        records = ws.get_all_records()
        for r in records:
            try:
                r["valor"] = float(str(r.get("valor", 0)).replace(",", "."))
            except:
                r["valor"] = 0.0
            try:
                r["id"] = int(r.get("id", 0))
            except:
                r["id"] = 0
        return records
    except Exception:
        return []


def save_proposta(proposta):
    """Adiciona uma nova proposta ao Google Sheets"""
    client = get_client()
    if not client:
        return False

    sp = get_spreadsheet(client)
    if not sp:
        return False

    init_sheets(sp)

    try:
        ws = sp.worksheet("Propostas")
        row = [proposta.get(h, "") for h in HEADERS_PROPOSTAS]
        ws.insert_rows([row], row=2)  # Insere após o header (mais recente primeiro)
        return True
    except Exception:
        return False


def update_proposta_status(proposta_id, novo_status):
    """Atualiza o status de uma proposta"""
    client = get_client()
    if not client:
        return False

    sp = get_spreadsheet(client)
    if not sp:
        return False

    try:
        ws = sp.worksheet("Propostas")
        cell = ws.find(str(proposta_id), in_column=1)
        if cell:
            status_col = HEADERS_PROPOSTAS.index("status") + 1
            ws.update_cell(cell.row, status_col, novo_status)
            return True
    except Exception:
        pass
    return False


def delete_proposta(proposta_id):
    """Exclui uma proposta"""
    client = get_client()
    if not client:
        return False

    sp = get_spreadsheet(client)
    if not sp:
        return False

    try:
        ws = sp.worksheet("Propostas")
        cell = ws.find(str(proposta_id), in_column=1)
        if cell:
            ws.delete_rows(cell.row)
            return True
    except Exception:
        pass
    return False


# ===== CONFIG =====

def load_config_sheets():
    """Carrega configurações do Google Sheets"""
    default = {"meta_mensal": 12000, "vendedores": ["Allan"]}

    client = get_client()
    if not client:
        return default

    sp = get_spreadsheet(client)
    if not sp:
        return default

    init_sheets(sp)

    try:
        ws = sp.worksheet("Config")
        records = ws.get_all_records()
        config = {}
        for r in records:
            chave = r.get("chave", "")
            valor = r.get("valor", "")
            if chave == "meta_mensal":
                try:
                    config["meta_mensal"] = float(str(valor).replace(",", "."))
                except:
                    config["meta_mensal"] = 12000
            elif chave == "vendedores":
                try:
                    config["vendedores"] = json.loads(valor)
                except:
                    config["vendedores"] = [valor] if valor else ["Allan"]
            else:
                config[chave] = valor

        for k, v in default.items():
            if k not in config:
                config[k] = v
        return config
    except Exception:
        return default


def save_config_sheets(config):
    """Salva configurações no Google Sheets"""
    client = get_client()
    if not client:
        return False

    sp = get_spreadsheet(client)
    if not sp:
        return False

    try:
        ws = sp.worksheet("Config")
        ws.clear()
        ws.append_row(HEADERS_CONFIG)
        ws.append_row(["meta_mensal", str(config.get("meta_mensal", 12000))])
        ws.append_row(["vendedores", json.dumps(config.get("vendedores", ["Allan"]))])
        return True
    except Exception:
        return False


# ===== VERIFICAÇÃO =====

def sheets_disponivel():
    """Verifica se o Google Sheets está configurado e acessível"""
    client = get_client()
    if not client:
        return False
    sp = get_spreadsheet(client)
    return sp is not None
