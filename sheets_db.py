"""
Módulo de persistência com Google Sheets
Azevedo Contabilidade
Com cache para performance otimizada
"""

import json
import streamlit as st
from time import time

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
    "vendedor", "servicos", "valor", "status", "obs", "motivo_perda", "historico",
    "servicos_detalhados"
]

HEADERS_CONFIG = ["chave", "valor"]

# Cache TTL em segundos (2 minutos)
CACHE_TTL = 120


# ===== CONEXÃO COM CACHE =====

@st.cache_resource(ttl=300)
def _get_client():
    """Retorna cliente gspread autenticado (cacheado por 5 min)"""
    if not GSPREAD_AVAILABLE:
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception:
        return None


def _get_spreadsheet():
    """Abre a planilha (usa cliente cacheado)"""
    client = _get_client()
    if not client:
        return None
    try:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        if sheet_id:
            return client.open_by_key(sheet_id)
        return client.open("Propostas_Azevedo")
    except Exception:
        return None


def _init_sheets_if_needed(sp):
    """Garante que as abas existam e que os headers estejam atualizados"""
    if st.session_state.get("_sheets_initialized") and st.session_state.get("_sheets_migrated_v2"):
        return True

    try:
        try:
            ws = sp.worksheet("Propostas")
            # Verificar se os headers estão atualizados (migração)
            existing_headers = ws.row_values(1)
            missing = [h for h in HEADERS_PROPOSTAS if h not in existing_headers]
            if missing:
                # Expandir a planilha se necessário
                total_cols_needed = len(existing_headers) + len(missing)
                if ws.col_count < total_cols_needed:
                    ws.resize(cols=total_cols_needed)
                # Adicionar colunas faltantes
                for col_name in missing:
                    next_col = len(existing_headers) + 1
                    ws.update_cell(1, next_col, col_name)
                    existing_headers.append(col_name)
        except gspread.WorksheetNotFound:
            ws = sp.add_worksheet("Propostas", rows=1000, cols=len(HEADERS_PROPOSTAS))
            ws.append_row(HEADERS_PROPOSTAS)

        try:
            sp.worksheet("Config")
        except gspread.WorksheetNotFound:
            ws_cfg = sp.add_worksheet("Config", rows=50, cols=2)
            ws_cfg.append_row(HEADERS_CONFIG)
            ws_cfg.append_row(["meta_mensal", "12000"])
            ws_cfg.append_row(["vendedores", json.dumps(["Allan"])])

        st.session_state["_sheets_initialized"] = True
        st.session_state["_sheets_migrated_v2"] = True
        return True
    except Exception as e:
        st.session_state["_sheets_init_error"] = str(e)
        return False


# ===== CACHE DE DADOS =====

def _cache_valid(key):
    """Verifica se o cache ainda é válido"""
    ts = st.session_state.get(f"_cache_ts_{key}", 0)
    return (time() - ts) < CACHE_TTL


def _set_cache(key, data):
    """Salva dados no cache da sessão"""
    st.session_state[f"_cache_{key}"] = data
    st.session_state[f"_cache_ts_{key}"] = time()


def _get_cache(key):
    """Retorna dados do cache se válido"""
    if _cache_valid(key):
        return st.session_state.get(f"_cache_{key}")
    return None


def invalidate_cache(key=None):
    """Invalida o cache (chamado após salvar/atualizar)"""
    if key:
        st.session_state.pop(f"_cache_{key}", None)
        st.session_state.pop(f"_cache_ts_{key}", None)
    else:
        # Invalida tudo
        keys_to_remove = [k for k in st.session_state if k.startswith("_cache_")]
        for k in keys_to_remove:
            del st.session_state[k]


# ===== PROPOSTAS =====

def load_propostas():
    """Carrega todas as propostas (com cache)"""
    cached = _get_cache("propostas")
    if cached is not None:
        return cached

    sp = _get_spreadsheet()
    if not sp:
        return []

    _init_sheets_if_needed(sp)

    try:
        ws = sp.worksheet("Propostas")
        records = ws.get_all_records()
        for r in records:
            try:
                val_str = str(r.get("valor", 0)).strip().replace("R$", "").replace(" ", "")
                # Formato brasileiro: 5.850,00 (ponto=milhar, vírgula=decimal)
                if "," in val_str and "." in val_str:
                    val_str = val_str.replace(".", "").replace(",", ".")
                elif "," in val_str:
                    val_str = val_str.replace(",", ".")
                r["valor"] = float(val_str) if val_str else 0.0
            except:
                r["valor"] = 0.0
            try:
                r["id"] = int(r.get("id", 0))
            except:
                r["id"] = 0
        _set_cache("propostas", records)
        return records
    except Exception:
        return []


def save_proposta(proposta):
    """Adiciona uma nova proposta ao Google Sheets"""
    sp = _get_spreadsheet()
    if not sp:
        return False

    _init_sheets_if_needed(sp)

    try:
        ws = sp.worksheet("Propostas")
        row = [proposta.get(h, "") for h in HEADERS_PROPOSTAS]
        ws.insert_rows([row], row=2)
        invalidate_cache("propostas")
        return True
    except Exception:
        return False


def update_proposta_status(proposta_id, novo_status, motivo="", historico_anterior=""):
    """Atualiza o status de uma proposta, com motivo de perda e histórico"""
    sp = _get_spreadsheet()
    if not sp:
        return False, "Não foi possível conectar ao Google Sheets"

    try:
        # Garantir que os headers estejam atualizados antes de atualizar
        _init_sheets_if_needed(sp)

        ws = sp.worksheet("Propostas")
        cell = ws.find(str(proposta_id), in_column=1)
        if not cell:
            return False, f"Proposta ID {proposta_id} não encontrada na planilha"

        status_col = HEADERS_PROPOSTAS.index("status") + 1
        ws.update_cell(cell.row, status_col, novo_status)

        # Salvar motivo de perda se status for "Não Fechou"
        if novo_status == "Não Fechou" and motivo:
            motivo_col = HEADERS_PROPOSTAS.index("motivo_perda") + 1
            ws.update_cell(cell.row, motivo_col, motivo)

            # Adicionar ao histórico
            from datetime import datetime
            data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
            nova_entrada = f"[{data_hora}] {motivo}"
            if historico_anterior:
                historico_completo = f"{historico_anterior} | {nova_entrada}"
            else:
                historico_completo = nova_entrada
            hist_col = HEADERS_PROPOSTAS.index("historico") + 1
            ws.update_cell(cell.row, hist_col, historico_completo)

        invalidate_cache("propostas")
        return True, "OK"
    except Exception as e:
        return False, f"Erro ao atualizar: {str(e)}"


def update_servicos_detalhados(proposta_id, servicos_detalhados_json):
    """Atualiza os serviços detalhados (status individual por item) de uma proposta.
    Automaticamente registra data_aprovacao quando item muda para Aprovado."""
    from datetime import datetime

    sp = _get_spreadsheet()
    if not sp:
        return False, "Não foi possível conectar ao Google Sheets"

    try:
        _init_sheets_if_needed(sp)
        ws = sp.worksheet("Propostas")
        cell = ws.find(str(proposta_id), in_column=1)
        if not cell:
            return False, f"Proposta ID {proposta_id} não encontrada"

        # Carregar serviços atuais para comparar e registrar data_aprovacao
        try:
            servicos = json.loads(servicos_detalhados_json)
            hoje = datetime.now().strftime("%Y-%m-%d")
            for s in servicos:
                if s.get("status") == "Aprovado" and not s.get("data_aprovacao"):
                    s["data_aprovacao"] = hoje
                elif s.get("status") != "Aprovado":
                    s.pop("data_aprovacao", None)
            servicos_detalhados_json = json.dumps(servicos, ensure_ascii=False)
        except Exception:
            pass

        sd_col = HEADERS_PROPOSTAS.index("servicos_detalhados") + 1
        ws.update_cell(cell.row, sd_col, servicos_detalhados_json)

        # Recalcular valor aprovado e atualizar coluna valor
        try:
            servicos = json.loads(servicos_detalhados_json)
            valor_aprovado = sum(
                s.get("valor", 0) for s in servicos if s.get("status") == "Aprovado"
            )
            valor_col = HEADERS_PROPOSTAS.index("valor") + 1
            ws.update_cell(cell.row, valor_col, valor_aprovado)
        except Exception:
            pass

        # Derivar status geral da proposta baseado nos itens
        try:
            servicos = json.loads(servicos_detalhados_json)
            statuses = [s.get("status", "Pendente") for s in servicos]
            if all(s == "Aprovado" for s in statuses):
                novo_status = "Fechou"
            elif all(s in ("Recusado", "Expirado") for s in statuses):
                novo_status = "Não Fechou"
            elif any(s == "Aprovado" for s in statuses):
                novo_status = "Fechou Parcial"
            else:
                novo_status = "Enviada"
            status_col = HEADERS_PROPOSTAS.index("status") + 1
            ws.update_cell(cell.row, status_col, novo_status)
        except Exception:
            pass

        invalidate_cache("propostas")
        return True, "OK"
    except Exception as e:
        return False, f"Erro ao atualizar serviços: {str(e)}"


def expirar_itens_pendentes(dias=30):
    """Marca como 'Expirado' itens pendentes de propostas com mais de X dias"""
    from datetime import datetime, timedelta

    propostas = load_propostas()
    if not propostas:
        return 0

    hoje = datetime.now().date()
    limite = hoje - timedelta(days=dias)
    count = 0

    for p in propostas:
        data_str = p.get("data", "")
        sd_str = p.get("servicos_detalhados", "")
        if not data_str or not sd_str:
            continue

        try:
            data_proposta = datetime.strptime(data_str, "%Y-%m-%d").date()
        except Exception:
            continue

        if data_proposta > limite:
            continue

        try:
            servicos = json.loads(sd_str)
        except Exception:
            continue

        alterou = False
        for s in servicos:
            if s.get("status") == "Pendente":
                s["status"] = "Expirado"
                alterou = True

        if alterou:
            ok, _ = update_servicos_detalhados(p["id"], json.dumps(servicos, ensure_ascii=False))
            if ok:
                count += 1

    return count


def delete_proposta(proposta_id):
    """Exclui uma proposta"""
    sp = _get_spreadsheet()
    if not sp:
        return False

    try:
        ws = sp.worksheet("Propostas")
        cell = ws.find(str(proposta_id), in_column=1)
        if cell:
            ws.delete_rows(cell.row)
            invalidate_cache("propostas")
            return True
    except Exception:
        pass
    return False


# ===== CONFIG =====

def load_config_sheets():
    """Carrega configurações (com cache)"""
    default = {"meta_mensal": 12000, "vendedores": ["Allan"]}

    cached = _get_cache("config")
    if cached is not None:
        return cached

    sp = _get_spreadsheet()
    if not sp:
        return default

    _init_sheets_if_needed(sp)

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
            elif chave == "vendedores_fotos":
                try:
                    config["vendedores_fotos"] = json.loads(valor)
                except:
                    config["vendedores_fotos"] = {}
            elif chave.startswith("foto_"):
                nome_vendedor = chave[5:]
                if "vendedores_fotos" not in config:
                    config["vendedores_fotos"] = {}
                config["vendedores_fotos"][nome_vendedor] = valor
            else:
                config[chave] = valor

        for k, v in default.items():
            if k not in config:
                config[k] = v

        _set_cache("config", config)
        return config
    except Exception:
        return default


def save_config_sheets(config):
    """Salva configurações no Google Sheets"""
    sp = _get_spreadsheet()
    if not sp:
        return False

    try:
        ws = sp.worksheet("Config")
        ws.clear()
        ws.append_row(HEADERS_CONFIG)
        ws.append_row(["meta_mensal", str(config.get("meta_mensal", 12000))])
        ws.append_row(["vendedores", json.dumps(config.get("vendedores", ["Allan"]))])

        # Salvar fotos dos vendedores (cada foto em uma linha separada)
        fotos = config.get("vendedores_fotos", {})
        if fotos:
            for nome, foto_b64 in fotos.items():
                if foto_b64:
                    ws.append_row([f"foto_{nome}", foto_b64])

        invalidate_cache("config")
        return True
    except Exception:
        return False


# ===== VERIFICAÇÃO =====

def sheets_disponivel():
    """Verifica se o Google Sheets está configurado e acessível"""
    sp = _get_spreadsheet()
    return sp is not None
