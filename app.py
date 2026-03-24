"""
Sistema de Propostas Comerciais - Azevedo Contabilidade
Rode com: streamlit run app.py
"""

import streamlit as st
import json
import os
from datetime import datetime, date
from gerar_proposta import gerar_docx
from sheets_db import (
    sheets_disponivel, load_propostas, save_proposta,
    update_proposta_status, delete_proposta,
    load_config_sheets, save_config_sheets,
    update_servicos_detalhados, expirar_itens_pendentes
)

# ===== CONFIG =====
st.set_page_config(
    page_title="Propostas - Azevedo Contabilidade",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "logo.png")
USING_SHEETS = sheets_disponivel()

# ===== DATABASE (fallback local + Google Sheets) =====
DB_FILE = os.path.join(os.path.dirname(__file__), "propostas_db.json")
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

def load_db():
    if USING_SHEETS:
        return load_propostas()
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_db(data):
    if not USING_SHEETS:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

def load_config():
    if USING_SHEETS:
        return load_config_sheets()
    default = {"meta_mensal": 12000, "vendedores": ["Allan"]}
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            for k, v in default.items():
                if k not in cfg:
                    cfg[k] = v
            return cfg
    save_config(default)
    return default

def save_config(cfg):
    if USING_SHEETS:
        save_config_sheets(cfg)
    else:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

def fc(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def get_avatar_html(nome, fotos_dict=None):
    """Gera HTML de avatar: foto se disponível, ou iniciais coloridas."""
    if fotos_dict and nome in fotos_dict and fotos_dict[nome]:
        return f'<img src="{fotos_dict[nome]}" class="ranking-avatar" alt="{nome}">'
    # Gerar cor baseada no nome
    cores = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#f97316']
    cor = cores[sum(ord(c) for c in nome) % len(cores)]
    iniciais = "".join(p[0].upper() for p in nome.split()[:2]) if nome else "?"
    return f'<div class="ranking-initials" style="background:{cor};">{iniciais}</div>'

# ===== CUSTOM CSS =====
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    .main .block-container { padding-top: 1rem; max-width: 1200px; }

    .stTabs [data-baseweb="tab-list"] {
        gap: 0;
        background: #f8f9fa;
        border-radius: 12px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 24px;
        font-weight: 600;
        border-radius: 8px;
        font-size: 14px;
    }
    .stTabs [aria-selected="true"] {
        background: white !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }

    .header-bar {
        background: linear-gradient(135deg, #1a2744 0%, #2a3a5a 100%);
        border-radius: 16px;
        padding: 24px 32px;
        margin-bottom: 24px;
        display: flex;
        align-items: center;
        gap: 20px;
    }
    .header-bar h1 {
        color: white;
        font-size: 22px;
        font-weight: 700;
        margin: 0;
        font-family: 'Inter', sans-serif;
    }
    .header-bar p {
        color: #b8960c;
        font-size: 13px;
        margin: 4px 0 0;
        font-weight: 500;
    }

    .metric-card {
        background: white;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
        border: 1px solid #f0f0f0;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }
    .metric-value {
        font-size: 28px;
        font-weight: 800;
        color: #1a2744;
        font-family: 'Inter', sans-serif;
        line-height: 1.2;
    }
    .metric-label {
        font-size: 11px;
        color: #9ca3af;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 600;
        margin-bottom: 8px;
    }
    .metric-icon {
        font-size: 20px;
        margin-bottom: 8px;
    }

    .meta-card {
        background: linear-gradient(135deg, #faf6e6 0%, #fff8e1 100%);
        border-radius: 16px;
        padding: 24px;
        border: 2px solid #b8960c33;
    }
    .meta-title {
        font-size: 14px;
        font-weight: 700;
        color: #b8960c;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 16px;
    }
    .meta-valor {
        font-size: 36px;
        font-weight: 800;
        color: #1a2744;
        font-family: 'Inter', sans-serif;
    }
    .meta-sub {
        font-size: 13px;
        color: #6b7280;
    }

    .progress-container {
        background: #e5e7eb;
        border-radius: 100px;
        height: 14px;
        overflow: hidden;
        margin: 12px 0;
    }
    .progress-bar {
        height: 100%;
        border-radius: 100px;
        transition: width 0.5s ease;
    }
    .progress-green { background: linear-gradient(90deg, #10b981, #34d399); }
    .progress-gold { background: linear-gradient(90deg, #b8960c, #d4af37); }
    .progress-red { background: linear-gradient(90deg, #ef4444, #f87171); }

    .ranking-row {
        display: flex;
        align-items: center;
        padding: 12px 16px;
        background: white;
        border-radius: 10px;
        margin-bottom: 8px;
        border: 1px solid #f0f0f0;
    }
    .ranking-pos {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 800;
        font-size: 14px;
        margin-right: 14px;
    }
    .ranking-1 { background: #fef3c7; color: #b8960c; }
    .ranking-2 { background: #e5e7eb; color: #6b7280; }
    .ranking-3 { background: #fde8d0; color: #c2703e; }
    .ranking-other { background: #f3f4f6; color: #9ca3af; }
    .ranking-avatar {
        width: 36px;
        height: 36px;
        border-radius: 50%;
        margin-right: 12px;
        object-fit: cover;
        border: 2px solid #e5e7eb;
    }
    .ranking-initials {
        width: 36px;
        height: 36px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 14px;
        color: white;
        margin-right: 12px;
        flex-shrink: 0;
    }
    .ranking-name {
        flex: 1;
        font-weight: 600;
        color: #1a2744;
        font-size: 14px;
    }
    .ranking-stats {
        text-align: right;
        font-size: 12px;
        color: #6b7280;
    }
    .ranking-valor {
        font-weight: 700;
        color: #1a2744;
        font-size: 15px;
    }

    .status-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 100px;
        font-size: 12px;
        font-weight: 600;
    }
    .status-enviada { background: #dbeafe; color: #2563eb; }
    .status-fechou { background: #d1fae5; color: #059669; }
    .status-fechou-parcial { background: #e0f2fe; color: #0284c7; }
    .status-nao-fechou { background: #fee2e2; color: #dc2626; }
    .status-pendente { background: #fef3c7; color: #d97706; }
    .status-expirado { background: #f3f4f6; color: #6b7280; }

    .svc-item {
        display: flex;
        align-items: center;
        padding: 10px 14px;
        background: #f9fafb;
        border-radius: 8px;
        margin-bottom: 6px;
        border: 1px solid #e5e7eb;
        gap: 12px;
    }
    .svc-item-desc { flex: 2; font-size: 13px; font-weight: 500; color: #374151; }
    .svc-item-valor { flex: 1; font-size: 13px; font-weight: 700; color: #1a2744; text-align: right; }
    .svc-item-per { flex: 1; font-size: 12px; color: #6b7280; text-align: center; }
    .svc-status-aprovado { color: #059669; font-weight: 700; }
    .svc-status-recusado { color: #dc2626; font-weight: 700; }
    .svc-status-pendente { color: #d97706; font-weight: 700; }
    .svc-status-expirado { color: #6b7280; font-weight: 700; font-style: italic; }

    .section-title {
        font-size: 13px;
        font-weight: 700;
        color: #b8960c;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        margin: 24px 0 16px;
        padding-bottom: 8px;
        border-bottom: 2px solid #f0f0f0;
    }

    div[data-testid="stForm"] {
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 24px;
        background: white;
    }
</style>
""", unsafe_allow_html=True)

# ===== HEADER =====
st.markdown("""
<div class="header-bar">
    <div>
        <h1>📄 Sistema de Propostas Comerciais</h1>
        <p>Azevedo Contabilidade — Contabilidade Estratégica & Planejamento Tributário</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ===== TABS =====
tab_dash, tab_nova, tab_hist, tab_config = st.tabs(["📊 Dashboard", "📄 Nova Proposta", "📋 Histórico", "⚙️ Configurações"])

# ==========================================
# TAB: DASHBOARD
# ==========================================
with tab_dash:
    db = load_db()
    config = load_config()
    meta_mensal = config.get("meta_mensal", 50000)

    # ===== HELPER: extrair itens aprovados com data_aprovacao =====
    def _parse_data(data_str):
        """Converte data para formato YYYY-MM-DD se necessário"""
        if not data_str:
            return ""
        data_str = str(data_str).strip()
        # Já no formato correto
        if len(data_str) >= 10 and data_str[4] == "-":
            return data_str[:10]
        # Formato DD/MM/YYYY
        if "/" in data_str:
            parts = data_str.split("/")
            if len(parts) == 3 and len(parts[2]) == 4:
                return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
        return data_str

    def extrair_itens_aprovados(propostas):
        """Retorna lista de dicts com: vendedor, valor, data_aprovacao, descricao"""
        itens = []
        for p in propostas:
            sd = p.get("servicos_detalhados", "")
            vendedor = p.get("vendedor", "Sem vendedor") or "Sem vendedor"
            data_proposta = _parse_data(p.get("data", ""))

            if sd and sd.strip().startswith("["):
                try:
                    svcs = json.loads(sd)
                    for s in svcs:
                        if s.get("status") == "Aprovado":
                            itens.append({
                                "vendedor": vendedor,
                                "valor": s.get("valor", 0),
                                "data_aprovacao": _parse_data(s.get("data_aprovacao", "")) or data_proposta,
                                "descricao": s.get("descricao", ""),
                                "cliente": p.get("cliente", "-"),
                                "periodicidade": s.get("periodicidade", "-")
                            })
                except Exception:
                    # Se o JSON der erro, trata como proposta antiga
                    if p.get("status") in ("Fechou", "Fechou Parcial") and p.get("valor", 0) > 0:
                        itens.append({
                            "vendedor": vendedor,
                            "valor": p.get("valor", 0),
                            "data_aprovacao": data_proposta,
                            "descricao": p.get("servicos", "-"),
                            "cliente": p.get("cliente", "-"),
                            "periodicidade": "-"
                        })
            else:
                # Proposta antiga sem servicos_detalhados — valor integral
                if p.get("status") in ("Fechou", "Fechou Parcial") and p.get("valor", 0) > 0:
                    itens.append({
                        "vendedor": vendedor,
                        "valor": p.get("valor", 0),
                        "data_aprovacao": data_proposta,
                        "descricao": p.get("servicos", "-"),
                        "cliente": p.get("cliente", "-"),
                        "periodicidade": "-"
                    })
        return itens

    todos_itens_aprov = extrair_itens_aprovados(db)

    # ===== SELETOR DE MÊS =====
    st.markdown('<div class="section-title">📅 Período</div>', unsafe_allow_html=True)
    col_sel_mes, col_sel_ano, _ = st.columns([1, 1, 3])
    meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                   "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    with col_sel_mes:
        mes_sel = st.selectbox("Mês", range(1, 13), index=date.today().month - 1,
                               format_func=lambda x: meses_nomes[x-1], key="dash_mes")
    with col_sel_ano:
        ano_sel = st.number_input("Ano", min_value=2024, max_value=2030, value=date.today().year, key="dash_ano")

    prefixo_mes = f"{ano_sel}-{mes_sel:02d}"

    # Filtrar itens aprovados pelo mês da APROVAÇÃO (não da proposta)
    itens_mes = [i for i in todos_itens_aprov if i["data_aprovacao"].startswith(prefixo_mes)]
    receita_mes = sum(i["valor"] for i in itens_mes)

    # Propostas criadas no mês selecionado (para contagem)
    db_mes = [p for p in db if _parse_data(p.get("data", "")).startswith(prefixo_mes)]
    total_mes = len(db_mes)

    # Métricas gerais (todas as propostas, não filtrado por mês)
    total = len(db)
    receita_total = sum(i["valor"] for i in todos_itens_aprov)
    propostas_com_aprovacao = set()
    for i in todos_itens_aprov:
        propostas_com_aprovacao.add(i["cliente"] + i.get("data_aprovacao", ""))
    fechou_count = sum(1 for p in db if p.get("status") in ("Fechou", "Fechou Parcial"))
    taxa = (fechou_count / total * 100) if total > 0 else 0

    # ===== MÉTRICAS GERAIS =====
    st.markdown('<div class="section-title">Visão Geral</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)

    cards = [
        ("📋", "Total Propostas", str(total), ""),
        ("✅", "Fechamentos", str(fechou_count), ""),
        ("📈", "Taxa Conversão", f"{taxa:.1f}%", ""),
        ("💰", "Receita Aprovada", fc(receita_total), "")
    ]
    for col, (icon, label, value, _) in zip([c1, c2, c3, c4], cards):
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-icon">{icon}</div>
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("")

    # ===== META + MÊS SELECIONADO =====
    col_meta, col_mes_info = st.columns([3, 2])

    with col_meta:
        st.markdown(f'<div class="section-title">🎯 Meta de Vendas — {meses_nomes[mes_sel-1]} {ano_sel}</div>', unsafe_allow_html=True)
        progresso = (receita_mes / meta_mensal * 100) if meta_mensal > 0 else 0
        progresso_cap = min(progresso, 100)
        falta = max(meta_mensal - receita_mes, 0)

        cache_key = f"meta_celebrada_{prefixo_mes}"
        if progresso < 100 and cache_key in st.session_state:
            del st.session_state[cache_key]

        if progresso >= 100:
            bar_class = "progress-green"
            emoji_status = "🏆"
            txt_status = "META BATIDA!"
        elif progresso >= 60:
            bar_class = "progress-gold"
            emoji_status = "🔥"
            txt_status = "Bom ritmo!"
        else:
            bar_class = "progress-red"
            emoji_status = "⚡"
            txt_status = "Vamos acelerar!"

        itens_aprov_mes = len(itens_mes)
        st.markdown(f"""
        <div class="meta-card">
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                <div>
                    <div class="meta-title">Progresso da Meta</div>
                    <div class="meta-valor">{fc(receita_mes)}</div>
                    <div class="meta-sub">de {fc(meta_mensal)}</div>
                </div>
                <div style="text-align:right;">
                    <div style="font-size:40px;">{emoji_status}</div>
                    <div style="font-weight:700; color:#1a2744; font-size:22px;">{progresso:.0f}%</div>
                    <div style="font-size:12px; color:#6b7280;">{txt_status}</div>
                </div>
            </div>
            <div class="progress-container">
                <div class="progress-bar {bar_class}" style="width:{progresso_cap}%;"></div>
            </div>
            <div style="display:flex; justify-content:space-between; font-size:12px; color:#6b7280;">
                <span>{total_mes} propostas criadas | {itens_aprov_mes} itens aprovados</span>
                <span>Falta: <strong style="color:#1a2744;">{fc(falta)}</strong></span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if progresso >= 100 and cache_key not in st.session_state:
            st.session_state[cache_key] = True
            st.balloons()

    with col_mes_info:
        st.markdown(f'<div class="section-title">📅 {meses_nomes[mes_sel-1]} {ano_sel}</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="metric-card" style="margin-bottom:12px;">
            <div class="metric-label">Propostas Criadas</div>
            <div class="metric-value">{total_mes}</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="metric-card" style="margin-bottom:12px;">
            <div class="metric-label">Itens Aprovados</div>
            <div class="metric-value" style="color:#10b981;">{itens_aprov_mes}</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Receita Aprovada</div>
            <div class="metric-value" style="color:#b8960c;">{fc(receita_mes)}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("")

    # ===== RANKING DE VENDEDORES (baseado na data_aprovacao do mês) =====
    col_rank, col_status = st.columns([3, 2])

    with col_rank:
        st.markdown(f'<div class="section-title">🏅 Ranking de Vendedores — {meses_nomes[mes_sel-1]} {ano_sel}</div>', unsafe_allow_html=True)
        fotos_vendedores = config.get("vendedores_fotos", {})
        vendedores_stats = {}
        for item in itens_mes:
            v = item["vendedor"]
            if v not in vendedores_stats:
                vendedores_stats[v] = {"receita": 0, "itens": 0}
            vendedores_stats[v]["receita"] += item["valor"]
            vendedores_stats[v]["itens"] += 1

        # Contagem de propostas por vendedor no mês
        propostas_por_vendedor = {}
        for p in db_mes:
            v = p.get("vendedor", "Sem vendedor") or "Sem vendedor"
            propostas_por_vendedor[v] = propostas_por_vendedor.get(v, 0) + 1

        ranking = sorted(vendedores_stats.items(), key=lambda x: x[1]["receita"], reverse=True)

        if ranking:
            for pos, (nome, stats) in enumerate(ranking[:5], 1):
                pos_class = f"ranking-{pos}" if pos <= 3 else "ranking-other"
                medal = {1: "🥇", 2: "🥈", 3: "🥉"}.get(pos, f"{pos}º")
                n_propostas = propostas_por_vendedor.get(nome, 0)
                avatar = get_avatar_html(nome, fotos_vendedores)
                st.markdown(f"""
                <div class="ranking-row">
                    <div class="ranking-pos {pos_class}">{medal if pos <= 3 else pos}</div>
                    {avatar}
                    <div class="ranking-name">{nome}</div>
                    <div class="ranking-stats">
                        {stats['itens']} item(ns) aprovado(s) · {n_propostas} proposta(s)<br>
                        <span class="ranking-valor">{fc(stats['receita'])}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("Nenhum item aprovado neste mês.")

    with col_status:
        st.markdown('<div class="section-title">📊 Status das Propostas</div>', unsafe_allow_html=True)
        counts = {"Enviada": 0, "Fechou": 0, "Fechou Parcial": 0, "Não Fechou": 0, "Pendente": 0}
        for p in db:
            s = p.get("status", "Enviada")
            if s in counts:
                counts[s] += 1

        colors_map = {"Enviada": "#3b82f6", "Fechou": "#10b981", "Fechou Parcial": "#0284c7", "Não Fechou": "#ef4444", "Pendente": "#f59e0b"}

        for status_name, count in counts.items():
            pct = (count / total * 100) if total > 0 else 0
            color = colors_map.get(status_name, "#6b7280")
            st.markdown(f"""
            <div style="margin-bottom:14px;">
                <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
                    <span style="font-weight:600; color:#374151; font-size:13px;">{status_name}</span>
                    <span style="font-weight:700; color:{color}; font-size:13px;">{count}</span>
                </div>
                <div style="background:#e5e7eb; border-radius:100px; height:8px; overflow:hidden;">
                    <div style="height:100%; width:{pct}%; background:{color}; border-radius:100px;"></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    # ===== ÚLTIMAS PROPOSTAS =====
    if db:
        st.markdown('<div class="section-title">🕐 Últimas Propostas</div>', unsafe_allow_html=True)
        for p in db[:5]:
            status = p.get("status", "Enviada")
            badge_class = {"Enviada": "status-enviada", "Fechou": "status-fechou", "Fechou Parcial": "status-fechou-parcial", "Não Fechou": "status-nao-fechou", "Pendente": "status-pendente"}.get(status, "status-enviada")
            st.markdown(f"""
            <div style="display:flex; align-items:center; padding:12px 16px; background:white; border-radius:10px; margin-bottom:8px; border:1px solid #f0f0f0;">
                <div style="flex:1;">
                    <strong style="color:#1a2744;">{p.get('cliente', '-')}</strong>
                    <span style="color:#9ca3af; font-size:12px; margin-left:8px;">{p.get('data', '')}</span>
                </div>
                <div style="margin-right:16px; font-weight:700; color:#1a2744;">{fc(p.get('valor', 0))}</div>
                <span class="status-badge {badge_class}">{status}</span>
            </div>
            """, unsafe_allow_html=True)


# ==========================================
# TAB: NOVA PROPOSTA
# ==========================================
with tab_nova:
    config = load_config()
    vendedores_lista = config.get("vendedores", ["Allan"])

    with st.form("proposta_form"):
        st.markdown("#### 👤 Dados do Cliente")
        col1, col2 = st.columns(2)
        with col1:
            tratamento = st.selectbox("Tratamento", ["Ao Sr.", "À Sra.", "A", "Ao", "Prezado(a)"])
            telefone = st.text_input("Telefone", placeholder="(84) 99999-0000")
        with col2:
            nome_cliente = st.text_input("Nome do Cliente *", placeholder="Ex: João da Silva")
            email_cliente = st.text_input("E-mail", placeholder="cliente@email.com")

        st.markdown("#### 👨‍💼 Vendedor")
        vendedor = st.selectbox("Vendedor responsável", vendedores_lista)

        st.markdown("#### 📝 Descrição da Proposta")
        introducao = st.text_area(
            "Introdução (referente a...)",
            placeholder="Ex: gestão contábil e fiscal da empresa NOME LTDA",
            height=68
        )

        st.markdown("#### 💰 Serviços e Honorários")
        col_desc, col_per = st.columns(2)
        with col_desc:
            desconto_pct = st.selectbox("Desconto fictício", [
                ("10%", 0.10), ("15%", 0.15), ("20%", 0.20), ("Sem desconto", 0.0)
            ], format_func=lambda x: x[0])
        with col_per:
            periodicidade_padrao = st.selectbox("Periodicidade padrão", ["Mensal", "Única Vez", "Trimestral", "Anual"])

        st.markdown("**Adicione os serviços:**")
        num_svcs = st.number_input("Quantidade de serviços", min_value=1, max_value=10, value=1)

        servicos = []
        for i in range(int(num_svcs)):
            st.markdown(f"**Serviço {i+1}**")
            cs1, cs2, cs3 = st.columns([3, 1, 1])
            with cs1:
                desc_svc = st.text_input(f"Descrição", key=f"svc_desc_{i}", placeholder="Ex: Gestão contábil mensal")
            with cs2:
                valor_svc = st.text_input(f"Valor Real (R$)", key=f"svc_val_{i}", placeholder="500,00")
            with cs3:
                per_svc = st.selectbox(f"Periodicidade", ["Mensal", "Única Vez", "Trimestral", "Anual"], key=f"svc_per_{i}",
                                       index=["Mensal", "Única Vez", "Trimestral", "Anual"].index(periodicidade_padrao))
            servicos.append({"descricao": desc_svc, "valor": valor_svc, "periodicidade": per_svc})

        st.markdown("#### 💳 Pagamento e Extras")
        pix_opcao = st.selectbox("Chave PIX para pagamento", [
            ("PF — 33.540.066/0001-23 (Allan Sayure)", "33.540.066/0001-23", "ALLAN SAYURE DE AZEVEDO BARBOSA"),
            ("PJ — 35.304.872/0001-28 (Azevedo Contabilidade)", "35.304.872/0001-28", "AZEVEDO CONTABILIDADE LTDA")
        ], format_func=lambda x: x[0])

        observacao = st.text_area("Observação (opcional)", placeholder="Ex: O valor refere-se exclusivamente a...", height=68)

        incluir_doc = st.checkbox("Incluir seção de Documentação Necessária")
        texto_doc = ""
        if incluir_doc:
            texto_doc = st.text_area(
                "Texto da documentação",
                value="Cópia da identidade do responsável de cada empresa para fazermos a procuração necessária e assim enviamos para assinatura pelo Gov.br ou certificado digital.",
                height=68
            )

        obs_internas = st.text_area("📌 Observações internas (não aparece na proposta)", placeholder="Anotações internas...", height=68)

        submitted = st.form_submit_button("📄 Gerar Proposta DOCX", use_container_width=True, type="primary")

    if submitted:
        if not nome_cliente.strip():
            st.error("Preencha o nome do cliente!")
        elif not any(s["descricao"].strip() for s in servicos):
            st.error("Adicione pelo menos um serviço com descrição!")
        else:
            def parse_valor(v):
                v = v.replace(".", "").replace(",", ".").strip()
                try:
                    return float(v)
                except:
                    return 0.0

            svcs_parsed = []
            for s in servicos:
                if s["descricao"].strip():
                    svcs_parsed.append({
                        "descricao": s["descricao"],
                        "valor": parse_valor(s["valor"]),
                        "periodicidade": s["periodicidade"]
                    })

            dados = {
                "tratamento": tratamento,
                "nome": nome_cliente,
                "telefone": telefone,
                "email": email_cliente,
                "vendedor": vendedor,
                "introducao": introducao or "prestação de serviços contábeis",
                "servicos": svcs_parsed,
                "desconto_pct": desconto_pct[1],
                "pix_cnpj": pix_opcao[1],
                "pix_titular": pix_opcao[2],
                "observacao": observacao,
                "incluir_doc": incluir_doc,
                "texto_doc": texto_doc,
                "logo_path": LOGO_PATH
            }

            with st.spinner("Gerando proposta..."):
                docx_bytes = gerar_docx(dados)

            filename = f"Proposta - {nome_cliente}.docx"
            st.success(f"✅ Proposta gerada com sucesso!")
            st.download_button(
                label="⬇️ Baixar Proposta DOCX",
                data=docx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

            total_valor = sum(s["valor"] for s in svcs_parsed)

            # Serviços detalhados com status individual (Pendente por padrão)
            svcs_detalhados = []
            for s in svcs_parsed:
                svcs_detalhados.append({
                    "descricao": s["descricao"],
                    "valor": s["valor"],
                    "periodicidade": s["periodicidade"],
                    "status": "Pendente"
                })

            nova_proposta = {
                "id": int(datetime.now().timestamp() * 1000),
                "data": datetime.now().strftime("%Y-%m-%d"),
                "cliente": nome_cliente,
                "tratamento": tratamento,
                "telefone": telefone,
                "email": email_cliente,
                "vendedor": vendedor,
                "servicos": "; ".join(s["descricao"] for s in svcs_parsed),
                "valor": total_valor,
                "status": "Enviada",
                "obs": obs_internas,
                "servicos_detalhados": json.dumps(svcs_detalhados, ensure_ascii=False)
            }

            if USING_SHEETS:
                save_proposta(nova_proposta)
            else:
                db = load_db()
                db.insert(0, nova_proposta)
                save_db(db)


# ==========================================
# TAB: HISTÓRICO
# ==========================================
with tab_hist:
    db = load_db()

    # Executar expiração automática de itens pendentes > 30 dias
    if not st.session_state.get("_expirado_check_done"):
        if USING_SHEETS:
            expirar_itens_pendentes(30)
        st.session_state["_expirado_check_done"] = True

    col_search, col_filter, col_export = st.columns([3, 2, 1])
    with col_search:
        busca = st.text_input("🔍 Buscar cliente ou serviço", placeholder="Digite para filtrar...")
    with col_filter:
        filtro_status = st.selectbox("Filtrar por status", ["Todos", "Enviada", "Fechou", "Fechou Parcial", "Não Fechou", "Pendente"])
    with col_export:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📥 Exportar CSV", use_container_width=True):
            if db:
                import csv
                import io
                output = io.StringIO()
                writer = csv.writer(output, delimiter=";")
                writer.writerow(["Data", "Cliente", "Serviços", "Valor Aprovado", "Valor Total", "Vendedor", "Status", "Observações"])
                for p in db:
                    # Calcular valor aprovado dos itens
                    val_aprovado = 0
                    val_total = p.get("valor", 0)
                    sd = p.get("servicos_detalhados", "")
                    if sd and sd.strip().startswith("["):
                        try:
                            svcs = json.loads(sd)
                            val_aprovado = sum(s.get("valor", 0) for s in svcs if s.get("status") == "Aprovado")
                            val_total = sum(s.get("valor", 0) for s in svcs)
                        except Exception:
                            val_aprovado = val_total if p.get("status") in ("Fechou", "Fechou Parcial") else 0
                    else:
                        # Proposta antiga — valor integral se fechou
                        if p.get("status") in ("Fechou", "Fechou Parcial"):
                            val_aprovado = val_total
                    writer.writerow([
                        p.get("data", ""), p.get("cliente", ""), p.get("servicos", ""),
                        str(val_aprovado).replace(".", ","),
                        str(val_total).replace(".", ","),
                        p.get("vendedor", ""), p.get("status", ""), p.get("obs", "")
                    ])
                csv_data = "\ufeff" + output.getvalue()
                st.download_button("⬇️ Baixar CSV", csv_data.encode("utf-8"), "propostas.csv", "text/csv")

    filtered = db
    if busca:
        busca_lower = busca.lower()
        filtered = [p for p in filtered if busca_lower in p.get("cliente", "").lower() or busca_lower in p.get("servicos", "").lower()]
    if filtro_status != "Todos":
        filtered = [p for p in filtered if p.get("status") == filtro_status]

    if not filtered:
        st.info("Nenhuma proposta encontrada.")
    else:
        for i, p in enumerate(filtered):
            status = p.get("status", "Enviada")
            status_color = {"Enviada": "🔵", "Fechou": "🟢", "Fechou Parcial": "🔵🟢", "Não Fechou": "🔴", "Pendente": "🟡"}.get(status, "⚪")

            # Calcular valor aprovado para exibição
            sd_str = p.get("servicos_detalhados", "")
            valor_aprovado = 0
            valor_total = p.get("valor", 0)
            if sd_str and sd_str.strip().startswith("["):
                try:
                    svcs_det = json.loads(sd_str)
                    valor_aprovado = sum(s.get("valor", 0) for s in svcs_det if s.get("status") == "Aprovado")
                    valor_total_calc = sum(s.get("valor", 0) for s in svcs_det)
                    if valor_total_calc > 0:
                        valor_total = valor_total_calc
                except Exception:
                    valor_aprovado = valor_total if status in ("Fechou", "Fechou Parcial") else 0
            else:
                # Proposta antiga — valor integral se fechou
                if status in ("Fechou", "Fechou Parcial"):
                    valor_aprovado = valor_total

            valor_display = fc(valor_aprovado) if valor_aprovado > 0 else fc(valor_total)
            label_extra = f" (aprovado: {fc(valor_aprovado)})" if valor_aprovado > 0 and valor_aprovado != valor_total else ""

            with st.expander(f"{status_color} **{p['cliente']}** — {p.get('data', '')} — {fc(valor_total)}{label_extra} — {status}"):
                c1, c2 = st.columns(2)
                with c1:
                    st.write(f"**Vendedor:** {p.get('vendedor', '-')}")
                    st.write(f"**Telefone:** {p.get('telefone', '-')}")
                with c2:
                    st.write(f"**E-mail:** {p.get('email', '-')}")
                    st.write(f"**Data:** {p.get('data', '-')}")

                # ===== SERVIÇOS COM STATUS INDIVIDUAL =====
                st.markdown("---")
                st.markdown("#### 📦 Serviços da Proposta")

                if sd_str:
                    try:
                        svcs_det = json.loads(sd_str)
                    except Exception:
                        svcs_det = []
                else:
                    # Proposta antiga sem servicos_detalhados — criar a partir do campo servicos
                    svcs_det = []
                    servicos_texto = p.get("servicos", "")
                    if servicos_texto:
                        for desc in servicos_texto.split(";"):
                            desc = desc.strip()
                            if desc:
                                svcs_det.append({
                                    "descricao": desc,
                                    "valor": 0,
                                    "periodicidade": "-",
                                    "status": "Pendente" if status == "Enviada" else ("Aprovado" if status == "Fechou" else "Recusado")
                                })

                if svcs_det:
                    alterou_servicos = False
                    for si, svc in enumerate(svcs_det):
                        svc_status = svc.get("status", "Pendente")

                        col_desc, col_val, col_per, col_st, col_dt = st.columns([2.5, 1.2, 0.8, 1.2, 1.3])
                        with col_desc:
                            st.markdown(f"**{svc.get('descricao', '-')}**")
                        with col_val:
                            st.markdown(f"**{fc(svc.get('valor', 0))}**")
                        with col_per:
                            st.markdown(f"*{svc.get('periodicidade', '-')}*")
                        with col_st:
                            opcoes_status = ["Pendente", "Aprovado", "Recusado"]
                            idx_atual = opcoes_status.index(svc_status) if svc_status in opcoes_status else 0
                            novo_svc_status = st.selectbox(
                                f"Status",
                                opcoes_status,
                                index=idx_atual,
                                key=f"svc_st_{p['id']}_{si}",
                                label_visibility="collapsed"
                            )
                            if novo_svc_status != svc_status:
                                svcs_det[si]["status"] = novo_svc_status
                                alterou_servicos = True
                        with col_dt:
                            if novo_svc_status == "Aprovado" or svc_status == "Aprovado":
                                # Data de aprovação: pré-preenchida com hoje, editável, só passado
                                data_aprov_atual = svc.get("data_aprovacao", "")
                                try:
                                    default_date = datetime.strptime(data_aprov_atual, "%Y-%m-%d").date() if data_aprov_atual else date.today()
                                except Exception:
                                    default_date = date.today()
                                nova_data = st.date_input(
                                    "Aprovado em",
                                    value=default_date,
                                    max_value=date.today(),
                                    key=f"svc_dt_{p['id']}_{si}",
                                    label_visibility="collapsed"
                                )
                                nova_data_str = nova_data.strftime("%Y-%m-%d")
                                if nova_data_str != data_aprov_atual:
                                    svcs_det[si]["data_aprovacao"] = nova_data_str
                                    alterou_servicos = True
                            else:
                                st.markdown("<span style='color:#9ca3af; font-size:11px;'>—</span>", unsafe_allow_html=True)

                    if alterou_servicos:
                        # Registrar data de aprovação nos itens aprovados
                        hoje = datetime.now().strftime("%Y-%m-%d")
                        for s in svcs_det:
                            if s.get("status") == "Aprovado" and not s.get("data_aprovacao"):
                                s["data_aprovacao"] = hoje
                            elif s.get("status") != "Aprovado":
                                s.pop("data_aprovacao", None)

                        novo_json = json.dumps(svcs_det, ensure_ascii=False)
                        if USING_SHEETS:
                            ok, msg = update_servicos_detalhados(p["id"], novo_json)
                            if not ok:
                                st.error(f"Erro ao salvar status do serviço: {msg}")
                        else:
                            for item in db:
                                if item["id"] == p["id"]:
                                    item["servicos_detalhados"] = novo_json
                                    # Recalcular status geral
                                    statuses = [s.get("status", "Pendente") for s in svcs_det]
                                    if all(s == "Aprovado" for s in statuses):
                                        item["status"] = "Fechou"
                                    elif all(s in ("Recusado", "Expirado") for s in statuses):
                                        item["status"] = "Não Fechou"
                                    elif any(s == "Aprovado" for s in statuses):
                                        item["status"] = "Fechou Parcial"
                                    else:
                                        item["status"] = "Enviada"
                                    # Valor = só aprovados
                                    item["valor"] = sum(s.get("valor", 0) for s in svcs_det if s.get("status") == "Aprovado")
                                    break
                            save_db(db)
                        st.rerun()

                    # Resumo visual
                    qtd_aprov = sum(1 for s in svcs_det if s.get("status") == "Aprovado")
                    qtd_recus = sum(1 for s in svcs_det if s.get("status") == "Recusado")
                    qtd_pend = sum(1 for s in svcs_det if s.get("status") == "Pendente")
                    qtd_exp = sum(1 for s in svcs_det if s.get("status") == "Expirado")
                    st.markdown(f"""
                    <div style="display:flex; gap:16px; margin-top:8px; font-size:12px;">
                        <span class="svc-status-aprovado">✅ {qtd_aprov} aprovado(s)</span>
                        <span class="svc-status-recusado">❌ {qtd_recus} recusado(s)</span>
                        <span class="svc-status-pendente">⏳ {qtd_pend} pendente(s)</span>
                        {"<span class='svc-status-expirado'>⌛ " + str(qtd_exp) + " expirado(s)</span>" if qtd_exp > 0 else ""}
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.write(f"**Serviços:** {p.get('servicos', '-')}")

                st.markdown("---")

                # Exibir observações gerais
                if p.get("obs"):
                    st.write(f"**Observações:** {p['obs']}")

                # Exibir motivo da perda e histórico
                if status in ("Não Fechou",):
                    if p.get("motivo_perda"):
                        st.markdown(f"🔴 **Último motivo:** {p['motivo_perda']}")

                    # Botão para adicionar/atualizar motivo
                    if st.session_state.get(f"add_motivo_{p['id']}", False):
                        motivo_add = st.text_area(
                            "Descreva o motivo da perda",
                            placeholder="Ex: Preço alto, foi para concorrente, não respondeu mais...",
                            key=f"motivo_add_{p['id']}"
                        )
                        col_s, col_c = st.columns(2)
                        with col_s:
                            if st.button("✅ Salvar", key=f"salvar_add_{p['id']}"):
                                if motivo_add.strip():
                                    historico_ant = str(p.get("historico", "") or "")
                                    if USING_SHEETS:
                                        ok, msg = update_proposta_status(p["id"], "Não Fechou", motivo_add.strip(), historico_ant)
                                        if not ok:
                                            st.error(f"Erro ao salvar motivo: {msg}")
                                        else:
                                            st.session_state.pop(f"add_motivo_{p['id']}", None)
                                            st.rerun()
                                    else:
                                        for item in db:
                                            if item["id"] == p["id"]:
                                                item["motivo_perda"] = motivo_add.strip()
                                                data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
                                                nova_entrada = f"[{data_hora}] {motivo_add.strip()}"
                                                item["historico"] = f"{historico_ant} | {nova_entrada}" if historico_ant else nova_entrada
                                                break
                                        save_db(db)
                                        st.session_state.pop(f"add_motivo_{p['id']}", None)
                                        st.rerun()
                                else:
                                    st.warning("Preencha o motivo.")
                        with col_c:
                            if st.button("❌ Cancelar", key=f"cancelar_add_{p['id']}"):
                                st.session_state.pop(f"add_motivo_{p['id']}", None)
                                st.rerun()
                    else:
                        if st.button("📝 Adicionar/Atualizar motivo", key=f"btn_add_motivo_{p['id']}"):
                            st.session_state[f"add_motivo_{p['id']}"] = True
                            st.rerun()

                if p.get("historico"):
                    with st.expander("📋 Histórico de observações"):
                        entradas = str(p["historico"]).split(" | ")
                        for entrada in entradas:
                            if entrada.strip():
                                st.markdown(f"- {entrada}")

                if st.button(f"🗑️ Excluir", key=f"del_{p['id']}"):
                    if USING_SHEETS:
                        delete_proposta(p["id"])
                    else:
                        db = [item for item in db if item["id"] != p["id"]]
                        save_db(db)
                    st.rerun()


# ==========================================
# TAB: CONFIGURAÇÕES
# ==========================================
with tab_config:
    config = load_config()

    st.markdown("#### ⚙️ Configurações do Sistema")

    st.markdown('<div class="section-title">🎯 Meta de Vendas</div>', unsafe_allow_html=True)
    nova_meta = st.number_input(
        "Meta mensal de vendas (R$)",
        min_value=0.0,
        value=float(config.get("meta_mensal", 50000)),
        step=1000.0,
        format="%.2f"
    )

    st.markdown('<div class="section-title">👥 Vendedores</div>', unsafe_allow_html=True)
    st.caption("Um vendedor por linha")
    vendedores_text = st.text_area(
        "Lista de vendedores",
        value="\n".join(config.get("vendedores", ["Allan"])),
        height=120
    )

    # 📸 Upload de fotos dos vendedores
    st.markdown('<div class="section-title">📸 Fotos dos Vendedores</div>', unsafe_allow_html=True)
    st.caption("Faça upload da foto de cada vendedor (aparece no ranking)")
    vendedores_atuais = [v.strip() for v in vendedores_text.split("\n") if v.strip()]
    fotos_atuais = config.get("vendedores_fotos", {})

    import base64
    fotos_novas = dict(fotos_atuais)
    for vend in vendedores_atuais:
        col_foto_preview, col_foto_upload = st.columns([1, 3])
        with col_foto_preview:
            avatar_html = get_avatar_html(vend, fotos_atuais)
            st.markdown(f'<div style="display:flex;align-items:center;height:60px;justify-content:center;">{avatar_html}</div>', unsafe_allow_html=True)
        with col_foto_upload:
            foto_file = st.file_uploader(
                f"Foto de {vend}",
                type=["png", "jpg", "jpeg"],
                key=f"foto_{vend}",
                label_visibility="collapsed"
            )
            if foto_file is not None:
                foto_bytes = foto_file.read()
                foto_b64 = base64.b64encode(foto_bytes).decode()
                ext = foto_file.type.split("/")[-1]
                fotos_novas[vend] = f"data:image/{ext};base64,{foto_b64}"

    if st.button("💾 Salvar Configurações", type="primary", use_container_width=True):
        vendedores_lista = [v.strip() for v in vendedores_text.split("\n") if v.strip()]
        config["meta_mensal"] = nova_meta
        config["vendedores"] = vendedores_lista
        config["vendedores_fotos"] = fotos_novas
        save_config(config)
        st.success("✅ Configurações salvas com sucesso!")
        st.rerun()
