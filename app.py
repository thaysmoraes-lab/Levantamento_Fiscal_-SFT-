"""
Configurador de Tributos Protheus — App Web (Streamlit)
Tema visual: TOTVS (roxo escuro + rosa pink)
"""
import streamlit as st
import io
import os
import base64
from datetime import datetime

from processador import gerar_planilha, PERFIL_PERGUNTAS


# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Configurador de Tributos Protheus",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# PALETA TOTVS
# ============================================================
COR_ROSA       = "#D52B7B"
COR_ROSA_HOVER = "#B0205F"
COR_ROXO       = "#352A47"
COR_ROXO_LIGHT = "#4A3B5E"
COR_CINZA_BG   = "#F4F4F4"
COR_CINZA_BORD = "#E0E0E0"
COR_TEXTO      = "#2D2D2D"
COR_TEXTO_SEC  = "#717171"


# ============================================================
# CSS CUSTOMIZADO — tema TOTVS
# ============================================================
def aplicar_tema_totvs():
    css = f"""
    <style>
        html, body, [class*="css"] {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }}

        .main .block-container {{
            padding-top: 1rem;
            padding-bottom: 3rem;
            max-width: 1100px;
        }}

        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        header[data-testid="stHeader"] {{
            background: transparent;
        }}

        /* === HEADER COM GRADIENTE TOTVS === */
        .totvs-header {{
            background: linear-gradient(135deg, {COR_ROXO} 0%, {COR_ROXO_LIGHT} 50%, {COR_ROSA} 100%);
            padding: 2.5rem 3rem;
            border-radius: 16px;
            margin-bottom: 2rem;
            box-shadow: 0 10px 30px rgba(53, 42, 71, 0.25);
            color: white;
        }}
        .totvs-header h1 {{
            color: white !important;
            font-size: 2.4rem !important;
            font-weight: 700 !important;
            margin: 0 0 0.5rem 0 !important;
            letter-spacing: -0.5px;
        }}
        .totvs-header p {{
            color: rgba(255, 255, 255, 0.85) !important;
            font-size: 1.05rem !important;
            margin: 0 !important;
            font-weight: 400;
        }}
        .totvs-header-logo {{
            display: flex;
            align-items: center;
            gap: 1rem;
            margin-bottom: 1rem;
        }}

        /* === CARDS === */
        .secao-card {{
            background: white;
            padding: 2rem;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
            border: 1px solid {COR_CINZA_BORD};
            margin-bottom: 1.5rem;
        }}

        /* === BOTÕES === */
        .stButton > button {{
            background: {COR_ROSA};
            color: white;
            border: none;
            padding: 0.7rem 1.8rem;
            border-radius: 8px;
            font-weight: 600;
            font-size: 0.95rem;
            transition: all 0.2s;
            box-shadow: 0 2px 8px rgba(213, 43, 123, 0.3);
        }}
        .stButton > button:hover {{
            background: {COR_ROSA_HOVER};
            box-shadow: 0 4px 14px rgba(213, 43, 123, 0.4);
            transform: translateY(-1px);
        }}
        .stButton > button:active {{
            transform: translateY(0);
        }}
        .stButton > button[kind="secondary"] {{
            background: white;
            color: {COR_ROXO};
            border: 2px solid {COR_CINZA_BORD};
            box-shadow: none;
        }}
        .stButton > button[kind="secondary"]:hover {{
            border-color: {COR_ROXO};
            background: {COR_CINZA_BG};
            box-shadow: none;
            transform: none;
        }}

        /* === DOWNLOAD BUTTON === */
        .stDownloadButton > button {{
            background: {COR_ROSA};
            color: white;
            border: none;
            padding: 1rem 2rem;
            border-radius: 8px;
            font-weight: 600;
            font-size: 1rem;
            width: 100%;
            box-shadow: 0 4px 14px rgba(213, 43, 123, 0.35);
            transition: all 0.2s;
        }}
        .stDownloadButton > button:hover {{
            background: {COR_ROSA_HOVER};
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(213, 43, 123, 0.5);
        }}

        /* === UPLOADER === */
        [data-testid="stFileUploader"] {{
            background: white;
            border: 2px dashed {COR_ROSA};
            border-radius: 12px;
            padding: 1.5rem;
            transition: all 0.2s;
        }}
        [data-testid="stFileUploader"]:hover {{
            background: #FFF5FA;
            border-color: {COR_ROSA_HOVER};
        }}
        [data-testid="stFileUploader"] section {{
            background: transparent;
        }}
        [data-testid="stFileUploader"] button {{
            background: {COR_ROXO} !important;
            color: white !important;
            border: none !important;
        }}

        /* === CHECKBOX === */
        .stCheckbox label {{
            font-size: 0.95rem;
            color: {COR_TEXTO};
        }}
        .stCheckbox label p {{
            font-weight: 400 !important;
        }}

        /* === ALERTS === */
        .stAlert {{
            border-radius: 10px;
            border-left-width: 4px;
        }}

        /* === SIDEBAR === */
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, {COR_ROXO} 0%, {COR_ROXO_LIGHT} 100%);
        }}
        section[data-testid="stSidebar"] * {{
            color: white !important;
        }}
        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {{
            color: white !important;
        }}
        section[data-testid="stSidebar"] .stButton > button {{
            background: rgba(255, 255, 255, 0.1);
            color: white;
            border: 1px solid rgba(255, 255, 255, 0.2);
            box-shadow: none;
        }}
        section[data-testid="stSidebar"] .stButton > button:hover {{
            background: {COR_ROSA};
            border-color: {COR_ROSA};
        }}
        section[data-testid="stSidebar"] hr {{
            border-color: rgba(255, 255, 255, 0.15);
        }}

        /* === SUBHEADER === */
        .stApp h2, .stApp h3 {{
            color: {COR_ROXO};
            font-weight: 600;
        }}

        /* === EXPANDER === */
        .streamlit-expanderHeader {{
            background: white;
            border-radius: 8px;
            border: 1px solid {COR_CINZA_BORD};
            font-weight: 500;
        }}

        /* === SPINNER === */
        .stSpinner > div > div {{
            border-top-color: {COR_ROSA} !important;
        }}

        [data-testid="stForm"] {{
            background: transparent;
            border: none;
            padding: 0;
        }}

        /* === STEP INDICATOR === */
        .step-item {{
            padding: 0.7rem 1rem;
            margin: 0.4rem 0;
            border-radius: 8px;
            font-size: 0.95rem;
            font-weight: 500;
        }}
        .step-done {{
            background: rgba(34, 197, 94, 0.2);
        }}
        .step-active {{
            background: {COR_ROSA};
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        }}
        .step-pending {{
            background: rgba(255, 255, 255, 0.05);
            color: rgba(255, 255, 255, 0.5) !important;
        }}

        /* === GRUPO DE PERGUNTAS === */
        .grupo-titulo {{
            color: {COR_ROXO};
            font-weight: 600;
            font-size: 1.1rem;
            margin: 1.5rem 0 0.5rem 0;
            padding-left: 0.8rem;
            border-left: 4px solid {COR_ROSA};
        }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


# ============================================================
# LOGO TOTVS
# ============================================================
def renderizar_logo_totvs():
    """
    Renderiza o logo TOTVS. Tenta carregar de data/logo_totvs.png ou .svg.
    Fallback: SVG inline com texto estilizado 'TOTVS'.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidatos = [
        os.path.join(base_dir, 'data', 'logo_totvs.png'),
        os.path.join(base_dir, 'data', 'logo_totvs.svg'),
        os.path.join(base_dir, 'logo_totvs.png'),
        os.path.join(base_dir, 'logo_totvs.svg'),
    ]
    for path in candidatos:
        if os.path.exists(path):
            ext = os.path.splitext(path)[1].lower()
            with open(path, 'rb') as f:
                data = f.read()
            if ext == '.svg':
                svg_text = data.decode('utf-8')
                return f'<div style="height: 50px; display: flex; align-items: center;">{svg_text}</div>'
            else:
                b64 = base64.b64encode(data).decode('ascii')
                return f'<img src="data:image/png;base64,{b64}" style="height: 50px; width: auto;" alt="TOTVS"/>'

    # Fallback: logo TOTVS estilizada em SVG inline
    return '''
    <svg height="48" viewBox="0 0 220 60" xmlns="http://www.w3.org/2000/svg">
      <text x="0" y="42" font-family="Arial Black, sans-serif" font-size="44" font-weight="900"
            fill="white" letter-spacing="2">TOTVS</text>
    </svg>
    '''


# ============================================================
# SESSION STATE
# ============================================================
if 'passo' not in st.session_state:
    st.session_state.passo = 1
if 'sft_bytes' not in st.session_state:
    st.session_state.sft_bytes = None
if 'sft_nome' not in st.session_state:
    st.session_state.sft_nome = None
if 'perfil' not in st.session_state:
    st.session_state.perfil = {}
if 'planilha_bytes' not in st.session_state:
    st.session_state.planilha_bytes = None
if 'logs' not in st.session_state:
    st.session_state.logs = []


def reset_app():
    for key in ['passo', 'sft_bytes', 'sft_nome', 'perfil', 'planilha_bytes', 'logs']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()


# ============================================================
# RENDERIZAÇÃO
# ============================================================
aplicar_tema_totvs()


# ===== SIDEBAR =====
with st.sidebar:
    st.markdown(
        f'<div style="text-align: center; padding: 1.5rem 0 1rem 0;">{renderizar_logo_totvs()}</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div style="text-align: center; padding-bottom: 1rem; '
        'border-bottom: 1px solid rgba(255,255,255,0.15); margin-bottom: 1.5rem;">'
        '<div style="font-size: 0.95rem; opacity: 0.9; font-weight: 500;">Configurador de Tributos</div>'
        '<div style="font-size: 0.8rem; opacity: 0.6;">Protheus — Livros Fiscais</div>'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown('<div style="font-size: 0.85rem; opacity: 0.7; margin-bottom: 0.5rem; '
                'text-transform: uppercase; letter-spacing: 1px;">Progresso</div>',
                unsafe_allow_html=True)

    passo = st.session_state.passo
    for n, label in [(1, "Upload da SFT"), (2, "Perfil da empresa"), (3, "Download")]:
        if n < passo:
            cls = "step-done"
            icon = "✓"
        elif n == passo:
            cls = "step-active"
            icon = "●"
        else:
            cls = "step-pending"
            icon = "○"
        st.markdown(
            f'<div class="step-item {cls}">'
            f'<span style="margin-right: 0.6rem; font-size: 1.1rem;">{icon}</span>'
            f'<strong>{n}.</strong> {label}'
            f'</div>',
            unsafe_allow_html=True
        )

    st.markdown(
        '<div style="margin-top: 2rem; padding: 1rem; background: rgba(255,255,255,0.05); '
        'border-radius: 8px; font-size: 0.85rem; line-height: 1.5;">'
        '💡 As tabelas <strong>CBENEF</strong> e <strong>NCM</strong> já estão '
        'embarcadas no app — basta enviar sua <strong>SFT</strong>.'
        '</div>',
        unsafe_allow_html=True
    )

    if st.session_state.passo > 1:
        st.markdown("&nbsp;", unsafe_allow_html=True)
        if st.button("🔄 Recomeçar", use_container_width=True):
            reset_app()


# ===== HEADER GRADIENTE =====
st.markdown(
    f'''
    <div class="totvs-header">
        <div class="totvs-header-logo">
            {renderizar_logo_totvs()}
        </div>
        <h1>Configurador de Tributos</h1>
        <p>Gerador de Planilha de Apoio — Cruzamento SFT × CBENEF × NCM</p>
    </div>
    ''',
    unsafe_allow_html=True
)


# ============================================================
# PASSO 1 — UPLOAD DA SFT
# ============================================================
if st.session_state.passo == 1:
    st.markdown(
        '<div class="secao-card">'
        '<h2 style="margin-top: 0;">Passo 1 — Upload da tabela SFT</h2>'
        '<p style="color: #717171; margin-bottom: 1.5rem;">'
        'Faça upload do arquivo <strong>SFT.xlsx</strong> extraído do Protheus '
        '(tabela de Livros Fiscais). O app detecta automaticamente a aba correta.'
        '</p>'
        '</div>',
        unsafe_allow_html=True
    )

    sft_file = st.file_uploader(
        "Selecione o arquivo SFT (.xlsx)",
        type=['xlsx', 'xls'],
        help="Arquivo Excel exportado da tabela SFT do Protheus.",
    )

    if sft_file is not None:
        st.session_state.sft_bytes = sft_file.read()
        st.session_state.sft_nome = sft_file.name
        st.success(f"✅ Arquivo carregado: **{sft_file.name}** ({len(st.session_state.sft_bytes):,} bytes)")

        st.markdown("&nbsp;", unsafe_allow_html=True)
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("Próximo →", type="primary", use_container_width=True):
                st.session_state.passo = 2
                st.rerun()


# ============================================================
# PASSO 2 — QUESTIONÁRIO DE PERFIL
# ============================================================
elif st.session_state.passo == 2:
    st.markdown(
        '<div class="secao-card">'
        '<h2 style="margin-top: 0;">Passo 2 — Perfil da empresa</h2>'
        '<p style="color: #717171; margin-bottom: 0.5rem;">'
        'Responda às perguntas sobre o perfil operacional. Suas respostas '
        '<strong>filtram automaticamente</strong> os códigos CBENEF que não se aplicam.'
        '</p>'
        '<p style="background: #FFF5FA; border-left: 4px solid #D52B7B; '
        'padding: 0.8rem 1rem; border-radius: 6px; font-size: 0.9rem; color: #2D2D2D; margin-top: 1rem;">'
        '💡 <strong>Dica:</strong> em caso de dúvida, deixe desmarcado — o código vai '
        'aparecer como ⚠ TALVEZ para revisar.'
        '</p>'
        '</div>',
        unsafe_allow_html=True
    )

    grupos = {
        "🔹 Operações comuns": [
            'exporta', 'zfm', 'doacao', 'orgao_publico', 'drawback',
            'ativo_imobilizado', 'mudanca_endereco', 'sucessor_legal',
            'arrendamento', 'comodato', 'demonstracao_mostr',
        ],
        "🔹 Regime e localização": [
            'agropecuario', 'simples_nacional', 'ride',
        ],
        "🔹 Setores específicos": [
            'aeroespacial', 'sucroenergetico', 'telecom', 'transporte_passag',
            'energia', 'farmaceutica', 'reporto', 'befiex', 'fomentar_produzir',
            'embarcacao_aeronave', 'vestuario_cama_mesa', 'minerador',
            'reciclagem', 'gas_combustivel', 'obra_arte', 'concessionaria',
            'artesao',
        ],
    }
    perguntas_dict = {flag: (pergunta, default) for (flag, pergunta, default) in PERFIL_PERGUNTAS}

    with st.form("form_perfil"):
        perfil = {}
        for nome_grupo, flags in grupos.items():
            st.markdown(f'<div class="grupo-titulo">{nome_grupo}</div>', unsafe_allow_html=True)
            cols = st.columns(2)
            for i, flag in enumerate(flags):
                if flag not in perguntas_dict:
                    continue
                pergunta, default = perguntas_dict[flag]
                default_bool = (default == 's')
                valor_salvo = st.session_state.perfil.get(flag, default_bool)
                with cols[i % 2]:
                    perfil[flag] = st.checkbox(
                        pergunta,
                        value=valor_salvo,
                        key=f"chk_{flag}",
                    )
            st.markdown("&nbsp;", unsafe_allow_html=True)

        st.markdown("---")
        col1, _, col3 = st.columns([1, 2, 1])
        with col1:
            voltar = st.form_submit_button("← Voltar", use_container_width=True)
        with col3:
            gerar = st.form_submit_button("Gerar planilha →", type="primary", use_container_width=True)

    if voltar:
        st.session_state.perfil = perfil
        st.session_state.passo = 1
        st.rerun()

    if gerar:
        st.session_state.perfil = perfil
        st.session_state.passo = 3
        st.rerun()


# ============================================================
# PASSO 3 — PROCESSAMENTO E DOWNLOAD
# ============================================================
elif st.session_state.passo == 3:
    if st.session_state.planilha_bytes is None:
        logs_buffer = []
        def coletar_log(msg):
            logs_buffer.append(str(msg))

        with st.spinner("⚙️ Processando SFT e gerando planilha..."):
            try:
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_sft:
                    tmp_sft.write(st.session_state.sft_bytes)
                    sft_tmp_path = tmp_sft.name

                try:
                    planilha_bytes = gerar_planilha(
                        sft_path=sft_tmp_path,
                        perfil_empresa=st.session_state.perfil,
                        log_fn=coletar_log,
                        output_path=None,
                    )
                    st.session_state.planilha_bytes = planilha_bytes
                    st.session_state.logs = logs_buffer
                finally:
                    try: os.unlink(sft_tmp_path)
                    except OSError: pass
            except Exception as e:
                st.error(f"❌ Erro ao processar: {type(e).__name__}: {e}")
                st.session_state.logs = logs_buffer + [f"ERRO: {e}"]
                with st.expander("📋 Ver log completo", expanded=True):
                    st.code("\n".join(st.session_state.logs), language=None)
                if st.button("← Voltar"):
                    st.session_state.passo = 2
                    st.rerun()
                st.stop()

    n_sim = sum(1 for v in st.session_state.perfil.values() if v)
    n_nao = sum(1 for v in st.session_state.perfil.values() if not v)

    st.markdown(
        f'<div class="secao-card" style="text-align: center; '
        f'background: linear-gradient(135deg, white 0%, #FFF5FA 100%); '
        f'border: 2px solid #D52B7B;">'
        f'<div style="font-size: 3.5rem; margin-bottom: 0.5rem;">🎉</div>'
        f'<h2 style="margin-top: 0; color: #D52B7B;">Planilha gerada com sucesso!</h2>'
        f'<p style="color: #717171; font-size: 1rem;">'
        f'Perfil capturado: <strong>{n_sim}</strong> características SIM, '
        f'<strong>{n_nao}</strong> NÃO'
        f'</p>'
        f'</div>',
        unsafe_allow_html=True
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    nome_arquivo = f"Planilha_Apoio_Configurador_Tributos_{timestamp}.xlsx"

    st.download_button(
        label="📥  Baixar Planilha de Apoio (.xlsx)",
        data=st.session_state.planilha_bytes,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.markdown("&nbsp;", unsafe_allow_html=True)

    with st.expander("📋 Ver log de processamento", expanded=False):
        st.code("\n".join(st.session_state.logs), language=None)

    st.markdown("&nbsp;", unsafe_allow_html=True)
    if st.button("🔄 Processar outra SFT", use_container_width=True, type="secondary"):
        reset_app()


# ===== FOOTER =====
st.markdown(
    '<div style="text-align: center; color: #717171; font-size: 0.8rem; '
    'padding: 2rem 0 1rem 0; margin-top: 3rem; border-top: 1px solid #E0E0E0;">'
    'Desenvolvido para automação fiscal — Protheus TOTVS · '
    'Cruzamento inteligente CBENEF × NCM × Perfil da Empresa'
    '</div>',
    unsafe_allow_html=True
)
