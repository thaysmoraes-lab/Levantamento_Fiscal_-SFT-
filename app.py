"""
Configurador de Tributos Protheus — App Web (Streamlit)

Interface web para gerar a Planilha de Apoio do Configurador de Tributos
TOTVS Protheus a partir de uma tabela SFT (Livros Fiscais).

Como rodar localmente:
    pip install -r requirements.txt
    streamlit run app.py
"""
import streamlit as st
import io
from datetime import datetime

from processador import gerar_planilha, PERFIL_PERGUNTAS


# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Configurador de Tributos Protheus",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="expanded",
)


# ============================================================
# SESSION STATE — controla o passo do wizard
# ============================================================
if 'passo' not in st.session_state:
    st.session_state.passo = 1   # 1=upload, 2=perfil, 3=resultado
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
    """Reseta tudo e volta ao início."""
    for key in ['passo', 'sft_bytes', 'sft_nome', 'perfil', 'planilha_bytes', 'logs']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()


# ============================================================
# SIDEBAR — informações fixas
# ============================================================
with st.sidebar:
    st.title("📊 Configurador de Tributos")
    st.markdown("**Protheus TOTVS — Livros Fiscais**")
    st.markdown("---")

    st.markdown("### Passos")
    passo = st.session_state.passo
    st.markdown(f"{'✅' if passo > 1 else '🟢' if passo == 1 else '⚪'} **1.** Upload da SFT")
    st.markdown(f"{'✅' if passo > 2 else '🟢' if passo == 2 else '⚪'} **2.** Perfil da empresa")
    st.markdown(f"{'✅' if passo > 3 else '🟢' if passo == 3 else '⚪'} **3.** Download da planilha")

    st.markdown("---")
    st.markdown("### Sobre")
    st.markdown(
        "Este app gera a planilha de apoio do **Configurador de Tributos** "
        "do Protheus a partir da sua tabela **SFT** (Livros Fiscais), "
        "cruzando com tabelas oficiais de **CBENEF (GO)** e **NCM**."
    )
    st.markdown(
        "As tabelas CBENEF e NCM já estão embarcadas no app — você só precisa "
        "fazer upload da sua SFT."
    )

    if st.session_state.passo > 1:
        st.markdown("---")
        if st.button("🔄 Recomeçar", use_container_width=True):
            reset_app()


# ============================================================
# CABEÇALHO
# ============================================================
st.title("Configurador de Tributos Protheus")
st.caption("Gerador de Planilha de Apoio | Cruzamento SFT × CBENEF × NCM")


# ============================================================
# PASSO 1 — UPLOAD DA SFT
# ============================================================
if st.session_state.passo == 1:
    st.subheader("Passo 1 — Upload da tabela SFT")
    st.markdown(
        "Faça upload do arquivo **SFT.xlsx** extraído do Protheus "
        "(tabela de Livros Fiscais). O app detecta automaticamente "
        "a aba correta."
    )

    sft_file = st.file_uploader(
        "Selecione o arquivo SFT (.xlsx)",
        type=['xlsx', 'xls'],
        help="Arquivo Excel exportado da tabela SFT do Protheus."
    )

    if sft_file is not None:
        # Guarda em session state
        st.session_state.sft_bytes = sft_file.read()
        st.session_state.sft_nome = sft_file.name
        st.success(f"✅ Arquivo carregado: **{sft_file.name}** ({len(st.session_state.sft_bytes):,} bytes)")

        st.markdown("---")
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("Próximo →", type="primary", use_container_width=True):
                st.session_state.passo = 2
                st.rerun()


# ============================================================
# PASSO 2 — QUESTIONÁRIO DE PERFIL
# ============================================================
elif st.session_state.passo == 2:
    st.subheader("Passo 2 — Perfil da empresa")
    st.markdown(
        "Responda às perguntas abaixo sobre o perfil operacional da empresa. "
        "Suas respostas vão **filtrar automaticamente** os códigos CBENEF que "
        "não se aplicam ao seu negócio."
    )
    st.info(
        "💡 **Dica**: marque apenas o que de fato a empresa faz. Em caso de "
        "dúvida, deixe desmarcado — o código vai aparecer como ⚠ TALVEZ para revisar."
    )

    # Agrupa as 31 perguntas em categorias
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
    # Indexa as perguntas por flag
    perguntas_dict = {flag: (pergunta, default) for (flag, pergunta, default) in PERFIL_PERGUNTAS}

    with st.form("form_perfil"):
        perfil = {}
        for nome_grupo, flags in grupos.items():
            st.markdown(f"#### {nome_grupo}")
            for flag in flags:
                if flag not in perguntas_dict:
                    continue
                pergunta, default = perguntas_dict[flag]
                default_bool = (default == 's')
                # Recupera valor salvo no session_state (em caso de voltar)
                valor_salvo = st.session_state.perfil.get(flag, default_bool)
                perfil[flag] = st.checkbox(
                    pergunta,
                    value=valor_salvo,
                    key=f"chk_{flag}",
                )

        st.markdown("---")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col1:
            voltar = st.form_submit_button("← Voltar", use_container_width=True)
        with col3:
            gerar = st.form_submit_button("Gerar planilha →", type="primary", use_container_width=True)

    if voltar:
        st.session_state.perfil = perfil  # salva o que está preenchido
        st.session_state.passo = 1
        st.rerun()

    if gerar:
        st.session_state.perfil = perfil
        # Avança e processa
        st.session_state.passo = 3
        st.rerun()


# ============================================================
# PASSO 3 — PROCESSAMENTO E DOWNLOAD
# ============================================================
elif st.session_state.passo == 3:
    st.subheader("Passo 3 — Resultado")

    # Se ainda não processou, processa agora
    if st.session_state.planilha_bytes is None:
        logs_buffer = []
        def coletar_log(msg):
            logs_buffer.append(str(msg))

        with st.spinner("Processando SFT e gerando planilha... isso pode levar alguns segundos"):
            try:
                # Cria arquivo temporário com a SFT
                import tempfile, os
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_sft:
                    tmp_sft.write(st.session_state.sft_bytes)
                    sft_tmp_path = tmp_sft.name

                try:
                    planilha_bytes = gerar_planilha(
                        sft_path=sft_tmp_path,
                        perfil_empresa=st.session_state.perfil,
                        log_fn=coletar_log,
                        output_path=None,   # retorna bytes
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

    # Sucesso!
    st.success("✅ Planilha gerada com sucesso!")

    # Resumo do perfil
    n_sim = sum(1 for v in st.session_state.perfil.values() if v)
    n_nao = sum(1 for v in st.session_state.perfil.values() if not v)
    st.markdown(f"**Perfil capturado**: {n_sim} características SIM, {n_nao} NÃO")

    st.markdown("---")

    # Botão de download
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    nome_arquivo = f"Planilha_Apoio_Configurador_Tributos_{timestamp}.xlsx"
    st.download_button(
        label="📥 Baixar Planilha de Apoio (.xlsx)",
        data=st.session_state.planilha_bytes,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.markdown("---")

    # Log de processamento
    with st.expander("📋 Ver log de processamento", expanded=False):
        st.code("\n".join(st.session_state.logs), language=None)

    st.markdown(" ")
    if st.button("🔄 Processar outra SFT", use_container_width=True):
        reset_app()
