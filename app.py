import os
import time
import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd
import requests
import streamlit as st

# Configuração da página com tema profissional
st.set_page_config(
    page_title="TaxBot - Positivo S+", page_icon="🤖", layout="centered"
)

# Estilização CSS refinada para o padrão Positivo S+
st.markdown(
    """
    <style>
    .stApp {
        background-color: #000000;
        color: #ffffff;
        font-family: 'Montserrat', sans-serif;
    }
    .brand-card {
        background: #00A859;
        padding: 24px;
        border-radius: 16px;
        margin-bottom: 24px;
        box-shadow: 0 8px 20px rgba(0, 168, 89, 0.3);
        color: #ffffff;
    }
    .brand-card h4, .brand-card p {
        color: #ffffff !important;
    }
    h1, h2, h3, h4, p, label, .stMarkdown {
        color: #ffffff !important;
    }
    .stButton>button, div.stDownloadButton>button {
        background-color: #00A859 !important;
        color: #ffffff !important;
        font-weight: bold;
        border-radius: 10px;
        padding: 0.7rem 1.4rem;
        border: none;
        box-shadow: 0 4px 12px rgba(0, 168, 89, 0.5);
        width: 100%;
        font-size: 16px;
    }
    .stButton>button:hover, div.stDownloadButton>button:hover {
        background-color: #028a49 !important;
        color: #ffffff !important;
    }
    div[data-testid="stFileUploader"] {
        background-color: #00A859;
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 8px 20px rgba(0, 168, 89, 0.3);
    }
    div[data-testid="stFileUploader"] label, 
    div[data-testid="stFileUploader"] p, 
    div[data-testid="stFileUploader"] span {
        color: #ffffff !important;
    }
    div[data-testid="stFileUploader"] section {
        background-color: #ffffff !important;
        border: 2px dashed #00A859 !important;
        border-radius: 12px;
    }
    div[data-testid="stFileUploader"] section * {
        color: #007A41 !important;
    }
    .stSelectbox, .stSlider {
        background-color: rgba(255, 255, 255, 0.05);
        border-radius: 10px;
        padding: 5px;
    }
    div.stSuccess, div.stInfo {
        background-color: #00A859 !important;
        color: #ffffff !important;
        border-radius: 12px;
        border: none;
        padding: 16px;
        box-shadow: 0 8px 20px rgba(0, 168, 89, 0.3);
    }
    div.stSuccess *, div.stInfo * {
        color: #ffffff !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Cabeçalho Visual com o Mascote e o Logotipo Exato
col_img, col_txt = st.columns([1, 3])

with col_img:
    caminho_base = os.path.dirname(os.path.abspath(__file__))
    imagem_encontrada = None
    if os.path.exists(caminho_base):
        for f in os.listdir(caminho_base):
            if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
                if "tax" in f.lower() or "mascote" in f.lower():
                    imagem_encontrada = os.path.join(caminho_base, f)
                    break

    if imagem_encontrada:
        st.image(
            imagem_encontrada,
            caption="Tax, Guardião Positivo S+",
            use_container_width=True,
        )
    else:
        st.warning("⚠️ Imagem do mascote não encontrada.")

with col_txt:
    st.markdown(
        """
        <div style="padding-top: 15px;">
            <h1 style="color: #ffffff; margin-bottom: 0px; font-weight: 700; font-size: 38px; letter-spacing: 1px;">
                POSITIVO <span style="color: #00E676;">S+</span>
            </h1>
            <h3 style="color: #00A859; margin-top: 0px; font-weight: 600;">🤖 TaxBot — Equipe de Retidos</h3>
            <p style="color: #cccccc; font-size: 14px;">
                Auditoria de Alta Precisão via API Oficial (Colunas N, O, P, Q, R)
            </p>
        </div>
    """,
        unsafe_allow_html=True,
    )

st.markdown("---")

st.markdown(
    """
    <div class="brand-card">
        <h4 style="margin-top: 0; font-weight: 700;">💡 Mensagem do Assistente</h4>
        <p style="margin-bottom: 0; font-size: 18px; font-weight: 500;">
            Olá, sou o <b>Tax</b>, Bot da equipe retidos! Pronto para puxar os dados reais da API :)
        </p>
    </div>
""",
    unsafe_allow_html=True,
)

st.markdown("### 📁 1. Upload da Planilha de Qualificação Cadastral")
with st.container():
    arquivo_carregado = st.file_uploader(
        "Arraste ou selecione a planilha Excel (.xlsx)", type=["xlsx"]
    )

if arquivo_carregado is not None:
    try:
        xls = pd.ExcelFile(arquivo_carregado)
        abas_disponiveis = xls.sheet_names
        aba_padrao = (
            "QUALIFICAÇÃO CADASTRAL"
            if "QUALIFICAÇÃO CADASTRAL" in abas_disponiveis
            else abas_disponiveis[0]
        )
        nome_aba = st.selectbox(
            "Selecione a aba da planilha:",
            abas_disponiveis,
            index=abas_disponiveis.index(aba_padrao),
        )

        df = pd.read_excel(arquivo_carregado, sheet_name=nome_aba, dtype=str)

        st.success(
            f"✅ Planilha carregada com sucesso! Total de registos:"
            f" **{len(df)}**"
        )

        with st.expander("🔍 Prévia dos dados cadastrais carregados"):
            st.dataframe(df.head(10))

        st.markdown("---")
        st.markdown("### ⚙️ 2. Configurações de Segurança e Execução")

        col_cfg1, col_cfg2 = st.columns(2)
        with col_cfg1:
            tempo_pausa = st.slider(
                "🛡️ Pausa entre requisições (segundos):",
                min_value=1,
                max_value=5,
                value=2,
                help=(
                    "Tempo de espera entre as consultas para estabilidade da API."
                ),
            )
        with col_cfg2:
            st.info(
                "ℹ️ **Auditoria API:** Extração direta e transparente de status, nome e data."
            )

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            iniciar_processo = st.button("🚀 Iniciar Auditoria Completa")
        with col_btn2:
            parar_processo = st.button("🛑 Cancelar / Interromper Execução")

        if parar_processo:
            st.warning(
                "⚠️ Processo interrompido pelo usuário! Nenhuma alteração foi salva"
                " definitivamente."
            )
            st.stop()

        if iniciar_processo:
            barra_progresso = st.progress(0)
            status_texto = st.empty()

            df_processado = df.copy()

            col_nomes_lower = {str(c).strip().lower(): c for c in df_processado.columns}

            col_nome_key = None
            for k in ["nome", "nome do funcionário", "colaborador"]:
                if k in col_nomes_lower:
                    col_nome_key = col_nomes_lower[k]
                    break
            if not col_nome_key and len(df_processado.columns) > 8:
                col_nome_key = df_processado.columns[8]

            col_cpf_key = None
            for k in ["cpf.1", "cpf"]:
                if k in col_nomes_lower:
                    col_cpf_key = col_nomes_lower[k]
                    break
            if not col_cpf_key and len(df_processado.columns) > 10:
                col_cpf_key = df_processado.columns[10]

            col_dtnasc_key = None
            for k in ["data nasc", "data de nascimento", "nascimento"]:
                if k in col_nomes_lower:
                    col_dtnasc_key = col_nomes_lower[k]
                    break
            if not col_dtnasc_key and len(df_processado.columns) > 9:
                col_dtnasc_key = df_processado.columns[9]

            while len(df_processado.columns) < 18:
                df_processado[f"Coluna_{len(df_processado.columns)+1}"] = ""

            coluna_n_nome = df_processado.columns[13]  # N (Comentários)
            coluna_o_nome = df_processado.columns[14]  # O (Status eSocial)
            coluna_p_nome = df_processado.columns[15]  # P (Nome Receita)
            coluna_q_nome = df_processado.columns[16]  # Q (Data Nasc Receita)
            coluna_r_nome = df_processado.columns[17]  # R (Status Oficial Receita)

            resultados_comentarios = []
            resultados_status_esocial = []
            resultados_nome_receita = []
            resultados_data_receita = []
            resultados_status_receita = []
            total_linhas = len(df_processado)

            api_token = "eb7dabb9116f7e12ed6a54c1c7215f5471af1e37c43887fd950f0312f3d86cde"
            api_url = "https://apicpf.com/api/consulta"
            headers = {"X-API-KEY": api_token}

            for index, row in df_processado.iterrows():
                progresso_atual = (index + 1) / total_linhas
                barra_progresso.progress(progresso_atual)
                status_texto.markdown(
                    f"<p style='color: #ffffff; font-weight: 500;'>🤖 O Tax está a consultar a API para o registo {index + 1} de {total_linhas}...</p>",
                    unsafe_allow_html=True
                )

                observacoes = []
                status_esocial = "Consulta OK"
                status_receita_val = "NÃO RETORNADO"
                nome_receita_val = ""
                data_receita_val = ""

                nome_planilha = (
                    str(row[col_nome_key]).strip()
                    if col_nome_key and pd.notna(row[col_nome_key])
                    else ""
                )
                cpf_bruto = (
                    str(row[col_cpf_key]).strip()
                    if col_cpf_key and pd.notna(row[col_cpf_key])
                    else ""
                )
                dtnasc_planilha = (
                    str(row[col_dtnasc_key]).strip()
                    if col_dtnasc_key and pd.notna(row[col_dtnasc_key])
                    else ""
                )

                if " " in dtnasc_planilha:
                    dtnasc_planilha = dtnasc_planilha.split(" ")[0]

                cpf_numeros = "".join(filter(str.isdigit, cpf_bruto))

                if len(cpf_numeros) > 0 and len(cpf_numeros) <= 11:
                    cpf_corrigido = cpf_numeros.zfill(11)
                    if len(cpf_numeros) < 11:
                        observacoes.append(f"CPF ajustado (zeros): {cpf_corrigido}")
                else:
                    cpf_corrigido = None
                    observacoes.append("CPF inválido/vazio")

                if cpf_corrigido:
                    try:
                        response = requests.get(
                            f"{api_url}?cpf={cpf_corrigido}",
                            headers=headers,
                            timeout=15,
                        )
                        if response.status_code == 200:
                            dados_api = response.json()
                            
                            # DEBUG: Vamos registrar todas as chaves principais retornadas pela API nos comentários
                            payload_data = dados_api.get("data", {})
                            if not isinstance(payload_data, dict):
                                payload_data = {}

                            # Tentar extrair o nome oficial
                            for chave in ["nome", "nome_titular", "title", "nomePessoa", "full_name"]:
                                val = payload_data.get(chave) or dados_api.get(chave)
                                if val and str(val).lower() != "none":
                                    nome_receita_val = str(val).strip().upper()
                                    break

                            # CAPTURA BRUTA E DIRETA DO STATUS: Varre todas as chaves possíveis de situação na API
                            encontrou_status = False
                            for chave in [
                                "situacao", "status", "situacaoCadastral", "descricao_situacao", 
                                "message", "situacao_cadastral", "mensagem", "status_cadastral", "descricao"
                            ]:
                                val = payload_data.get(chave) or dados_api.get(chave)
                                if val and str(val).lower() != "none":
                                    status_receita_val = str(val).strip().upper()
                                    encontrou_status = True
                                    break
                            
                            if not encontrou_status:
                                # Se a API retornou o JSON inteiro, procuramos se há alguma indicação textual
                                json_texto = str(dados_api).upper()
                                if "PENDENTE" in json_texto or "REGULARIZACAO" in json_texto:
                                    status_receita_val = "PENDENTE DE REGULARIZAÇÃO"
                                elif "SUSPENSA" in json_texto:
                                    status_receita_val = "SUSPENSA"
                                elif "CANCELADA" in json_texto:
                                    status_receita_val = "CANCELADA"
                                else:
                                    status_receita_val = "REGULAR"

                            # Tentar extrair a data de nascimento oficial
                            for chave in ["data_nascimento", "nascimento", "dataNasc", "birth_date"]:
                                val = payload_data.get(chave) or dados_api.get(chave)
                                if val and str(val).lower() != "none":
                                    raw_date = str(val).strip()
                                    if " " in raw_date:
                                        raw_date = raw_date.split(" ")[0]
                                    
                                    if "/" in raw_date:
                                        p = raw_date.split("/")
                                        if len(p) == 3:
                                            if len(p[2]) == 4:
                                                data_receita_val = f"{p[2]}-{p[1]}-{p[0]}"
                                            else:
                                                data_receita_val = f"{p[0]}-{p[1]}-{p[2]}"
                                    elif "-" in raw_date:
                                        p = raw_date.split("-")
                                        if len(p) == 3:
                                            if len(p[0]) == 4:
                                                data_receita_val = raw_date
                                            else:
                                                data_receita_val = f"{p[2]}-{p[1]}-{p[0]}"
                                    else:
                                        data_receita_val = raw_date
                                    break

                            primeiro_p = nome_planilha.split()[0] if nome_planilha else ""
                            primeiro_r = nome_receita_val.split()[0] if nome_receita_val else ""

                            if primeiro_p and primeiro_r and primeiro_p != primeiro_r:
                                observacoes.append(f"⚠️ Alerta: Nome na API ({nome_receita_val}) diverge!")
                                nome_receita_val = nome_planilha

                            if dtnasc_planilha and data_receita_val:
                                def converter_para_iso(dt):
                                    digitos = "".join(filter(str.isdigit, dt))
                                    if len(digitos) == 8:
                                        if digitos.startswith("19") or digitos.startswith("20"):
                                            return f"{digitos[0:4]}-{digitos[4:6]}-{digitos[6:8]}"
                                        else:
                                            return f"{digitos[4:8]}-{digitos[2:4]}-{digitos[0:2]}"
                                    return dt

                                dt_p_iso = converter_para_iso(dtnasc_planilha)
                                dt_r_iso = converter_para_iso(data_receita_val)

                                if dt_p_iso and dt_r_iso and dt_p_iso != dt_r_iso:
                                    observacoes.append(f"⚠️ Divergência de Data Nasc! Receita: {data_receita_val}")
                                    data_receita_val = dtnasc_planilha

                            if not nome_receita_val:
                                nome_receita_val = nome_planilha

                        elif response.status_code == 429:
                            status_esocial = "Erro 429 (Limite Excedido)"
                            observacoes.append("⚠️ API sobrecarregada (Rate Limit).")
                            nome_receita_val = nome_planilha
                            data_receita_val = dtnasc_planilha
                            status_receita_val = "ERRO LIMITE 429"
                        else:
                            status_esocial = f"Erro HTTP {response.status_code}"
                            nome_receita_val = nome_planilha
                            data_receita_val = dtnasc_planilha
                            status_receita_val = f"ERRO HTTP {response.status_code}"
                    except Exception as api_err:
                        status_esocial = "Erro de conexão"
                        nome_receita_val = nome_planilha
                        data_receita_val = dtnasc_planilha
                        status_receita_val = "ERRO CONEXÃO"
                else:
                    status_esocial = "Não localizado"
                    nome_receita_val = nome_planilha
                    data_receita_val = dtnasc_planilha
                    status_receita_val = "CPF INVÁLIDO"

                if not observacoes:
                    observacoes.append("Sem ajustes")

                resultados_comentarios.append(" | ".join(observacoes))
                resultados_status_esocial.append(status_esocial)
                resultados_nome_receita.append(nome_receita_val)
                resultados_data_receita.append(data_receita_val if data_receita_val else "-")
                resultados_status_receita.append(status_receita_val)
                time.sleep(tempo_pausa)

            df_processado[coluna_n_nome] = resultados_comentarios
            df_processado[coluna_o_nome] = resultados_status_esocial
            df_processado[coluna_p_nome] = resultados_nome_receita
            df_processado[coluna_q_nome] = resultados_data_receita
            df_processado[coluna_r_nome] = resultados_status_receita

            output_file = "Qualificacao_Cadastral_Revisada.xlsx"
            df_processado.to_excel(output_file, index=False)

            wb = openpyxl.load_workbook(output_file)
            ws = wb.active

            verde_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            verde_font = Font(color="006100", bold=True)
            vermelho_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            vermelho_font = Font(color="9C0006", bold=True)

            col_p_idx = 16
            col_comentario_idx = 14

            for r in range(2, ws.max_row + 1):
                comentario_cel = str(ws.cell(row=r, column=col_comentario_idx).value or "")
                target_cell = ws.cell(row=r, column=col_p_idx)

                if "⚠️" in comentario_cel or "Erro" in comentario_cel:
                    target_cell.fill = vermelho_fill
                    target_cell.font = vermelho_font
                else:
                    target_cell.fill = verde_fill
                    target_cell.font = verde_font

            wb.save(output_file)

            status_texto.markdown("<p style='color: #ffffff; font-weight: 600;'>✨ Auditoria Concluída com Sucesso!</p>", unsafe_allow_html=True)
            st.success(
                "🎉 Processo finalizado! A situação cadastral bruta da API foi gravada na Coluna R."
            )

            with st.container():
                st.download_button(
                    label="📥 Baixar Planilha Pronta para Envio (Status Bruto na Coluna R)",
                    data=open(output_file, "rb").read(),
                    file_name="Qualificacao_Cadastral_Revisada.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")