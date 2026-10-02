import io
import time
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st

# ------------------------------------------------------------------
# CONFIGURAÇÕES PADRÃO
# ------------------------------------------------------------------
BASE_URL = "https://mingle-ionapi.inforcloudsuite.com/US45PBYRE7XKA5QB_PRD/WM/wmwebservice_rest"
WAREHOUSE = "US45PBYRE7XKA5QB_PRD_COBALTLIKABLECAT_PRD_SCE_PRD_4_wmwhse1"

DEFAULT_FROMLOC = "FALTAS"
DEFAULT_FROMID = " "  # espaço literal

st.set_page_config(page_title="Movimentação de Estoque por SKU", layout="wide")
st.title("📦 Movimentação de Estoque por SKU")


# ------------------------------------------------------------------
# FUNÇÃO UNIVERSAL DE CONVERSÃO NUMÉRICA
# ------------------------------------------------------------------
def parse_numero(valor):
    """
    Converte qualquer formato numérico para float, independente da
    configuração regional do usuário. Retorna None se vazio ou inválido.

    Exemplos:
        12.00000      -> 12.0
        10,50000      -> 10.5
        10.50000      -> 10.5
        2,323.00000   -> 2323.0
        2.323,00000   -> 2323.0
        1.234.567     -> 1234567.0
        1,234,567     -> 1234567.0
    """
    if valor is None:
        return None

    s = str(valor).strip().replace(" ", "").replace("\u00a0", "")
    if not s or s.lower() in ("nan", "none", "nat"):
        return None

    tem_ponto = "." in s
    tem_virgula = "," in s

    if tem_ponto and tem_virgula:
        # o ÚLTIMO separador que aparece é o decimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")  # 2.323,00000
        else:
            s = s.replace(",", "")  # 2,323.00000
    elif tem_virgula:
        if s.count(",") > 1:
            s = s.replace(",", "")  # 1,234,567 (milhar)
        else:
            s = s.replace(",", ".")  # 10,50000 (decimal)
    elif tem_ponto:
        if s.count(".") > 1:
            s = s.replace(".", "")  # 1.234.567 (milhar)
        # um ponto só: já é decimal (12.00000)

    try:
        return float(s)
    except ValueError:
        return None


# ------------------------------------------------------------------
# 1. AUTENTICAÇÃO FIXA
# ------------------------------------------------------------------
st.header("1. Autenticação automática")

token_url = "https://mingle-sso.inforcloudsuite.com:443/US45PBYRE7XKA5QB_PRD/as/token.oauth2"

client_id = st.secrets["ci"]
client_secret = st.secrets["cs"]
username = st.secrets["saak"]
password = st.secrets["sask"]

token_payload = {
    "grant_type": "password",
    "client_id": client_id,
    "client_secret": client_secret,
    "username": username,
    "password": password,
}

with st.spinner("Autenticando automaticamente..."):
    resp = requests.post(token_url, data=token_payload)

if resp.status_code == 200:
    access_token = resp.json().get("access_token")
    st.success("Autenticado com sucesso.")
else:
    st.error(f"Falha na autenticação ({resp.status_code}): {resp.text}")
    access_token = None


# ------------------------------------------------------------------
# FUNÇÃO DE DETECÇÃO AUTOMÁTICA DE COLUNAS
# ------------------------------------------------------------------
def detectar_coluna(df, candidatos):
    cols_norm = df.columns.str.lower().str.strip()
    for nome in candidatos:
        nome_norm = nome.lower().strip()
        if nome_norm in cols_norm.values:
            return df.columns[cols_norm == nome_norm][0]
    return df.columns[0]  # fallback


# ------------------------------------------------------------------
# 2. PLANILHA XLSM
# ------------------------------------------------------------------
st.header("2. Planilha XLSM")

planilha_file = st.file_uploader(
    "Planilha .xlsm (abas: BIPAGEM e SALDOS)",
    type=["xlsm", "xlst", "xlsx", "xls", "xltx"],
)

df_pedido = None
df_estoque = None

if planilha_file is not None:
    xls = pd.ExcelFile(planilha_file)
    st.write(f"Abas encontradas: {xls.sheet_names}")

    aba_pedido_default = "BIPAGEM" if "BIPAGEM" in xls.sheet_names else xls.sheet_names[0]
    aba_estoque_default = "SALDOS" if "SALDOS" in xls.sheet_names else xls.sheet_names[-1]

    col1, col2 = st.columns(2)
    with col1:
        aba_pedido = st.selectbox(
            "Aba do PEDIDO (BIPAGEM)", xls.sheet_names, index=xls.sheet_names.index(aba_pedido_default)
        )
    with col2:
        aba_estoque = st.selectbox(
            "Aba do ESTOQUE (SALDOS)", xls.sheet_names, index=xls.sheet_names.index(aba_estoque_default)
        )

    df_pedido = pd.read_excel(xls, sheet_name=aba_pedido, dtype=str)
    df_estoque = pd.read_excel(xls, sheet_name=aba_estoque, dtype=str)

    # remove linhas totalmente vazias (comuns no fim de planilhas exportadas)
    df_pedido = df_pedido.dropna(how="all")
    df_estoque = df_estoque.dropna(how="all")

    def corrigir_zero_esquerda(serie: pd.Series) -> pd.Series:
        """
        Colunas que no Excel são código (ex: endereço) mas vêm com tipo numérico
        perdem zero à esquerda (número não tem zero à esquerda). Detecta a
        largura correta pelos valores mais frequentes e preenche com zfill.
        """
        limpa = serie.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        larguras = limpa.dropna().str.len()
        if larguras.empty:
            return limpa
        largura_alvo = larguras.mode().iloc[0]  # largura mais comum
        return limpa.str.zfill(int(largura_alvo))

    st.subheader("Prévia - BIPAGEM (Pedido)")
    st.dataframe(df_pedido, use_container_width=True)

    st.subheader("Prévia - SALDOS (Estoque)")
    st.dataframe(df_estoque, use_container_width=True)

    # ------------------------------------------------------------------
    # MAPEAMENTO AUTOMÁTICO
    # ------------------------------------------------------------------
    st.subheader("Mapeamento automático de colunas")

    col_sku_pedido = detectar_coluna(df_pedido, ["sku", "item"])
    col_qty_pedido = detectar_coluna(df_pedido, ["quantidade", "qty"])
    col_toid = detectar_coluna(df_pedido, ["lpn", "toid"])
    col_toloc = detectar_coluna(df_pedido, ["endereço", "toloc", "local"])

    col_sku_estoque = detectar_coluna(df_estoque, ["item", "sku", "codigo"])
    col_lote = detectar_coluna(df_estoque, ["lote"])
    col_qty_estoque = detectar_coluna(df_estoque, ["qtd. física", "quantidade", "saldo", "disponível"])

    # corrige zero à esquerda perdido (ENDEREÇO/toloc costuma vir como número no Excel)
    df_pedido[col_toloc] = corrigir_zero_esquerda(df_pedido[col_toloc])

    colA, colB, colC = st.columns(3)
    with colA:
        st.write("**Pedido**")
        st.write("SKU:", col_sku_pedido)
        st.write("Qty:", col_qty_pedido)
    with colB:
        st.write("**Pedido**")
        st.write("LPN (toid):", col_toid)
        st.write("Endereço (toloc):", col_toloc)
    with colC:
        st.write("**Estoque**")
        st.write("SKU:", col_sku_estoque)
        st.write("Lote:", col_lote)
        st.write("Qtd. física:", col_qty_estoque)

    # planilha nova -> limpa resultados antigos pra nunca mostrar dado desatualizado
    assinatura_atual = (planilha_file.name, planilha_file.size, aba_pedido, aba_estoque)
    if st.session_state.get("_assinatura_planilha") != assinatura_atual:
        st.session_state.pop("alocacoes", None)
        st.session_state.pop("avisos", None)
        st.session_state.pop("resultados", None)
        st.session_state["_assinatura_planilha"] = assinatura_atual


# ------------------------------------------------------------------
# 3. ALOCAÇÃO (com parsing robusto do saldo)
# ------------------------------------------------------------------
def parse_saldo(serie_bruta: pd.Series) -> pd.Series:
    """
    Converte a coluna de quantidade para float usando parse_numero
    (aceita 12.00000, 10,50000, 2,323.00000, 2.323,00000 etc.).
    Valores inválidos/vazios viram 0 em vez de travar o app ou virar NaN
    (NaN em comparações silenciosamente deixa passar qty sem limite).
    """
    return serie_bruta.map(parse_numero).astype(float).fillna(0.0)


def montar_alocacoes(df_pedido, df_estoque, cols):
    alocacoes = []
    avisos = []
    erros_linha = []

    estoque = df_estoque.copy()
    estoque["_saldo"] = parse_saldo(estoque[cols["qty_estoque"]])

    for i, pedido_row in df_pedido.iterrows():
        sku_raw = pedido_row[cols["sku_pedido"]]
        if pd.isna(sku_raw) or str(sku_raw).strip() == "":
            continue  # linha vazia, ignora

        sku = str(sku_raw).strip()

        qty_necessaria = parse_numero(pedido_row[cols["qty_pedido"]])
        if qty_necessaria is None:
            erros_linha.append(
                f"Linha {i + 2} do pedido: quantidade inválida ('{pedido_row[cols['qty_pedido']]}'), linha ignorada."
            )
            continue

        toid = str(pedido_row[cols["toid"]]).strip()
        toloc = str(pedido_row[cols["toloc"]]).strip()

        restante = qty_necessaria
        candidatos = estoque[estoque[cols["sku_estoque"]] == sku]

        for idx in candidatos.index:
            if restante <= 0:
                break

            saldo_linha = estoque.at[idx, "_saldo"]
            if pd.isna(saldo_linha) or saldo_linha <= 0:
                continue

            qty_mover = min(restante, saldo_linha)

            alocacoes.append(
                {
                    "sku": sku,
                    "fromlot": str(estoque.at[idx, cols["lote"]]).strip(),
                    "toloc": toloc,
                    "toid": toid,
                    "qty": qty_mover,
                }
            )

            estoque.at[idx, "_saldo"] = saldo_linha - qty_mover
            restante -= qty_mover

        if restante > 0:
            avisos.append(
                f"SKU {sku}: pedido {qty_necessaria}, faltou mover {restante} unidade(s) "
                f"(saldo insuficiente) para toloc={toloc} / toid={toid}."
            )

    return alocacoes, avisos, erros_linha


def executar_movimentacao(alocacao, fromloc_fixo, fromid_fixo, token):
    fromloc_enc = quote(str(fromloc_fixo), safe="")
    fromid_enc = quote(str(fromid_fixo), safe="")
    toloc_enc = quote(str(alocacao["toloc"]), safe="")

    url = f"{BASE_URL}/{WAREHOUSE}/inventory/{fromloc_enc}/{fromid_enc}/{toloc_enc}/move"

    params = {
        "toid": str(alocacao["toid"]),
        "fromlot": str(alocacao["fromlot"]),
        "qty": alocacao["qty"],
    }

    headers = {"Authorization": f"Bearer {token}"} if token else {}

    try:
        resp = requests.post(url, params=params, headers=headers, timeout=30)
        return {
            "status_code": resp.status_code,
            "sucesso": resp.status_code in (200, 201, 204),
            "resposta": resp.text,
        }
    except Exception as e:
        return {"status_code": None, "sucesso": False, "resposta": str(e)}


# ------------------------------------------------------------------
# 4. BOTÕES
# ------------------------------------------------------------------
st.header("3. Executar movimentação")

if df_pedido is not None and df_estoque is not None:
    cols = {
        "sku_pedido": col_sku_pedido,
        "qty_pedido": col_qty_pedido,
        "toid": col_toid,
        "toloc": col_toloc,
        "sku_estoque": col_sku_estoque,
        "lote": col_lote,
        "qty_estoque": col_qty_estoque,
    }

    if st.button("Calcular alocação (sem enviar ainda)"):
        try:
            alocacoes, avisos, erros_linha = montar_alocacoes(df_pedido, df_estoque, cols)
            st.session_state["alocacoes"] = alocacoes
            st.session_state["avisos"] = avisos
            st.session_state["erros_linha"] = erros_linha
            st.session_state.pop("resultados", None)  # invalida resultado antigo de envio
        except Exception as e:
            st.session_state.pop("alocacoes", None)
            st.error(f"Erro ao calcular alocação: {e}")

    if "alocacoes" in st.session_state:
        alocacoes = st.session_state["alocacoes"]
        avisos = st.session_state["avisos"]
        erros_linha = st.session_state.get("erros_linha", [])

        st.subheader("Alocação calculada")
        st.dataframe(pd.DataFrame(alocacoes), use_container_width=True)

        if erros_linha:
            st.error("Linhas do pedido ignoradas por erro:")
            for e in erros_linha:
                st.write(f"- {e}")

        if avisos:
            st.warning("Avisos de saldo insuficiente:")
            for a in avisos:
                st.write(f"- {a}")

        if access_token is None:
            st.info("Autenticação falhou, não é possível enviar.")

        if st.button("🚀 Enviar movimentações para o WMS", disabled=(access_token is None or len(alocacoes) == 0)):
            resultados = []
            barra = st.progress(0)
            total = len(alocacoes)

            for i, aloc in enumerate(alocacoes):
                resultado = executar_movimentacao(aloc, DEFAULT_FROMLOC, DEFAULT_FROMID, access_token)
                resultado.update(aloc)
                resultados.append(resultado)
                barra.progress((i + 1) / total)
                time.sleep(0.1)

            st.session_state["resultados"] = pd.DataFrame(resultados)

    if "resultados" in st.session_state:
        df_resultados = st.session_state["resultados"]

        st.subheader("Resultado")
        sucesso = df_resultados[df_resultados["sucesso"] == True]
        erro = df_resultados[df_resultados["sucesso"] == False]

        st.success(f"{len(sucesso)} movimentação(ões) com sucesso.")
        if len(erro) > 0:
            st.error(f"{len(erro)} movimentação(ões) com erro.")
            st.dataframe(erro, use_container_width=True)

        st.dataframe(df_resultados, use_container_width=True)

        csv_buffer = io.StringIO()
        df_resultados.to_csv(csv_buffer, index=False)
        st.download_button(
            "⬇️ Baixar resultado completo (CSV)",
            data=csv_buffer.getvalue(),
            file_name="resultado_movimentacao.csv",
            mime="text/csv",
        )

        if len(erro) > 0:
            erro_csv_buffer = io.StringIO()
            erro.to_csv(erro_csv_buffer, index=False)
            st.download_button(
                "⬇️ Baixar apenas erros (CSV)",
                data=erro_csv_buffer.getvalue(),
                file_name="erros_movimentacao.csv",
                mime="text/csv",
            )
else:
    st.info("Envie a planilha com as abas BIPAGEM e SALDOS para continuar.")
