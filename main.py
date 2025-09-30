import streamlit as st
import pandas as pd
import unicodedata
import re
from collections import Counter
import io
import time

# Configuração da página
st.set_page_config(
    page_title="Gerador de Excel para Análise",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
* {
    font-size: 18px !important;
}
</style>
""", unsafe_allow_html=True)

# ---------------- Utilidades ----------------
def _normalize_text(s: str) -> str:
    if s is None: return ""
    if not isinstance(s, str): s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s

def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm_map = {_normalize_text(c): c for c in df.columns}
    for cand in candidates:
        nc = _normalize_text(cand)
        if nc in norm_map:
            return norm_map[nc]
    for want in candidates:
        nw = _normalize_text(want)
        for nc, orig in norm_map.items():
            if nw in nc:
                return orig
    return None

def _to_int_or_none(x):
    if pd.isna(x): return None
    try:
        return int(float(str(x).replace(",", ".")))
    except:
        digits = re.findall(r"\d+", str(x))
        if digits: return int(digits[0])
    return None

# ---------------- Leitura de arquivos ----------------
def carregar_plano(uploaded_file) -> pd.DataFrame | None:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)  # força leitura como texto
    except Exception as e:
        st.error(f"Erro ao carregar o plano: {e}")
        return None

    col_cep_ini = _find_col(df, ["cep inicial", "inicio"])
    col_cep_fim = _find_col(df, ["cep final", "fim"])
    col_direcao = _find_col(df, ["direção de triagem", "direcao"])
    col_saida = _find_col(df, ["saída principal", "rampa"])
    col_tipo = _find_col(df, ["tipo de objeto", "objeto"])

    if not all([col_cep_ini, col_cep_fim, col_direcao, col_saida]):
        st.error("Colunas necessárias não encontradas no arquivo do plano")
        return None

    # mantém CEP como texto com zeros à esquerda
    df["_cep_ini"] = df[col_cep_ini].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
    df["_cep_fim"] = df[col_cep_fim].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
    df["_direcao"] = df[col_direcao].fillna("").astype(str).str.strip()
    df["_saida_num"] = df[col_saida].apply(_to_int_or_none)

    if col_tipo:
        df["_tipo_objeto"] = df[col_tipo].astype(str).str.strip()
    else:
        df["_tipo_objeto"] = "Desconhecido"

    df = df.dropna(subset=["_cep_ini", "_cep_fim"]).copy()
    return df

def carregar_ceps(uploaded_file) -> list[str] | None:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)  # lê como texto
    except Exception as e:
        st.error(f"Erro ao carregar CEPS: {e}")
        return None

    col_cep = _find_col(df, ["cep"])
    if col_cep is None:
        st.error("Coluna 'CEP' não encontrada no arquivo")
        return None

    try:
        return df[col_cep].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8).tolist()
    except Exception as e:
        st.error(f"Erro processando coluna 'cep': {e}")
        return None

# ---------------- Simulação ----------------
def simular_triagem(plano: pd.DataFrame, ceps: list[str]) -> list[dict]:
    faixas = [(int(r["_cep_ini"]), int(r["_cep_fim"]),
               (r["_direcao"] or "").strip(), r["_saida_num"], r["_tipo_objeto"])
              for _, r in plano.iterrows()]
    resultados = []
    for cep in ceps:
        try:
            cep_num = int(cep)  # só para comparação
        except:
            continue
        direcao, saida, tipo = "Não encontrado", None, "Desconhecido"
        for ini, fim, d, s, t in faixas:
            if ini <= cep_num <= fim:
                direcao = d if d else "Sem direção"
                saida = s
                tipo = t
                break
        resultados.append({
            "CEP": cep,  # mantém string com zeros à esquerda
            "Direção": direcao,
            "Saída Principal": saida,
            "Tipo de Objeto": tipo
        })
    return resultados

# ---------------- Painel de Alas ----------------
def montar_painel_alas(resultados: list[dict]):
    direcoes_por_rampa: dict[int, Counter] = {}
    for r in resultados:
        rampa = r.get("Saída Principal")
        direcao = r.get("Direção") or ""
        if rampa is None or direcao in ("", "Não encontrado"):
            continue
        direcoes_por_rampa.setdefault(int(rampa), Counter())[direcao] += 1

    alas = [("Ala A", list(range(109, 153))),
            ("Ala B", list(range(77, 109))),
            ("Ala C", list(range(44, 77))),
            ("Ala D", list(range(1, 44)))]

    cols = st.columns(4)
    for idx, (nome_ala, rampas) in enumerate(alas):
        with cols[idx]:
            st.subheader(nome_ala)
            for rampa in rampas:
                counter = direcoes_por_rampa.get(rampa, Counter())
                total = sum(counter.values())
                bg_color = "#97e657" if total > 0 else "#f9f9f9"
                html_content = f"""
                <div style="border: 2px solid #ddd; border-radius: 8px; padding: 10px; margin: 5px 0; background-color: {bg_color};">
                    <div style="font-weight: bold; margin-bottom: 5px;">Rampa {rampa}</div>
                """
                if total > 0:
                    for d, q in sorted(counter.items()):
                        html_content += f"<div>{d}: {q}</div>"
                    html_content += f"<div><i>Total de objetos: {total}</i></div>"
                else:
                    html_content += "<div>Nenhuma direção</div>"
                    html_content += f"<div><i>Total de objetos: 0</i></div>"
                html_content += "</div>"
                st.markdown(html_content, unsafe_allow_html=True)
def montar_resumo_blocos_por_ala(resultados: list[dict], ala: str):
    df = pd.DataFrame(resultados)
    df = df.dropna(subset=["Saída Principal"]).copy()
    df["Saída Principal"] = df["Saída Principal"].astype(int)

    blocos = {
        "Ala D": [(1, 5), (6, 11), (12, 16), (17, 22), (23, 27),
                  (28, 33), (34, 39), (40, 43)],
        "Ala C": [(44, 48), (49, 54), (55, 59), (60, 65), (66, 71), (72, 76)],
        "Ala B": [(77, 80), (81, 86), (87, 91), (92, 97), (98, 103), (104, 108)],
        "Ala A": [(109, 113), (114, 119), (120, 124), (125, 130),
                  (131, 136), (137, 142), (143, 148), (149, 152)]
    }

    st.header(f"📊 Resumo - {ala}")
    resumo_blocos = []

    # Total da ala inteira
    total_ala = df[df["Saída Principal"].between(blocos[ala][0][0], blocos[ala][-1][1])].shape[0]

    for ini, fim in blocos[ala]:
        df_bloco = df[(df["Saída Principal"] >= ini) & (df["Saída Principal"] <= fim)]
        if df_bloco.empty:
            continue

        # Total de objetos por rampa
        resumo = (
            df_bloco.groupby("Saída Principal")
            .size()
            .reset_index(name="Total")
            .sort_values("Saída Principal")
        )

        # Total de direções diferentes por rampa
        total_direcoes = (
            df_bloco.groupby("Saída Principal")["Direção"]
            .nunique()
            .reset_index(name="Qt Direções")
        )

        # Junta com o resumo
        resumo = resumo.merge(total_direcoes, on="Saída Principal")

        # Porcentagem pelo total da ala
        resumo["% na Ala"] = (resumo["Total"] / total_ala * 100).map("{:.2f}%".format)

        # Linha TOTAL do bloco
        total_bloco = resumo["Total"].sum()
        total_direcoes_bloco = resumo["Qt Direções"].sum()
        resumo_total = pd.DataFrame({
            "Saída Principal": ["TOTAL"],
            "Total": [total_bloco],
            "Qt Direções": [total_direcoes_bloco],
            "% na Ala": [f"{(total_bloco / total_ala * 100):.2f}%"]
        })

        resumo = pd.concat([resumo, resumo_total], ignore_index=True)

        # Exibe no Streamlit (sem índice e maior)
        st.subheader(f"Bloco {ini}-{fim}")
        st.dataframe(
            resumo.rename(columns={"Saída Principal": "Rampa"}).reset_index(drop=True),
            use_container_width=True,

        )

        resumo["Bloco"] = f"{ini}-{fim}"
        resumo_blocos.append(resumo)

    if resumo_blocos:
        return pd.concat(resumo_blocos, ignore_index=True)
    return pd.DataFrame()


# ---------------- Exportação ----------------
def exportar_triagem_excel(plano, ceps, resumo_blocos_df):
    if plano is None or ceps is None:
        st.error("⚠️ Por favor, carregue os dois arquivos antes de exportar.")
        return None

    progress_bar = st.progress(0)
    for i in range(100):
        time.sleep(0.01)
        progress_bar.progress(i + 1)

    resultados = simular_triagem(plano, ceps)
    df_res = pd.DataFrame(resultados)

    contagem_por_direcao = (
        df_res.groupby(["Direção", "Saída Principal", "Tipo de Objeto"])
        .size()
        .reset_index(name="Quantidade")
    )

    resumo = contagem_por_direcao.pivot_table(
        index=["Direção", "Saída Principal"],
        columns="Tipo de Objeto",
        values="Quantidade",
        fill_value=0
    ).reset_index()

    if "Pacote" not in resumo.columns:
        resumo["Pacote"] = 0
    if "Envelope" not in resumo.columns:
        resumo["Envelope"] = 0

    resumo["Quantidade Total"] = resumo["Pacote"] + resumo["Envelope"]
    resumo = resumo[["Direção", "Saída Principal", "Quantidade Total", "Pacote", "Envelope"]]

    df_sem_direcao = df_res[df_res["Direção"].isin(["Não encontrado", "Sem direção"])]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        df_sem_direcao.to_excel(writer, sheet_name="CEPs sem direção", index=False)
        if not resumo_blocos_df.empty:
            resumo_blocos_df.to_excel(writer, sheet_name="Resumo Blocos", index=False)

    output.seek(0)
    progress_bar.empty()
    st.success("💾 Arquivo preparado para download!")

    return output

# ---------------- Interface Principal ----------------
def main():
    st.title("📊 Gerador de Excel para Análise")

    with st.sidebar:
        st.header("📁 Upload de Arquivos")
        uploaded_plano = st.file_uploader("Carregar Plano (Excel)", type=["xlsx", "xls"])
        uploaded_ceps = st.file_uploader("Carregar Arquivo de Triagem (Excel)", type=["xlsx", "xls"])

        st.header("⚙️ Configurações")
        if st.button("🔄 Processar Triagem", use_container_width=True):
            if uploaded_plano and uploaded_ceps:
                st.session_state.plano = carregar_plano(uploaded_plano)
                st.session_state.ceps = carregar_ceps(uploaded_ceps)
                if st.session_state.plano is not None and st.session_state.ceps is not None:
                    st.session_state.resultados = simular_triagem(st.session_state.plano, st.session_state.ceps)
                    st.success("✅ Processamento concluído com sucesso!")
            else:
                st.error("⚠️ Por favor, carregue os dois arquivos antes de processar.")

    if 'resultados' in st.session_state:
        st.header("📈 Painel de Alas")
        montar_painel_alas(st.session_state.resultados)

        st.header("📊 Resumo por Blocos - Escolha a Ala")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            if st.button("📌 Ala A", use_container_width=True):
                st.session_state.resumo_blocos_df = montar_resumo_blocos_por_ala(st.session_state.resultados, "Ala A")
        with col2:
            if st.button("📌 Ala B", use_container_width=True):
                st.session_state.resumo_blocos_df = montar_resumo_blocos_por_ala(st.session_state.resultados, "Ala B")
        with col3:
            if st.button("📌 Ala C", use_container_width=True):
                st.session_state.resumo_blocos_df = montar_resumo_blocos_por_ala(st.session_state.resultados, "Ala C")
        with col4:
            if st.button("📌 Ala D", use_container_width=True):
                st.session_state.resumo_blocos_df = montar_resumo_blocos_por_ala(st.session_state.resultados, "Ala D")

    if 'plano' in st.session_state and 'ceps' in st.session_state:
        if st.button("📥 Exportar para Excel", use_container_width=True):
            excel_file = exportar_triagem_excel(
                st.session_state.plano,
                st.session_state.ceps,
                st.session_state.get("resumo_blocos_df", pd.DataFrame())
            )
            if excel_file:
                st.download_button(
                    label="⬇️ Baixar Resultados",
                    data=excel_file,
                    file_name="resultados_triagem.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

if __name__ == "__main__":
    main()

