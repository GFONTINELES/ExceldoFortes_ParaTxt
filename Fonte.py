import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
import os

st.set_page_config(page_title="Gerador TXT - Fortes Seguro v3", layout="wide")
st.title("📄 Gerador de TXT - Fortes (Somente linhas com código, texto e valor)")

# estilo botão
st.markdown("""
<style>
div.stDownloadButton > button {
    background-color: #FFD60A !important;
    color: #000000 !important;
    border: none !important;
    font-weight: 700 !important;
    font-size: 17px !important;
    padding: 0.6em 1.5em !important;
    border-radius: 12px !important;
    box-shadow: 0px 3px 6px rgba(0,0,0,0.25) !important;
    transition: all 0.3s ease-in-out !important;
}
div.stDownloadButton > button:hover {
    background-color: #E6B800 !important;
    color: #000000 !important;
    transform: scale(1.03);
}
</style>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Envie a planilha da folha (.xls ou .xlsx)", type=["xls", "xlsx"])

def normalize_number_str(s: str):
    """Converte string numérica em float"""
    if s is None:
        return None
    s = str(s).strip()
    if s == "":
        return None
    s = s.replace(" ", "")
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    s_clean = re.sub(r"[^0-9\.]", "", s)
    if s_clean == "":
        return None
    try:
        val = float(s_clean)
        return -val if neg else val
    except:
        return None

def line_has_text(row_cells):
    """Retorna True se a linha tiver ao menos uma célula contendo letras."""
    pattern_letters = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ]")
    for cell in row_cells:
        if cell is None:
            continue
        c = str(cell).strip()
        if c == "":
            continue
        if pattern_letters.search(c):
            return True
    return False

if uploaded_file:
    try:
        # Detectar extensão e engine
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        if ext in ['.xls', '.xlsx']:
            try:
                df = pd.read_excel(uploaded_file, header=None, dtype=str, engine="openpyxl")
            except:
                df = pd.read_excel(uploaded_file, header=None, dtype=str, engine="xlrd")
        else:
            st.error("Formato de arquivo não suportado. Envie .xls ou .xlsx.")
            st.stop()

        total_rows = len(df)

        # -------------------------------
        # Buscar CNPJ e Mês/Ano (VARRE TODA A PLANILHA E PEGA A 2ª OCORRÊNCIA)
        # -------------------------------
        # texto de topo (mantido para mes/ano)
        top_area = df.iloc[:20, :20].fillna("").astype(str)
        text_join = " ".join(top_area.values.flatten())

        # varre toda a planilha para localizar TODAS as ocorrências possíveis de CNPJ
        all_text = " ".join(df.fillna("").astype(str).values.flatten())

        # padrões possíveis: com label "CNPJ", formatado ou apenas 14 dígitos
        matches_label = re.findall(r"CNPJ[:\- ]*\s*([\d\.\-/]+)", all_text, re.IGNORECASE)
        matches_formatted = re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", all_text)
        matches_plain14 = re.findall(r"\b(\d{14})\b", all_text)

        # combinar mantendo ordem de aparição: procurar por qualquer ocorrência usando um regex que capture qualquer dos padrões
        combined_pattern = re.compile(r"(CNPJ[:\- ]*\s*[\d\.\-/]+)|(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})|(\b\d{14}\b)", re.IGNORECASE)
        combined = []
        for m in combined_pattern.finditer(all_text):
            txt = m.group(0)
            # extrai apenas dígitos
            digits = re.sub(r"\D", "", txt)
            if len(digits) == 14:
                combined.append(digits)

        # Agora escolhe a segunda ocorrência (índice 1)
        if len(combined) >= 2:
            cnpj = combined[1]
        elif len(combined) == 1:
            cnpj = combined[0]
        else:
            # fallback: tentar no text_join (top area) como antes
            cnpjs_top = re.findall(r"CNPJ[:\- ]*\s*([\d\.\-/]+)", text_join, re.IGNORECASE)
            if len(cnpjs_top) >= 2:
                cnpj = re.sub(r"\D", "", cnpjs_top[1])
            elif len(cnpjs_top) == 1:
                cnpj = re.sub(r"\D", "", cnpjs_top[0])
            else:
                st.warning("⚠️ Não foi possível localizar dois CNPJs; usando CNPJ padrão.")
                cnpj = "00000000000000"

        # Mês/Ano (mantém seu código original)
        mesano_match = re.search(r"M[eê]s/?Ano[:\- ]*\s*([0-9]{2}/[0-9]{4})", text_join, re.IGNORECASE)

        if mesano_match:
            mesano = mesano_match.group(1)
            mes, ano = [int(x) for x in mesano.split('/')]
            data_ini = datetime(ano, mes, 1)
            data_fim = datetime(ano, mes + 1, 1) - pd.Timedelta(days=1) if mes < 12 else datetime(ano, 12, 31)
        else:
            st.warning("⚠️ Não foi possível identificar Mês/Ano automaticamente.")
            data_ini = datetime.now()
            data_fim = datetime.now()

        # Procurar blocos "TOTAL GERAL"
        total_geral_idxs = []
        for idx in range(total_rows):
            row_text = " ".join([str(x) for x in df.iloc[idx, :20].fillna("").astype(str)])
            if re.search(r"total\s*geral", row_text, re.IGNORECASE):
                total_geral_idxs.append(idx)

        if not total_geral_idxs:
            st.error("❌ Nenhuma linha contendo 'TOTAL GERAL' encontrada.")
        else:
            found_all = []

            for t_idx in total_geral_idxs:
                start_idx = t_idx + 1
                for idx in range(start_idx, total_rows):
                    row_cells = df.iloc[idx].fillna("").astype(str).tolist()

                    if all(str(x).strip() == "" for x in row_cells):
                        break

                    first_non_empty = None
                    for cell in row_cells:
                        if str(cell).strip() != "":
                            first_non_empty = str(cell).strip()
                            break

                    if not first_non_empty:
                        continue

                    m_code = re.match(r'^(\d{3})\b', first_non_empty)
                    if not m_code:
                        continue

                    codigo = m_code.group(1)

                    if not line_has_text(row_cells):
                        continue

                    valores_na_linha = []
                    for cell in row_cells:
                        v = normalize_number_str(cell)
                        if v is not None and 0.01 <= abs(v) < 1e8:
                            valores_na_linha.append(v)

                    if not valores_na_linha:
                        continue

                    valor = valores_na_linha[-1]
                    found_all.append((codigo, valor))

            if not found_all:
                st.error("❌ Nenhum código/valor válido encontrado após 'TOTAL GERAL'.")
            else:
                df_found = pd.DataFrame(found_all, columns=["codigo", "valor"])
                df_group = df_found.groupby("codigo", as_index=False)["valor"].sum()
                df_group = df_group.sort_values("codigo")

                df_group["valor_fmt"] = df_group["valor"].apply(lambda x: f"{x:.2f}".replace(".", ","))

                header_line = f"{cnpj}|{data_ini.strftime('%d%m%Y')}|{data_fim.strftime('%d%m%Y')}|"
                txt_lines = [header_line] + [f"{r['codigo']}|{r['valor_fmt']}|" for _, r in df_group.iterrows()]
                txt_output = "\n".join(txt_lines) + "\n"

                st.success("✅ TXT gerado (apenas linhas com código, descrição e valor).")
                st.dataframe(df_group[["codigo", "valor_fmt"]].rename(columns={"valor_fmt": "valor"}), use_container_width=True)
                st.text_area("📄 Prévia do TXT:", txt_output, height=300)

                buffer = BytesIO()
                buffer.write(txt_output.encode("utf-8"))
                buffer.seek(0)

                st.download_button(
                    "💾 Baixar Arquivo TXT",
                    data=buffer,
                    file_name=f"{cnpj[:8]}-{data_ini.strftime('%m%Y')}.txt",
                    mime="text/plain"
                )

    except Exception as e:
        st.error(f"⚠️ Erro ao processar o arquivo: {e}")
