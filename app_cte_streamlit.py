import io
import re
import xml.etree.ElementTree as ET
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP

import pandas as pd
import streamlit as st

NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}


def brl(v: float) -> str:
    s = f"{v:,.2f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def numero_br(v: float, casas=3) -> str:
    s = f"{v:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def parse_nfe(uploaded_file):
    """Extrai os campos necessários de uma NF-e XML (layout 4.00)."""
    data = uploaded_file.getvalue()
    root = ET.fromstring(data)

    def text(path, default=""):
        el = root.find(path, NS)
        return (el.text or "").strip() if el is not None else default

    peso_bruto = Decimal("0")
    pesos = root.findall(".//nfe:transp/nfe:vol/nfe:pesoB", NS)
    for el in pesos:
        if el.text:
            try:
                peso_bruto += Decimal(el.text.strip().replace(",", "."))
            except Exception:
                pass

    inf_nfe = root.find(".//nfe:infNFe", NS)
    chave_inf = ""
    if inf_nfe is not None:
        chave_inf = re.sub(r"\D", "", inf_nfe.attrib.get("Id", "").replace("NFe", ""))

    return {
        "arquivo": uploaded_file.name,
        "nf": text(".//nfe:ide/nfe:nNF"),
        "chave": text(".//nfe:protNFe/nfe:infProt/nfe:chNFe") or chave_inf,
        "emitente_cidade": text(".//nfe:emit/nfe:enderEmit/nfe:xMun"),
        "origem_uf": text(".//nfe:emit/nfe:enderEmit/nfe:UF"),
        "destinatario": text(".//nfe:dest/nfe:xNome"),
        "destino_cidade": text(".//nfe:dest/nfe:enderDest/nfe:xMun"),
        "destino_uf": text(".//nfe:dest/nfe:enderDest/nfe:UF"),
        "peso_bruto": float(peso_bruto),
        "valor_nf": float(Decimal(text(".//nfe:total/nfe:ICMSTot/nfe:vNF", "0") or "0")),
    }


def allocate_total_by_weight(total_value: float, weights):
    """Rateia em centavos, garantindo que a soma final seja exatamente o total informado."""
    total = Decimal(str(total_value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    total_cents = int(total * 100)

    dec_weights = [Decimal(str(w)) for w in weights]
    sum_w = sum(dec_weights)
    if sum_w <= 0:
        raise ValueError("O peso bruto total precisa ser maior que zero.")

    raw_cents = [Decimal(total_cents) * w / sum_w for w in dec_weights]
    base_cents = [int(x.to_integral_value(rounding=ROUND_DOWN)) for x in raw_cents]
    missing = total_cents - sum(base_cents)

    order = sorted(
        range(len(raw_cents)),
        key=lambda i: raw_cents[i] - Decimal(base_cents[i]),
        reverse=True,
    )
    for i in order[:missing]:
        base_cents[i] += 1

    return [c / 100 for c in base_cents]


def excel_bytes(detail_df, city_df, total_frete, aliquota):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        detail_df.to_excel(writer, sheet_name="Rateio por NF", index=False)
        city_df.to_excel(writer, sheet_name="Resumo por cidade", index=False)
        resumo = pd.DataFrame(
            {
                "Campo": ["Frete total", "Alíquota ICMS", "Peso bruto total", "Qtde NF-e"],
                "Valor": [
                    total_frete,
                    aliquota,
                    float(detail_df["Peso bruto (kg)"].sum()),
                    int(len(detail_df)),
                ],
            }
        )
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
    buffer.seek(0)
    return buffer.getvalue()


st.set_page_config(
    page_title="Rateio CT-e por Peso",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        :root {
            --ink: #EDEEF0;
            --ink-soft: #9A9FA8;
            --ink-faint: #6B7078;
            --line: #2A2D33;
            --surface: #16181C;
            --surface-raised: #1C1F24;
            --bg: #0E1013;
            --accent: #34D399;
            --accent-ink: #06281B;
        }

        html, body, [class*="css"] {
            font-family: -apple-system, "Inter", "Segoe UI", sans-serif;
        }

        .stApp {
            background: var(--bg);
            color: var(--ink);
        }

        .block-container {
            padding-top: 2.4rem;
            padding-bottom: 4rem;
            max-width: 1240px;
        }

        h1, h2, h3, h4, h5, p, span, label, li {
            color: var(--ink);
        }

        /* Sidebar */
        [data-testid="stSidebar"] {
            background: var(--surface);
            border-right: 1px solid var(--line);
        }

        [data-testid="stSidebar"] * {
            color: var(--ink);
        }

        [data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] small {
            color: var(--ink-soft) !important;
        }

        /* Header */
        .app-header {
            border-bottom: 1px solid var(--line);
            padding-bottom: 20px;
            margin-bottom: 32px;
        }

        .app-kicker {
            font-size: 0.72rem;
            letter-spacing: .14em;
            text-transform: uppercase;
            color: var(--ink-soft);
            font-weight: 600;
            margin-bottom: 6px;
        }

        .app-title {
            font-size: 1.65rem;
            line-height: 1.2;
            font-weight: 700;
            color: var(--ink);
            margin: 0 0 6px 0;
        }

        .app-subtitle {
            font-size: 0.92rem;
            line-height: 1.5;
            color: var(--ink-soft);
            max-width: 760px;
            margin: 0;
        }

        /* Section labels */
        .section-label {
            font-size: 0.72rem;
            color: var(--ink-soft);
            text-transform: uppercase;
            letter-spacing: .1em;
            font-weight: 700;
            margin: 28px 0 12px 0;
            padding-top: 4px;
        }

        .section-label:first-of-type {
            margin-top: 0;
        }

        /* Text inputs, number inputs, selects, multiselects */
        input, textarea {
            background-color: var(--surface-raised) !important;
            color: var(--ink) !important;
            border-color: var(--line) !important;
        }

        [data-testid="stTextInput"] > div,
        [data-testid="stNumberInput"] > div,
        div[data-baseweb="select"] > div,
        div[data-baseweb="base-input"] {
            background-color: var(--surface-raised) !important;
            border-color: var(--line) !important;
            border-radius: 8px !important;
        }

        [data-testid="stNumberInput"] button {
            background-color: var(--surface-raised) !important;
            border-color: var(--line) !important;
            color: var(--ink-soft) !important;
        }

        [data-testid="stNumberInput"] button:hover {
            color: var(--ink) !important;
            border-color: var(--ink-faint) !important;
        }

        div[data-baseweb="select"] svg {
            fill: var(--ink-soft) !important;
        }

        [data-baseweb="tag"] {
            background-color: var(--surface) !important;
            border: 1px solid var(--line) !important;
        }

        [data-baseweb="tag"] span {
            color: var(--ink) !important;
        }

        [data-baseweb="popover"] li,
        ul[data-testid="stSelectboxVirtualDropdown"] li {
            background-color: var(--surface-raised) !important;
            color: var(--ink) !important;
        }

        /* Metrics */
        [data-testid="stMetric"] {
            background: var(--surface);
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 9px 12px;
            min-height: 82px;
        }

        [data-testid="stMetricLabel"] {
            color: var(--ink-soft);
            font-weight: 600;
            font-size: 0.72rem;
        }

        [data-testid="stMetricValue"] {
            color: var(--ink);
            font-weight: 700;
            font-size: 1.35rem !important;
            line-height: 1.15 !important;
            white-space: nowrap !important;
        }

        /* Buttons */
        .stButton > button {
            background: var(--surface-raised);
            color: var(--ink);
            border-radius: 8px;
            min-height: 42px;
            font-weight: 600;
            border: 1px solid var(--line);
            box-shadow: none;
        }

        .stDownloadButton > button {
            background: var(--accent);
            color: var(--accent-ink);
            border-radius: 8px;
            min-height: 42px;
            font-weight: 700;
            border: none;
            box-shadow: none;
        }

        .stDownloadButton > button:hover {
            background: #2BBE87;
            color: var(--accent-ink);
        }

        /* File uploader */
        [data-testid="stFileUploader"] {
            background: transparent;
            border-radius: 10px;
            padding: 0;
        }

        [data-testid="stFileUploaderDropzone"] {
            background: var(--surface) !important;
            border: 1px dashed var(--line) !important;
            border-radius: 10px !important;
        }

        [data-testid="stFileUploaderDropzone"] * {
            color: var(--ink) !important;
        }

        [data-testid="stFileUploaderDropzone"] small,
        [data-testid="stFileUploaderDropzoneInstructions"] span {
            color: var(--ink-soft) !important;
        }

        [data-testid="stBaseButton-secondary"] {
            background: var(--surface-raised) !important;
            color: var(--ink) !important;
            border: 1px solid var(--line) !important;
        }

        /* Dataframes */
        div[data-testid="stDataFrame"] {
            border: 1px solid var(--line);
            border-radius: 10px;
            overflow: hidden;
        }

        /* Alerts (info / warning / error) */
        [data-testid="stAlert"] {
            background: var(--surface) !important;
            border: 1px solid var(--line) !important;
            border-radius: 10px !important;
        }

        [data-testid="stAlert"] * {
            color: var(--ink) !important;
        }

        [data-testid="stAlertContentInfo"] {
            border-left: 3px solid #5B9BD5 !important;
        }

        [data-testid="stAlertContentWarning"] {
            border-left: 3px solid #E0A845 !important;
        }

        [data-testid="stAlertContentError"] {
            border-left: 3px solid #E06767 !important;
        }

        .status-ok {
            background: rgba(52, 211, 153, 0.08);
            color: var(--accent);
            border: 1px solid rgba(52, 211, 153, 0.28);
            border-radius: 10px;
            padding: 13px 16px;
            font-weight: 600;
            font-size: 0.92rem;
        }

        .mini-note {
            color: var(--ink-soft);
            font-size: .84rem;
            line-height: 1.5;
        }

        .sidebar-brand {
            font-weight: 700;
            font-size: 0.98rem;
            color: var(--ink);
            margin-bottom: 2px;
        }

        .sidebar-sub {
            font-size: .8rem;
            color: var(--ink-soft);
            margin-bottom: 18px;
        }

        /* Tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 4px;
            border-bottom: 1px solid var(--line);
        }

        .stTabs [data-baseweb="tab"] {
            height: 42px;
            border-radius: 8px 8px 0 0;
            padding-left: 16px;
            padding-right: 16px;
            font-weight: 600;
            font-size: 0.88rem;
            color: var(--ink-soft);
        }

        .stTabs [aria-selected="true"] {
            color: var(--ink) !important;
        }

        .stTabs [data-baseweb="tab-highlight"] {
            background-color: var(--accent) !important;
        }

        /* Expander */
        [data-testid="stExpander"] {
            background: var(--surface);
            border: 1px solid var(--line) !important;
            border-radius: 10px;
        }

        [data-testid="stExpander"] summary {
            color: var(--ink) !important;
        }

        hr {
            border-color: var(--line) !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown('<div class="sidebar-brand">🚚 Calculadora CT-e</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">Rateio de frete por peso bruto</div>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("**Parâmetros**")
    origem_esperada = st.text_input("UF de origem", value="PE", max_chars=2).upper().strip()
    aliquota_pct = st.number_input(
        "ICMS (%)",
        min_value=0.0,
        max_value=100.0,
        value=12.0,
        step=0.1,
    )
    aliquota = aliquota_pct / 100

st.markdown(
    """
    <div class="app-header">
        <div class="app-kicker">WS Transporte · Operação Fiscal</div>
        <div class="app-title">Rateio de CT-e por peso bruto</div>
        <p class="app-subtitle">
            Importe os XMLs das NF-e, informe o valor total do frete e obtenha automaticamente o rateio por nota,
            cidade e ICMS, com conferência do valor final.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="section-label">1. Dados do frete</div>', unsafe_allow_html=True)
col_frete, col_upload = st.columns([0.34, 0.66], gap="large")

with col_frete:
    frete_total = st.number_input(
        "Valor total do frete (já com ICMS incluso)",
        min_value=0.0,
        value=None,
        step=100.0,
        format="%.2f",
        placeholder="Informe o valor do frete",
        help="Informe o valor total que deverá ser distribuído entre as NF-e pelo peso bruto.",
    )

with col_upload:
    uploads = st.file_uploader(
        "Importe uma ou várias NF-e em XML",
        type=["xml"],
        accept_multiple_files=True,
        help="Você pode selecionar vários XMLs de uma vez.",
    )

if uploads:
    registros = []
    erros = []
    for f in uploads:
        try:
            registros.append(parse_nfe(f))
        except Exception as exc:
            erros.append(f"{f.name}: {exc}")

    if erros:
        st.warning("Alguns arquivos não puderam ser lidos:\n\n" + "\n".join(f"- {e}" for e in erros))

    if registros:
        df = pd.DataFrame(registros)
        ufs_disponiveis = sorted([x for x in df["destino_uf"].dropna().unique().tolist() if x])
        default_ufs_12 = [uf for uf in ufs_disponiveis if uf != origem_esperada]

        st.markdown('<div class="section-label">2. Regra de ICMS</div>', unsafe_allow_html=True)
        ufs_12 = st.multiselect(
            "UFs de destino com ICMS de 12%",
            options=ufs_disponiveis,
            default=default_ufs_12,
            help="Selecione as UFs em que o ICMS deve ser aplicado.",
        )

        elegivel = (
            (df["origem_uf"].str.upper() == origem_esperada)
            & (df["peso_bruto"] > 0)
        )
        calc = df.loc[elegivel].copy()
        ignoradas = df.loc[~elegivel].copy()

        if calc.empty:
            st.error("Nenhuma NF-e ficou elegível para o rateio com os filtros atuais.")
            st.stop()

        if frete_total is None or frete_total <= 0:
            st.warning("Informe o valor total do frete para realizar o cálculo.")
            st.stop()

        calc["frete_rateado"] = allocate_total_by_weight(frete_total, calc["peso_bruto"].tolist())
        calc["percentual_peso"] = calc["peso_bruto"] / calc["peso_bruto"].sum()
        calc["aplica_icms_12"] = calc["destino_uf"].isin(ufs_12)
        calc["icms"] = 0.0
        calc.loc[calc["aplica_icms_12"], "icms"] = (
            calc.loc[calc["aplica_icms_12"], "frete_rateado"] * aliquota
        ).round(2)
        calc["valor_sem_icms"] = (calc["frete_rateado"] - calc["icms"]).round(2)

        total_icms = round(float(calc["icms"].sum()), 2)
        peso_total = float(calc["peso_bruto"].sum())
        cidades_total = int(calc[["destino_cidade", "destino_uf"]].drop_duplicates().shape[0])

        st.markdown('<div class="section-label">3. Resultado do rateio</div>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Frete total", brl(frete_total))
        c2.metric("Peso bruto", f"{numero_br(peso_total)} kg")
        c3.metric("NF-e no rateio", len(calc))
        c4.metric("Cidades", cidades_total)
        c5.metric("ICMS destacado", brl(total_icms))

        detalhe = calc[
            [
                "nf",
                "chave",
                "destinatario",
                "destino_cidade",
                "destino_uf",
                "aplica_icms_12",
                "peso_bruto",
                "percentual_peso",
                "frete_rateado",
                "icms",
                "valor_sem_icms",
                "valor_nf",
            ]
        ].copy()
        detalhe.columns = [
            "NF-e",
            "Chave de acesso",
            "Destinatário",
            "Cidade",
            "UF",
            "ICMS 12%?",
            "Peso bruto (kg)",
            "% do peso",
            "Valor CT-e / frete",
            "ICMS",
            "Valor sem ICMS",
            "Valor NF-e",
        ]
        detalhe["ICMS 12%?"] = detalhe["ICMS 12%?"].map({True: "Sim", False: "Não"})

        resumo_cidade = (
            detalhe.groupby(["Cidade", "UF"], as_index=False)
            .agg(
                **{
                    "Qtde NF-e": ("NF-e", "count"),
                    "Peso bruto (kg)": ("Peso bruto (kg)", "sum"),
                    "Valor CT-e / frete": ("Valor CT-e / frete", "sum"),
                    "ICMS": ("ICMS", "sum"),
                    "Valor sem ICMS": ("Valor sem ICMS", "sum"),
                }
            )
            .sort_values(["UF", "Cidade"])
        )

        tab1, tab2, tab3 = st.tabs(["Rateio por NF-e", "Resumo por cidade", "Conferência"])

        with tab1:
            st.dataframe(
                detalhe.style.format(
                    {
                        "Peso bruto (kg)": "{:,.3f}",
                        "% do peso": "{:.4%}",
                        "Valor CT-e / frete": "R$ {:,.2f}",
                        "ICMS": "R$ {:,.2f}",
                        "Valor sem ICMS": "R$ {:,.2f}",
                        "Valor NF-e": "R$ {:,.2f}",
                    }
                ),
                use_container_width=True,
                hide_index=True,
                height=420,
            )

        with tab2:
            st.dataframe(
                resumo_cidade.style.format(
                    {
                        "Peso bruto (kg)": "{:,.3f}",
                        "Valor CT-e / frete": "R$ {:,.2f}",
                        "ICMS": "R$ {:,.2f}",
                        "Valor sem ICMS": "R$ {:,.2f}",
                    }
                ),
                use_container_width=True,
                hide_index=True,
                height=360,
            )

        with tab3:
            total_rateado = float(detalhe["Valor CT-e / frete"].sum())
            diferenca = round(total_rateado - float(frete_total), 2)
            st.markdown(
                f"""
                <div class="status-ok">
                    Rateio conferido: {brl(total_rateado)} — diferença de {brl(diferenca)} em relação ao frete informado.
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.write("")
            conf1, conf2, conf3 = st.columns(3)
            conf1.metric("Total rateado", brl(total_rateado))
            conf2.metric("Total informado", brl(frete_total))
            conf3.metric("Diferença", brl(diferenca))

        st.write("")
        arquivo_excel = excel_bytes(detalhe, resumo_cidade, frete_total, aliquota)
        _, d2 = st.columns([0.72, 0.28])
        with d2:
            st.download_button(
                "Baixar resultado em Excel",
                data=arquivo_excel,
                file_name="rateio_cte_por_peso.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        if not ignoradas.empty:
            with st.expander(f"NF-e fora do cálculo ({len(ignoradas)})"):
                st.dataframe(
                    ignoradas[["nf", "origem_uf", "destino_cidade", "destino_uf", "peso_bruto"]].rename(
                        columns={
                            "nf": "NF-e",
                            "origem_uf": "Origem UF",
                            "destino_cidade": "Cidade",
                            "destino_uf": "Destino UF",
                            "peso_bruto": "Peso bruto (kg)",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

else:
    st.markdown("### Comece importando os XMLs")
    st.info("Selecione as NF-e em XML acima. Assim que os arquivos forem carregados, o cálculo aparece automaticamente.")
