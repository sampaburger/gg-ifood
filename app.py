import re
from io import BytesIO
from typing import List, Optional, Tuple

import altair as alt
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Analytics - iFood @gianmirante", layout="wide")

# -----------------------------
# Utilitários de formatação
# -----------------------------
def brl(valor):
    if pd.isna(valor):
        valor = 0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(valor):
    if pd.isna(valor):
        valor = 0
    return f"{valor:.2%}".replace(".", ",")


def normalizar(txt):
    txt = str(txt).strip().lower()
    mapa = str.maketrans("áàãâäéèêëíìîïóòõôöúùûüç", "aaaaaeeeeiiiiooooouuuuc")
    txt = txt.translate(mapa)
    txt = re.sub(r"\s+", " ", txt)
    return txt


def to_number(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return 0.0
    s = s.replace("R$", "").replace(" ", "")
    # formato brasileiro 1.234,56
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^0-9\.-]", "", s)
    try:
        return float(s)
    except Exception:
        return 0.0


def ler_abas_excel(uploaded_file):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="openpyxl")


def inferir_tabela(uploaded_file, palavras_chave: List[str]) -> pd.DataFrame:
    """Tenta encontrar a linha de cabeçalho em qualquer aba."""
    abas = ler_abas_excel(uploaded_file)
    melhores = []
    for nome, bruto in abas.items():
        bruto = bruto.dropna(how="all").dropna(axis=1, how="all")
        for i in range(min(len(bruto), 80)):
            linha_norm = [normalizar(x) for x in bruto.iloc[i].tolist()]
            score = sum(any(chave in cel for cel in linha_norm) for chave in palavras_chave)
            if score > 0:
                melhores.append((score, nome, i, bruto))
    if not melhores:
        # fallback simples
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, engine="openpyxl")
    melhores.sort(key=lambda x: x[0], reverse=True)
    _, _, idx, bruto = melhores[0]
    header = bruto.iloc[idx].tolist()
    df = bruto.iloc[idx + 1:].copy()
    df.columns = [str(c).strip() for c in header]
    df = df.dropna(how="all")
    return df


def achar_coluna(df: pd.DataFrame, termos: List[str]) -> Optional[str]:
    cols_norm = {col: normalizar(col) for col in df.columns}
    for termo in termos:
        termo_n = normalizar(termo)
        for col, col_n in cols_norm.items():
            if termo_n in col_n:
                return col
    return None


def soma_coluna(df: pd.DataFrame, col: Optional[str]) -> float:
    if col is None or col not in df.columns:
        return 0.0
    return df[col].apply(to_number).sum()


# -----------------------------
# Extração: relatório de pedidos
# -----------------------------
def extrair_pedidos(uploaded_file):
    df = inferir_tabela(uploaded_file, ["pedido", "status", "forma", "pagamento", "total", "item"])

    col_status = achar_coluna(df, ["status final", "status do pedido", "status"])
    col_forma = achar_coluna(df, ["forma de pagamento", "meio de pagamento", "pagamento"])
    col_total_cliente = achar_coluna(df, ["total pago pelo cliente", "valor pago pelo cliente"])
    col_itens = achar_coluna(df, ["valor dos itens", "valor itens", "itens"])
    col_pedido = achar_coluna(df, ["id do pedido", "pedido"])

    df_calc = df.copy()
    if col_status:
        status_norm = df_calc[col_status].map(normalizar)
        mask_concluido = status_norm.str.contains("concluido|concluído|entregue|finalizado", regex=True, na=False)
        # Se não encontrar concluído, usa tudo exceto cancelados.
        if mask_concluido.sum() == 0:
            mask_concluido = ~status_norm.str.contains("cancel", regex=True, na=False)
    else:
        mask_concluido = pd.Series(True, index=df_calc.index)

    faturamento_itens = soma_coluna(df_calc[mask_concluido], col_itens)
    pedidos_concluidos = int(mask_concluido.sum()) if col_pedido is None else int(df_calc.loc[mask_concluido, col_pedido].nunique())

    recebido_direto = 0.0
    if col_forma and col_total_cliente:
        forma_norm = df_calc[col_forma].map(normalizar)
        # Recebido direto pela loja: apenas bandeiras externas pagas diretamente à loja.
        # IMPORTANTE: não considerar "iFood Meal Voucher" nem outros vouchers do próprio iFood,
        # porque esses entram no repasse do iFood e inflam o recebido direto.
        termos_direto = r"sodexo|ticket|\bvr\b|alelo"
        mask_direto = forma_norm.str.contains(termos_direto, regex=True, na=False)
        mask_direto = mask_direto & ~forma_norm.str.contains("ifood", regex=False, na=False)
        recebido_direto = soma_coluna(df_calc[mask_concluido & mask_direto], col_total_cliente)

    return {
        "df": df,
        "faturamento_itens": faturamento_itens,
        "pedidos_concluidos": pedidos_concluidos,
        "recebido_direto_loja": recebido_direto,
    }


# -----------------------------
# Extração: relatório desempenho/vendas
# -----------------------------
def extrair_desempenho(uploaded_file):
    abas = ler_abas_excel(uploaded_file)
    total_vendas = 0.0
    total_pedidos = 0

    for nome, bruto in abas.items():
        bruto = bruto.dropna(how="all").dropna(axis=1, how="all")
        # Procura linha de cabeçalho com logística/pedidos/vendas.
        for i in range(min(len(bruto), 80)):
            row_norm = [normalizar(x) for x in bruto.iloc[i].tolist()]
            if any("logistica" in x for x in row_norm) and any("venda" in x for x in row_norm):
                df = bruto.iloc[i + 1:].copy()
                df.columns = [str(c).strip() for c in bruto.iloc[i].tolist()]
                df = df.dropna(how="all")
                col_log = achar_coluna(df, ["logistica"])
                col_ped = achar_coluna(df, ["pedidos", "quantidade"])
                col_vendas = achar_coluna(df, ["valor total de vendas", "vendas", "valor"])
                if col_vendas:
                    # Se existir linha Total, usa ela. Senão soma linhas de logística.
                    if col_log:
                        mask_total = df[col_log].map(normalizar).str.contains("total", na=False)
                        if mask_total.any():
                            total_vendas = df.loc[mask_total, col_vendas].apply(to_number).sum()
                            if col_ped:
                                total_pedidos = int(df.loc[mask_total, col_ped].apply(to_number).sum())
                            return {"faturamento_comercial": total_vendas, "pedidos_desempenho": total_pedidos}
                    total_vendas = df[col_vendas].apply(to_number).sum()
                    if col_ped:
                        total_pedidos = int(df[col_ped].apply(to_number).sum())
                    return {"faturamento_comercial": total_vendas, "pedidos_desempenho": total_pedidos}

    # Fallback: tenta achar um número perto de uma célula Total.
    for nome, bruto in abas.items():
        for r in range(len(bruto)):
            for c in range(len(bruto.columns)):
                if "total" in normalizar(bruto.iat[r, c]):
                    nums = []
                    for cc in range(c + 1, min(c + 6, len(bruto.columns))):
                        n = to_number(bruto.iat[r, cc])
                        if n > 0:
                            nums.append(n)
                    if nums:
                        total_vendas = max(nums)
                        total_pedidos = int(min(nums)) if len(nums) > 1 and min(nums) < 100000 else 0
                        return {"faturamento_comercial": total_vendas, "pedidos_desempenho": total_pedidos}

    return {"faturamento_comercial": 0.0, "pedidos_desempenho": 0}


# -----------------------------
# Extração: financeiro/conciliação
# -----------------------------
def extrair_financeiro(uploaded_file):
    df = inferir_tabela(uploaded_file, ["fato", "gerador", "valor", "descricao", "lançamento", "lancamento"])

    col_fato = achar_coluna(df, ["fato gerador", "fato"])
    col_valor = achar_coluna(df, ["valor", "valor liquido", "valor líquido", "total"])
    col_desc = achar_coluna(df, ["descrição", "descricao", "lançamento", "lancamento", "tipo", "categoria"])

    # Repasses: usar a própria coluna do iFood "impacto_no_repasse" quando existir.
    # Esta regra bate com o painel financeiro oficial do iFood.
    repasse = 0.0
    col_impacto = achar_coluna(df, ["impacto_no_repasse", "impacto no repasse", "impacta repasse"])
    if col_impacto and col_valor:
        impacto_norm = df[col_impacto].map(normalizar)
        mask_repasse = impacto_norm.eq("sim")
        repasse = soma_coluna(df[mask_repasse], col_valor)
    elif col_fato and col_valor:
        # Fallback para exportações antigas sem a coluna impacto_no_repasse.
        fato_norm = df[col_fato].map(normalizar)
        mask_repasse = fato_norm.str.fullmatch("venda|cancelamento total|cancelamento parcial", na=False)
        if mask_repasse.sum() == 0:
            mask_repasse = fato_norm.str.contains("venda|cancelamento total|cancelamento parcial", regex=True, na=False)
        repasse = soma_coluna(df[mask_repasse], col_valor)

    texto_busca = ""
    if col_desc:
        texto_busca = df[col_desc].map(normalizar)
    elif col_fato:
        texto_busca = df[col_fato].map(normalizar)
    else:
        texto_busca = pd.Series(["" for _ in range(len(df))], index=df.index)

    valor_series = df[col_valor].apply(to_number) if col_valor else pd.Series([0.0] * len(df), index=df.index)

    # Investimento Comercial: SOMENTE promoções custeadas pela loja + anúncios.
    # Importante: não incluir promoções custeadas pelo iFood, taxas ou comissões.
    # Usamos abs(soma líquida) para respeitar estornos/créditos do próprio iFood.
    mask_promocoes_loja = texto_busca.str.contains(
        "promocao custeada pela loja|promoção custeada pela loja",
        regex=True,
        na=False,
    )
    mask_anuncios = texto_busca.str.contains("anuncio|anuncios|pacote de anuncio|pacote de anúncios", regex=True, na=False)
    promocoes_loja = abs(valor_series[mask_promocoes_loja].sum())
    anuncios = abs(valor_series[mask_anuncios].sum())
    investimento = promocoes_loja + anuncios

    # Taxas e comissões conforme card do iFood: comissões + taxa de transação.
    # Exclui taxa de entrega, taxa de serviço e parcelamento para não distorcer o painel gerencial.
    mask_taxa_comissao = texto_busca.str.contains("comiss|taxa de transa", regex=True, na=False)
    # O card oficial do iFood considera o valor líquido dessas linhas;
    # por isso usamos abs(soma líquida), e não soma dos absolutos linha a linha.
    taxas_comissoes = abs(valor_series[mask_taxa_comissao].sum())

    return {
        "df": df,
        "repasse_liquido": repasse,
        "promocoes_loja": promocoes_loja,
        "anuncios": anuncios,
        "taxas_comissoes": taxas_comissoes,
        "total_gasto_plataforma": promocoes_loja + anuncios + taxas_comissoes,
    }


# -----------------------------
# Interface
# -----------------------------
st.title("📊 Analytics - Resumo iFood @gianmirante")
st.caption("Upload dos relatórios iFood para gerar o resumo gerencial da unidade. Este programa apenas lê as planilhas enviadas e não substitui o portal do iFood")

with st.expander("📖 Como utilizar"):
    st.markdown("""
### 1. Relatório de Pedidos iFood
Entre no Portal do iFood:

**Menu → Pedidos → Selecione o período desejado → Exportar**

### 2. Relatório Financeiro / Conciliação iFood
Entre no Portal do iFood:

**Financeiro → Selecione o mês desejado → Exportar**

### 3. Relatório de Desempenho / Vendas iFood
Entre no Portal do iFood:

**Desempenho → Vendas → Selecione o período desejado → Exportar**

### 4. Atenção aos períodos
Os três arquivos devem possuir exatamente o mesmo período de análise.

Exemplo correto: todos os arquivos de **Abril/2026**.
""")

col_up1, col_up2, col_up3 = st.columns(3)
with col_up1:
    arquivo_pedidos = st.file_uploader("Relatório de Pedidos", type=["xlsx", "xls"], key="pedidos")
with col_up2:
    arquivo_financeiro = st.file_uploader("Financeiro / Conciliação", type=["xlsx", "xls"], key="financeiro")
with col_up3:
    arquivo_desempenho = st.file_uploader("Desempenho / Vendas", type=["xlsx", "xls"], key="desempenho")

if arquivo_pedidos and arquivo_financeiro and arquivo_desempenho:
    try:
        pedidos = extrair_pedidos(arquivo_pedidos)
        financeiro = extrair_financeiro(arquivo_financeiro)
        desempenho = extrair_desempenho(arquivo_desempenho)

        faturamento_operacional = pedidos["faturamento_itens"]
        faturamento_comercial = desempenho["faturamento_comercial"]
        pedidos_qtd = desempenho["pedidos_desempenho"] or pedidos["pedidos_concluidos"]
        ticket_medio = faturamento_comercial / pedidos_qtd if pedidos_qtd else 0.0

        repasse_liquido = financeiro["repasse_liquido"]
        recebido_direto = pedidos["recebido_direto_loja"]
        repasse_total = repasse_liquido + recebido_direto
        pct_repasse_total = repasse_total / faturamento_operacional if faturamento_operacional else 0.0

        promocoes_loja = financeiro["promocoes_loja"]
        anuncios = financeiro["anuncios"]
        taxa = financeiro["taxas_comissoes"]
        total_gasto_plataforma = financeiro["total_gasto_plataforma"]

        pct_promocoes = promocoes_loja / faturamento_operacional if faturamento_operacional else 0.0
        pct_anuncios = anuncios / faturamento_operacional if faturamento_operacional else 0.0
        pct_taxa = taxa / faturamento_operacional if faturamento_operacional else 0.0
        pct_total_gasto = total_gasto_plataforma / faturamento_operacional if faturamento_operacional else 0.0

        st.divider()
        st.subheader("Resumo principal")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Faturamento Comercial", brl(faturamento_comercial))
        c2.metric("Faturamento Operacional", brl(faturamento_operacional))
        c3.metric("Pedidos", f"{pedidos_qtd:,}".replace(",", "."))
        c4.metric("Ticket Médio", brl(ticket_medio))

        st.subheader("Financeiro (não inclui taxa de antecipação)")
        f1, f2, f3, f4 = st.columns(4)
        f1.metric("Repasse Líquido Total", brl(repasse_total))
        f2.metric("Repasse Líquido", brl(repasse_liquido))
        f3.metric("Recebido Direto pela Loja", brl(recebido_direto))
        f4.metric("% Repasse Total / Fat. Itens", pct(pct_repasse_total))

        st.subheader("Eficiência iFood")
        e1, e2, e3, e4 = st.columns(4)
        e1.metric("Promoções Custeadas pela Loja", f"{brl(promocoes_loja)} ({pct(pct_promocoes)})")
        e2.metric("Anúncios", f"{brl(anuncios)} ({pct(pct_anuncios)})")
        e3.metric("Taxas e Comissões", f"{brl(taxa)} ({pct(pct_taxa)})")
        e4.metric("Total Gasto na Plataforma", f"{brl(total_gasto_plataforma)} ({pct(pct_total_gasto)})")

        st.caption("Distribuição visual sobre o faturamento operacional: a barra inteira representa 100% do faturamento de itens.")

        restante_operacional = max(faturamento_operacional - total_gasto_plataforma, 0)
        grafico_df = pd.DataFrame([
            {"Base": "Faturamento Operacional", "Categoria": "Promoções", "Valor": promocoes_loja, "Percentual": pct_promocoes * 100},
            {"Base": "Faturamento Operacional", "Categoria": "Anúncios", "Valor": anuncios, "Percentual": pct_anuncios * 100},
            {"Base": "Faturamento Operacional", "Categoria": "Taxas e Comissões", "Valor": taxa, "Percentual": pct_taxa * 100},
            {"Base": "Faturamento Operacional", "Categoria": "Restante", "Valor": restante_operacional, "Percentual": (restante_operacional / faturamento_operacional * 100) if faturamento_operacional else 0},
        ])

        chart = (
            alt.Chart(grafico_df)
            .mark_bar(size=44, cornerRadius=8)
            .encode(
                x=alt.X("Percentual:Q", stack="normalize", title="100% do faturamento operacional", axis=alt.Axis(format="%")),
                y=alt.Y("Base:N", title=None, axis=alt.Axis(labels=False, ticks=False)),
                color=alt.Color("Categoria:N", title=None),
                order=alt.Order("Categoria:N", sort="ascending"),
                tooltip=[
                    alt.Tooltip("Categoria:N", title="Indicador"),
                    alt.Tooltip("Valor:Q", title="Valor", format=",.2f"),
                    alt.Tooltip("Percentual:Q", title="% do faturamento", format=".2f"),
                ],
            )
            .properties(height=120)
        )
        st.altair_chart(chart, use_container_width=True)

        resumo = pd.DataFrame([
            ["Faturamento Comercial", faturamento_comercial],
            ["Faturamento Operacional", faturamento_operacional],
            ["Pedidos", pedidos_qtd],
            ["Ticket Médio", ticket_medio],
            ["Repasse Líquido Total", repasse_total],
            ["Repasse Líquido", repasse_liquido],
            ["Recebido Direto pela Loja", recebido_direto],
            ["% Repasse Total / Fat. Itens", pct_repasse_total],
            ["Promoções Custeadas pela Loja", promocoes_loja],
            ["% Promoções / Fat. Itens", pct_promocoes],
            ["Anúncios", anuncios],
            ["% Anúncios / Fat. Itens", pct_anuncios],
            ["Taxas e Comissões", taxa],
            ["% Taxas e Comissões / Fat. Itens", pct_taxa],
            ["Total Gasto na Plataforma", total_gasto_plataforma],
            ["% Total Gasto / Fat. Itens", pct_total_gasto],
        ], columns=["Indicador", "Valor"])

        st.divider()
        st.subheader("Indicadores (KPI'S)")
        tabela_view = resumo.copy()
        tabela_view["Valor"] = tabela_view.apply(
            lambda r: pct(r["Valor"]) if "%" in r["Indicador"] else (f"{int(r['Valor']):,}".replace(",", ".") if r["Indicador"] == "Pedidos" else brl(r["Valor"])),
            axis=1,
        )
        st.dataframe(tabela_view, use_container_width=True, hide_index=True)

        whatsapp = f"""📊 Resumo iFood

📈 Faturamento Comercial: {brl(faturamento_comercial)}
🍔 Faturamento Operacional: {brl(faturamento_operacional)}
📦 Pedidos: {str(f'{pedidos_qtd:,}').replace(',', '.')}
🎟️ Ticket Médio: {brl(ticket_medio)}

💰 Repasse Líquido Total: {brl(repasse_total)}
🏦 Repasse Líquido: {brl(repasse_liquido)}
💳 Recebido Direto pela Loja: {brl(recebido_direto)}
📊 Repasse Total / Fat. Itens: {pct(pct_repasse_total)}

🎯 Promoções Custeadas pela Loja: {brl(promocoes_loja)} ({pct(pct_promocoes)})
📢 Anúncios: {brl(anuncios)} ({pct(pct_anuncios)})
🏷️ Taxas e Comissões: {brl(taxa)} ({pct(pct_taxa)})
📊 Total Gasto na Plataforma: {brl(total_gasto_plataforma)} ({pct(pct_total_gasto)})"""

        st.subheader("Resumo para WhatsApp")
        st.text_area("Copiar resumo", whatsapp, height=260)

        csv = resumo.to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button("Baixar resumo em CSV", data=csv, file_name="resumo_ifood_ggk.csv", mime="text/csv")

        with st.expander("🔍 Auditoria das fórmulas"):
            st.markdown("""
- **Faturamento Comercial:** relatório de Desempenho / Vendas.
- **Faturamento Operacional:** valor dos itens do relatório de Pedidos.
- **Recebido Direto pela Loja:** `Total pago pelo cliente` filtrando `Forma de pagamento` contendo **SODEXO**, **TICKET**, **VR** ou **ALELO**, apenas pedidos concluídos. O sistema exclui vouchers do próprio iFood, como **iFood Meal Voucher**, porque eles entram no repasse do iFood.
- **Repasse Líquido:** conciliação financeira somando os lançamentos onde `impacto_no_repasse` = **SIM**. Se a coluna não existir, usa fallback por `Fato Gerador` = **Venda**, **Cancelamento Total** e **Cancelamento Parcial**.
- **Promoções Custeadas pela Loja:** promoções financiadas pela loja na conciliação financeira.
- **Anúncios:** pacote/anúncios iFood na conciliação financeira.
- **Taxas e Comissões:** comissões + taxa de transação, usando soma líquida para bater com o card oficial do iFood.
- **Total Gasto na Plataforma:** promoções custeadas pela loja + anúncios + taxas e comissões.
""")

    except Exception as e:
        st.error("Não foi possível processar os arquivos. Confira se os três relatórios são exportações originais do iFood e do mesmo período.")
        st.exception(e)
else:
    st.info("Faça upload dos 3 arquivos para gerar o relatório.")
