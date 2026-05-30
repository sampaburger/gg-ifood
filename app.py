import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Resumo iFood GGK", page_icon="🍔", layout="wide")

# =========================
# Utilidades
# =========================
def read_ifood_xlsx(uploaded_file):
    """Lê arquivos XLSX exportados pelo iFood mesmo quando a dimensão interna vem como A1."""
    data = uploaded_file.read()
    wb = load_workbook(BytesIO(data), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    rows = list(ws.iter_rows(values_only=True))
    rows = [r for r in rows if any(v is not None for v in r)]
    if not rows:
        return pd.DataFrame()
    header = [str(x).strip() if x is not None else f"col_{i}" for i, x in enumerate(rows[0])]
    return pd.DataFrame(rows[1:], columns=header)

def money(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def pct(v):
    try:
        return f"{float(v):.1f}%".replace(".", ",")
    except Exception:
        return "0,0%"

def num(v):
    try:
        return f"{int(round(float(v), 0)):,}".replace(",", ".")
    except Exception:
        return "0"

def to_number(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

def sum_col(df, col, mask=None):
    if col not in df.columns:
        return 0.0
    s = to_number(df[col])
    if mask is not None:
        s = s[mask]
    return float(s.sum())

def sum_financeiro(fin, termos, absolute=True):
    if fin.empty or "descricao_lancamento" not in fin.columns or "valor" not in fin.columns:
        return 0.0
    desc = fin["descricao_lancamento"].fillna("").astype(str).str.lower()
    mask = pd.Series(False, index=fin.index)
    for termo in termos:
        mask = mask | desc.str.contains(termo.lower(), regex=False)
    valor = to_number(fin.loc[mask, "valor"]).sum()
    return float(abs(valor) if absolute else valor)

def montar_resumo(pedidos, financeiro):
    status_col = "STATUS FINAL DO PEDIDO"
    if status_col in pedidos.columns:
        status = pedidos[status_col].fillna("").astype(str).str.upper().str.strip()
        mask_concluido = status.eq("CONCLUIDO")
        pedidos_concluidos = int(mask_concluido.sum())
        pedidos_cancelados = int(status.eq("CANCELADO").sum())
        cancelamentos_parciais = int(status.eq("CANCELAMENTO PARCIAL").sum())
    else:
        mask_concluido = pd.Series(True, index=pedidos.index)
        pedidos_concluidos = len(pedidos)
        pedidos_cancelados = 0
        cancelamentos_parciais = 0

    pedidos_totais = len(pedidos)
    fat_itens = sum_col(pedidos, "VALOR DOS ITENS (R$)", mask_concluido)
    pago_cliente_ped = sum_col(pedidos, "TOTAL PAGO PELO CLIENTE (R$)", mask_concluido)
    taxa_entrega_cliente = sum_col(pedidos, "TAXA DE ENTREGA PAGA PELO CLIENTE (R$)", mask_concluido)
    incentivo_ifood_ped = sum_col(pedidos, "INCENTIVO PROMOCIONAL DO IFOOD (R$)", mask_concluido)
    incentivo_loja_ped = sum_col(pedidos, "INCENTIVO PROMOCIONAL DA LOJA (R$)", mask_concluido)
    taxa_servico_ped = sum_col(pedidos, "TAXA DE SERVIÇO (R$)", mask_concluido)
    taxas_comissoes_ped = abs(sum_col(pedidos, "TAXAS E COMISSOES (R$)", mask_concluido))
    valor_liquido_ped = sum_col(pedidos, "VALOR LIQUIDO (R$)", mask_concluido)

    # Financeiro: usa as descrições do extrato de conciliação
    entrada_financeira = sum_financeiro(financeiro, ["Entrada Financeira"], absolute=False)
    incentivos_ifood = sum_financeiro(financeiro, ["Promoção custeada pelo iFood"], absolute=True)
    ressarcimentos = sum_financeiro(financeiro, ["Ressarcimento"], absolute=True)
    promocoes_loja = sum_financeiro(financeiro, ["Promoção custeada pela loja"], absolute=True)
    anuncios = sum_financeiro(financeiro, ["pacote de anúncios"], absolute=True)
    comissoes = sum_financeiro(financeiro, ["Comissão"], absolute=True)
    taxas = sum_financeiro(financeiro, ["Taxa entrega iFood", "Taxa de transação", "Taxa de serviço iFood", "Taxa de conveniência"], absolute=True)

    # Repasse líquido: soma geral do financeiro, mas exclui linhas que não impactam o repasse se a coluna existir
    if "impacto_no_repasse" in financeiro.columns and "valor" in financeiro.columns:
        impacto = financeiro["impacto_no_repasse"].fillna("").astype(str).str.upper().str.strip()
        base_repasse = financeiro.loc[impacto.eq("SIM"), "valor"]
        repasse_liquido = float(to_number(base_repasse).sum())
    elif "valor" in financeiro.columns:
        repasse_liquido = float(to_number(financeiro["valor"]).sum())
    else:
        repasse_liquido = 0.0

    ticket_medio = fat_itens / pedidos_concluidos if pedidos_concluidos else 0
    pedidos_dia = pedidos_concluidos / 30
    fat_dia = fat_itens / 30
    repasse_dia = repasse_liquido / 30
    repasse_pedido = repasse_liquido / pedidos_concluidos if pedidos_concluidos else 0

    resumo = {
        "Pedidos totais": pedidos_totais,
        "Pedidos concluídos": pedidos_concluidos,
        "Pedidos cancelados": pedidos_cancelados,
        "Cancelamentos parciais": cancelamentos_parciais,
        "Ticket médio": ticket_medio,
        "Faturamento bruto operacional (itens)": fat_itens,
        "Valor pago pelos clientes (pedidos)": pago_cliente_ped,
        "Venda recebida dos clientes (financeiro)": entrada_financeira,
        "Taxa de entrega paga pelos clientes": taxa_entrega_cliente,
        "Incentivos iFood (financeiro)": incentivos_ifood,
        "Incentivos iFood (pedidos)": incentivo_ifood_ped,
        "Promoções da loja (financeiro)": promocoes_loja,
        "Promoções da loja (pedidos)": incentivo_loja_ped,
        "Ressarcimentos": ressarcimentos,
        "Anúncios iFood": anuncios,
        "Comissões iFood": comissoes,
        "Taxas iFood": taxas,
        "Taxa de serviço (pedidos)": taxa_servico_ped,
        "Taxas e comissões (pedidos)": taxas_comissoes_ped,
        "Valor líquido (pedidos)": valor_liquido_ped,
        "Repasse líquido recebido": repasse_liquido,
        "Pedidos/dia": pedidos_dia,
        "Faturamento/dia": fat_dia,
        "Repasse/dia": repasse_dia,
        "Repasse/pedido": repasse_pedido,
        "Promoções % fat.": (promocoes_loja / fat_itens * 100) if fat_itens else 0,
        "Anúncios % fat.": (anuncios / fat_itens * 100) if fat_itens else 0,
        "Comissões % fat.": (comissoes / fat_itens * 100) if fat_itens else 0,
        "Taxas % fat.": (taxas / fat_itens * 100) if fat_itens else 0,
        "Promoções + anúncios % fat.": ((promocoes_loja + anuncios) / fat_itens * 100) if fat_itens else 0,
        "Repasse líquido % fat.": (repasse_liquido / fat_itens * 100) if fat_itens else 0,
    }
    return resumo

# =========================
# Interface
# =========================
st.title("🍔 Resumo iFood - GGK / Sampa Burger")
st.caption("Suba o Relatório de Pedidos e o Extrato/Relatório Financeiro do iFood para gerar o resumo consolidado.")

with st.expander("Como usar", expanded=False):
    st.write("""
    1. Faça upload do **Relatório de Pedidos iFood**.  
    2. Faça upload do **Relatório Financeiro / Conciliação iFood**.  
    3. O sistema calcula o resumo usando o **faturamento de itens** como faturamento bruto operacional.
    """)

col1, col2 = st.columns(2)
with col1:
    arq_pedidos = st.file_uploader("Relatório de Pedidos iFood (.xlsx)", type=["xlsx"], key="pedidos")
with col2:
    arq_financeiro = st.file_uploader("Relatório Financeiro iFood (.xlsx)", type=["xlsx"], key="financeiro")

if arq_pedidos and arq_financeiro:
    try:
        pedidos = read_ifood_xlsx(arq_pedidos)
        financeiro = read_ifood_xlsx(arq_financeiro)
        resumo = montar_resumo(pedidos, financeiro)

        st.success("Resumo gerado com sucesso.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Faturamento bruto (itens)", money(resumo["Faturamento bruto operacional (itens)"]))
        c2.metric("Pedidos concluídos", num(resumo["Pedidos concluídos"]))
        c3.metric("Ticket médio", money(resumo["Ticket médio"]))
        c4.metric("Repasse líquido", money(resumo["Repasse líquido recebido"]))

        st.subheader("📊 Painel Consolidado")
        painel = pd.DataFrame([
            ["Pedidos totais", resumo["Pedidos totais"]],
            ["Pedidos concluídos", resumo["Pedidos concluídos"]],
            ["Pedidos cancelados", resumo["Pedidos cancelados"]],
            ["Cancelamentos parciais", resumo["Cancelamentos parciais"]],
            ["Ticket médio", money(resumo["Ticket médio"])],
            ["Faturamento bruto operacional (itens vendidos)", money(resumo["Faturamento bruto operacional (itens)"])],
            ["Valor pago pelos clientes (relatório de pedidos)", money(resumo["Valor pago pelos clientes (pedidos)"])],
            ["Venda recebida dos clientes (financeiro)", money(resumo["Venda recebida dos clientes (financeiro)"])],
            ["Taxa de entrega paga pelos clientes", money(resumo["Taxa de entrega paga pelos clientes"])],
            ["Incentivos iFood", money(resumo["Incentivos iFood (financeiro)"])],
            ["Ressarcimentos", money(resumo["Ressarcimentos"])],
            ["Promoções subsidiadas pela loja", money(resumo["Promoções da loja (financeiro)"])],
            ["Pacote de anúncios iFood", money(resumo["Anúncios iFood"])],
            ["Comissões iFood", money(resumo["Comissões iFood"])],
            ["Taxas iFood", money(resumo["Taxas iFood"])],
            ["Repasse líquido recebido", money(resumo["Repasse líquido recebido"])]
        ], columns=["Indicador", "Valor"])
        st.dataframe(painel, use_container_width=True, hide_index=True)

        st.subheader("📈 Indicadores Gerenciais")
        kpis = pd.DataFrame([
            ["Pedidos por dia", num(resumo["Pedidos/dia"])],
            ["Faturamento bruto por dia", money(resumo["Faturamento/dia"])],
            ["Repasse líquido por dia", money(resumo["Repasse/dia"])],
            ["Repasse líquido por pedido", money(resumo["Repasse/pedido"])],
            ["Promoções sobre faturamento", pct(resumo["Promoções % fat."])],
            ["Anúncios sobre faturamento", pct(resumo["Anúncios % fat."])],
            ["Comissões sobre faturamento", pct(resumo["Comissões % fat."])],
            ["Taxas sobre faturamento", pct(resumo["Taxas % fat."])],
            ["Investimento em crescimento (promoções + anúncios)", pct(resumo["Promoções + anúncios % fat."])],
            ["Repasse líquido sobre faturamento", pct(resumo["Repasse líquido % fat."])],
        ], columns=["KPI", "Valor"])
        st.dataframe(kpis, use_container_width=True, hide_index=True)

        st.subheader("💰 Visão Financeira Simplificada")
        receita_apos_promocoes = resumo["Faturamento bruto operacional (itens)"] - resumo["Promoções da loja (financeiro)"]
        visao = pd.DataFrame([
            ["Faturamento bruto (itens)", money(resumo["Faturamento bruto operacional (itens)"])],
            ["(-) Promoções da loja", money(resumo["Promoções da loja (financeiro)"])],
            ["Receita após promoções", money(receita_apos_promocoes)],
            ["(+) Incentivos iFood", money(resumo["Incentivos iFood (financeiro)"])],
            ["(+) Ressarcimentos", money(resumo["Ressarcimentos"])],
            ["(-) Comissões", money(resumo["Comissões iFood"])],
            ["(-) Taxas", money(resumo["Taxas iFood"])],
            ["(-) Anúncios", money(resumo["Anúncios iFood"])],
            ["Repasse líquido final", money(resumo["Repasse líquido recebido"])],
        ], columns=["Etapa", "Valor"])
        st.dataframe(visao, use_container_width=True, hide_index=True)

        whatsapp = f"""📊 Resumo iFood\n\n🍔 Faturamento Bruto (Itens): {money(resumo['Faturamento bruto operacional (itens)'])}\n📦 Pedidos concluídos: {num(resumo['Pedidos concluídos'])}\n🎟️ Ticket médio: {money(resumo['Ticket médio'])}\n\n💸 Promoções da loja: {money(resumo['Promoções da loja (financeiro)'])}\n📢 Anúncios: {money(resumo['Anúncios iFood'])}\n🏦 Comissões: {money(resumo['Comissões iFood'])}\n📋 Taxas: {money(resumo['Taxas iFood'])}\n\n🤝 Incentivos iFood: {money(resumo['Incentivos iFood (financeiro)'])}\n🔄 Ressarcimentos: {money(resumo['Ressarcimentos'])}\n\n✅ Repasse líquido recebido: {money(resumo['Repasse líquido recebido'])}\n\n📈 Investimento em crescimento (promoções + anúncios): {money(resumo['Promoções da loja (financeiro)'] + resumo['Anúncios iFood'])} ({pct(resumo['Promoções + anúncios % fat.'])} do faturamento)"""
        st.subheader("📱 Resumo para WhatsApp")
        st.text_area("Copiar e colar", whatsapp, height=260)

        csv = pd.concat([painel, kpis.rename(columns={"KPI":"Indicador"}), visao.rename(columns={"Etapa":"Indicador"})], ignore_index=True)
        st.download_button("Baixar resumo em CSV", csv.to_csv(index=False).encode("utf-8-sig"), file_name="resumo_ifood.csv", mime="text/csv")

        with st.expander("Conferência técnica"):
            st.write(f"Linhas lidas no relatório de pedidos: {len(pedidos)}")
            st.write(f"Linhas lidas no financeiro: {len(financeiro)}")
            st.write("Observação: para CMV, royalties e produtividade, use o faturamento bruto operacional de itens, não a venda recebida dos clientes.")

    except Exception as e:
        st.error("Não consegui processar os arquivos. Confira se os dois arquivos são relatórios XLSX do iFood.")
        st.exception(e)
else:
    st.info("Aguardando upload dos 2 arquivos.")
