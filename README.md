# Resumo iFood GGK / Sampa Burger

Aplicação simples em Streamlit para o franqueado subir 2 planilhas do iFood e receber um resumo consolidado.

## Arquivos necessários

- `app.py` — aplicação principal
- `requirements.txt` — bibliotecas necessárias
- `README.md` — instruções

## Como rodar no computador

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Como publicar no Streamlit Cloud

1. Crie uma conta em https://streamlit.io/cloud
2. Crie um repositório no GitHub com estes 3 arquivos.
3. No Streamlit Cloud, clique em **New app**.
4. Selecione o repositório.
5. Main file path: `app.py`
6. Clique em Deploy.

O Streamlit vai gerar um link externo para os franqueados acessarem.

## Observação gerencial

O app usa o **faturamento bruto operacional de itens** como base de faturamento para análise de CMV, royalties e produtividade.
A venda recebida dos clientes e o repasse líquido aparecem separadamente para não distorcer o CMV.
