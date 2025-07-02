import pandas as pd
import matplotlib.pyplot as plt

# caminho base dos arquivos
p = r"C:/Users/gmarc/OneDrive/Área de Trabalho/Projeto análise anual/Vendas/7. Vendas"

# leitura das planilhas de vendas de 2020, 2021 e 2022
vendas_20 = pd.read_excel(p + "/Base Vendas - 2020.xlsx", usecols=["Data da Venda", "ID Cliente", "SKU", "Qtd Vendida"])
vendas_21 = pd.read_excel(p + "/Base Vendas - 2021.xlsx", usecols=["Data da Venda", "ID Cliente", "SKU", "Qtd Vendida"])
vendas_22 = pd.read_excel(p + "/Base Vendas - 2022.xlsx", usecols=["Data da Venda", "ID Cliente", "SKU", "Qtd Vendida"])

# junta todas as vendas em um único dataframe
vendas = pd.concat([vendas_20, vendas_21, vendas_22])

# leitura das devoluções e cadastro de produtos
devolucoes = pd.read_excel(p + "/Base Devoluções.xlsx", usecols=["Data Devolução", "SKU", "Qtd Devolvida"])
produtos = pd.read_excel(p + "/Cadastro Produtos.xlsx", usecols=["SKU", "Preço Unitario", "Custo Unitario"])

# calcula o lucro por unidade
produtos["lucro_unit"] = produtos["Preço Unitario"] - produtos["Custo Unitario"]

# cria colunas de ano e mês nas vendas
vendas["Ano"] = vendas["Data da Venda"].dt.year
vendas["Mes"] = vendas["Data da Venda"].dt.month

# encontra a data da primeira compra de cada cliente
primeira_compra = vendas.groupby("ID Cliente")["Data da Venda"].min().reset_index()
primeira_compra.columns = ["ID Cliente", "Data Primeira Compra"]

# junta essa informação na base de vendas
vendas = vendas.merge(primeira_compra, on="ID Cliente", how="left")



# ---------------- novos compradores até junho ----------------

# filtra vendas de 2021 e 2022 até o mês 6
vendas_filtrada = vendas[(vendas["Ano"].isin([2021, 2022])) & (vendas["Mes"] <= 6)]

# clientes que compraram pela primeira vez até junho de 2021
novos_21 = vendas_filtrada[
    (vendas_filtrada["Ano"] == 2021) &
    (vendas_filtrada["Data Primeira Compra"].between("2021-01-01", "2021-06-30"))
]["ID Cliente"].nunique()

# clientes que compraram pela primeira vez até junho de 2022
novos_22 = vendas_filtrada[
    (vendas_filtrada["Ano"] == 2022) &
    (vendas_filtrada["Data Primeira Compra"].between("2022-01-01", "2022-06-30"))
]["ID Cliente"].nunique()

# mostra o resultado
print("Novos compradores em 2021 (até jun):", novos_21)
print("Novos compradores em 2022 (até jun):", novos_22)
print("Diferença percentual: {:.2f}%".format((novos_22 - novos_21) / novos_21 * 100) + "\n")



# ---------------- ticket médio ----------------

# junta os dados de produto na base de vendas
vendas = vendas.merge(produtos, on="SKU", how="left")

# calcula o valor total e lucro por venda
vendas["valor_venda"] = vendas["Preço Unitario"] * vendas["Qtd Vendida"]
vendas["lucro_bruto"] = vendas["lucro_unit"] * vendas["Qtd Vendida"]

# separa vendas até junho por ano
v21 = vendas[(vendas["Ano"] == 2021) & (vendas["Mes"] <= 6)]
v22 = vendas[(vendas["Ano"] == 2022) & (vendas["Mes"] <= 6)]

# calcula receita total e quantidade de vendas
receita_21 = v21["valor_venda"].sum()
receita_22 = v22["valor_venda"].sum()
qtd_vendas_21 = len(v21)
qtd_vendas_22 = len(v22)

# calcula o ticket médio (valor médio por venda)
ticket_21 = receita_21 / qtd_vendas_21
ticket_22 = receita_22 / qtd_vendas_22

# mostra o resultado
print("Ticket médio em 2021 (até jun): R$ {:.2f}".format(ticket_21))
print("Ticket médio em 2022 (até jun): R$ {:.2f}".format(ticket_22))
print("Diferença percentual: {:.2f}%".format((ticket_22 - ticket_21) / ticket_21 * 100) + "\n")



# ---------------- lucro líquido ----------------

# separa devoluções até junho de cada ano
devolucoes["Ano"] = devolucoes["Data Devolução"].dt.year
devolucoes["Mes"] = devolucoes["Data Devolução"].dt.month
devolucoes_21 = devolucoes[(devolucoes["Ano"] == 2021) & (devolucoes["Mes"] <= 6)]
devolucoes_22 = devolucoes[(devolucoes["Ano"] == 2022) & (devolucoes["Mes"] <= 6)]

# junta os produtos nas devoluções e calcula prejuízo
dev_21 = devolucoes_21.merge(produtos, on="SKU", how="left")
dev_22 = devolucoes_22.merge(produtos, on="SKU", how="left")
dev_21["prejuizo"] = dev_21["lucro_unit"] * dev_21["Qtd Devolvida"]
dev_22["prejuizo"] = dev_22["lucro_unit"] * dev_22["Qtd Devolvida"]

# calcula o lucro líquido: lucro bruto - prejuízo
lucro_21 = v21["lucro_bruto"].sum() - dev_21["prejuizo"].sum()
lucro_22 = v22["lucro_bruto"].sum() - dev_22["prejuizo"].sum()

# mostra o resultado
print("Lucro líquido em 2021 (até jun): R$ {:.2f}".format(lucro_21))
print("Lucro líquido em 2022 (até jun): R$ {:.2f}".format(lucro_22))
print("Diferença absoluta: R$ {:.2f}".format(lucro_22 - lucro_21))
print("Diferença percentual: {:.2f}%".format((lucro_22 - lucro_21) / lucro_21 * 100) + "\n")



# ---------------- margem operacional ----------------

# calcula a margem (lucro sobre receita)
margem_21 = lucro_21 / receita_21
margem_22 = lucro_22 / receita_22

# mostra o resultado
print("Margem operacional em 2021 (até jun): {:.2%}".format(margem_21))
print("Margem operacional em 2022 (até jun): {:.2%}".format(margem_22))
print("Diferença: {:.2f} pontos percentuais".format((margem_22 - margem_21) * 100) + "\n")

# mostra devoluções totais
qtd_dev_21 = dev_21["Qtd Devolvida"].count()
qtd_dev_22 = dev_22["Qtd Devolvida"].count()
print("Devoluções em 2021 (até jun):", qtd_dev_21)
print("Devoluções em 2022 (até jun):", qtd_dev_22)
print("Diferença percentual: {:.2f}%".format((qtd_dev_22 - qtd_dev_21) / qtd_dev_21 * 100) + "\n")



# ---------------- projeção de lucro até dezembro ----------------

# pega vendas do ano todo e cria coluna ano-mês
v21_full = vendas[vendas["Ano"] == 2021]
v22_full = vendas[vendas["Ano"] == 2022]
v21_full["ano_mes"] = v21_full["Data da Venda"].dt.to_period("M")
v22_full["ano_mes"] = v22_full["Data da Venda"].dt.to_period("M")

# pega devoluções do ano todo e junta produtos
dev_21_full = devolucoes[devolucoes["Ano"] == 2021].copy()
dev_22_full = devolucoes[devolucoes["Ano"] == 2022].copy()
dev_21_full = dev_21_full.merge(produtos, on="SKU", how="left")
dev_22_full = dev_22_full.merge(produtos, on="SKU", how="left")
dev_21_full["prejuizo"] = dev_21_full["lucro_unit"] * dev_21_full["Qtd Devolvida"]
dev_22_full["prejuizo"] = dev_22_full["lucro_unit"] * dev_22_full["Qtd Devolvida"]
dev_21_full["ano_mes"] = dev_21_full["Data Devolução"].dt.to_period("M")
dev_22_full["ano_mes"] = dev_22_full["Data Devolução"].dt.to_period("M")

# calcula lucro líquido mês a mês
lucro_mensal_21_full = (
    v21_full.groupby("ano_mes")["lucro_bruto"].sum()
    - dev_21_full.groupby("ano_mes")["prejuizo"].sum()
).reset_index()

lucro_mensal_22_full = (
    v22_full.groupby("ano_mes")["lucro_bruto"].sum()
    - dev_22_full.groupby("ano_mes")["prejuizo"].sum()
).reset_index()

# renomeia e adiciona a coluna do número do mês
lucro_mensal_21_full.columns = ["ano_mes", "lucro_liquido"]
lucro_mensal_21_full["mes"] = lucro_mensal_21_full["ano_mes"].dt.month
lucro_mensal_22_full.columns = ["ano_mes", "lucro_liquido"]
lucro_mensal_22_full["mes"] = lucro_mensal_22_full["ano_mes"].dt.month

# calcula a variação percentual mês a mês de 2021
lucro_21_pct_full = lucro_mensal_21_full.set_index("mes")["lucro_liquido"].pct_change()

# pega os valores reais de 2022 e inicia a projeção
proj_22_full = lucro_mensal_22_full.set_index("mes")["lucro_liquido"].copy()

# vê até qual mês temos dados reais
ultimo_mes_real = proj_22_full.index.max()

# projeta de julho até dezembro com base no crescimento de 2021
for m in range(ultimo_mes_real + 1, 13):
    ultimo_valor = proj_22_full[m - 1]
    variacao_pct = lucro_21_pct_full.get(m, 0)
    proj_22_full[m] = ultimo_valor * (1 + variacao_pct)

# ajusta índice para mostrar como ano-mês (2022-01, etc)
proj_22_full.index = pd.period_range("2022-01", "2022-12", freq="M")


# ---------------- gráfico: lucro líquido por mês até junho ----------------

# filtra lucro mensal até junho
lucro_mensal_21_jun = lucro_mensal_21_full[lucro_mensal_21_full["mes"] <= 6]
lucro_mensal_22_jun = lucro_mensal_22_full[lucro_mensal_22_full["mes"] <= 6]

# aplica estilo escuro
plt.style.use("dark_background")
plt.figure(figsize=(10, 5))

# plota os dados com mês como eixo X
plt.plot(
    lucro_mensal_21_jun["mes"],
    lucro_mensal_21_jun["lucro_liquido"] / 1_000_000,
    label="2021",
    color="lightgray",
    linewidth=2
)
plt.plot(
    lucro_mensal_22_jun["mes"],
    lucro_mensal_22_jun["lucro_liquido"] / 1_000_000,
    label="2022",
    color="yellow",
    linewidth=2
)

# ajustes do gráfico
plt.title("Lucro líquido por mês - até junho", fontsize=14, color="white")
plt.xlabel("Mês", color="white")
plt.ylabel("Lucro Líquido (R$ Milhão)", color="white")
plt.xticks(ticks=range(1, 7), labels=["Jan", "Fev", "Mar", "Abr", "Mai", "Jun"], color="white")
plt.yticks(color="white")
plt.grid(True, linestyle='--', alpha=0.3)
plt.legend()
plt.tight_layout()
plt.show()

# ---------------- análise descritiva ----------------

print("""
A partir dos números apresentados e das análises extraídas do dashboard disponível, 
chegamos à conclusão de que a loja passou por uma fase de foco em expansão e escalabilidade em 2021, 
aumentando consideravelmente seus investimentos em novas filiais. Além disso, é possível que tenha 
implementado uma estratégia específica de marketing a partir de julho de 2021, o que justifica o crescimento expressivo.

O ticket médio menor sugere um aumento no volume de compradores, o que permite o lançamento de 
novas estratégias — como cupons de desconto a partir de determinado valor de compra — que visam 
aumentar ou manter esse ticket médio.

Em contrapartida, torna-se necessário reforçar os investimentos em operações logísticas, como o 
processo de devolução e a revisão da qualidade dos produtos, já que o índice de devolução também 
cresceu proporcionalmente.

Com base nos dados apresentados, a tendência é de crescimento — caso o comportamento de 2021 se repita — 
e novas estratégias podem ser lançadas, focando em públicos específicos, também descritos no dashboard.
""")


# ---------------- gráfico: projeção de lucro líquido ----------------

plt.style.use("dark_background")
plt.figure(figsize=(10, 5))

# plota a linha da projeção
plt.plot(
    proj_22_full.index.astype(str),
    proj_22_full.values / 1_000_000,
    marker='o',
    color='yellow',
    linewidth=2,
    label="Projeção 2022"
)

# linha vertical marcando o fim dos dados reais
plt.axvline(str(lucro_mensal_22_full["ano_mes"].iloc[-1]), color='gray', linestyle='--', label="Início da projeção")

# ajustes do gráfico
plt.title("Projeção de Lucro Líquido até Dez/2022 (com base nas variações de 2021)", fontsize=13, color="white")
plt.xlabel("Mês", color="white")
plt.ylabel("Lucro Líquido (R$ Milhão)", color="white")
plt.xticks(rotation=45, color="white")
plt.yticks(color="white")
plt.grid(True, linestyle='--', alpha=0.3)
plt.legend()
plt.tight_layout()
plt.show()
