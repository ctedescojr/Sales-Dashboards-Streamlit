import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# Configuração da página
st.set_page_config(page_title="Dashboard de Análise de Vendas", layout="wide")

# Título do dashboard
st.title("Dashboard de Análise de Vendas")


# Função para processar os dados já normalizados da worksheet Tabela1
def processar_dados(df):
    # Os dados já estão normalizados, apenas precisamos fazer alguns ajustes
    df_processado = df.copy()

    # Renomear colunas para manter compatibilidade com o restante do código
    mapeamento_colunas = {
        "Data": "Data",
        "ID": "Pedido",
        "NF": "Codigo",
        "Cliente": "Cliente",
        "Produto ID": "Codigo_Produto",
        "Produto": "Produto",
        "Quantidade": "Quantidade",
        "Preço R$": "Preco",
        "Total": "Valor",
    }

    # Renomear as colunas
    df_processado = df_processado.rename(columns=mapeamento_colunas)

    # Garantir que a coluna de data esteja no formato datetime
    if not pd.api.types.is_datetime64_any_dtype(df_processado["Data"]):
        df_processado["Data"] = pd.to_datetime(df_processado["Data"], errors="coerce")

    # Adicionar colunas de ano e mês para facilitar a filtragem
    df_processado["Ano"] = df_processado["Data"].dt.year
    df_processado["Mes"] = df_processado["Data"].dt.month

    # Definir a ordem dos meses
    ordem_meses = [
        "Janeiro",
        "Fevereiro",
        "Março",
        "Abril",
        "Maio",
        "Junho",
        "Julho",
        "Agosto",
        "Setembro",
        "Outubro",
        "Novembro",
        "Dezembro",
    ]
    # Create a mapping from month number to month name
    mapa_meses = {i + 1: mes for i, mes in enumerate(ordem_meses)}
    df_processado["Mes_Nome"] = df_processado["Mes"].map(mapa_meses)

    df_processado["Mes_Nome"] = pd.Categorical(
        df_processado["Mes_Nome"], categories=ordem_meses, ordered=True
    )

    return df_processado


# Função para calcular a curva ABC (por valor ou quantidade)
def calcular_curva_abc(df, metrica="Valor"):
    # Agrupar por produto e somar os valores ou quantidades
    if metrica == "Valor":
        produto_metrica = df.groupby("Produto")["Valor"].sum().reset_index()
        coluna_metrica = "Valor"
    else:  # metrica == 'Quantidade'
        produto_metrica = df.groupby("Produto")["Quantidade"].sum().reset_index()
        coluna_metrica = "Quantidade"

    # Ordenar por valor/quantidade em ordem decrescente
    produto_metrica = produto_metrica.sort_values(coluna_metrica, ascending=False)

    # Calcular o total
    total = produto_metrica[coluna_metrica].sum()

    # Calcular a porcentagem de cada produto
    produto_metrica["Porcentagem"] = produto_metrica[coluna_metrica] / total * 100

    # Calcular a porcentagem acumulada
    produto_metrica["Porcentagem_Acumulada"] = produto_metrica["Porcentagem"].cumsum()

    # Classificar os produtos em A, B ou C
    def classificar(porcentagem_acumulada):
        if porcentagem_acumulada <= 80:
            return "A"
        elif porcentagem_acumulada <= 95:
            return "B"
        else:
            return "C"

    produto_metrica["Classificacao"] = produto_metrica["Porcentagem_Acumulada"].apply(
        classificar
    )

    return produto_metrica


# Função para analisar produtos comprados juntos
def produtos_comprados_juntos(df):
    # Agrupar por pedido
    pedidos_produtos = (
        df.groupby(["Pedido", "Produto"])["Quantidade"].sum().reset_index()
    )

    # Criar um dicionário para armazenar os pares de produtos
    pares_produtos = {}

    # Iterar pelos pedidos
    for pedido in pedidos_produtos["Pedido"].unique():
        # Obter os produtos do pedido
        produtos = pedidos_produtos[pedidos_produtos["Pedido"] == pedido][
            "Produto"
        ].tolist()

        # Criar pares de produtos
        for i in range(len(produtos)):
            for j in range(i + 1, len(produtos)):
                par = tuple(sorted([produtos[i], produtos[j]]))
                if par in pares_produtos:
                    pares_produtos[par] += 1
                else:
                    pares_produtos[par] = 1

    # Converter o dicionário para DataFrame
    pares_df = pd.DataFrame(
        [(p1, p2, count) for (p1, p2), count in pares_produtos.items()],
        columns=["Produto1", "Produto2", "Frequencia"],
    )

    # Ordenar por frequência
    pares_df = pares_df.sort_values("Frequencia", ascending=False)

    return pares_df


# Função para analisar o perfil do consumidor
def perfil_consumidor(df):
    # Agrupar por cliente e somar os valores
    cliente_valor = df.groupby("Cliente")["Valor"].sum().reset_index()

    # Ordenar por valor em ordem decrescente
    cliente_valor = cliente_valor.sort_values("Valor", ascending=False)

    # Calcular o número de pedidos por cliente
    cliente_pedidos = df.groupby("Cliente")["Pedido"].nunique().reset_index()
    cliente_pedidos.columns = ["Cliente", "Num_Pedidos"]

    # Mesclar os DataFrames
    cliente_analise = pd.merge(cliente_valor, cliente_pedidos, on="Cliente")

    # Calcular o valor médio por pedido
    cliente_analise["Valor_Medio_Pedido"] = (
        cliente_analise["Valor"] / cliente_analise["Num_Pedidos"]
    )

    return cliente_analise


# Carregar os dados
try:
    # Carregar o arquivo Excel da worksheet Tabela1
    @st.cache_data
    def load_data():
        df = pd.read_excel("source/frRelPedidosEmitidos.xls", sheet_name="Tabela1")
        return processar_dados(df)

    df = load_data()

    # Verificar se há dados
    if df.empty:
        st.error("Não foi possível processar os dados do arquivo.")
    else:
        # Sidebar para filtros
        st.sidebar.header("Filtros")

        # Filtro de ano
        anos = sorted(df["Ano"].unique())
        anos_selecionados = st.sidebar.multiselect(
            "Selecione o(s) Ano(s)", anos, default=anos
        )

        # Filtro de mês
        meses = sorted(df["Mes"].unique())
        mes_selecionado = st.sidebar.selectbox(
            "Selecione o Mês", ["Todos"] + list(meses)
        )

        # Filtro de clientes
        clientes = sorted(df["Cliente"].unique())
        clientes_excluidos = st.sidebar.multiselect("Excluir Clientes", clientes)

        # Aplicar filtros
        df_filtrado = df[df["Ano"].isin(anos_selecionados)]

        if mes_selecionado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Mes"] == mes_selecionado]

        # Excluir clientes selecionados
        if clientes_excluidos:
            df_filtrado = df_filtrado[~df_filtrado["Cliente"].isin(clientes_excluidos)]

        # Métricas principais
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            valor_total = df_filtrado["Valor"].sum()
            st.metric(
                "Total de Vendas",
                f"R$ {valor_total:,.2f}".replace(",", ".").replace(".", ",", 1),
            )

        with col2:
            st.metric(
                "Número de Pedidos",
                f"{df_filtrado['Pedido'].nunique():,}".replace(",", "."),
            )

        with col3:
            st.metric(
                "Número de Clientes",
                f"{df_filtrado['Cliente'].nunique():,}".replace(",", "."),
            )

        with col4:
            ticket_medio = (
                df_filtrado["Valor"].sum() / df_filtrado["Pedido"].nunique()
                if df_filtrado["Pedido"].nunique() > 0
                else 0
            )
            st.metric(
                "Ticket Médio",
                f"R$ {ticket_medio:,.2f}".replace(",", ".").replace(".", ",", 1),
            )

        # Análise de Curva ABC
        st.header("Análise de Curva ABC")

        # Opções para escolher entre valor e quantidade
        col1, col2 = st.columns(2)
        with col1:
            metrica_abc = st.radio(
                "Métrica para Curva ABC", ["Valor", "Quantidade"], horizontal=True
            )
        with col2:
            filtro_abc = st.multiselect(
                "Filtrar por Classificação", ["A", "B", "C"], default=["A", "B", "C"]
            )

        curva_abc = calcular_curva_abc(df_filtrado, metrica=metrica_abc)

        # Filtrar por classificação selecionada
        curva_abc_filtrada = curva_abc[curva_abc["Classificacao"].isin(filtro_abc)]

        # Gráfico de Pareto
        fig_pareto = go.Figure()

        # Definir cores para cada classificação
        cores_abc = {"A": "#1f77b4", "B": "#ff7f0e", "C": "#2ca02c"}

        # Adicionar barras com cores por classificação
        for classificacao in curva_abc_filtrada["Classificacao"].unique():
            df_class = curva_abc_filtrada[
                curva_abc_filtrada["Classificacao"] == classificacao
            ]
            fig_pareto.add_trace(
                go.Bar(
                    x=df_class["Produto"],
                    y=df_class[metrica_abc],
                    name=f"Produtos {classificacao}",
                    marker_color=cores_abc[classificacao],
                )
            )

        # Adicionar linha de porcentagem acumulada apenas se tiver todos os dados
        if set(filtro_abc) == set(["A", "B", "C"]):
            fig_pareto.add_trace(
                go.Scatter(
                    x=curva_abc["Produto"],
                    y=curva_abc["Porcentagem_Acumulada"],
                    name="% Acumulada",
                    marker_color="red",
                    yaxis="y2",
                )
            )
        else:
            # Recalcular porcentagem acumulada para os dados filtrados
            total_filtrado = curva_abc_filtrada[metrica_abc].sum()
            curva_abc_filtrada = curva_abc_filtrada.sort_values(
                metrica_abc, ascending=False
            ).copy()
            curva_abc_filtrada["Porcentagem_Filtrada"] = (
                curva_abc_filtrada[metrica_abc] / total_filtrado * 100
            )
            curva_abc_filtrada["Porcentagem_Acumulada_Filtrada"] = curva_abc_filtrada[
                "Porcentagem_Filtrada"
            ].cumsum()

            fig_pareto.add_trace(
                go.Scatter(
                    x=curva_abc_filtrada["Produto"],
                    y=curva_abc_filtrada["Porcentagem_Acumulada_Filtrada"],
                    name="% Acumulada (Filtrada)",
                    marker_color="red",
                    yaxis="y2",
                )
            )

        # Atualizar layout
        fig_pareto.update_layout(
            title=f"Gráfico de Pareto - Curva ABC por {metrica_abc}",
            xaxis_title="Produto",
            yaxis_title=f"{metrica_abc}" + (" (R$)" if metrica_abc == "Valor" else ""),
            yaxis2=dict(
                title="Porcentagem Acumulada",
                overlaying="y",
                side="right",
                range=[0, 100],
            ),
            legend=dict(x=0.01, y=0.99, bgcolor="rgba(255, 255, 255, 0.5)"),
        )

        st.plotly_chart(fig_pareto, use_container_width=True)

        # Tabela de Curva ABC
        st.subheader("Tabela de Classificação ABC")

        # Formatar a tabela
        curva_abc_display = curva_abc_filtrada.copy()
        if "Valor" in curva_abc_display.columns:
            curva_abc_display["Valor"] = curva_abc_display["Valor"].apply(
                lambda x: f"R$ {x:,.2f}".replace(",", ".").replace(".", ",", 1)
            )
        if "Quantidade" in curva_abc_display.columns:
            curva_abc_display["Quantidade"] = curva_abc_display["Quantidade"].apply(
                lambda x: f"{x:,.0f}".replace(",", ".")
            )
        curva_abc_display["Porcentagem"] = curva_abc_display["Porcentagem"].apply(
            lambda x: f"{x:.2f}%".replace(".", ",")
        )
        curva_abc_display["Porcentagem_Acumulada"] = curva_abc_display[
            "Porcentagem_Acumulada"
        ].apply(lambda x: f"{x:.2f}%".replace(".", ","))

        # Mostrar a tabela
        st.dataframe(curva_abc_display)

        # Resumo da Curva ABC
        col1, col2, col3 = st.columns(3)

        with col1:
            produtos_a = curva_abc[curva_abc["Classificacao"] == "A"]
            st.metric(
                "Produtos A",
                len(produtos_a),
                f"{len(produtos_a)/len(curva_abc)*100:.1f}%".replace(".", ",")
                + " do total",
            )

            if metrica_abc == "Valor":
                valor_a = produtos_a["Valor"].sum()
                valor_total = curva_abc["Valor"].sum()
                st.metric(
                    f"{metrica_abc} Produtos A",
                    f"R$ {valor_a:,.2f}".replace(",", ".").replace(".", ",", 1),
                    f"{valor_a/valor_total*100:.1f}%".replace(".", ",") + " do total",
                )
            else:  # metrica_abc == 'Quantidade'
                qtd_a = produtos_a["Quantidade"].sum()
                qtd_total = curva_abc["Quantidade"].sum()
                st.metric(
                    f"{metrica_abc} Produtos A",
                    f"{qtd_a:,.0f}".replace(",", "."),
                    f"{qtd_a/qtd_total*100:.1f}%".replace(".", ",") + " do total",
                )

        with col2:
            produtos_b = curva_abc[curva_abc["Classificacao"] == "B"]
            st.metric(
                "Produtos B",
                len(produtos_b),
                f"{len(produtos_b)/len(curva_abc)*100:.1f}%".replace(".", ",")
                + " do total",
            )

            if metrica_abc == "Valor":
                valor_b = produtos_b["Valor"].sum()
                valor_total = curva_abc["Valor"].sum()
                st.metric(
                    f"{metrica_abc} Produtos B",
                    f"R$ {valor_b:,.2f}".replace(",", ".").replace(".", ",", 1),
                    f"{valor_b/valor_total*100:.1f}%".replace(".", ",") + " do total",
                )
            else:  # metrica_abc == 'Quantidade'
                qtd_b = produtos_b["Quantidade"].sum()
                qtd_total = curva_abc["Quantidade"].sum()
                st.metric(
                    f"{metrica_abc} Produtos B",
                    f"{qtd_b:,.0f}".replace(",", "."),
                    f"{qtd_b/qtd_total*100:.1f}%".replace(".", ",") + " do total",
                )

        with col3:
            produtos_c = curva_abc[curva_abc["Classificacao"] == "C"]
            st.metric(
                "Produtos C",
                len(produtos_c),
                f"{len(produtos_c)/len(curva_abc)*100:.1f}%".replace(".", ",")
                + " do total",
            )

            if metrica_abc == "Valor":
                valor_c = produtos_c["Valor"].sum()
                valor_total = curva_abc["Valor"].sum()
                st.metric(
                    f"{metrica_abc} Produtos C",
                    f"R$ {valor_c:,.2f}".replace(",", ".").replace(".", ",", 1),
                    f"{valor_c/valor_total*100:.1f}%".replace(".", ",") + " do total",
                )
            else:  # metrica_abc == 'Quantidade'
                qtd_c = produtos_c["Quantidade"].sum()
                qtd_total = curva_abc["Quantidade"].sum()
                st.metric(
                    f"{metrica_abc} Produtos C",
                    f"{qtd_c:,.0f}".replace(",", "."),
                    f"{qtd_c/qtd_total*100:.1f}%".replace(".", ",") + " do total",
                )

        # Análise de Produtos Comprados Juntos
        st.header("Análise de Produtos Comprados Juntos")

        pares_produtos = produtos_comprados_juntos(df_filtrado)

        if not pares_produtos.empty:
            # Mostrar os top 10 pares de produtos
            st.subheader("Top 10 Pares de Produtos Comprados Juntos")
            st.dataframe(pares_produtos.head(10))

            # Gráfico de barras para os top 10 pares
            top_10_pares = pares_produtos.head(10).sort_values(
                "Frequencia", ascending=True
            )
            fig_pares = px.bar(
                top_10_pares,
                x="Frequencia",
                y=top_10_pares.apply(
                    lambda x: f"{x['Produto1']} + {x['Produto2']}", axis=1
                ),
                orientation="h",
                title="Top 10 Pares de Produtos Comprados Juntos",
                labels={"y": "Par de Produtos", "x": "Frequência"},
            )

            st.plotly_chart(fig_pares, use_container_width=True)
        else:
            st.info(
                "Não há dados suficientes para análise de produtos comprados juntos."
            )

        # Análise do Perfil do Consumidor
        st.header("Análise do Perfil do Consumidor")

        perfil = perfil_consumidor(df_filtrado)

        # Mostrar os top 10 clientes
        st.subheader("Top 10 Clientes por Valor de Compra")

        # Formatar a tabela
        perfil_display = perfil.head(10).copy()
        perfil_display["Valor"] = perfil_display["Valor"].apply(
            lambda x: f"R$ {x:,.2f}".replace(",", ".").replace(".", ",", 1)
        )
        perfil_display["Valor_Medio_Pedido"] = perfil_display[
            "Valor_Medio_Pedido"
        ].apply(lambda x: f"R$ {x:,.2f}".replace(",", ".").replace(".", ",", 1))
        perfil_display["Num_Pedidos"] = perfil_display["Num_Pedidos"].apply(
            lambda x: f"{x:,}".replace(",", ".")
        )

        st.dataframe(perfil_display)

        # Gráfico de barras para os top 10 clientes
        fig_clientes = px.bar(
            perfil.head(10),
            x="Cliente",
            y="Valor",
            title="Top 10 Clientes por Valor de Compra",
            labels={"Cliente": "Cliente", "Valor": "Valor Total (R$)"},
            color="Valor",
            color_continuous_scale="Blues",
        )

        st.plotly_chart(fig_clientes, use_container_width=True)

        # Gráfico de dispersão: Número de Pedidos vs Valor Total
        fig_dispersao = px.scatter(
            perfil,
            x="Num_Pedidos",
            y="Valor",
            size="Valor_Medio_Pedido",
            color="Valor_Medio_Pedido",
            hover_name="Cliente",
            title="Relação entre Número de Pedidos e Valor Total",
            labels={
                "Num_Pedidos": "Número de Pedidos",
                "Valor": "Valor Total (R$)",
                "Valor_Medio_Pedido": "Valor Médio por Pedido (R$)",
            },
            color_continuous_scale="Viridis",
        )

        st.plotly_chart(fig_dispersao, use_container_width=True)

        # Evolução das vendas ao longo do tempo
        st.header("Evolução das Vendas ao Longo do Tempo")

        # Agrupar por data e somar os valores
        vendas_tempo = df.groupby("Data")["Valor"].sum().reset_index()

        # Gráfico de linha
        fig_tempo = px.line(
            vendas_tempo,
            x="Data",
            y="Valor",
            title="Evolução das Vendas ao Longo do Tempo",
            labels={"Data": "Data", "Valor": "Valor Total (R$)"},
            markers=True,
        )

        st.plotly_chart(fig_tempo, use_container_width=True)

        # Vendas por mês
        vendas_mes = (
            df.groupby(["Ano", "Mes", "Mes_Nome"], observed=False)["Valor"]
            .sum()
            .reset_index()
        )
        vendas_mes = vendas_mes.sort_values(["Ano", "Mes"])

        # Gráfico de barras por mês
        fig_mes = px.bar(
            vendas_mes,
            x="Mes_Nome",
            y="Valor",
            color="Ano",
            title="Vendas por Mês",
            labels={"Mes_Nome": "", "Valor": "Valor Total (R$)", "Ano": "Ano"},
            barmode="group",
        )

        st.plotly_chart(fig_mes, use_container_width=True)

        # Dados brutos
        st.header("Dados Processados")
        st.dataframe(df_filtrado)

except Exception as e:
    st.error(f"Ocorreu um erro ao carregar ou processar os dados: {e}")
