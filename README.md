# Dashboard de Análise de Vendas / Sales Analysis Dashboard

[English](#english) | [Português](#português)

<a id="english"></a>
## English

### Overview

This is a sales analysis dashboard built with Streamlit that provides comprehensive insights into sales data. The dashboard processes data from an Excel file and displays various metrics, charts, and analyses to help understand sales performance.

### Features

- **Main Metrics**: Total Sales, Number of Orders, Number of Customers, and Average Ticket
- **ABC Curve Analysis**: Categorizes products into A, B, and C classes based on their contribution to total sales or quantity
  - Option to filter by classification (A, B, C)
  - Color-coded visualization for each category
- **Products Purchased Together**: Analysis of product pairs frequently bought together
- **Customer Profile Analysis**: Top customers by purchase value, order frequency, and average order value
- **Sales Evolution**: Time-based analysis of sales trends
- **Filtering Options**: Year, Month, and Customer exclusion filters

### Installation

#### Prerequisites

- Python 3.8 or higher
- Git (optional, for cloning the repository)

#### Step 1: Clone or download the repository

```bash
git clone <repository-url>
cd "Persona - Karen"
```

Or download and extract the ZIP file, then navigate to the extracted folder.

#### Step 2: Create a virtual environment (recommended)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

#### Step 3: Install dependencies

```bash
pip install -r requirements.txt
```

### Running the Dashboard

After installing the dependencies, run the dashboard with:

```bash
streamlit run dashboard.py
```

The dashboard will open in your default web browser at http://localhost:8501.

### Data Source

The dashboard reads data from an Excel file located in the `source` directory. The file should contain a worksheet named "Tabela1" with the following columns:

- Data (Date)
- ID (Order ID)
- NF (Invoice Number)
- Cliente (Customer)
- Produto ID (Product ID)
- Produto (Product)
- Quantidade (Quantity)
- Preço R$ (Price)
- Total (Total Value)

---

<a id="português"></a>
## Português

### Visão Geral

Este é um dashboard de análise de vendas construído com Streamlit que fornece insights abrangentes sobre dados de vendas. O dashboard processa dados de um arquivo Excel e exibe várias métricas, gráficos e análises para ajudar a entender o desempenho de vendas.

### Funcionalidades

- **Métricas Principais**: Total de Vendas, Número de Pedidos, Número de Clientes e Ticket Médio
- **Análise de Curva ABC**: Categoriza produtos em classes A, B e C com base em sua contribuição para o total de vendas ou quantidade
  - Opção para filtrar por classificação (A, B, C)
  - Visualização com cores diferentes para cada categoria
- **Produtos Comprados Juntos**: Análise de pares de produtos frequentemente comprados juntos
- **Análise do Perfil do Consumidor**: Top clientes por valor de compra, frequência de pedidos e valor médio de pedido
- **Evolução das Vendas**: Análise de tendências de vendas ao longo do tempo
- **Opções de Filtragem**: Filtros por Ano, Mês e exclusão de Clientes

### Instalação

#### Pré-requisitos

- Python 3.8 ou superior
- Git (opcional, para clonar o repositório)

#### Passo 1: Clonar ou baixar o repositório

```bash
git clone <url-do-repositório>
cd "Persona - Karen"
```

Ou baixe e extraia o arquivo ZIP, depois navegue até a pasta extraída.

#### Passo 2: Criar um ambiente virtual (recomendado)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

#### Passo 3: Instalar dependências

```bash
pip install -r requirements.txt
```

### Executando o Dashboard

Após instalar as dependências, execute o dashboard com:

```bash
streamlit run dashboard.py
```

O dashboard será aberto em seu navegador padrão em http://localhost:8501.

### Fonte de Dados

O dashboard lê dados de um arquivo Excel localizado no diretório `source`. O arquivo deve conter uma planilha chamada "Tabela1" com as seguintes colunas:

- Data
- ID (ID do Pedido)
- NF (Número da Nota Fiscal)
- Cliente
- Produto ID (ID do Produto)
- Produto
- Quantidade
- Preço R$
- Total (Valor Total)