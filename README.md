🇧🇷 **PT-BR**

## 🧠 O que faz?

Este projeto é um **scraper automatizado do PNCP (Portal Nacional de Contratações Públicas)** desenvolvido em **Node.js**, que permite pesquisar editais ou contratações diretas a partir de um termo informado pelo usuário e **exportar os resultados organizados em uma planilha Excel**.

A ferramenta realiza todo o processo de forma automática:

* Acessa o PNCP com base no termo de busca informado
* Navega pelos editais encontrados
* Extrai informações gerais de cada edital, como:

  * Nome do edital ou contratação direta
  * Data de início do recebimento de propostas
  * Data final do recebimento de propostas
* Acessa a tabela de itens de cada edital e coleta:

  * Número do item
  * Descrição
  * Quantidade
  * Valor unitário estimado
  * Valor total estimado

Os dados coletados são **filtrados por palavras-chave**, organizados e exportados automaticamente para um arquivo **Excel (.xlsx)**, já formatado para facilitar análise, leitura e uso posterior.

---

### 🔧 Recursos adicionais

* Execução via **linha de comando (CLI)**
* Parâmetros dinâmicos:

  * Termo de busca
  * Quantidade de páginas/editais a processar
* Escolha do diretório de saída dos arquivos

  * O local é salvo automaticamente e reutilizado nas próximas execuções
  * Pode ser alterado a qualquer momento via flag (`--arquivo`)
* Backups de progresso: o programa cria um arquivo de backup a cada **15 páginas processadas**. O backup tem nome fixo e é sobrescrito a cada 15 páginas (ex.: `protetor_solar_backup.xlsx`).
* O usuário pode escolher executar até **3 processos simultâneos** para acelerar o scraping (mais processos consomem mais recursos).

* Exportação para Excel com:

  * Cabeçalhos fixos
  * Filtros automáticos
  * Formatação de valores monetários
  * Ajuste automático de colunas e linhas

---

## ▶️ Como usar

### 1️⃣ Pré-requisitos

Antes de tudo, você precisa ter instalado na máquina:

* **Node.js** (versão 18 ou superior recomendada)
* **npm** (vem junto com o Node)

Para verificar:

```bash
node -v
npm -v
```

---

### 2️⃣ Instalação

Clone o repositório e instale as dependências:

```bash
git clone <URL_DO_REPOSITORIO>
cd <NOME_DO_PROJETO>
npm install
```

> O repositório já inclui `node_modules` para garantir o funcionamento correto.

OU

Baixe usando o botão de "Download ZIP" e extraia.

---

### 3️⃣ Executando o scraper

O script é executado via **linha de comando**, informando:

```bash
node index.js "<termo de busca>" <numero de paginas>
```

#### 🔹 Parâmetros

| Parâmetro           | Obrigatório | Descrição                                     |
| ------------------- | ----------- | --------------------------------------------- |
| `termo de busca`    | ✅ Sim       | Texto usado para pesquisar editais no PNCP    |
| `numero de paginas` | ❌ Não       | Quantidade de editais a processar (padrão: 1) |

> ⚠️ O termo de busca **deve estar entre aspas**.

---

### 4️⃣ Exemplos

Buscar por “protetor solar” e processar até 10 editais:

```bash
node index.js "protetor solar" 10
```

Buscar apenas por “luvas”, processando 1 edital (padrão quando não informado):

```bash
node index.js "luvas"
```

---

### 5️⃣ Escolhendo onde salvar o arquivo

Na **primeira execução**, o programa perguntará em qual pasta os arquivos Excel devem ser salvos:

```text
Informe o diretório onde os arquivos devem ser salvos:
>
```

Esse caminho será **salvo automaticamente** e reutilizado nas próximas execuções.

#### 🔁 Alterar o diretório de saída

Para escolher um novo local de salvamento, execute o script com a flag:

```bash
node index.js "protetor solar" 3 --arquivo
```

"protetor solar" e "3" são apenas exemplos de palavras.

---

### 6️⃣ Resultado

Ao final da execução:

* Um arquivo **Excel (.xlsx)** será gerado no diretório escolhido
* O nome do arquivo definitivo inclui data e hora da execução
* A planilha vem formatada, com:

  * Filtros automáticos
  * Valores monetários no formato **R$**
  * Dados organizados por edital

---

🇺🇸 **EN-US**

## 🧠 What does it do?

This project is an **automated scraper for the PNCP (National Public Procurement Portal of Brazil)** developed in **Node.js**. It allows users to search for **public notices or direct procurements** based on a keyword and **export the organized results to an Excel spreadsheet**.

The tool performs the entire process automatically:

* Accesses the PNCP using the provided search term
* Navigates through the found public notices
* Extracts general information from each public notice, such as:

  * Public notice name or direct procurement name
  * Start date for proposal submission
  * End date for proposal submission
* Accesses the items table of each public notice and collects:

  * Item number
  * Description
  * Quantity
  * Estimated unit price
  * Estimated total price

The collected data is **filtered by keywords**, organized, and automatically exported to an **Excel (.xlsx)** file, already formatted to facilitate analysis, reading, and further use.

---

### 🔧 Additional Features

* Command-line interface (CLI)
* Dynamic parameters:

  * Search term
  * Number of pages/public notices to process
* Output directory selection:

  * The chosen location is automatically saved and reused in future runs
  * It can be changed at any time using a flag (`--arquivo`)
* Progress backups: the program creates a backup file every **15 pages processed**. The backup uses a fixed name and is overwritten at each checkpoint (e.g. `sunscreen_backup.xlsx`).
* Concurrency: you can run up to **3 simultaneous processes** to speed up scraping (more processes use more resources).
* Error handling:

  * Clear message when no results are found
  * Prevents empty files or silent failures
* Excel export with:

  * Fixed headers
  * Automatic filters
  * Currency formatting
  * Automatic column and row resizing

---

## ▶️ How to Use

### 1️⃣ Prerequisites

Before starting, make sure you have installed:

* **Node.js** (version 18 or higher recommended)
* **npm** (comes bundled with Node.js)

To check:

```bash
node -v
npm -v
```

---

### 2️⃣ Installation

Clone the repository and install dependencies:

```bash
git clone <REPOSITORY_URL>
cd <PROJECT_NAME>
npm install
```

> The repository already includes `node_modules` to ensure proper functionality.

OR

Download the ZIP and extract.

---

### 3️⃣ Running the Scraper

The script is executed via the **command line**, providing:

```bash
node index.js "<search term>" <number of pages>
```

#### 🔹 Parameters

| Parameter         | Required | Description                               |
| ----------------- | -------- | ----------------------------------------- |
| `search term`     | ✅ Yes    | Text used to search public notices on PNCP       |
| `number of pages` | ❌ No     | Number of public notices to process (default: 1) |

> ⚠️ The search term **must be enclosed in quotes**.

---

### 4️⃣ Examples

Search for “sunscreen” and process up to 10 public notices:

```bash
node index.js "sunscreen" 10
```

Search only for “gloves”, processing 1 public notice (default when not provided):

```bash
node index.js "gloves"
```

---

### 5️⃣ Choosing Where to Save the File

On the **first execution**, the program will ask where the Excel files should be saved:

```text
Enter the directory where the files should be saved:
>
```

This path will be **automatically saved** and reused in future executions.

#### 🔁 Changing the Output Directory

To choose a new save location, run the script with the flag:

```bash
node index.js "sunscreen" 3 --arquivo
```

---

### 6️⃣ Output

At the end of execution:

* An **Excel (.xlsx)** file will be generated in the selected directory
* The file name includes the execution date and time (for the final export)
* The spreadsheet comes preformatted, with:

  * Automatic filters
  * Currency values formatted in **R$**
  * Data organized by public notice
