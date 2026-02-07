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
* Tratamento de erros:

  * Mensagem clara quando nenhum resultado é encontrado
  * Evita arquivos vazios ou falhas silenciosas
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
> O repositório já vêm com o node_modules, já que em alguns testes, instalar ele usando 
```npm install node_modules```, dava erro.

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
| `numero de paginas` | ❌ Não       | Quantidade de editais a processar (padrão: 5) |

> ⚠️ O termo de busca **deve estar entre aspas**.

---

### 4️⃣ Exemplos

Buscar por “protetor solar” e processar até 5 editais:

```bash
node index.js "protetor solar" 5
```

Buscar apenas por “luvas”, processando 5 editais (se não digitar número, padrão de 5 páginas):

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

---

### 6️⃣ Ajuda

Para exibir as instruções de uso no terminal:

```bash
node index.js --help
```

ou

```bash
node index.js -h
```

---

### 7️⃣ Resultado

Ao final da execução:

* Um arquivo **Excel (.xlsx)** será gerado no diretório escolhido
* O nome do arquivo inclui data e hora da execução
* A planilha vem formatada, com:

  * Filtros automáticos
  * Valores monetários no formato **R$**
  * Dados organizados por edital

