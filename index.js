import puppeteer from "puppeteer";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import readline from "readline";

const CONFIG_FILE = path.resolve(process.cwd(), "config.json");
const args = process.argv.slice(2);

const forceAskPath = args.includes("--arquivo");

const rawSearchTerm = args[0];
const numberOfPagesToSearch = Number(args[1] || 1);

if (!rawSearchTerm) {
  console.error(
    'Informe o termo de busca.\nEx: node index.js "protetor solar" 3'
  );
  process.exit(1);
}

if (!Number.isInteger(numberOfPagesToSearch) || numberOfPagesToSearch <= 0) {
  console.error("Número de páginas inválido.");
  process.exit(1);
}

const config = loadConfig();
let outputDir = config.outputDir;

// se não existir config ou se o user pediu pra trocar
if (!outputDir || forceAskPath) {
  const answer = await askDirectory(
    "Informe o diretório onde os arquivos devem ser salvos:\n> "
  );

  if (!answer) {
    console.error("Diretório inválido.");
    process.exit(1);
  }

  const resolvedPath = path.resolve(answer);

  // cria a pasta se não existir
  fs.mkdirSync(resolvedPath, { recursive: true });

  outputDir = resolvedPath;

  saveConfig({ ...config, outputDir });

  console.log("Diretório salvo:", outputDir);
} else {
  console.log("Usando diretório salvo:", outputDir);
}

function askDirectory(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

function loadConfig() {
  if (fs.existsSync(CONFIG_FILE)) {
    try {
      return JSON.parse(fs.readFileSync(CONFIG_FILE, "utf-8"));
    } catch {
      return {};
    }
  }
  return {};
}

// salva config
function saveConfig(config) {
  fs.writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2), "utf-8");
}

if (
  args.includes("--ajuda") ||
  args.includes("-a") ||
  args.includes("--help") ||
  args.includes("-h")
) {
  console.log(`
Uso:
  node index.js "<termo de busca>" <numero de paginas>

Exemplo:
  node index.js "protetor solar" 5 (o texto deve ser com aspas)

Observações:
- O termo de busca é obrigatório
- O número de páginas é opcional (não digitar nada = 5)
`);
  process.exit(0);
}

if (!rawSearchTerm || typeof rawSearchTerm !== "string") {
  console.error("Erro: informe um termo de busca.");
  console.error('Exemplo: node index.js "protetor solar" 3');
  process.exit(1);
}

if (!Number.isInteger(numberOfPagesToSearch) || numberOfPagesToSearch <= 0) {
  console.error("Erro: o número de páginas deve ser um inteiro positivo.");
  console.error('Exemplo: node index.js "protetor solar" 3');
  process.exit(1);
}

console.log("Termo de busca:", rawSearchTerm);
console.log("Páginas a processar:", numberOfPagesToSearch);

function searchLink(receivedUserInput) {
  const search = (receivedUserInput || "").trim();
  if (!search) {
    throw new Error("A busca não pode ficar vazia.");
  }
  const encodedInput = encodeURIComponent(receivedUserInput);
  return `https://pncp.gov.br/app/editais?q=${encodedInput}&status=recebendo_proposta&pagina=1&ordenacao=relevancia&tam_pagina=${numberOfPagesToSearch}`;
}

const userInput = searchLink(rawSearchTerm);

const sleep = (ms) => new Promise((res) => setTimeout(res, ms));

const normalize = (s) =>
  (s || "")
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();

async function clickNextBtnFirst(page) {
  const els = await page.$$("#btn-next-page");
  if (!els || els.length === 0) return false;
  const btn = els[0];
  const disabled = await page.evaluate((e) => {
    return (
      e.hasAttribute("disabled") ||
      e.classList.contains("disabled") ||
      e.getAttribute("aria-disabled") === "true"
    );
  }, btn);
  if (disabled) return false;

  try {
    await btn.click();
    return true;
  } catch (err) {
    try {
      await page.evaluate((e) => e.scrollIntoView({ block: "center" }), btn);
      await btn.click();
      return true;
    } catch {
      return false;
    }
  }
}

const browser = await puppeteer.launch({ headless: true });
const page = await browser.newPage();
await page.setViewport({ width: 1920, height: 1080 });

const aggregatedResults = [];

const tokens = (rawSearchTerm || "")
  .split(/\s+/)
  .map((t) => t.replace(/[^\p{L}\p{N}_-]+/gu, ""))
  .filter(Boolean)
  .filter((t) => !/^\d+$/.test(t))
  .filter((t) => t.length >= 2);
const sortedByLen = tokens.sort((a, b) => b.length - a.length);
const selectedKeywords = sortedByLen.slice(0, 2).map(normalize);
const keywordCount = Math.min(2, selectedKeywords.length);

// console.log("Search URL:", userInput);
// console.log("Selected keywords:", selectedKeywords);
// console.log("Keyword matches required:", keywordCount);

await page.goto(userInput);

// pega lista de links da busca e processa até maxEditais
for (let index = 0; index < numberOfPagesToSearch; index++) {
  // garantir que a página de busca está carregada e os links visíveis
  await page.goto(userInput, { waitUntil: "networkidle2" });

  let hasResults = true;

  try {
    await page.waitForSelector("a.br-item", { timeout: 10000 });
  } catch (error) {
    hasResults = false;
  }

  if (!hasResults) {
    console.error(
      "\nNenhum resultado encontrado.\nTente usar outros termos de busca."
    );
    await browser.close();
    process.exit(0);
  }

  const links = await page.$$("a.br-item");
  if (!links.length) {
    console.log("Nenhum edital encontrado na busca. Encerrando.");
    break;
  }

  if (index >= links.length) {
    console.log(
      `Index ${index} fora do alcance dos links (${links.length}). Encerrando.`
    );
    break;
  }

  console.log(
    `\n===== Processando ${index + 1} de ${numberOfPagesToSearch} =====`
  );

  await links[index].click();

  await page.waitForFunction(
    () =>
      Array.from(document.querySelectorAll(".ng-star-inserted")).some((el) =>
        el.innerText?.includes("Data de início de recebimento de propostas")
      ),
    { timeout: 10000 }
  );

  const editalInfo = await page.$$eval(".ng-star-inserted", (elements) => {
    const extractDateAndTimeOnly = (text) => {
      const match = text.match(/(\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2})/);
      return match ? match[1] : null;
    };

    const result = { edital: null, dataInicio: null, dataFim: null };

    if (elements.length >= 9) {
      const rawEdital = elements[8].innerText || "";
      result.edital = rawEdital.replace(/\s+/g, " ").trim();
    }

    elements.forEach((el) => {
      const text = el.innerText?.replace(/\s+/g, " ").trim();
      if (!text) return;
      if (text.includes("Data de início de recebimento de propostas")) {
        result.dataInicio = extractDateAndTimeOnly(text);
      }
      if (text.includes("Data fim de recebimento de propostas")) {
        result.dataFim = extractDateAndTimeOnly(text);
      }
    });

    return result;
  });

  // console.log("Edital/Contr. Direta:", editalInfo.edital);
  // console.log("Data início:", editalInfo.dataInicio);
  // console.log("Data fim:", editalInfo.dataFim);
  // console.log("URL atual do edital:", page.url());

  await page.waitForSelector("ng-select .ng-select-container", {
    timeout: 15000
  });
  await page.click("ng-select .ng-select-container");
  await page.waitForSelector(".ng-option", { timeout: 10000 });

  const options = await page.$$(".ng-option");
  for (const option of options) {
    const text = await page.evaluate((el) => el.innerText.trim(), option);
    if (text === "50") {
      await option.click();
      break;
    }
  }

  await page.waitForSelector(".datatable-body-row", { timeout: 15000 });
  await page.waitForSelector(".pagination-information.d-none.d-sm-flex", {
    visible: true,
    timeout: 15000
  });

  async function collectRenderedItems() {
    return page.$$eval(".datatable-body-row", (rows) => {
      const results = [];
      for (const row of rows) {
        const cells = row.querySelectorAll(".datatable-body-cell-label");
        if (!cells || cells.length < 3) continue;

        const getText = (index) => {
          const span =
            cells[index]?.querySelector("span[title]") ||
            cells[index]?.querySelector("span");
          return span
            ? (span.getAttribute("title") || span.innerText)
                .replace(/\s+/g, " ")
                .trim()
            : "";
        };

        const numero = getText(0);
        const descricao = getText(1);
        const quantidadeRaw = getText(2);
        const quantidadeNum = Number(
          quantidadeRaw
            .replace(/\./g, "")
            .replace(",", ".")
            .replace(/[^\d\.\-]/g, "") || NaN
        );

        if (
          !numero ||
          isNaN(Number(numero)) ||
          !descricao ||
          isNaN(quantidadeNum)
        ) {
          continue;
        }

        results.push({
          numero,
          descricao,
          quantidade: quantidadeRaw,
          valorUnitarioEstimado: getText(3),
          valorTotalEstimado: getText(4)
        });
      }

      return results;
    });
  }

  const totalItems = await page.$eval(
    ".pagination-information.d-none.d-sm-flex",
    (el) => {
      const text = el.innerText || "";
      const match = text.match(/de\s+(\d+)\s+itens?/i);
      return match ? Number(match[1]) : 0;
    }
  );

  // console.log("Total de itens declarado no edital:", totalItems);

  const collectedMap = new Map();

  async function ensureCollectPageTarget(desiredGlobalCount, maxAttempts = 20) {
    let attempts = 0;
    while (attempts < maxAttempts) {
      attempts++;
      const rendered = await collectRenderedItems();
      for (const it of rendered) {
        if (it.numero && !collectedMap.has(it.numero)) {
          collectedMap.set(it.numero, it);
        }
      }
      if (collectedMap.size >= desiredGlobalCount) return true;
      await page.evaluate(() => {
        const body = document.querySelector(".datatable-body");
        if (body) body.scrollTop += 800;
        else window.scrollBy(0, 800);
      });
      await sleep(300);
    }
    return collectedMap.size >= desiredGlobalCount;
  }

  let remaining = totalItems > 0 ? totalItems : Infinity;
  let pageIndex = 0;
  while (remaining > 0) {
    pageIndex++;
    const pageTargetRelative = Math.min(
      50,
      remaining === Infinity ? 50 : remaining
    );
    const desiredGlobal = collectedMap.size + pageTargetRelative;

    // console.log(
    //   ` -> Página interna ${pageIndex}: alvo desta página = ${pageTargetRelative} (restantes=${remaining})`
    // );

    await ensureCollectPageTarget(desiredGlobal, 20);
    // console.log(
    //   ` -> Após coleta página ${pageIndex}: total acumulado no edital = ${collectedMap.size}`
    // );

    if (totalItems > 0 && collectedMap.size >= totalItems) break;
    if (remaining === Infinity) break;

    const clicked = await clickNextBtnFirst(page);
    if (!clicked) {
      console.log(
        " -> Não foi possível clicar no botão próxima do edital — parando paginação interna."
      );
      break;
    }

    const prev = await page.$eval(
      ".pagination-information.d-none.d-sm-flex",
      (el) => el.innerText || ""
    );
    try {
      await page.waitForFunction(
        (prevText) => {
          const el = document.querySelector(
            ".pagination-information.d-none.d-sm-flex"
          );
          return !!el && el.innerText.trim() !== prevText;
        },
        { timeout: 10000 },
        prev
      );
    } catch {
      await sleep(800);
    }

    await page.waitForSelector(".datatable-body-row", { timeout: 10000 });
    remaining =
      totalItems > 0 ? Math.max(0, totalItems - collectedMap.size) : Infinity;
    if (pageIndex >= 50) {
      // console.log(
      //   " -> Limite de páginas internas alcançado. Parando para segurança."
      // );
      break;
    }
  }

  let finalItems = Array.from(collectedMap.values());
  finalItems.sort((a, b) => Number(a.numero) - Number(b.numero));
  if (totalItems > 0) finalItems = finalItems.slice(0, totalItems);

  // console.log(`-> Itens coletados nesse edital: ${finalItems.length}`);

  if (keywordCount === 0) {
    console.log(
      "Nenhuma palavra-chave válida para filtrar. Mantendo todos os itens coletados."
    );
  }
  const filtered = finalItems.filter((it) => {
    if (keywordCount === 0) return true;
    const desc = normalize(it.descricao);
    let matches = 0;
    for (const k of selectedKeywords) {
      if (!k) continue;
      if (desc.includes(k)) matches++;
    }
    return matches >= Math.min(keywordCount, selectedKeywords.length);
  });

  // console.log(
  //   `-> Itens filtrados do edital (após keywords): ${filtered.length}`
  // );
  // console.log(JSON.stringify(filtered, null, 2));

  aggregatedResults.push({
    sourceEdital: editalInfo.edital,
    dataInicio: editalInfo.dataInicio,
    dataFim: editalInfo.dataFim,
    url: page.url(),
    filteredItems: filtered
  });

  await page.goto(userInput, { waitUntil: "networkidle2" });
  await sleep(500); // pequeno respiro antes do próximo loop
}

console.log("\n===== Resumo final de todos os editais processados =====");
console.log(`Editais processados: ${aggregatedResults.length}`);
// console.log(JSON.stringify(aggregatedResults, null, 2));

async function exportToExcel(aggregatedResults, outputDir, filename = null) {
  if (!Array.isArray(aggregatedResults) || aggregatedResults.length === 0) {
    console.log("Nenhum dado para exportar.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("PNCP Results");

  sheet.columns = [
    { header: "URL / Link", key: "url", width: 45 },
    { header: "Edital / Contratação Direta", key: "edital", width: 40 },
    {
      header: "Data de início de receb. de propostas",
      key: "dataInicio",
      width: 20
    },
    { header: "Data fim de receb. de propostas", key: "dataFim", width: 20 },
    { header: "Número", key: "numero", width: 12 },
    { header: "Descrição", key: "descricao", width: 100 },
    { header: "Quantidade", key: "quantidade", width: 14 },
    {
      header: "Valor unitário estimado",
      key: "valorUnitarioEstimado",
      width: 20
    },
    { header: "Valor total estimado", key: "valorTotalEstimado", width: 20 }
  ];

  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center" };
  sheet.views = [{ state: "frozen", ySplit: 1 }]; // congela header
  sheet.autoFilter = {
    from: "A1",
    to: "I1"
  };

  sheet.getColumn("descricao").alignment = { wrapText: true };

  const centerCols = [
    "url",
    "edital",
    "dataInicio",
    "dataFim",
    "numero",
    "quantidade",
    "valorUnitarioEstimado",
    "valorTotalEstimado"
  ];

  centerCols.forEach((key) => {
    const col = sheet.getColumn(key);
    col.alignment = { horizontal: "center", vertical: "middle" };
  });

  const parseItemNumber = (raw) => {
    if (raw === null || raw === undefined) return null;
    const s = String(raw).trim();

    if (!/^\d+$/.test(s)) return null;

    const n = Number(s);
    return Number.isNaN(n) ? null : n;
  };

  const parseNumber = (raw) => {
    if (raw === null || raw === undefined) return null;
    let s = String(raw).trim();
    if (!s) return null;
    s = s.replace(/[^\d\.\,\-]/g, "");
    if (!s) return null;

    const hasDot = s.indexOf(".") !== -1;
    const hasComma = s.indexOf(",") !== -1;
    if (hasDot && hasComma) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else if (hasComma && !hasDot) {
      s = s.replace(",", ".");
    } else if (hasDot && !hasComma) {
      const dotCount = (s.match(/\./g) || []).length;
      if (dotCount > 1) {
        s = s.replace(/\./g, "");
      }
    }
    const n = Number(s);
    return Number.isNaN(n) ? null : n;
  };

  const parseCurrency = (raw) => {
    return parseNumber(raw);
  };

  for (const ed of aggregatedResults) {
    const url = ed.url || "";
    const edital = ed.sourceEdital || ed.edital || "";
    const dataInicio = ed.dataInicio || "";
    const dataFim = ed.dataFim || "";

    const items = Array.isArray(ed.filteredItems) ? ed.filteredItems : [];

    if (items.length === 0) {
      const r = sheet.addRow({
        url,
        edital,
        dataInicio,
        dataFim,
        numero: "",
        descricao: "",
        quantidade: "",
        valorUnitarioEstimado: "",
        valorTotalEstimado: ""
      });

      for (let c = 1; c <= sheet.columnCount; c++) {
        r.getCell(c).border = {
          bottom: { style: "thin", color: { argb: "FF999999" } }
        };
      }
      sheet.addRow({});
      continue;
    }

    for (let i = 0; i < items.length; i++) {
      const it = items[i] || {};

      const numeroItem = parseItemNumber(it.numero);
      const quantidadeNum = parseNumber(it.quantidade);
      const valorUnitNum = parseCurrency(it.valorUnitarioEstimado);
      const valorTotalNum = parseCurrency(it.valorTotalEstimado);

      const rowObj =
        i === 0
          ? {
              url,
              edital,
              dataInicio,
              dataFim,
              numero: numeroItem !== null ? numeroItem : it.numero || "",
              descricao: it.descricao || "",
              quantidade:
                quantidadeNum !== null ? quantidadeNum : it.quantidade || "",
              valorUnitarioEstimado:
                valorUnitNum !== null
                  ? valorUnitNum
                  : it.valorUnitarioEstimado || "",
              valorTotalEstimado:
                valorTotalNum !== null
                  ? valorTotalNum
                  : it.valorTotalEstimado || ""
            }
          : {
              url: "",
              edital: "",
              dataInicio: "",
              dataFim: "",
              numero: numeroItem !== null ? numeroItem : it.numero || "",
              descricao: it.descricao || "",
              quantidade:
                quantidadeNum !== null ? quantidadeNum : it.quantidade || "",
              valorUnitarioEstimado:
                valorUnitNum !== null
                  ? valorUnitNum
                  : it.valorUnitarioEstimado || "",
              valorTotalEstimado:
                valorTotalNum !== null
                  ? valorTotalNum
                  : it.valorTotalEstimado || ""
            };

      const newRow = sheet.addRow(rowObj);

      newRow.eachCell((cell, colNumber) => {
        if (colNumber === 6) {
          cell.alignment = {
            vertical: "top",
            wrapText: true
          };
        }
      });

      const startDateCell = newRow.getCell(3);
      if (startDateCell.value instanceof Date) {
        startDateCell.numFmt = "dd/mm/yyyy hh:mm";
      }

      const endDateCell = newRow.getCell(4);
      if (endDateCell.value instanceof Date) {
        endDateCell.numFmt = "dd/mm/yyyy hh:mm";
      }

      const numCell = newRow.getCell(5);
      if (typeof numCell.value === "number") {
        numCell.numFmt = "0";
      }

      const qtyCell = newRow.getCell(7);
      if (typeof qtyCell.value === "number") {
        qtyCell.numFmt = "#,##0";
      }

      const unitCell = newRow.getCell(8);
      if (typeof unitCell.value === "number") {
        unitCell.numFmt = '"R$" #,##0.00;[Red]\-"R$" #,##0.00';
      }

      const totalCell = newRow.getCell(9);
      if (typeof totalCell.value === "number") {
        totalCell.numFmt = '"R$" #,##0.00;[Red]\-"R$" #,##0.00';
      }
    }

    const lastRowNumber = sheet.lastRow.number;
    const lastRow = sheet.getRow(lastRowNumber);

    for (let c = 1; c <= sheet.columnCount; c++) {
      const cell = lastRow.getCell(c);
      cell.border = {
        bottom: { style: "thin", color: { argb: "FF666666" } }
      };
    }
  }

  for (let i = 2; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);
    const desc = row.getCell("descricao").value;

    if (typeof desc === "string") {
      const lines = Math.ceil(desc.length / 80);
      row.height = Math.min(140, Math.max(20, lines * 18));
    } else {
      row.height = 20;
    }
  }

  sheet.columns.forEach((column) => {
    let maxLength = 10;

    column.eachCell({ includeEmpty: true }, (cell) => {
      const cellValue = cell.value;
      if (cellValue == null) return;

      let text = "";
      if (typeof cellValue === "string") text = cellValue;
      else if (typeof cellValue === "number") text = cellValue.toString();
      else if (cellValue instanceof Date) text = cellValue.toLocaleString();
      else if (typeof cellValue === "object" && cellValue.text)
        text = cellValue.text;

      maxLength = Math.max(maxLength, text.length);
    });

    column.width = Math.min(Math.max(maxLength + 2, 10), 60);
  });

  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const day = pad(now.getDate());
  const month = pad(now.getMonth() + 1);
  const year = String(now.getFullYear()).slice(-2);
  const hour = pad(now.getHours());
  const minute = pad(now.getMinutes());
  const ts = `${day}-${month}-${year}_${hour}-${minute}`;

  const outFile = filename || `PNCP_resultados_${ts}.xlsx`;
  const outPath = path.join(outputDir, outFile);

  await workbook.xlsx.writeFile(outPath);
  console.log("Arquivo Excel gravado em:", outPath);

  return outPath;
}

await exportToExcel(aggregatedResults, outputDir);
await browser.close();
