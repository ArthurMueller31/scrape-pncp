import puppeteer from "puppeteer";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import readline from "readline";

const CONFIG_FILE = path.resolve(process.cwd(), "config.json");
const args = process.argv.slice(2);

const forceAskPath = args.includes("--arquivo");

const rawSearchTerm = args[0];
let numberOfPagesToSearch = Number(args[1] || 1);

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

const ITEMS_PER_PAGE = 50;
const config = loadConfig();
let outputDir = config.outputDir;

if (!outputDir || forceAskPath) {
  const answer = await askDirectory(
    "Informe o diretório onde os arquivos devem ser salvos:\n> "
  );

  if (!answer) {
    console.error("Diretório inválido.");
    process.exit(1);
  }

  const resolvedPath = path.resolve(answer);
  fs.mkdirSync(resolvedPath, { recursive: true });
  outputDir = resolvedPath;
  saveConfig({ ...config, outputDir });
  console.log("Usando diretório salvo:", outputDir);
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
function promptEnter(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer);
    });
  });
}
async function promptYesNo(question, defaultYes = false) {
  while (true) {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    const answer = await new Promise((resolve) => {
      rl.question(question, (ans) => {
        rl.close();
        resolve(ans);
      });
    });
    const v = (answer || "").trim().toLowerCase();
    if (v === "") return defaultYes ? "s" : "n";
    if (v === "s" || v === "n") return v;
    console.log("Por favor digite 'S' ou 'N'.");
  }
}
function promptNumber(question, min = 1, max = 10, defaultVal = 1) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      const v = (answer || "").trim();
      if (v === "") return resolve(defaultVal);
      const n = Number(v);
      if (!Number.isInteger(n) || n < min || n > max) return resolve(null);
      return resolve(n);
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
function saveConfig(config) {
  fs.writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2), "utf-8");
}

console.log("Termo de busca:", rawSearchTerm);
console.log("Páginas solicitadas:", numberOfPagesToSearch);

function buildSearchUrl(
  receivedUserInput,
  tamPagina = ITEMS_PER_PAGE,
  page = 1
) {
  const search = (receivedUserInput || "").trim();
  if (!search) throw new Error("A busca não pode ficar vazia.");
  const encodedInput = encodeURIComponent(receivedUserInput);
  return `https://pncp.gov.br/app/editais?q=${encodedInput}&status=recebendo_proposta&pagina=${page}&ordenacao=relevancia&tam_pagina=${tamPagina}`;
}

const sleep = (ms) => new Promise((res) => setTimeout(res, ms));

const normalize = (s) =>
  (s || "")
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();

async function safeWaitForSelector(page, selector, timeout = 10000) {
  try {
    await page.waitForSelector(selector, { timeout });
    return true;
  } catch {
    return false;
  }
}
async function safeWaitForFunction(page, fnOrStr, timeout = 10000, ...args) {
  try {
    await page.waitForFunction(fnOrStr, { timeout }, ...args);
    return true;
  } catch {
    return false;
  }
}
async function safeClickHandle(handle) {
  try {
    await handle.click();
    return true;
  } catch {
    try {
      await handle.evaluate((el) => el.click());
      return true;
    } catch {
      return false;
    }
  }
}
async function safeClickByIndexOnPage(page, indexInPage) {
  try {
    const clicked = await page.evaluate((i) => {
      const els = Array.from(document.querySelectorAll("a.br-item"));
      if (!els || els.length <= i) return false;
      els[i].scrollIntoView({ block: "center" });
      els[i].click();
      return true;
    }, indexInPage);
    return Boolean(clicked);
  } catch {
    return false;
  }
}

function parseDateFromString(s) {
  if (!s || typeof s !== "string") return null;
  const trimmed = s.trim();
  let m = trimmed.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
  if (m) {
    const day = Number(m[1]),
      month = Number(m[2]),
      year = Number(m[3]);
    const hour = Number(m[4]),
      minute = Number(m[5]);
    return new Date(year, month - 1, day, hour, minute, 0, 0);
  }
  m = trimmed.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) {
    const day = Number(m[1]),
      month = Number(m[2]),
      year = Number(m[3]);
    return new Date(year, month - 1, day, 0, 0, 0, 0);
  }
  return null;
}

const detectionBrowser = await puppeteer.launch({ headless: true });
const detectionPage = await detectionBrowser.newPage();
await detectionPage.setViewport({ width: 1200, height: 900 });

await detectionPage
  .goto(buildSearchUrl(rawSearchTerm, ITEMS_PER_PAGE, 1), {
    waitUntil: "networkidle2",
    timeout: 30000
  })
  .catch(() => {});

let detectedTotalItems = null;
try {
  await detectionPage
    .waitForSelector(".pagination-information.d-none.d-sm-flex", {
      timeout: 5000
    })
    .catch(() => {});
  const paginationText = await detectionPage.evaluate(() => {
    const el = document.querySelector(
      ".pagination-information.d-none.d-sm-flex"
    );
    return el ? el.innerText : "";
  });

  let match = (paginationText || "").match(/de\s+(\d+)\s+itens?/i);
  if (!match)
    match = (paginationText || "").match(/Página\s+\d+\s+de\s+(\d+)/i);
  if (!match) match = (paginationText || "").match(/de\s+(\d+)/i);
  if (match && match[1]) {
    const n = Number(match[1].replace(/\D/g, ""));
    if (!Number.isNaN(n) && n > 0) detectedTotalItems = n;
  }
} catch {}
await detectionPage.close().catch(() => {});
await detectionBrowser.close().catch(() => {});

console.log(
  "Total de páginas encontradas para essa pesquisa:",
  detectedTotalItems !== null ? detectedTotalItems : "desconhecido"
);

let totalToProcess = numberOfPagesToSearch;
if (detectedTotalItems !== null && numberOfPagesToSearch > detectedTotalItems) {
  totalToProcess = detectedTotalItems;
}

const answerConfirm = await promptEnter(
  `Continuar com ${totalToProcess}? (Aperte Enter para continuar, qualquer outra tecla + Enter para cancelar) `
);
if (answerConfirm.trim() !== "") process.exit(0);

const yn = await promptYesNo(
  "Deseja usar mais processos simultâneos? (S - sim / N (ou Enter) - não) - Mais processos vão usar mais recursos do PC, mas pode ir mais rápido\n",
  false
);
let concurrency = yn === "s" ? Math.min(3, totalToProcess) : 1;

function splitRanges(total, parts) {
  const ranges = [];
  if (total <= 0) return ranges;
  parts = Math.min(parts, total);
  const base = Math.floor(total / parts);
  const rem = total % parts;
  let cursor = 1;
  for (let i = 0; i < parts; i++) {
    const size = base + (i < rem ? 1 : 0);
    if (size <= 0) continue;
    const start = cursor;
    const end = cursor + size - 1;
    ranges.push({ start, end });
    cursor = end + 1;
  }
  return ranges;
}

const ranges = splitRanges(totalToProcess, concurrency);

const tokens = (rawSearchTerm || "")
  .split(/\s+/)
  .map((t) => t.replace(/[^\p{L}\p{N}_-]+/gu, ""))
  .filter(Boolean)
  .filter((t) => !/^\d+$/.test(t))
  .filter((t) => t.length >= 2);
const sortedByLen = tokens.sort((a, b) => b.length - a.length);
const selectedKeywords = sortedByLen.slice(0, 2).map(normalize);
const keywordCount = Math.min(2, selectedKeywords.length);

let processedCount = 0;
const total = totalToProcess;
const step = concurrency;

const firstBatch = Math.min(step, total);
console.log(`\n===== Processando ${firstBatch} de ${total} =====`);

const interimResults = [];
let lastBackupPath = null;
let lastBackupMultiple = 0;

function getBaseName(raw) {
  const base =
    (raw || "search")
      .toString()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 3)
      .join("-") || "search";
  return base;
}

async function workerWithRange(range, workerId) {
  const workerResults = [];
  let browser = null;
  let pageTab = null;
  try {
    browser = await puppeteer.launch({ headless: true });
    pageTab = await browser.newPage();
    await pageTab.setViewport({ width: 1200, height: 900 });

    for (
      let globalIndex = range.start;
      globalIndex <= range.end;
      globalIndex++
    ) {
      try {
        const pageNum = Math.floor((globalIndex - 1) / ITEMS_PER_PAGE) + 1;
        const indexInPage = (globalIndex - 1) % ITEMS_PER_PAGE;

        const searchUrl = buildSearchUrl(
          rawSearchTerm,
          ITEMS_PER_PAGE,
          pageNum
        );
        await pageTab
          .goto(searchUrl, { waitUntil: "networkidle2", timeout: 30000 })
          .catch(() => {});

        const hasResults = await safeWaitForSelector(
          pageTab,
          "a.br-item",
          8000
        );
        if (!hasResults) continue;

        const linksCount = await pageTab
          .$$eval("a.br-item", (els) => els.length)
          .catch(() => 0);
        if (linksCount === 0 || indexInPage >= linksCount) continue;

        const clicked = await safeClickByIndexOnPage(pageTab, indexInPage);
        if (!clicked) {
          const handles = await pageTab.$$("a.br-item");
          if (handles[indexInPage]) {
            const ok = await safeClickHandle(handles[indexInPage]);
            if (!ok) continue;
          } else continue;
        }

        const hasEditalContent = await safeWaitForFunction(
          pageTab,
          () =>
            Array.from(document.querySelectorAll(".ng-star-inserted")).some(
              (el) =>
                el.innerText?.includes(
                  "Data de início de recebimento de propostas"
                )
            ),
          10000
        );
        if (!hasEditalContent) continue;

        let editalInfo = { edital: null, dataInicio: null, dataFim: null };
        try {
          editalInfo = await pageTab.$$eval(".ng-star-inserted", (elements) => {
            const extractDateAndTimeOnly = (text) => {
              const m = text.match(/(\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2})/);
              if (m) return m[1];
              const m2 = text.match(/(\d{2}\/\d{2}\/\d{4})/);
              return m2 ? m2[1] : null;
            };
            const result = { edital: null, dataInicio: null, dataFim: null };
            if (elements.length >= 9) {
              const rawEdital = elements[8].innerText || "";
              result.edital = rawEdital.replace(/\s+/g, " ").trim();
            }
            elements.forEach((el) => {
              const text = el.innerText?.replace(/\s+/g, " ").trim();
              if (!text) return;
              if (text.includes("Data de início de recebimento de propostas"))
                result.dataInicio = extractDateAndTimeOnly(text);
              if (text.includes("Data fim de recebimento de propostas"))
                result.dataFim = extractDateAndTimeOnly(text);
            });
            return result;
          });
        } catch {}

        const dataInicioDate = parseDateFromString(editalInfo.dataInicio);
        const dataFimDate = parseDateFromString(editalInfo.dataFim);

        const collectedMap = new Map();
        async function collectRenderedItemsSafe() {
          try {
            return await pageTab.$$eval(".datatable-body-row", (rows) => {
              const results = [];
              for (const row of rows) {
                const cells = row.querySelectorAll(
                  ".datatable-body-cell-label"
                );
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
                )
                  continue;
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
          } catch {
            return [];
          }
        }

        async function ensureCollect(desiredGlobalCount, maxAttempts = 15) {
          let attempts = 0;
          while (attempts < maxAttempts) {
            attempts++;
            const rendered = await collectRenderedItemsSafe();
            for (const it of rendered) {
              if (it.numero && !collectedMap.has(it.numero)) {
                collectedMap.set(it.numero, it);
              }
            }
            if (collectedMap.size >= desiredGlobalCount) return true;
            await pageTab
              .evaluate(() => {
                const b = document.querySelector(".datatable-body");
                if (b) b.scrollTop += 800;
                else window.scrollBy(0, 800);
              })
              .catch(() => {});
            await sleep(300);
          }
          return collectedMap.size >= desiredGlobalCount;
        }

        let totalItems = 0;
        try {
          const infoPresent = await pageTab
            .$eval(
              ".pagination-information.d-none.d-sm-flex",
              (el) => el.innerText
            )
            .catch(() => "");
          const m = (infoPresent || "").match(/de\s+(\d+)\s+itens?/i);
          if (m && m[1]) totalItems = Number(m[1]);
        } catch {}

        const desiredGlobal = totalItems > 0 ? totalItems : ITEMS_PER_PAGE;
        await ensureCollect(desiredGlobal, 15).catch(() => {});

        if (totalItems > 0) {
          let attemptsPage = 0;
          while (collectedMap.size < totalItems && attemptsPage < 30) {
            attemptsPage++;
            const nextClicked = await pageTab
              .evaluate(() => {
                const b = document.querySelector("#btn-next-page");
                if (!b) return false;
                if (
                  b.hasAttribute("disabled") ||
                  b.classList.contains("disabled") ||
                  b.getAttribute("aria-disabled") === "true"
                )
                  return false;
                b.click();
                return true;
              })
              .catch(() => false);
            if (!nextClicked) break;
            await safeWaitForSelector(pageTab, ".datatable-body-row", 8000);
            await ensureCollect(desiredGlobal, 8).catch(() => {});
          }
        }

        let finalItems = Array.from(collectedMap.values()).sort(
          (a, b) => Number(a.numero) - Number(b.numero)
        );
        if (totalItems > 0) finalItems = finalItems.slice(0, totalItems);

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

        if (filtered.length > 0) {
          const entry = {
            sourceEdital: editalInfo.edital,
            dataInicio: dataInicioDate,
            dataFim: dataFimDate,
            url: pageTab.url(),
            filteredItems: filtered
          };
          workerResults.push(entry);
          interimResults.push(entry);
        }

        await pageTab
          .goto(buildSearchUrl(rawSearchTerm, ITEMS_PER_PAGE, pageNum), {
            waitUntil: "networkidle2"
          })
          .catch(() => {});
        await sleep(200);
      } catch (err) {
        try {
          await pageTab
            .goto(
              buildSearchUrl(
                rawSearchTerm,
                ITEMS_PER_PAGE,
                Math.floor((globalIndex - 1) / ITEMS_PER_PAGE) + 1
              ),
              { waitUntil: "networkidle2" }
            )
            .catch(() => {});
        } catch {}
        await sleep(200);
        continue;
      } finally {
        processedCount++;
        if (processedCount % step === 0 || processedCount === total) {
          const nextBatch = Math.min(processedCount + step, total);
          if (processedCount < total) {
            console.log(`\n===== Processando ${nextBatch} de ${total} =====`);
          }
        }
        if (processedCount % 15 === 0 && processedCount > lastBackupMultiple) {
          lastBackupMultiple = processedCount;
          if (interimResults.length > 0) {
            try {
              const backupPath = await exportToExcel(
                interimResults,
                outputDir,
                rawSearchTerm,
                true
              );
              if (backupPath) {
                lastBackupPath = backupPath;
                console.log("Backup de progresso criado em:", backupPath);
              }
            } catch (e) {}
          }
        }
      }
    }

    return workerResults;
  } finally {
    try {
      if (pageTab) await pageTab.close().catch(() => {});
    } catch {}
    try {
      if (browser) await browser.close().catch(() => {});
    } catch {}
  }
}

const startTimeMs = Date.now();
const workerPromises = ranges.map((r, i) => workerWithRange(r, i + 1));
const workerResults = await Promise.all(workerPromises);
const aggregatedResults = workerResults.flat();

const totalFilteredItems = aggregatedResults.reduce(
  (s, e) => s + (Array.isArray(e.filteredItems) ? e.filteredItems.length : 0),
  0
);
const uniqueDescSet = new Set();
for (const ed of aggregatedResults) {
  const items = Array.isArray(ed.filteredItems) ? ed.filteredItems : [];
  for (const it of items) {
    const descNorm = normalize(it.descricao || "");
    if (descNorm) uniqueDescSet.add(descNorm);
  }
}
const uniqueProductsCount = uniqueDescSet.size;

const outPath = await exportToExcel(
  aggregatedResults,
  outputDir,
  rawSearchTerm
);

try {
  const backupFile = path.join(
    outputDir,
    `${getBaseName(rawSearchTerm)}_backup.xlsx`
  );
  if (fs.existsSync(backupFile)) fs.unlinkSync(backupFile);
} catch {}

function formatDuration(ms) {
  const totalSec = Math.floor(ms / 1000);
  const hours = Math.floor(totalSec / 3600);
  const minutes = Math.floor((totalSec % 3600) / 60);
  const seconds = totalSec % 60;
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
}

console.log(
  `\nEditais gravados: Dos ${totalToProcess}, ${aggregatedResults.length} continham o termo de busca`
);
console.log("Total de itens filtrados:", totalFilteredItems);
console.log("Produtos diferentes:", uniqueProductsCount);
console.log("Tempo total decorrido:", formatDuration(Date.now() - startTimeMs));
console.log("Arquivo Excel gravado em:", outPath);

async function exportToExcel(
  aggregatedResults,
  outputDir,
  rawSearchTermArg = "",
  isBackup = false
) {
  if (!Array.isArray(aggregatedResults) || aggregatedResults.length === 0) {
    console.log("Nenhum dado para exportar.");
    return "";
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

  for (const col of sheet.columns) {
    if (col && col.key === "descricao") {
      col.alignment = { vertical: "top", wrapText: true };
    } else if (col) {
      col.alignment = { vertical: "middle", horizontal: "center" };
    }
  }

  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center" };
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = { from: "A1", to: "I1" };

  const parseItemNumber = (raw) => {
    if (raw == null) return null;
    const s = String(raw).trim();
    if (!/^\d+$/.test(s)) return null;
    const n = Number(s);
    return Number.isNaN(n) ? null : n;
  };
  const parseNumber = (raw) => {
    if (raw == null) return null;
    let s = String(raw).trim();
    if (!s) return null;
    s = s.replace(/[^\d\.\,\-]/g, "");
    if (!s) return null;
    const hasDot = s.indexOf(".") !== -1;
    const hasComma = s.indexOf(",") !== -1;
    if (hasDot && hasComma) s = s.replace(/\./g, "").replace(",", ".");
    else if (hasComma && !hasDot) s = s.replace(",", ".");
    else if (hasDot && !hasComma) {
      const dotCount = (s.match(/\./g) || []).length;
      if (dotCount > 1) s = s.replace(/\./g, "");
    }
    const n = Number(s);
    return Number.isNaN(n) ? null : n;
  };
  const parseCurrency = (raw) => parseNumber(raw);

  for (const ed of aggregatedResults) {
    const url = ed.url || "";
    const edital = ed.sourceEdital || ed.edital || "";
    const dataInicio = ed.dataInicio instanceof Date ? ed.dataInicio : "";
    const dataFim = ed.dataFim instanceof Date ? ed.dataFim : "";

    const items = Array.isArray(ed.filteredItems) ? ed.filteredItems : [];
    if (!items.length) continue;

    for (let i = 0; i < items.length; i++) {
      const it = items[i] || {};
      const numeroItem = parseItemNumber(it.numero);
      const quantidadeNum = parseNumber(it.quantidade);
      const valorUnitNum = parseCurrency(it.valorUnitarioEstimado);
      const valorTotalNum = parseCurrency(it.valorTotalEstimado);

      const rowObj = {
        url,
        edital,
        dataInicio,
        dataFim,
        numero: numeroItem !== null ? numeroItem : it.numero || "",
        descricao: it.descricao || "",
        quantidade:
          quantidadeNum !== null ? quantidadeNum : it.quantidade || "",
        valorUnitarioEstimado:
          valorUnitNum !== null ? valorUnitNum : it.valorUnitarioEstimado || "",
        valorTotalEstimado:
          valorTotalNum !== null ? valorTotalNum : it.valorTotalEstimado || ""
      };

      const newRow = sheet.addRow(rowObj);

      const startDateCell = newRow.getCell(3);
      if (startDateCell.value instanceof Date)
        startDateCell.numFmt = "dd/mm/yyyy";
      const endDateCell = newRow.getCell(4);
      if (endDateCell.value instanceof Date) endDateCell.numFmt = "dd/mm/yyyy";

      const numCell = newRow.getCell(5);
      if (typeof numCell.value === "number") numCell.numFmt = "0";
      const qtyCell = newRow.getCell(7);
      if (typeof qtyCell.value === "number") qtyCell.numFmt = "#,##0";
      const unitCell = newRow.getCell(8);
      if (typeof unitCell.value === "number")
        unitCell.numFmt = '"R$" #,##0.00;[Red]\-"R$" #,##0.00';
      const totalCell = newRow.getCell(9);
      if (typeof totalCell.value === "number")
        totalCell.numFmt = '"R$" #,##0.00;[Red]\-"R$" #,##0.00';
    }

    const lastRowNumber = sheet.lastRow.number;
    const lastRow = sheet.getRow(lastRowNumber);
    for (let c = 1; c <= sheet.columnCount; c++)
      lastRow.getCell(c).border = {
        bottom: { style: "thin", color: { argb: "FF666666" } }
      };
  }

  for (let i = 2; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);
    const desc = row.getCell("descricao").value;
    if (typeof desc === "string") {
      const lines = Math.ceil(desc.length / 80);
      row.height = Math.min(140, Math.max(20, lines * 18));
    } else row.height = 20;
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
  const base = getBaseName(rawSearchTermArg);
  let outFile;
  if (isBackup) outFile = `${base}_backup.xlsx`;
  else outFile = `${base}_${ts}.xlsx`;
  const outPath = path.join(outputDir, outFile);
  await workbook.xlsx.writeFile(outPath);
  return outPath;
}
