import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import XLSX from "xlsx";
import type { Page } from "puppeteer-core";
import type { BankMovement, BankScraper, CardOwner, CreditCardBalance, MovementSource, ScrapeResult, ScraperOptions } from "../types.js";
import { MOVEMENT_SOURCE } from "../types.js";
import { closePopups, delay, deduplicateMovements, deduplicateAcrossSources, monthYearLabel, normalizeDate, normalizeOwner, normalizeInstallments, parseChileanAmount } from "../utils.js";
import { runScraper } from "../infrastructure/scraper-runner.js";
import type { BrowserSession } from "../infrastructure/browser.js";
import { fillRut, fillPassword, clickSubmit } from "../actions/login.js";
import { clickByText, dismissBanners } from "../actions/navigation.js";
import { extractAccountMovements } from "../actions/extraction.js";
import { paginateAndExtract } from "../actions/pagination.js";
import {
  createTempDownloadDir,
  setupPageDownload,
  snapshotDir,
  waitForDownloadedFile,
  cleanupTempDir,
} from "../infrastructure/downloader.js";

// ─── Constants ───────────────────────────────────────────────────

const BANK_URL = "https://www.bancofalabella.cl";
const SHADOW_HOST = "credit-card-movements";
const DOWNLOADS_DIR = path.join(os.homedir(), "Downloads");
const INVOICED_RADIO_ID = "invoicedMovements";
const MAX_PAGES = 20;
const PAGE_CHANGE_TIMEOUT_MS = 15000;
const CMR_MOVEMENTS_TIMEOUT_MS = 30000;

// ─── Excel parsing ───────────────────────────────────────────────

function deriveInstallments(pending: number, monto: number, valorCuota: number): string {
  if (pending === 0) return "01/01";
  // ceil: interest plans have monto < valorCuota×total (fractional ratio → round up)
  const total = valorCuota > 0 ? Math.ceil(monto / valorCuota) : 1;
  const current = Math.max(1, total - pending);
  return `${String(current).padStart(2, "0")}/${String(total).padStart(2, "0")}`;
}

function excelSerialToDate(serial: number): string {
  const d = XLSX.SSF.parse_date_code(serial);
  if (!d) return "pendiente";
  return `${String(d.d).padStart(2, "0")}-${String(d.m).padStart(2, "0")}-${String(d.y).padStart(4, "0")}`;
}

/**
 * Parses a Falabella Excel export (no_facturados or facturados) into BankMovements.
 * Columns: FECHA | DESCRIPCION | TITULAR/ADICIONAL | MONTO | CUOTAS PENDIENTES | VALOR CUOTA
 *
 * installments logic:
 *   cuotasPendientes = 0 → "01/01" (cargo único)
 *   cuotasPendientes > 0 → undefined (no tenemos suficiente info para derivar XX/YY)
 */
function parseExcelMovements(filePath: string, source: MovementSource, debugLog: string[] = []): BankMovement[] {
  try {
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][];

    const movements: BankMovement[] = [];
    for (const row of rows.slice(1)) {
      const dateRaw = row[0];
      const description = String(row[1] ?? "").trim();
      const ownerRaw = String(row[2] ?? "").trim();
      const monto = typeof row[3] === "number" ? row[3] : parseChileanAmount(String(row[3] ?? ""));
      const cuotasPendientes = Number(row[4]) || 0;
      // VALOR CUOTA is negative for credits/payments, positive for purchases.
      // Cells may arrive as numbers or as currency strings ("-$1.769.390").
      const valorCuota = typeof row[5] === "number" ? row[5] : parseChileanAmount(String(row[5] ?? ""));

      if (!description || monto === 0) continue;

      const date =
        typeof dateRaw === "number"
          ? excelSerialToDate(dateRaw)
          : normalizeDate(String(dateRaw));

      // VALOR CUOTA is the installment amount billed this period (negative for payments).
      const amount = -valorCuota;
      const totalAmount = Math.abs(monto);
      movements.push({
        date,
        description,
        amount,
        balance: 0,
        source,
        owner: normalizeOwner(ownerRaw),
        installments: deriveInstallments(cuotasPendientes, Math.abs(monto), Math.abs(valorCuota)),
        totalAmount,
      });
    }
    return movements;
  } catch (err) {
    debugLog.push(`  [Excel] Parse error: ${err instanceof Error ? err.message : String(err)}`);
    return [];
  }
}

// ─── CMR billing info extraction ─────────────────────────────────

interface UnbilledPeriodInfo {
  nextBillingDate?: string;
  nextDueDate?: string;
  periodExpenses?: number;
}

async function extractUnbilledPeriodInfo(page: Page): Promise<UnbilledPeriodInfo> {
  return page.evaluate((host: string) => {
    const shadowEl = document.querySelector(host);
    const topRoot = (shadowEl as Element & { shadowRoot?: ShadowRoot })?.shadowRoot || document;

    function collectAllRoots(root: ShadowRoot | Element): Array<ShadowRoot | Element> {
      const found: Array<ShadowRoot | Element> = [root];
      for (const el of Array.from((root as Element).querySelectorAll("*"))) {
        const sr = (el as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
        if (sr) found.push(...collectAllRoots(sr));
      }
      return found;
    }

    // The period info is in divs whose textContent starts with "Próxima facturación",
    // "Próximo vencimiento", or "Gastos del periodo" — label and value are in the SAME div.
    function extractFromSameDiv(root: ShadowRoot | Element, label: string): string | undefined {
      for (const div of Array.from((root as Element).querySelectorAll("div"))) {
        const text = div.textContent?.trim() || "";
        if (text.toLowerCase().startsWith(label.toLowerCase())) {
          const rest = text.slice(label.length).trim();
          if (rest) return rest;
        }
      }
      return undefined;
    }

    function parseAmount(text?: string): number | undefined {
      if (!text) return undefined;
      const m = text.match(/\$([\d.,]+)/);
      if (!m) return undefined;
      return parseInt(m[1].replace(/\./g, "").replace(",", ""), 10) || undefined;
    }

    function extractDate(text?: string): string | undefined {
      if (!text) return undefined;
      const m = text.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
      return m ? m[1] : undefined;
    }

    let billingRaw: string | undefined;
    let dueRaw: string | undefined;
    let expensesRaw: string | undefined;

    for (const root of collectAllRoots(topRoot)) {
      if (!billingRaw) billingRaw = extractFromSameDiv(root, "Pr\u00f3xima facturaci\u00f3n");
      if (!dueRaw) dueRaw = extractFromSameDiv(root, "Pr\u00f3ximo vencimiento");
      if (!expensesRaw) expensesRaw = extractFromSameDiv(root, "Gastos del periodo");
    }

    return {
      nextBillingDate: extractDate(billingRaw),
      nextDueDate: extractDate(dueRaw),
      periodExpenses: parseAmount(expensesRaw),
    };
  }, SHADOW_HOST);
}

interface BilledStatementInfo {
  billingDate?: string;
  billedAmount?: number;
  dueDate?: string;
  minimumPayment?: number;
}

/**
 * Extracts lastStatement fields from the "movimientos facturados" tab.
 * The HTML structure has label and value in ADJACENT sibling divs:
 *   <div>Fecha de facturación</div><div>19/03/2026</div>
 *   <div>Monto facturado</div><div>$3.449.845</div>
 *   <div>Fecha de vencimiento</div><div>05/04/2026</div>
 *   <div>Pago minimo</div><div>$518.208</div>
 */
async function extractBilledStatementInfo(page: Page): Promise<BilledStatementInfo> {
  return page.evaluate((host: string) => {
    const shadowEl = document.querySelector(host);
    const topRoot = (shadowEl as Element & { shadowRoot?: ShadowRoot })?.shadowRoot || document;

    function collectAllRoots(root: ShadowRoot | Element): Array<ShadowRoot | Element> {
      const found: Array<ShadowRoot | Element> = [root];
      for (const el of Array.from((root as Element).querySelectorAll("*"))) {
        const sr = (el as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
        if (sr) found.push(...collectAllRoots(sr));
      }
      return found;
    }

    function findNextSiblingValue(root: ShadowRoot | Element, labelText: string): string | undefined {
      const divs = Array.from((root as Element).querySelectorAll<HTMLElement>("div"));
      for (let i = 0; i < divs.length - 1; i++) {
        const text = divs[i].textContent?.trim() || "";
        if (text.toLowerCase() === labelText.toLowerCase()) {
          const val = divs[i + 1]?.textContent?.trim() || "";
          if (val) return val;
        }
      }
      return undefined;
    }

    function parseAmount(text?: string): number | undefined {
      if (!text) return undefined;
      const m = text.match(/\$([\d.,]+)/);
      if (!m) return undefined;
      return parseInt(m[1].replace(/\./g, "").replace(",", ""), 10) || undefined;
    }

    function extractDate(text?: string): string | undefined {
      if (!text) return undefined;
      const m = text.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
      return m ? m[1] : undefined;
    }

    let billingDate: string | undefined;
    let billedAmount: number | undefined;
    let dueDate: string | undefined;
    let minimumPayment: number | undefined;

    for (const root of collectAllRoots(topRoot)) {
      if (!billingDate) billingDate = extractDate(findNextSiblingValue(root, "Fecha de facturaci\u00f3n"));
      if (!billedAmount) billedAmount = parseAmount(findNextSiblingValue(root, "Monto facturado"));
      if (!dueDate) dueDate = extractDate(findNextSiblingValue(root, "Fecha de vencimiento"));
      if (!minimumPayment) minimumPayment = parseAmount(findNextSiblingValue(root, "Pago minimo"));
    }

    return { billingDate, billedAmount, dueDate, minimumPayment };
  }, SHADOW_HOST);
}

// ─── CMR Shadow DOM helpers ───────────────────────────────────────

async function waitForCmrMovements(page: Page, timeoutMs = CMR_MOVEMENTS_TIMEOUT_MS): Promise<void> {
  try {
    await page.waitForFunction((host: string) => {
      const el = document.querySelector(host);
      if (!el?.shadowRoot) return false;
      // Recursively check all shadow roots — app-invoiced-movements lives several levels deep
      function collectAll(root: ShadowRoot | Element): Array<ShadowRoot | Element> {
        const found: Array<ShadowRoot | Element> = [root];
        for (const child of Array.from((root as Element).querySelectorAll("*"))) {
          const sr = (child as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
          if (sr) found.push(...collectAll(sr));
        }
        return found;
      }
      return collectAll(el.shadowRoot).some(
        (r) => (r as Element).querySelectorAll("table tbody tr td").length > 0,
      );
    }, { timeout: timeoutMs }, SHADOW_HOST);
  } catch { /* timeout */ }
  await delay(500);
}

async function extractCupos(page: Page, debugLog: string[]): Promise<CreditCardBalance | null> {
  try {
    const cupoData = await page.evaluate(() => {
      const text = document.body?.innerText || "";
      const labelMatch = text.match(/(CMR\s+\w+(?:\s+\w+)?)\s*\n?\s*[•·*\s]+\s*(\d{4})/i);
      const label = labelMatch ? `${labelMatch[1]} ****${labelMatch[2]}` : "";
      const cupoMatch = text.match(/\$([\d.,]+)\s*\n?\s*Cupo de compras/i);
      const usadoMatch = text.match(/\$([\d.,]+)\s*\n?\s*Cupo utilizado/i);
      const disponibleMatch = text.match(/\$([\d.,]+)\s*\n?\s*Cupo disponible/i);
      return { label, cupo: cupoMatch?.[1], usado: usadoMatch?.[1], disponible: disponibleMatch?.[1] };
    });
    if (!cupoData.cupo && !cupoData.disponible) return null;
    const total = cupoData.cupo ? parseChileanAmount(cupoData.cupo) : 0;
    const used = cupoData.usado ? parseChileanAmount(cupoData.usado) : 0;
    const available = cupoData.disponible ? parseChileanAmount(cupoData.disponible) : 0;
    debugLog.push(`  CMR cupos: total=$${total}, used=$${used}, available=$${available}`);
    return { label: cupoData.label || "CMR", national: { total, used, available } };
  } catch (err) {
    debugLog.push(`  [CMR] Error extracting cupos: ${err instanceof Error ? err.message : String(err)}`);
    return null;
  }
}

async function clickCmrTab(page: Page, tabText: string, debugLog: string[]): Promise<boolean> {
  // Angular mounts tab-specific components (e.g. app-invoiced-movements) only
  // when the radio INPUT is programmatically activated first (checked + change
  // event) AND the LABEL is then clicked. Doing only the label click is not
  // reliable after an Excel download because Angular's change detection may not
  // re-trigger unless the underlying model value actually changes.
  // Strategy: (1) find and activate the radio input via checked+change+click,
  //           (2) also click the corresponding label for good measure.
  const result = await page.evaluate((text: string, host: string, radioId: string) => {
    const shadowEl = document.querySelector(host);
    const roots: Array<Document | ShadowRoot> = [];
    if (shadowEl?.shadowRoot) roots.push(shadowEl.shadowRoot);
    roots.push(document);

    // Step 1: try the well-known radio id for the facturados tab
    for (const root of roots) {
      const radio = root.querySelector(`#${radioId}`) as HTMLInputElement | null;
      if (radio) {
        radio.checked = true;
        radio.dispatchEvent(new Event("change", { bubbles: true }));
        radio.click();
        // Also click the label for this radio (Angular needs both)
        const label = root.querySelector(`label[for="${radio.id}"]`) as HTMLElement | null
          ?? (radio.closest("label") as HTMLElement | null);
        if (label) label.click();
        return `radio#${radio.id}`;
      }
    }

    // Step 2: fallback — find any radio whose sibling/parent label contains tabText
    for (const root of roots) {
      for (const label of Array.from(root.querySelectorAll<HTMLLabelElement>("label"))) {
        if (!label.innerText?.trim().toLowerCase().includes(text.toLowerCase())) continue;
        // Activate the associated radio first
        const forId = label.getAttribute("for");
        const radio = forId
          ? (root.querySelector(`#${forId}`) as HTMLInputElement | null)
          : (label.querySelector("input[type='radio']") as HTMLInputElement | null);
        if (radio) {
          radio.checked = true;
          radio.dispatchEvent(new Event("change", { bubbles: true }));
          radio.click();
        }
        label.click();
        return `label: "${label.innerText.trim()}"`;
      }
    }
    return null;
  }, tabText, SHADOW_HOST, INVOICED_RADIO_ID);

  if (result) debugLog.push(`  CMR: Clicked tab "${tabText}" via ${result}`);
  return result !== null;
}

// ─── Excel download via CDP ───────────────────────────────────────

/**
 * Clicks the "Exportar a Excel" button inside the CMR shadow DOM and waits
 * for the downloaded .xlsx file to appear.
 *
 * Monitors both the CDP-redirected temp dir AND ~/Downloads as fallback,
 * since Page.setDownloadBehavior is deprecated in Chrome 131+ and may
 * not redirect correctly in all environments.
 *
 * Returns the file path (always inside downloadDir) on success, null on failure.
 */
async function downloadCmrExcel(page: Page, downloadDir: string, debugLog: string[]): Promise<string | null> {
  try {
    await setupPageDownload(page, downloadDir);

    const beforeTemp = snapshotDir(downloadDir);
    const beforeDownloads = snapshotDir(DOWNLOADS_DIR);

    const clicked = await page.evaluate((host: string) => {
      const shadowEl = document.querySelector(host);
      const topRoot = (shadowEl as Element & { shadowRoot?: ShadowRoot })?.shadowRoot || document;

      // Recursively collect ALL shadow roots at any depth (app-invoiced-movements
      // lives behind multiple shadow boundaries after label click).
      function collectAllRoots(root: ShadowRoot | Element): Array<ShadowRoot | Element> {
        const found: Array<ShadowRoot | Element> = [root];
        for (const el of Array.from((root as Element).querySelectorAll("*"))) {
          const sr = (el as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
          if (sr) found.push(...collectAllRoots(sr));
        }
        return found;
      }
      // Deepest roots first so app-invoiced-movements button is preferred over
      // the shared app-last-movements button (which always exports unbilled data).
      const nested = collectAllRoots(topRoot).filter((r) => r !== topRoot);
      const roots: Array<ShadowRoot | Document | Element> = [...nested, topRoot];

      function isVisible(el: HTMLElement): boolean {
        // getBoundingClientRect is reliable for all positions (including fixed/sticky);
        // offsetParent returns null for position:fixed elements even when visible.
        const r = el.getBoundingClientRect();
        return r.width > 0 || r.height > 0;
      }

      function findExcelBtn(root: ShadowRoot | Document | Element): HTMLButtonElement | null {
        const buttons = Array.from((root as Element).querySelectorAll<HTMLButtonElement>("button.btn-doc-export"));
        return buttons.find(
          (btn) =>
            isVisible(btn) &&
            (btn.querySelector('img[alt*="Excel"]') ||
             btn.querySelector(".button-label")?.textContent?.toLowerCase().includes("excel")),
        ) ?? null;
      }

      for (const root of roots) {
        const btn = findExcelBtn(root);
        if (btn) { btn.click(); return true; }
      }
      return false;
    }, SHADOW_HOST);

    if (!clicked) {
      debugLog.push("  [Excel] Botón 'Exportar a Excel' no encontrado");
      return null;
    }

    const filePath = await waitForDownloadedFile(
      downloadDir,
      beforeTemp,
      ".xlsx",
      30000,
      [{ dir: DOWNLOADS_DIR, before: beforeDownloads }],
    );

    if (filePath) {
      debugLog.push(`  [Excel] Descargado: ${path.basename(filePath)}`);
    } else {
      debugLog.push("  [Excel] Timeout esperando descarga");
    }
    return filePath;
  } catch (err) {
    debugLog.push(`  [Excel] Error: ${err instanceof Error ? err.message : String(err)}`);
    return null;
  }
}

// ─── DOM extraction (fallback) ────────────────────────────────────

async function extractCmrMovementsFromTable(page: Page): Promise<BankMovement[]> {
  return page.evaluate((host: string) => {
    const movements: BankMovement[] = [];
    const shadowEl = document.querySelector(host);
    const topRoot = shadowEl?.shadowRoot || document;

    function collectAll(r: ShadowRoot | Element | Document): Array<ShadowRoot | Element> {
      const found: Array<ShadowRoot | Element> = r instanceof Document ? [] : [r];
      for (const el of Array.from((r as ParentNode).querySelectorAll("*"))) {
        const sr = (el as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
        if (sr) found.push(...collectAll(sr));
      }
      return found;
    }
    const roots = collectAll(topRoot);

    for (const root of roots) {
      for (const table of Array.from((root as Element).querySelectorAll("table"))) {
        // Skip tables that are not rendered (hidden inactive tabs)
        const rect = table.getBoundingClientRect();
        if (rect.width === 0 && rect.height === 0) continue;
        for (const row of Array.from(table.querySelectorAll("tbody tr"))) {
          const cells = row.querySelectorAll("td");
          if (cells.length < 4) continue;
          const texts = Array.from(cells).map((c) => (c as HTMLElement).innerText?.trim() || "");
          const dateMatch = texts[0]?.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
          const pendingImg = row.querySelector("td:first-child img[alt*='pendiente'], td:first-child .td-time-img");
          if (!dateMatch && !pendingImg && texts[0] !== "") continue;
          const date = dateMatch ? dateMatch[1].replace(/\//g, "-") : "pendiente";
          const description = texts[1] || "";
          // texts[3]=Monto total, texts[5]=Cuota a pagar — prefer Cuota a pagar (installment amount)
          const totalText = texts[3] || "";
          const cuotaText = texts[5] || "";
          const montoText = cuotaText || totalText;
          const isNeg = montoText.includes("-$");
          const amountMatch = montoText.match(/\$\s*([\d.,]+)/);
          let amount = 0;
          if (amountMatch) {
            const value = parseInt(amountMatch[1].replace(/\./g, "").replace(",", "."), 10) || 0;
            amount = isNeg ? value : -value;
          }
          const totalAmountMatch = totalText.match(/\$\s*([\d.,]+)/);
          const totalAmount = totalAmountMatch
            ? parseInt(totalAmountMatch[1].replace(/\./g, "").replace(",", "."), 10) || undefined
            : undefined;
          if (description && amount !== 0)
            movements.push({ date, description, amount, balance: 0, source: "credit_card_unbilled" as MovementSource, owner: (texts[2] || undefined) as CardOwner | undefined, installments: texts[4] || undefined, totalAmount });
        }
      }
    }
    return movements;
  }, SHADOW_HOST);
}

/** Returns the text of the first data row across all shadow roots — used to detect page changes. */
async function getTablePageSignature(page: Page): Promise<string> {
  return page.evaluate((host: string) => {
    const el = document.querySelector(host);
    const topRoot = el?.shadowRoot || document;
    function collectAll(r: ShadowRoot | Element | Document): Array<ShadowRoot | Element> {
      const found: Array<ShadowRoot | Element> = r instanceof Document ? [] : [r];
      for (const child of Array.from((r as ParentNode).querySelectorAll("*"))) {
        const sr = (child as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
        if (sr) found.push(...collectAll(sr));
      }
      return found;
    }
    for (const root of collectAll(topRoot)) {
      const cells = (root as Element).querySelectorAll("table tbody tr:first-child td");
      if (cells.length > 0)
        return Array.from(cells).map((c) => (c as HTMLElement).innerText?.trim()).join("|");
    }
    return "";
  }, SHADOW_HOST);
}

async function paginateCmrMovements(page: Page, debugLog: string[]): Promise<BankMovement[]> {
  const all: BankMovement[] = [];
  for (let i = 0; i < MAX_PAGES; i++) {
    const pageRows = await extractCmrMovementsFromTable(page);
    all.push(...pageRows);

    const sigBefore = await getTablePageSignature(page);

    const hasNext = await page.evaluate((host: string) => {
      const shadowEl = document.querySelector(host);
      if (!shadowEl) return false;
      const root = shadowEl.shadowRoot || document;

      function collectAll(r: ShadowRoot | Element | Document): Array<ShadowRoot | Element> {
        const found: Array<ShadowRoot | Element> = r instanceof Document ? [] : [r];
        for (const el of Array.from((r as ParentNode).querySelectorAll("*"))) {
          const sr = (el as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
          if (sr) found.push(...collectAll(sr));
        }
        return found;
      }
      const roots = collectAll(root);

      for (const r of roots) {
        const candidates = Array.from(
          (r as Element).querySelectorAll<HTMLButtonElement>(
            ".btn-pagination, [class*='pagination'] button, button[aria-label], button",
          ),
        );
        for (const btn of candidates) {
          if (btn.disabled) continue;
          const img = btn.querySelector("img");
          const label = (btn.getAttribute("aria-label") || btn.innerText || "").toLowerCase();
          const imgAlt = (img?.getAttribute("alt") || "").toLowerCase();
          const imgSrc = img?.getAttribute("src") || "";
          const isNext =
            label.includes("siguiente") ||
            label.includes("next") ||
            label.includes("avanzar") ||
            imgAlt.includes("avanzar") ||
            imgAlt.includes("siguiente") ||
            imgAlt.includes("next") ||
            imgSrc.includes("right-arrow") ||
            imgSrc.includes("arrow-right") ||
            imgSrc.includes("next");
          if (isNext) { btn.click(); return true; }
        }
      }
      return false;
    }, SHADOW_HOST);

    if (!hasNext) break;

    // Wait for first-row content to change (not just for rows to exist)
    const changed = await page.waitForFunction((host: string, prevSig: string) => {
      const el = document.querySelector(host);
      const topRoot = el?.shadowRoot || document;
      function collectAll(r: ShadowRoot | Element | Document): Array<ShadowRoot | Element> {
        const found: Array<ShadowRoot | Element> = r instanceof Document ? [] : [r];
        for (const child of Array.from((r as ParentNode).querySelectorAll("*"))) {
          const sr = (child as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
          if (sr) found.push(...collectAll(sr));
        }
        return found;
      }
      for (const root of collectAll(topRoot)) {
        const cells = (root as Element).querySelectorAll("table tbody tr:first-child td");
        if (cells.length > 0) {
          const sig = Array.from(cells).map((c) => (c as HTMLElement).innerText?.trim()).join("|");
          return sig !== prevSig && sig !== "";
        }
      }
      return false;
    }, { timeout: PAGE_CHANGE_TIMEOUT_MS }, SHADOW_HOST, sigBefore).then(() => true, () => false);
    if (!changed) break;

    if (i === 19) debugLog.push(`  [warn] paginateCmrMovements: hit 20-page cap — may be truncated`);
  }
  return deduplicateMovements(
    all.map((m) => ({
      ...m,
      date: normalizeDate(m.date),
      owner: normalizeOwner(m.owner),
      installments: normalizeInstallments(m.installments),
    })),
  );
}

/**
 * Paginates through app-invoiced-movements (the facturados component).
 *
 * The paginator uses two arrow buttons (left/right) with class "btn-pagination".
 * The right-arrow has img alt="boton avanzar" and is disabled on the last page.
 * There are NO numbered page buttons — only a single <span class="m-2">N</span>.
 *
 * Change detection anchors on the first row of the "fecha de compra" table because
 * the "pendientes de confirmación" sub-table rows never change across pages.
 *
 * All extract + click logic is batched into a single page.evaluate per iteration
 * to avoid redundant shadow-DOM traversals.
 */
async function paginateInvoicedMovements(page: Page, debugLog: string[]): Promise<BankMovement[]> {
  const all: BankMovement[] = [];
  for (let i = 0; i < MAX_PAGES; i++) {
    type PageResult = {
      rows: BankMovement[];
      pageIndicator: string;
      firstDatedRow: string;
      clicked: boolean;
    };
    const result: PageResult = await page.evaluate((host: string): PageResult => {
      const shadowEl = document.querySelector(host);
      const topRoot = shadowEl?.shadowRoot;
      if (!topRoot) return { rows: [], pageIndicator: "", firstDatedRow: "", clicked: false };

      // app-invoiced-movements lives behind multiple shadow boundaries — must traverse recursively
      function collectAllRoots(root: ShadowRoot | Element): Array<ShadowRoot | Element> {
        const found: Array<ShadowRoot | Element> = [root];
        for (const el of Array.from((root as Element).querySelectorAll("*"))) {
          const sr = (el as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
          if (sr) found.push(...collectAllRoots(sr));
        }
        return found;
      }
      const roots = collectAllRoots(topRoot);

      // ── 1. Extract billed movements ────────────────────────────────────────
      const allTables: HTMLTableElement[] = roots.flatMap(
        (root) => Array.from((root as Element).querySelectorAll<HTMLTableElement>("table")),
      );
      // Include "fecha de compra", "monto total", and "cuota a pagar" tables
      // ("pendientes de confirmación" uses "cuota a pagar" header — include it too)
      function isVisible(t: HTMLTableElement): boolean {
        const r = t.getBoundingClientRect();
        return r.width > 0 || r.height > 0;
      }
      const matchingTables = allTables.filter((t) => {
        if (!isVisible(t)) return false;
        const hdr = (t.querySelector("thead, tr:first-child") as HTMLElement | null)?.innerText?.toLowerCase() ?? "";
        return hdr.includes("fecha de compra") || hdr.includes("monto total") || hdr.includes("cuota a pagar");
      });
      const tablesToExtract = matchingTables.length > 0
        ? matchingTables
        : allTables.filter((t) => isVisible(t) && !t.closest("app-last-movements"));

      const rows: BankMovement[] = [];
      for (const table of tablesToExtract) {
        for (const row of Array.from(table.querySelectorAll("tbody tr"))) {
          const cells = row.querySelectorAll("td");
          if (cells.length < 4) continue;
          const texts = Array.from(cells).map((c) => (c as HTMLElement).innerText?.trim() || "");
          const dateMatch = texts[0]?.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
          if (!dateMatch && texts[0] !== "") continue; // skip separator rows
          const date = dateMatch ? dateMatch[1].replace(/\//g, "-") : "pendiente";
          const description = texts[1] || "";
          // "Cuota a pagar" (col 5) takes precedence over "Monto total" (col 3) as amount
          const totalText = texts[3] || "";
          const montoText = (texts[5] && texts[5] !== texts[4]) ? texts[5] : totalText;
          const isNeg = montoText.includes("-$");
          const amountMatch = montoText.match(/\$\s*([\d.,]+)/);
          let amount = 0;
          if (amountMatch) {
            const value = parseInt(amountMatch[1].replace(/\./g, "").replace(",", "."), 10) || 0;
            amount = isNeg ? value : -value;
          }
          const totalAmountMatch = totalText.match(/\$\s*([\d.,]+)/);
          const totalAmount = totalAmountMatch
            ? parseInt(totalAmountMatch[1].replace(/\./g, "").replace(",", "."), 10) || undefined
            : undefined;
          if (description && amount !== 0) {
            rows.push({
              date, description, amount, balance: 0,
              source: "credit_card_billed" as MovementSource,
              owner: (texts[2] || undefined) as CardOwner | undefined,
              installments: texts[4] || undefined,
              totalAmount,
            });
          }
        }
      }

      // ── 2. Page indicator (paginator span.m-2) ─────────────────────────────
      function findPageIndicator(): string {
        for (const root of roots) {
          for (const span of Array.from((root as Element).querySelectorAll<HTMLElement>("span.m-2, span[class*='m-']"))) {
            const t = span.innerText?.trim() || "";
            if (/^\d+$/.test(t)) return t;
          }
        }
        for (const root of roots) {
          for (const span of Array.from((root as Element).querySelectorAll<HTMLElement>("span"))) {
            if (span.querySelector("*")) continue;
            const t = span.innerText?.trim() || "";
            if (/^\d+$/.test(t) && Number(t) <= 100) return t;
          }
        }
        return "";
      }
      const pageIndicator = findPageIndicator();

      // ── 3. First dated row (for change detection after click) ──────────────
      function findFirstDatedRow(): string {
        // Prefer "fecha de compra" table — "pendientes de confirmación" rows don't change across pages
        for (const r of roots) {
          for (const tbl of Array.from((r as Element).querySelectorAll<HTMLTableElement>("table"))) {
            const hdr = (tbl.querySelector("thead, tr:first-child") as HTMLElement | null)?.innerText?.toLowerCase() ?? "";
            if (!hdr.includes("fecha de compra")) continue;
            const cells = tbl.querySelectorAll("tbody tr:first-child td");
            if (cells.length > 0)
              return Array.from(cells).map((c) => (c as HTMLElement).innerText?.trim()).join("|");
          }
        }
        // Fallback: first visible table with data rows
        for (const r of roots) {
          for (const tbl of Array.from((r as Element).querySelectorAll<HTMLTableElement>("table"))) {
            if (!isVisible(tbl)) continue;
            const cells = tbl.querySelectorAll("tbody tr:first-child td");
            if (cells.length > 0)
              return Array.from(cells).map((c) => (c as HTMLElement).innerText?.trim()).join("|");
          }
        }
        return "";
      }
      const firstDatedRow = findFirstDatedRow();

      // ── 4. Click the "avanzar" (right-arrow) button ────────────────────────
      let clicked = false;
      for (const root of roots) {
        if (clicked) break;
        for (const btn of Array.from((root as Element).querySelectorAll<HTMLButtonElement>(".btn-pagination, button"))) {
          if (btn.disabled || btn.getAttribute("disabled") !== null) continue;
          const img = btn.querySelector("img");
          const imgAlt = (img?.getAttribute("alt") || "").toLowerCase();
          const imgSrc = img?.getAttribute("src") || "";
          const label = (btn.getAttribute("aria-label") || btn.innerText || "").toLowerCase();
          const isNext =
            imgAlt.includes("avanzar") || imgAlt.includes("siguiente") || imgAlt.includes("next") ||
            imgSrc.includes("right-arrow") || imgSrc.includes("arrow-right") || imgSrc.includes("next") ||
            label.includes("siguiente") || label.includes("next") || label.includes("avanzar");
          if (isNext) { btn.click(); clicked = true; break; }
        }
      }

      return { rows, pageIndicator, firstDatedRow, clicked };
    }, SHADOW_HOST);

    debugLog.push(`  [inv pag] page ${result.pageIndicator || i + 1}: ${result.rows.length} rows (total: ${all.length + result.rows.length})`);
    all.push(...result.rows);

    if (!result.clicked) {
      debugLog.push(`  [inv pag] no more pages (right-arrow disabled or absent)`);
      break;
    }

    const changed = await page.waitForFunction(
      (host: string, prevRow: string, prevIndicator: string) => {
        const shadowEl = document.querySelector(host);
        const topRoot = shadowEl?.shadowRoot;
        if (!topRoot) return false;
        function collectAll(root: ShadowRoot | Element): Array<ShadowRoot | Element> {
          const found: Array<ShadowRoot | Element> = [root];
          for (const el of Array.from((root as Element).querySelectorAll("*"))) {
            const sr = (el as Element & { shadowRoot?: ShadowRoot }).shadowRoot;
            if (sr) found.push(...collectAll(sr));
          }
          return found;
        }
        const allRoots = collectAll(topRoot);
        // Signal 1: page indicator changed (most reliable — avoids row-content false negatives)
        if (prevIndicator) {
          for (const r of allRoots) {
            for (const span of Array.from((r as Element).querySelectorAll<HTMLElement>("span.m-2, span[class*='m-']"))) {
              const t = span.innerText?.trim() || "";
              if (/^\d+$/.test(t) && t !== prevIndicator) return true;
            }
          }
        }
        function rowSig(tbl: HTMLTableElement): string {
          const cells = tbl.querySelectorAll("tbody tr:first-child td");
          return cells.length > 0 ? Array.from(cells).map((c) => (c as HTMLElement).innerText?.trim()).join("|") : "";
        }
        // Signal 2: "fecha de compra" table first row changed
        for (const r of allRoots) {
          for (const tbl of Array.from((r as Element).querySelectorAll<HTMLTableElement>("table"))) {
            const hdr = (tbl.querySelector("thead, tr:first-child") as HTMLElement | null)?.innerText?.toLowerCase() ?? "";
            if (!hdr.includes("fecha de compra")) continue;
            const sig = rowSig(tbl);
            if (sig) return sig !== prevRow;
          }
        }
        // Signal 3: any visible table first row changed
        for (const r of allRoots) {
          for (const tbl of Array.from((r as Element).querySelectorAll<HTMLTableElement>("table"))) {
            const rect = tbl.getBoundingClientRect();
            if (rect.width === 0 && rect.height === 0) continue;
            const sig = rowSig(tbl);
            if (sig) return sig !== prevRow;
          }
        }
        return false;
      },
      { timeout: PAGE_CHANGE_TIMEOUT_MS },
      SHADOW_HOST,
      result.firstDatedRow,
      result.pageIndicator,
    ).then(() => true, () => false);
    if (!changed) { debugLog.push(`  [inv pag] table content unchanged after next click, stopping`); break; }
    await delay(300);

    if (i === 19) debugLog.push(`  [warn] paginateInvoicedMovements: hit 20-page cap — may be truncated`);
  }

  return deduplicateMovements(
    all.map((m) => ({
      ...m,
      date: normalizeDate(m.date),
      owner: normalizeOwner(m.owner),
      installments: normalizeInstallments(m.installments),
    })),
  );
}

// ─── Navigation helpers ───────────────────────────────────────────

async function tryExpandDateRange(page: Page, debugLog: string[]): Promise<void> {
  try {
    const selectInfo = await page.evaluate(() => {
      const selects = Array.from(document.querySelectorAll("select"));
      return selects.map((sel, i) => ({
        index: i, name: sel.name || sel.id || `select-${i}`,
        options: Array.from(sel.querySelectorAll("option")).map((o) => ({ text: o.text.trim(), value: o.value })),
      }));
    });
    for (const sel of selectInfo) {
      for (const opt of sel.options) {
        const text = opt.text.toLowerCase();
        if (text.includes("todos") || text.includes("último mes") || text.includes("30 día") || text.includes("mes anterior")) {
          await page.evaluate((selIdx: number, optValue: string) => {
            const selects = document.querySelectorAll("select");
            const select = selects[selIdx] as HTMLSelectElement;
            if (select) { select.value = optValue; select.dispatchEvent(new Event("change", { bubbles: true })); }
          }, sel.index, opt.value);
          debugLog.push(`  Changed [${sel.name}] to "${opt.text}"`);
          await delay(3000);
          break;
        }
      }
    }
    const clickedSearch = await page.evaluate(() => {
      const buttons = Array.from(document.querySelectorAll("button, input[type='submit']"));
      for (const btn of buttons) {
        const text = (btn as HTMLElement).innerText?.trim().toLowerCase();
        if (text === "buscar" || text === "consultar" || text === "filtrar") { (btn as HTMLElement).click(); return text; }
      }
      return null;
    });
    if (clickedSearch) { debugLog.push(`  Clicked "${clickedSearch}"`); await delay(3000); }
  } catch (err) {
    debugLog.push(`  [dateRange] Error: ${err instanceof Error ? err.message : String(err)}`);
  }
}

async function clickNavTarget(page: Page, debugLog: string[]): Promise<boolean> {
  const targets = [
    { text: "cartola", exact: false },
    { text: "últimos movimientos", exact: false },
    { text: "movimientos", exact: true },
    { text: "estado de cuenta", exact: false },
  ];
  for (const target of targets) {
    const result = await page.evaluate((t: { text: string; exact: boolean }) => {
      const elements = Array.from(document.querySelectorAll("a, button, [role='tab'], [role='menuitem'], li, span"));
      for (const el of elements) {
        const text = (el as HTMLElement).innerText?.trim().toLowerCase() || "";
        const href = (el as HTMLAnchorElement).href || "";
        if (href.includes("cc-nuevos") || href.includes("comenzar")) continue;
        if (text.includes("historial de transferencia")) continue;
        const match = t.exact ? text === t.text : text.includes(t.text);
        if (match && text.length < 50) { (el as HTMLElement).click(); return `Clicked: "${text}"`; }
      }
      return null;
    }, target);
    if (result) { debugLog.push(`  ${result}`); await delay(4000); return true; }
  }
  return false;
}

// ─── Main scrape function ─────────────────────────────────────────

async function scrapeFalabella(session: BrowserSession, options: ScraperOptions): Promise<ScrapeResult> {
  const { rut, password, saveScreenshots: doScreenshots, owner = "B" } = options;
  const { onProgress } = options;
  const progress = onProgress || (() => {});
  const { page, debugLog, screenshot: doSave } = session;
  const bank = "falabella";

  const downloadDir = createTempDownloadDir(bank);

  try {
    // 1. Navigate
    debugLog.push("1. Navigating to bank homepage...");
    progress("Abriendo sitio del banco...");
    await page.goto(BANK_URL, { waitUntil: "networkidle2", timeout: 30000 });
    await delay(2000);
    await dismissBanners(page);
    // Extra wait for Falabella's Angular SPA to fully initialize before login
    await delay(1000);
    await doSave(page, "01-homepage");

    // 2-6. Login.
    //
    // PATCH (miaufinanzas, 2026 redesign): Banco Falabella's homepage no longer
    // navigates to a /login page. Clicking "Mi cuenta" (button #btn-auth-normal)
    // opens an INLINE auth widget on the same homepage. The old flow
    // (clickByText "Mi cuenta" → expect navigation → generic fillRut/fillPassword/
    // clickSubmit) silently failed: it stayed on the public homepage, typed the
    // RUT into the SEARCH box, never submitted (the real submit button
    // #desktop-login has no visible text and no type="submit" attribute, so the
    // generic clickSubmit missed it and fell back to Enter, which does not submit
    // this React form), then declared "Login OK" because no error element was
    // present — returning 0 movements. The structure below was verified against
    // the live site:
    //   - RUT field:    input[placeholder*="RUT"] (visible), maxlength=10. Typing
    //                   the CLEAN rut (digits + DV) lets the bank auto-insert the
    //                   dash, e.g. "111111111" -> "11111111-1".
    //   - Clave field:  input[type=password] placeholder "Clave Internet", max 6.
    //   - Submit:       #desktop-login, `disabled` until both fields validate.
    debugLog.push("2. Opening 'Mi cuenta' inline auth form...");
    progress("Abriendo formulario de ingreso...");

    // Dismiss the welcome modal ("Entendido", #btn-login-client-nuevo) and let
    // dismissBanners clear any cookie/promo banner that could intercept clicks.
    await page.evaluate(() => {
      const el = document.getElementById("btn-login-client-nuevo");
      if (el) (el as HTMLElement).click();
    }).catch(() => {});
    await dismissBanners(page).catch(() => {});
    await delay(500);

    // Open the auth widget via a DOM click (bypasses any visual overlay).
    await page.evaluate(() => {
      const byId = document.getElementById("btn-auth-normal");
      if (byId) { (byId as HTMLElement).click(); return; }
      for (const b of Array.from(document.querySelectorAll("button, a"))) {
        if (((b as HTMLElement).innerText || "").trim().toLowerCase() === "mi cuenta") {
          (b as HTMLElement).click();
          return;
        }
      }
    }).catch(() => {});

    // Wait for the inline auth drawer to render.
    //
    // PATCH (miaufinanzas, ago 2026 redesign): we used to wait for / match "an
    // input whose placeholder mentions RUT". The drawer's RUT input LOST that
    // word — it is now `input#document` with placeholder "Ej: 12345678-9" — so
    // the match started landing on the LOAN CALCULATOR input further down the
    // public homepage ("Ingresa tu RUT", inside .loanCalculator_wrapper). That
    // element lives outside the drawer and gets re-rendered, so the captured
    // handle went stale and `handle.click()` threw "Node is either not clickable
    // or not an Element" — the login never even started. We now scope every
    // lookup to the login drawer's own <form> and keep the placeholder match
    // only as a last resort (explicitly excluding the loan calculator).
    await page.waitForFunction(() => {
      const doc = document.getElementById("document") as HTMLInputElement | null;
      if (doc && doc.offsetParent !== null) return true;
      return Array.from(document.querySelectorAll('input[type="password"]'))
        .some((p) => (p as HTMLInputElement).offsetParent !== null);
    }, { timeout: 15000 }).catch(() => {});
    await delay(800);
    await doSave(page, "02-login-form");

    // 3. Mark the drawer's RUT + clave inputs and its submit button so every
    //    later step can grab the exact elements by data attribute.
    debugLog.push("3. Locating login drawer fields...");
    progress("Ingresando RUT...");
    const cleanRut = String(rut).replace(/[.\-]/g, "").toUpperCase();
    const markLoginFields = () => page.evaluate(() => {
      const vis = (el: Element | null) => !!el && (el as HTMLElement).offsetParent !== null;
      for (const a of ["data-obc-rut", "data-obc-pwd", "data-obc-submit"]) {
        document.querySelectorAll("[" + a + "]").forEach((e) => e.removeAttribute(a));
      }

      // The login drawer is the form that owns a password field.
      const forms = Array.from(document.querySelectorAll("form")) as HTMLFormElement[];
      const drawer =
        forms.find((f) => /DrawerFormLogin|FormLogin/i.test(f.className || "") && !!f.querySelector('input[type="password"]')) ||
        forms.find((f) => !!f.querySelector('input[type="password"]:not([type="hidden"])') && vis(f)) ||
        null;
      const scope: ParentNode = drawer || document;

      const pwd = (Array.from(scope.querySelectorAll('input[type="password"]')) as HTMLInputElement[])
        .find((p) => vis(p)) || null;

      // RUT field: #document today; else the drawer's non-password input; else
      // the text input sharing a form with the password; else placeholder match.
      let rutEl = (scope.querySelector("input#document") as HTMLInputElement | null);
      if (!rutEl || !vis(rutEl)) rutEl = null;
      if (!rutEl && drawer) {
        rutEl = (Array.from(drawer.querySelectorAll("input")) as HTMLInputElement[])
          .find((i) => vis(i) && i.type !== "password" && i.type !== "hidden") || null;
      }
      if (!rutEl && pwd && pwd.form) {
        rutEl = (Array.from(pwd.form.querySelectorAll("input")) as HTMLInputElement[])
          .find((i) => vis(i) && i !== pwd && i.type !== "password" && i.type !== "hidden") || null;
      }
      if (!rutEl) {
        rutEl = (Array.from(document.querySelectorAll("input")) as HTMLInputElement[])
          .find((i) => vis(i) && /rut/i.test(i.placeholder || "") && !i.closest('[class*="loanCalculator" i]')) || null;
      }

      // Submit: the old #desktop-login is gone; it is now the drawer's
      // button[type=submit] ("Ingresar"), disabled until both fields validate.
      let submit = (document.getElementById("desktop-login") ||
        document.querySelector('[data-testid="desktop-login"]')) as HTMLButtonElement | null;
      if (!submit && drawer) {
        submit = drawer.querySelector('button[type="submit"]') as HTMLButtonElement | null;
        if (!submit) {
          submit = (Array.from(drawer.querySelectorAll("button")) as HTMLButtonElement[])
            .find((b) => /ingresar/i.test((b.innerText || "").trim())) || null;
        }
      }

      if (rutEl) rutEl.setAttribute("data-obc-rut", "1");
      if (pwd) pwd.setAttribute("data-obc-pwd", "1");
      if (submit) submit.setAttribute("data-obc-submit", "1");
      return {
        rut: !!rutEl, pwd: !!pwd, submit: !!submit,
        scopedToDrawer: !!drawer,
        rutId: rutEl ? (rutEl.id || rutEl.placeholder || "?") : "",
      };
    });

    let marks = await markLoginFields();
    if (!marks.pwd) {
      // Some variants reveal the clave step only after Enter on the RUT field.
      await page.keyboard.press("Enter").catch(() => {});
      await page.waitForFunction(() => {
        return Array.from(document.querySelectorAll('input[type="password"]'))
          .some((p) => (p as HTMLInputElement).offsetParent !== null);
      }, { timeout: 15000, polling: 500 }).catch(() => {});
      marks = await markLoginFields();
    }
    debugLog.push(`  Fields: rut=${marks.rut} (${marks.rutId}) pwd=${marks.pwd} submit=${marks.submit} drawer=${marks.scopedToDrawer}`);
    if (!marks.rut) {
      const ss = await page.screenshot({ encoding: "base64" });
      return { success: false, bank, accounts: [], error: "No se encontró el campo de RUT (el formulario de ingreso no se abrió)", screenshot: ss as string, debug: debugLog.join("\n") };
    }
    if (!marks.pwd) {
      const ss = await page.screenshot({ encoding: "base64" });
      return { success: false, bank, accounts: [], error: "No se encontró el campo de clave", screenshot: ss as string, debug: debugLog.join("\n") };
    }
    let pwdHandle = await page.$('input[data-obc-pwd="1"]');

    // Robust fill with READ-BACK verification. Falabella's RUT field auto-formats
    // (inserts the dash) and these React inputs occasionally drop/reorder typed
    // characters, producing a valid-LOOKING but WRONG rut/clave — the bank then
    // returns USER_OR_PASSWORD_NOT_VALID, or the submit button never enables.
    // This was the root cause of the intermittent login failures. So we type,
    // read the value back, and re-type until BOTH fields hold exactly what we
    // intended AND the submit button has enabled.
    const expectedPwd = String(password);
    const fillAndRead = async (handle: any, attr: string, value: string): Promise<string> => {
      // Focus by DOM first: a coordinate click needs a box model and blows up
      // ("Node is either not clickable or not an Element") whenever React has
      // just re-rendered the input.
      const selected = await page.evaluate((a: string) => {
        const el = document.querySelector('input[' + a + '="1"]') as HTMLInputElement | null;
        if (!el) return false;
        el.focus();
        el.select?.();
        return true;
      }, attr);
      if (!selected) return "";
      await handle.click({ clickCount: 3 }).catch(() => {});
      await page.keyboard.press("Backspace").catch(() => {});
      await delay(150);
      await handle.type(value, { delay: 80 });
      await delay(350);
      return await page.evaluate((a: string) => {
        const el = document.querySelector('input[' + a + '="1"]') as HTMLInputElement | null;
        return el ? el.value : "";
      }, attr);
    };
    const submitEnabled = () => page.evaluate(() => {
      const b = document.querySelector('[data-obc-submit="1"]') as HTMLButtonElement | null;
      return !!b && !b.disabled;
    });

    debugLog.push("4. Filling RUT + clave (with read-back verification)...");
    let filledOk = false;
    for (let attempt = 1; attempt <= 4 && !filledOk; attempt++) {
      // Re-mark and re-acquire handles on every attempt: these React inputs are
      // re-rendered as you type, which detaches previously captured handles.
      if (attempt > 1) marks = await markLoginFields();
      const rutHandle = await page.$('input[data-obc-rut="1"]');
      pwdHandle = await page.$('input[data-obc-pwd="1"]');
      if (!rutHandle || !pwdHandle) { await delay(500); continue; }

      const rutVal = await fillAndRead(rutHandle, "data-obc-rut", cleanRut);
      const pwdVal = await fillAndRead(pwdHandle, "data-obc-pwd", expectedPwd);
      const rutMatches = rutVal.replace(/[.\-\s]/g, "").toUpperCase() === cleanRut;
      const pwdMatches = pwdVal === expectedPwd;
      await page.waitForFunction(() => {
        const b = document.querySelector('[data-obc-submit="1"]') as HTMLButtonElement | null;
        return !!b && !b.disabled;
      }, { timeout: 4000 }).catch(() => {});
      const btnEnabled = await submitEnabled();
      debugLog.push(`  Fill attempt ${attempt}: rutMatch=${rutMatches} pwdLen=${pwdVal.length} pwdMatch=${pwdMatches} btnEnabled=${btnEnabled}`);
      filledOk = rutMatches && pwdMatches && btnEnabled;
    }

    // 5. Submit by clicking the drawer's real login button.
    debugLog.push("5. Submitting login...");
    progress("Iniciando sesión...");
    const submitState = await page.evaluate(() => {
      const b = document.querySelector('[data-obc-submit="1"]') as HTMLButtonElement | null;
      if (!b) return "not-found";
      if (b.disabled) return "disabled";
      b.click();
      return "clicked";
    });
    debugLog.push(`  Submit button: ${submitState} (fields verified: ${filledOk})`);
    if (submitState !== "clicked" && pwdHandle) {
      await pwdHandle.press("Enter");
    }

    // Wait for the post-login transition (SPA navigation or DOM swap).
    await page.waitForNavigation({ waitUntil: "networkidle2", timeout: 25000 }).catch(() => {});
    await delay(5000);
    await doSave(page, "03-after-login");

    await closePopups(page).catch(() => {});
    const dashboardUrl = page.url();

    // 6. POSITIVE login verification — checked BEFORE any 2FA/error heuristic.
    //    The old code declared "Login OK" merely from the ABSENCE of an error
    //    element, so when the scraper stayed on the public homepage it still
    //    reported success and returned 0 movements silently. We now require
    //    evidence of an authenticated session. Doing this FIRST also avoids a
    //    false 2FA trigger: the logged-in portal's menus contain words like
    //    "clave"/"código" that a naive page-text match would mistake for a 2FA
    //    wall (this actually wasted attempt #1 in testing).
    const loginEvidence = await page.evaluate(() => {
      const bodyText = (document.body?.innerText || "").toLowerCase();
      const hasLogout = /cerrar sesi[oó]n/.test(bodyText)
        || !!document.querySelector('[id*="logout" i],[class*="logout" i],[href*="logout" i],[href*="cerrarSesion" i]');
      const hasPrivateNav = /mis productos|mis cuentas|saldo disponible|[uú]ltimos movimientos|cupo (utilizado|disponible)|estado de cuenta|transferir a terceros/.test(bodyText);
      const looksPublic = /atr[eé]vete a ser genial|abre tu cuenta corriente|cr[eé]dito de consumo que|hazte cliente|aprovecha beneficios/.test(bodyText);
      return { url: location.href, hasLogout, hasPrivateNav, looksPublic };
    });
    // The authenticated portal lives on web/web2.bancofalabella.cl/web-clientes;
    // the public marketing site is www.bancofalabella.cl.
    const onPrivatePortal = /\/web-clientes\b/i.test(loginEvidence.url) || /\/\/web2?\.bancofalabella\.cl/i.test(loginEvidence.url);
    const onPublicHome = /^https?:\/\/(www\.)?bancofalabella\.cl\/?$/i.test(loginEvidence.url);
    const loggedIn = onPrivatePortal
      || ((loginEvidence.hasLogout || loginEvidence.hasPrivateNav) && !(onPublicHome && loginEvidence.looksPublic && !loginEvidence.hasPrivateNav));
    debugLog.push(`6. Login check: url=${loginEvidence.url} portal=${onPrivatePortal} logout=${loginEvidence.hasLogout} privateNav=${loginEvidence.hasPrivateNav} public=${loginEvidence.looksPublic}`);

    if (!loggedIn) {
      // Not authenticated — diagnose why so the connection shows a useful error.
      const pageContent = (await page.content()).toLowerCase();
      if (/clave din[aá]mica|segundo factor|c[oó]digo de verificaci[oó]n/.test(pageContent)) {
        const ss = await page.screenshot({ encoding: "base64" });
        return { success: false, bank, accounts: [], error: "El banco pide clave dinámica (2FA).", screenshot: ss as string, debug: debugLog.join("\n") };
      }
      const errorCheck = await page.evaluate(() => {
        const els = document.querySelectorAll('[class*="error"], [class*="alert"], [role="alert"]');
        for (const el of els) {
          const t = (el as HTMLElement).innerText?.trim();
          if (t && t.length > 5 && t.length < 200 && /(rut|clave|incorrect|inv[aá]lid|bloque|credencial|no coincide|intenta nuevamente)/i.test(t)) return t;
        }
        return null;
      });
      const ss = await page.screenshot({ encoding: "base64" });
      return {
        success: false, bank, accounts: [],
        error: errorCheck ? `Error del banco: ${errorCheck}` : "No se pudo confirmar el inicio de sesión (quedó en la página pública). Revisa credenciales o posible cambio del sitio del banco.",
        screenshot: ss as string, debug: debugLog.join("\n"),
      };
    }

    debugLog.push("6. Login OK! (verified)");
    progress("Sesión iniciada correctamente");
    await doSave(page, "04-post-login");

    // ── Phase 1: Account movements ──────────────────────────────
    //
    // PATCH (miaufinanzas): rewrote this block to prioritise clicking on a
    // REAL "Cuenta Corriente" product entry (the row with a $ balance) before
    // falling back to clickNavTarget. The upstream behaviour was to call
    // clickNavTarget first, which matched "Estado de Cuenta CMR" via the
    // substring "estado de cuenta" — that navigates to the credit card
    // statement page, NOT the checking account movements. For users with an
    // actual checking account (e.g. Paola with "Cuenta Corriente + CMR Elite")
    // this meant account=0 movements despite a $3M balance. We also add a
    // handler for the "Seleccione una cuenta" modal that appears when the
    // user has multiple checking accounts (CLP + USD).
    debugLog.push("7. [Cuenta] Looking for Cartola/Movimientos...");
    progress("Buscando cartola de cuenta...");
    let navigated = false;

    // Step 7a: try to click a REAL "Cuenta Corriente / Vista" product entry.
    //
    // STRICT matching rules to avoid false positives (we previously matched a
    // "Te puede interesar… Crédito pre aprobado … Con Cuenta Corriente" promo
    // banner because it happened to contain the substring "cuenta corriente"
    // and a $ amount):
    //   - Text must START with "cuenta corriente" or "cuenta vista" after
    //     trim/lowercase. That excludes promos/banners that mention the
    //     product name mid-sentence.
    //   - Text must contain a $ (real products always render a balance).
    //   - Text length < 150 (real product rows are short).
    //   - Blacklist obvious promo keywords.
    // Hide any blocking modal/overlay. We previously tried to find a
    // "close" button inside the modal, but that was fragile (buttons can
    // be SVG icons, custom web components, inside shadow DOM, etc). The
    // reliable approach is to detect overlay-like elements by their
    // COMPUTED STYLE and force display:none on them.
    //
    // An overlay is identified by ALL of:
    //   - position: fixed or absolute
    //   - high z-index (>100)
    //   - occupies a large chunk of the viewport (>20% of window area)
    //   - is currently visible
    //
    // Those three criteria together reliably distinguish a modal/promo
    // from regular layout containers (which are static/relative). And
    // because we're NOT clicking anything, there's no risk of
    // accidentally triggering logout.
    const overlayKillResult = await page.evaluate(() => {
      const vw = window.innerWidth;
      const vh = window.innerHeight;
      const vArea = vw * vh;
      const hidden: Array<{ tag: string; cls: string; area: number; z: string }> = [];
      const all = Array.from(document.querySelectorAll<HTMLElement>("*"));
      for (const el of all) {
        const s = window.getComputedStyle(el);
        if (s.position !== "fixed" && s.position !== "absolute") continue;
        const zIndex = parseInt(s.zIndex || "0", 10);
        if (!Number.isFinite(zIndex) || zIndex < 100) continue;
        if (s.display === "none" || s.visibility === "hidden" || s.opacity === "0") continue;
        const r = el.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0) continue;
        const area = r.width * r.height;
        if (area < vArea * 0.2) continue;
        // Don't hide anything that contains the main product list. Check
        // for "cuenta corriente" text inside — if the overlay is actually
        // the dashboard, we'd be shooting ourselves in the foot.
        const innerText = (el.innerText || "").toLowerCase();
        if (innerText.includes("mis productos") || /cuenta corriente\s*\n\s*\d/i.test(el.innerText || "")) continue;
        el.style.setProperty("display", "none", "important");
        hidden.push({ tag: el.tagName.toLowerCase(), cls: (el.className || "").slice(0, 60), area: Math.round(area), z: s.zIndex });
      }
      return { hidden };
    });
    if (overlayKillResult.hidden.length > 0) {
      debugLog.push(`  [Overlay] Hidden ${overlayKillResult.hidden.length} blocking overlay(s):`);
      for (const h of overlayKillResult.hidden) {
        debugLog.push(`    ${h.tag}.${h.cls} (area=${h.area}, z=${h.z})`);
      }
      await delay(1500);
    }

    // DIAGNOSTIC: print all <a> elements whose href or text hints at account
    // movements. Used to discover how to reach the right page in Falabella's
    // Angular app without depending on the "Estado de Cuenta CMR" sidebar.
    const linkSurvey = await page.evaluate(() => {
      const anchors = Array.from(document.querySelectorAll<HTMLAnchorElement>("a[href]"));
      const hits: Array<{ text: string; href: string }> = [];
      for (const a of anchors) {
        const href = a.href || "";
        const text = (a.innerText || "").trim().slice(0, 60);
        if (!href) continue;
        if (/cartola|movimiento|cuenta|consultar|saldos/i.test(href) || /cartola|movimiento|cuenta corriente|saldos/i.test(text)) {
          hits.push({ text, href: href.slice(0, 140) });
        }
      }
      return hits.slice(0, 20);
    });
    debugLog.push(`  [Survey] account-related links (${linkSurvey.length}):`);
    for (const h of linkSurvey) {
      debugLog.push(`    "${h.text}" → ${h.href}`);
    }

    // Click strategy (in order of preference):
    //   Pass 1: <a> / <button> / [role='button'] whose text starts with
    //           "Cuenta Corriente/Vista". NO $ required — the real link
    //           often has just the name + account number.
    //   Pass 2: any DOM element whose text starts with the product name
    //           AND contains a $. Walk up to the closest clickable
    //           ancestor. This is the fallback for dashboards where the
    //           actionable element is a wrapper <div> or custom component.
    //
    // We mark the chosen element with data-miau-click="1" and then call
    // page.click() on that selector — Puppeteer moves a real mouse and
    // dispatches a proper hover/focus/click sequence, which Angular
    // handlers accept.
    const clickTarget = await page.evaluate(() => {
      const BLACKLIST = /(te puede interesar|pre aprobado|crédito hipotecario|línea de crédito|avance|súper avance|simula aquí|abre tu|solicita|contrata|conoce más)/i;
      // Must be visible: non-zero box + not display:none + not visibility:hidden.
      const isVisible = (el: HTMLElement) => {
        const r = el.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0) return false;
        const s = window.getComputedStyle(el);
        if (s.display === "none" || s.visibility === "hidden" || s.opacity === "0") return false;
        return true;
      };
      // Pass 1: prefer a direct <a>/<button>/[role="button"] that is VISIBLE.
      // Invisible anchors almost always come from collapsed dropdown menus
      // in the header nav, not from the dashboard product list.
      const pass1 = Array.from(document.querySelectorAll<HTMLElement>("a, button, [role='button']"));
      for (const el of pass1) {
        const text = (el.innerText || "").trim();
        if (text.length === 0 || text.length > 200) continue;
        const lower = text.toLowerCase();
        if (!/^(cuenta corriente|cuenta vista)/.test(lower)) continue;
        if (BLACKLIST.test(text)) continue;
        if (!isVisible(el)) continue;
        const href = (el as HTMLAnchorElement).href || "";
        if (href.includes("cc-nuevos") || href.includes("comenzar")) continue;
        el.setAttribute("data-miau-click", "1");
        const tag = el.tagName.toLowerCase();
        return { marked: true, text: text.slice(0, 80), tag, href: href.slice(0, 140), via: "pass1" };
      }
      // Pass 2: any VISIBLE element starting with the product name AND
      // having a $ amount.
      const candidates: Array<{ el: HTMLElement; depth: number; text: string }> = [];
      const walk = (node: Element, depth: number) => {
        const htmlEl = node as HTMLElement;
        const text = (htmlEl.innerText || "").trim();
        if (text.length > 0 && text.length <= 200) {
          const lower = text.toLowerCase();
          if (/^(cuenta corriente|cuenta vista)/.test(lower) && /\$/.test(text) && !BLACKLIST.test(text) && isVisible(htmlEl)) {
            candidates.push({ el: htmlEl, depth, text });
          }
        }
        for (const child of Array.from(node.children)) walk(child, depth + 1);
      };
      walk(document.documentElement, 0);
      if (candidates.length === 0) return { marked: false, reason: "no-candidate" };
      candidates.sort((a, b) => b.depth - a.depth);
      const innermost = candidates[0];
      let target: HTMLElement | null = innermost.el;
      for (let hop = 0; hop < 8 && target; hop++) {
        const style = window.getComputedStyle(target);
        const tag = target.tagName.toLowerCase();
        if (tag === "a" || tag === "button" || target.getAttribute("role") === "button" || target.hasAttribute("onclick") || style.cursor === "pointer") {
          break;
        }
        target = target.parentElement;
      }
      if (!target) target = innermost.el;
      target.setAttribute("data-miau-click", "1");
      const href = (target as HTMLAnchorElement).href || "";
      const tag = target.tagName.toLowerCase();
      return { marked: true, text: innermost.text.slice(0, 80), tag, href: href.slice(0, 140), via: "pass2-depth" + innermost.depth };
    });
    const clickedRealAccount: { clicked: boolean; text?: string; via?: string } = { clicked: false };
    if (clickTarget.marked) {
      debugLog.push(`  [Account] Marked clickable via ${clickTarget.via}: tag=${clickTarget.tag} href=${clickTarget.href || "(none)"} text="${clickTarget.text}"`);
      // Diagnostic: report the <a>'s computed style and bounding box so we
      // can tell whether it's visible/clickable at all.
      const meta = await page.evaluate(() => {
        const el = document.querySelector<HTMLElement>("[data-miau-click='1']");
        if (!el) return null;
        el.scrollIntoView({ block: "center", inline: "center" });
        const r = el.getBoundingClientRect();
        const s = window.getComputedStyle(el);
        return {
          rect: { x: r.left, y: r.top, w: r.width, h: r.height },
          display: s.display,
          visibility: s.visibility,
          opacity: s.opacity,
          pointerEvents: s.pointerEvents,
          parentTag: el.parentElement?.tagName.toLowerCase() || null,
          parentClass: el.parentElement?.className || null,
        };
      });
      debugLog.push(`  [Account] element meta: ${JSON.stringify(meta)}`);

      if (meta && meta.rect.w > 0 && meta.rect.h > 0) {
        // Use page.mouse.click() with real coordinates. This dispatches a
        // TRUSTED event sequence (isTrusted=true) which Angular routers
        // accept. Unlike page.click(selector), this doesn't do hit-testing,
        // so hidden/overlay issues don't reject the click.
        const cx = meta.rect.x + meta.rect.w / 2;
        const cy = meta.rect.y + meta.rect.h / 2;
        try {
          await page.mouse.click(cx, cy);
          clickedRealAccount.clicked = true;
          clickedRealAccount.text = clickTarget.text;
          clickedRealAccount.via = `mouse.click-${clickTarget.tag}`;
        } catch (err) {
          debugLog.push(`  [Account] mouse.click failed: ${err instanceof Error ? err.message : String(err)}`);
        }
      }

      if (!clickedRealAccount.clicked) {
        // Fallback: try Puppeteer's page.click() then in-browser el.click().
        try {
          await page.click("[data-miau-click='1']");
          clickedRealAccount.clicked = true;
          clickedRealAccount.text = clickTarget.text;
          clickedRealAccount.via = `puppeteer-${clickTarget.tag}`;
        } catch (err) {
          debugLog.push(`  [Account] page.click() fallback also failed: ${err instanceof Error ? err.message : String(err)}`);
          const inBrowserClick = await page.evaluate(() => {
            const el = document.querySelector<HTMLElement>("[data-miau-click='1']");
            if (!el) return false;
            el.click();
            return true;
          });
          if (inBrowserClick) {
            clickedRealAccount.clicked = true;
            clickedRealAccount.text = clickTarget.text;
            clickedRealAccount.via = `in-browser-${clickTarget.tag}`;
          }
        }
      }

      // Clean up the marker so it doesn't interfere with later queries.
      await page.evaluate(() => {
        document.querySelectorAll("[data-miau-click]").forEach((el) => el.removeAttribute("data-miau-click"));
      });
    } else {
      debugLog.push(`  [Account] No product candidate found (reason: ${clickTarget.reason})`);
    }
    if (clickedRealAccount.clicked) {
      debugLog.push(`  [Account] Clicked product via ${clickedRealAccount.via}: ${clickedRealAccount.text}`);
      await delay(4000);
      debugLog.push(`  [Account] Post-click URL: ${page.url()}`);

      // Step 7b: handle the "Seleccione una cuenta" modal that appears when
      // the user has more than one checking account (e.g. CLP + USD).
      const modalResult = await page.evaluate(() => {
        const bodyText = document.body?.innerText || "";
        if (!/seleccione una cuenta/i.test(bodyText)) return { handled: false, reason: "no-modal" };
        // Scope to the modal container so we don't accidentally click items
        // outside of it.
        let modalRoot: HTMLElement | null = null;
        const allContainers = Array.from(document.querySelectorAll<HTMLElement>("[role='dialog'], .modal, .mat-dialog-container, div, section"));
        for (const el of allContainers) {
          const t = el.innerText || "";
          if (/seleccione una cuenta/i.test(t) && t.length < 3000) {
            modalRoot = el;
            break;
          }
        }
        const searchRoot: ParentNode = modalRoot || document;
        const candidates = Array.from(searchRoot.querySelectorAll<HTMLElement>("label, li, tr, option, [role='option'], [role='radio'], .option, .account-option, div, a, button"));
        const clpOptions: HTMLElement[] = [];
        for (const el of candidates) {
          const t = (el.innerText || "").trim();
          if (!t) continue;
          if (!/cuenta (corriente|vista)/i.test(t)) continue;
          if (/usd|d[oó]lar|moneda extranjera|m\/e/i.test(t)) continue;
          if (t.length > 200) continue;
          clpOptions.push(el);
        }
        if (clpOptions.length === 0) return { handled: false, reason: "no-clp-option" };
        const target = clpOptions[0];
        target.click();
        for (const inp of Array.from(target.querySelectorAll<HTMLInputElement>("input[type='radio'], input[type='checkbox']"))) {
          inp.click();
        }
        const btnRoot: ParentNode = modalRoot || document;
        const btns = Array.from(btnRoot.querySelectorAll<HTMLElement>("button, [role='button'], a.btn, a"));
        for (const btn of btns) {
          const bt = (btn.innerText || "").trim().toLowerCase();
          if (/^(aceptar|continuar|seleccionar|confirmar|ver movimientos|ok)$/i.test(bt)) {
            btn.click();
            return { handled: true, selected: target.innerText.slice(0, 80), button: bt };
          }
        }
        return { handled: true, selected: target.innerText.slice(0, 80), button: null };
      });
      if (modalResult.handled) {
        debugLog.push(`  [Modal] Picked CLP account: ${modalResult.selected} (button=${modalResult.button})`);
        await delay(5000);
      } else if (modalResult.reason && modalResult.reason !== "no-modal") {
        debugLog.push(`  [Modal] Detected but could not handle: ${modalResult.reason}`);
      } else {
        debugLog.push(`  [Modal] No picker modal shown (single account user)`);
      }

      // Step 7c: on the checking-account landing page, click the tab that
      // shows the movements. We DELIBERATELY do not call the full
      // clickNavTarget() here because it matches "estado de cuenta" (which is
      // the CMR credit-card statement link in the global sidebar) and would
      // unwind all our progress. This is a safer version that accepts
      // movement-like tab names while BLOCKING anything mentioning "CMR" or
      // "credito" or "estado de cuenta".
      const clickedMovementsTab = await page.evaluate(() => {
        const ACCEPT = [
          /saldos? y movimientos/i,
          /últimos movimientos/i,
          /\bmovimientos\b/i,
          /cartola/i,
        ];
        const BLOCK = /(cmr|crédito|credito|estado de cuenta|facturad|pagar|transfer|pr[eé]stamo|inversi|seguro)/i;
        const elements = Array.from(document.querySelectorAll<HTMLElement>("a, button, [role='tab'], [role='menuitem'], li"));
        const availableTabs: string[] = [];
        for (const pat of ACCEPT) {
          for (const el of elements) {
            const text = (el.innerText || "").trim();
            if (text.length === 0 || text.length > 60) continue;
            availableTabs.push(text);
            if (!pat.test(text)) continue;
            if (BLOCK.test(text)) continue;
            const href = (el as HTMLAnchorElement).href || "";
            if (href.includes("cc-nuevos") || href.includes("comenzar")) continue;
            el.click();
            return { clicked: true, text, availableTabs: availableTabs.slice(0, 20) };
          }
        }
        return { clicked: false, availableTabs: availableTabs.slice(0, 20) };
      });
      if (clickedMovementsTab.clicked) {
        debugLog.push(`  [Account] Clicked sub-tab: "${clickedMovementsTab.text}"`);
        await delay(4000);
        navigated = true;
      } else {
        debugLog.push(`  [Account] No movements sub-tab found. Available tabs: ${clickedMovementsTab.availableTabs.join(" | ")}`);
        navigated = true;
      }
    } else {
      // Step 7c: No real checking account visible — fall back to the original
      // clickNavTarget path. This is the correct behaviour for users who only
      // have a CMR credit card (no checking account at all).
      debugLog.push(`  [Account] No checking-account product with balance found — falling back to clickNavTarget`);
      navigated = await clickNavTarget(page, debugLog);
    }

    await tryExpandDateRange(page, debugLog);
    progress("Extrayendo movimientos de cuenta...");
    const accountMovements = await paginateAndExtract(page, extractAccountMovements, debugLog);
    debugLog.push(`8. [Cuenta] Extracted ${accountMovements.length} movements`);
    progress(`Cuenta: ${accountMovements.length} movimientos encontrados`);

    let balance: number | undefined;
    if (accountMovements.length > 0 && accountMovements[0].balance > 0) balance = accountMovements[0].balance;
    if (balance === undefined) {
      balance = await page.evaluate(() => {
        const match = (document.body?.innerText || "").match(/Saldo disponible[\s\S]{0,50}\$\s*([\d.]+)/i);
        if (match) return parseInt(match[1].replace(/[^0-9]/g, ""), 10);
        return undefined;
      });
    }

    // ── Phase 2: CMR credit card movements ─────────────────────
    debugLog.push("9. [CMR] Navigating back to authenticated dashboard...");
    progress("Navegando a tarjeta de crédito...");
    await page.goto(dashboardUrl, { waitUntil: "networkidle2", timeout: 30000 });
    await delay(2000);
    await closePopups(page);

    debugLog.push("10. [CMR] Extracting credit card cupos...");
    const cmrBalance = await extractCupos(page, debugLog);
    const creditCardData: CreditCardBalance = cmrBalance ?? { label: "CMR" };

    debugLog.push("11. [CMR] Looking for CMR card product...");
    const cardClicked = await page.evaluate(() => {
      for (const sel of ["#cardDetail0", "[id^='cardDetail']", "app-credit-cards .card", "[class*='credit-card'] .card", "[class*='creditCard']"]) {
        const el = document.querySelector(sel);
        if (el) { (el as HTMLElement).click(); return true; }
      }
      for (const el of Array.from(document.querySelectorAll("a, button, div, li, [role='button']"))) {
        if ((el as HTMLElement).innerText?.trim().toLowerCase().includes("cmr") && (el as HTMLElement).innerText!.length < 100) { (el as HTMLElement).click(); return true; }
      }
      return false;
    });
    if (cardClicked) await waitForCmrMovements(page);
    await doSave(page, "06-cmr-card");

    // Owner filter
    if (owner !== "B") {
      await page.evaluate((host: string, value: string) => {
        const root = (document.querySelector(host) as Element & { shadowRoot?: ShadowRoot })?.shadowRoot || document;
        const select = root.querySelector("select[name='searchownership']") as HTMLSelectElement | null;
        if (select) { select.value = value; select.dispatchEvent(new Event("change", { bubbles: true })); }
      }, SHADOW_HOST, owner);
      await waitForCmrMovements(page);
    }

    // ── Tab: No facturados (default) ────────────────────────────
    debugLog.push("12. [CMR] Extracting TC por facturar...");
    progress("Extrayendo movimientos TC por facturar...");

    // Extract period info from "últimos movimientos" tab
    const unbilledInfo = await extractUnbilledPeriodInfo(page);
    if (unbilledInfo.nextBillingDate) {
      creditCardData.nextBillingDate = normalizeDate(unbilledInfo.nextBillingDate);
      debugLog.push(`  Próxima facturación: ${creditCardData.nextBillingDate}`);
    }
    if (unbilledInfo.nextDueDate) {
      creditCardData.nextDueDate = normalizeDate(unbilledInfo.nextDueDate);
      debugLog.push(`  Próximo vencimiento: ${creditCardData.nextDueDate}`);
    }
    if (unbilledInfo.periodExpenses !== undefined) {
      creditCardData.periodExpenses = unbilledInfo.periodExpenses;
      debugLog.push(`  Gastos del período: $${unbilledInfo.periodExpenses}`);
    }

    // DOM pagination first (preserves real installment counters); Excel as fallback
    let unbilledMovements: BankMovement[] = await paginateCmrMovements(page, debugLog);
    debugLog.push(`  TC por facturar (DOM): ${unbilledMovements.length}`);
    if (unbilledMovements.length === 0) {
      const excelPathUnbilled = await downloadCmrExcel(page, downloadDir, debugLog);
      if (excelPathUnbilled) {
        unbilledMovements = parseExcelMovements(excelPathUnbilled, MOVEMENT_SOURCE.credit_card_unbilled, debugLog);
        debugLog.push(`  TC por facturar (Excel fallback): ${unbilledMovements.length}`);
        // Remove so it doesn't collide with the facturados export filename.
        try { fs.unlinkSync(excelPathUnbilled); } catch { /* best-effort */ }
      }
    }

    await doSave(page, "07-cmr-no-facturados");

    // ── Tab: Facturados ─────────────────────────────────────────
    debugLog.push("13. [CMR] Extracting TC facturados...");
    progress("Extrayendo movimientos TC facturados...");

    let billedMovements: BankMovement[] = [];
    if (await clickCmrTab(page, "movimientos facturados", debugLog)) {
      // After clicking the facturados radio/label, Angular shows the billed
      // movements component (which may already be in the DOM but hidden, or
      // may be mounted fresh). We wait for table rows to appear in any shadow
      // root — the same strategy the debug script uses — rather than waiting
      // specifically for app-invoiced-movements to appear as a new element.
      // This handles both "lazy mount" and "show/hide" Angular patterns.
      await delay(2000);
      await waitForCmrMovements(page, 30000);
      // Extra settle time to let Angular finish rendering the full table.
      await delay(3000);
      await doSave(page, "07-cmr-facturados");

      // Extract lastStatement fields (billing date, billed amount, due date, minimum payment)
      const billedInfo = await extractBilledStatementInfo(page);
      if (billedInfo.billingDate && billedInfo.billedAmount && billedInfo.dueDate) {
        const billingDate = normalizeDate(billedInfo.billingDate);
        creditCardData.lastStatement = {
          billingDate,
          billedAmount: billedInfo.billedAmount,
          dueDate: normalizeDate(billedInfo.dueDate),
          minimumPayment: billedInfo.minimumPayment,
        };
        creditCardData.billingPeriod = monthYearLabel(billingDate);
        debugLog.push(`  lastStatement: facturado=${billingDate}, monto=$${billedInfo.billedAmount}, vence=${creditCardData.lastStatement.dueDate}, minimo=$${billedInfo.minimumPayment}`);
      }

      // DOM pagination first (preserves real installment counters); Excel as fallback
      billedMovements = await paginateInvoicedMovements(page, debugLog);
      debugLog.push(`  TC facturados (DOM): ${billedMovements.length}`);
      if (billedMovements.length === 0) {
        const excelPathBilled = await downloadCmrExcel(page, downloadDir, debugLog);
        if (excelPathBilled) {
          billedMovements = parseExcelMovements(excelPathBilled, MOVEMENT_SOURCE.credit_card_billed, debugLog);
          debugLog.push(`  TC facturados (Excel fallback): ${billedMovements.length}`);
        }
      }
    }

    // Note: "Estado de Cuenta CMR" tab renders the billing statement as PDF in an <iframe>.
    // The key billing fields (billingDate, billedAmount, dueDate, minimumPayment) are extracted
    // directly from the "movimientos facturados" tab HTML above.

    const cardMask = creditCardData.label.match(/\*{4}\d{4}/)?.[0];
    const withCard = (ms: BankMovement[]) => cardMask ? ms.map(m => ({ ...m, card: cardMask })) : ms;
    creditCardData.movements = deduplicateAcrossSources(
      deduplicateMovements([...withCard(unbilledMovements), ...withCard(billedMovements)]),
    );

    debugLog.push(`14. Total: ${accountMovements.length} account + ${creditCardData.movements.length} TC movements`);
    progress(`Listo — ${accountMovements.length + creditCardData.movements.length} movimientos totales`);

    await doSave(page, "08-final");
    const ss = doScreenshots ? (await page.screenshot({ encoding: "base64", fullPage: true })) as string : undefined;

    return {
      success: true,
      bank,
      accounts: [{ balance, movements: deduplicateMovements(accountMovements) }],
      creditCards: [creditCardData],
      screenshot: ss,
      debug: debugLog.join("\n"),
    };
  } finally {
    cleanupTempDir(downloadDir);
  }
}

// ─── Export ──────────────────────────────────────────────────────

const falabella: BankScraper = {
  id: "falabella",
  name: "Banco Falabella",
  url: BANK_URL,
  scrape: (options) => runScraper("falabella", options, { extraArgs: ["--disable-notifications"] }, scrapeFalabella),
};

export default falabella;
