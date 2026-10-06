import { createInterface } from 'node:readline/promises';
import { resolve } from 'node:path';
import { stdin, stdout } from 'node:process';
import ExcelJS from 'exceljs';
import { chromium, type Browser, type Locator, type Page } from '@playwright/test';

const CDP_ENDPOINT = process.env.VOLZA_CDP_ENDPOINT || 'http://localhost:9222';
const TABLE_READY_TIMEOUT = 90_000;
const DEFAULT_ENTRY_URL =
  'https://app.volza.com/workspace/search/25226929#Shipments';
const PAGE_SIZE = 50;

interface SearchCriterion {
  name: string;
  value: string;
}

interface ShipmentPage {
  headers: string[];
  rows: string[][];
}

function validateEntryUrl(value: string): string {
  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error('Enter a valid Volza workspace URL.');
  }

  if (
    !['http:', 'https:'].includes(url.protocol) ||
    url.hostname !== 'app.volza.com'
  ) {
    throw new Error('The URL must be on app.volza.com.');
  }

  return url.toString();
}

async function readEntryUrl() {
  const rl = createInterface({ input: stdin, output: stdout });
  try {
    const answer = (await rl.question(
      `Enter the Volza workspace URL (press Enter for ${DEFAULT_ENTRY_URL}): `,
    )).trim();
    return validateEntryUrl(answer || DEFAULT_ENTRY_URL);
  } finally {
    rl.close();
  }
}

// Capture the active criteria shown in the search controls, not shipment values.
async function readSearchCriteria(page: Page): Promise<SearchCriterion[]> {
  return page.evaluate<SearchCriterion[]>(() => {
    const criteria: SearchCriterion[] = [];

    const searchTerms = [
      ...document.querySelectorAll<HTMLElement>('.react-tags__selected-tag'),
    ]
      .map((tag) => tag.innerText.trim())
      .filter(Boolean);
    if (searchTerms.length) {
      criteria.push({
        name: 'Selected search terms',
        value: searchTerms.join(', '),
      });
    }

    for (const input of document.querySelectorAll<HTMLInputElement>(
      'input[placeholder="Start date"], input[placeholder="End date"]',
    )) {
      const value = input.value.trim().replace(/\s+/g, ' ');
      if (value) criteria.push({ name: input.placeholder, value });
    }

    for (const tag of document.querySelectorAll<HTMLElement>('.ant-tag')) {
      const value = tag.innerText.trim().replace(/\s+/g, ' ');
      if (value) criteria.push({ name: 'Applied filter', value });
    }

    if (criteria.length === 0) {
      throw new Error(
        'Could not read search criteria from the visible Volza search controls.',
      );
    }

    return criteria;
  });
}

// Volza uses Ant Design's page-size selector; confirm the UI accepts 50/page.
async function setPageSize(page: Page): Promise<void> {
  const pager = page.locator('.ant-pagination').first();
  const sizeChanger = pager.locator('.ant-pagination-options-size-changer');
  if ((await sizeChanger.count()) === 0) {
    throw new Error('Could not find the results page-size selector.');
  }

  const currentSize = sizeChanger.locator('.ant-select-selection-item');
  if ((await currentSize.innerText()).trim() !== `${PAGE_SIZE} / page`) {
    await sizeChanger.click();
    const option = page
      .locator('.ant-select-dropdown:visible .ant-select-item-option-content')
      .filter({ hasText: new RegExp(`^${PAGE_SIZE}\\s*/\\s*page$`) })
      .first();
    await option.waitFor({ state: 'visible', timeout: TABLE_READY_TIMEOUT });
    await option.click();
    await page.waitForFunction(
      (expectedSize) =>
        [...document.querySelectorAll<HTMLElement>('.ant-pagination-options-size-changer .ant-select-selection-item')]
          .some((item) => item.textContent.trim() === `${expectedSize} / page`),
      PAGE_SIZE,
      { timeout: TABLE_READY_TIMEOUT },
    );
  }
}

async function waitForShipmentTable(page: Page): Promise<void> {
  await page.waitForFunction(
    () =>
      [...document.querySelectorAll('table')].some(
        (table) =>
          table.querySelector('tbody') &&
          (table.querySelectorAll('tbody tr.ant-table-row').length > 0 ||
            table.parentElement?.classList.contains('ant-table-body')),
      ),
    { timeout: TABLE_READY_TIMEOUT },
  );
}

async function readCurrentPage(page: Page): Promise<ShipmentPage> {
  return page.evaluate<ShipmentPage>(() => {
    const tables = [...document.querySelectorAll('table')];
    const dataTables = tables
      .map((table) => {
        const rows = [
          ...table.querySelectorAll<HTMLTableRowElement>(
            'tbody tr.ant-table-row',
          ),
        ];
        const width = rows.reduce(
          (maximum, row) => Math.max(maximum, row.cells.length),
          0,
        );
        return { table, rows, width };
      })
      .filter(
        ({ table, rows }) =>
          rows.length > 0 ||
          (table.querySelector('tbody') &&
            table.parentElement?.classList.contains('ant-table-body')),
      )
      .sort(
        (left, right) =>
          right.width - left.width || right.rows.length - left.rows.length,
      );

    if (dataTables.length === 0) {
      throw new Error('Could not find the shipment results table.');
    }

    const dataTable = dataTables[0];
    const headerTables = tables
      .filter((table) => table.querySelector('thead th'))
      .sort(
        (left, right) =>
          right.querySelectorAll('thead th').length -
          left.querySelectorAll('thead th').length,
      );
    const headerCells = headerTables[0]
      ? [...headerTables[0].querySelectorAll<HTMLTableCellElement>('thead th')]
      : [];

    if (headerCells.length === 0) {
      throw new Error('Could not read column headings from the shipment table.');
    }

    const rowWidth = dataTable.rows[0]?.cells.length ?? headerCells.length;
    if (headerCells.length !== rowWidth) {
      throw new Error(
        `Column heading count (${headerCells.length}) does not match ` +
          `shipment cell count (${rowWidth}); refusing an incomplete export.`,
      );
    }
    const names = Array.from({ length: rowWidth }, (_, index) => {
      const header = headerCells[index];
      const label = header?.innerText.trim().replace(/\s+/g, ' ');
      return label;
    });
    const columnIndexes = names.flatMap((name, index) =>
      name ? [{ name, index }] : [],
    );
    if (columnIndexes.length === 0) {
      throw new Error('The shipment table has no named columns.');
    }
    const counts = new Map();
    const headers = columnIndexes.map(({ name }) => {
      const count = (counts.get(name) || 0) + 1;
      counts.set(name, count);
      return count === 1 ? name : `${name} (${count})`;
    });

    if (dataTable.rows.some((row) => row.cells.length !== rowWidth)) {
      throw new Error(
        'Shipment rows have inconsistent cell counts; refusing an incomplete export.',
      );
    }

    return {
      headers,
      rows: dataTable.rows.map((row) =>
        columnIndexes.map(({ index }) =>
          (row.cells[index]?.innerText || '').trim().replace(/\s+/g, ' '),
        ),
      ),
    };
  });
}

async function getPageNumber(pager: Locator): Promise<string> {
  return (await pager.locator('.ant-pagination-item-active').innerText()).trim();
}

async function goToNextPage(page: Page, pager: Locator): Promise<boolean> {
  const next = pager.locator('.ant-pagination-next');
  const nextButton = next.locator('button');
  if (
    (await next.getAttribute('aria-disabled')) === 'true' ||
    (await nextButton.isDisabled())
  ) {
    return false;
  }

  const currentPage = await getPageNumber(pager);
  await nextButton.evaluate((button: HTMLButtonElement) => button.click());
  await page.waitForFunction(
    (previousPage) =>
      document.querySelector('.ant-pagination-item-active')?.textContent.trim() !==
      previousPage,
    currentPage,
    { timeout: TABLE_READY_TIMEOUT },
  );
  return true;
}

function createOutputPath(criteria: SearchCriterion[]): string {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const relevance = criteria
    .filter(({ name }) => name !== 'Search URL')
    .map(({ value }) => value)
    .join(' ')
    .normalize('NFKD')
    .replace(/[^\w]+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 80)
    .toLowerCase();
  const context = relevance ? `${relevance}-` : '';
  return resolve(
    process.cwd(),
    `volza-shipments-${context}${timestamp}.xlsx`,
  );
}

async function main() {
  const entryUrl = await readEntryUrl();
  let browser: Browser | undefined;
  let page: Page | undefined;

  try {
    browser = await chromium.connectOverCDP(CDP_ENDPOINT);
    const context = browser.contexts()[0];
    if (!context) {
      throw new Error('No browser context is available at the CDP endpoint.');
    }

    page = await context.newPage();
    await page.goto(entryUrl, { waitUntil: 'domcontentloaded' });
    await waitForShipmentTable(page);

    if (new URL(page.url()).hostname !== 'app.volza.com') {
      throw new Error('Volza redirected away from the workspace; check your login.');
    }

    const searchCriteria = await readSearchCriteria(page);
    searchCriteria.unshift({ name: 'Search URL', value: entryUrl });
    await setPageSize(page);

    const worksheetData: string[][] = [];
    const pager = page.locator('.ant-pagination').first();
    let pageNumber = 1;
    let headers: string[] | undefined;

    do {
      const currentPage = await readCurrentPage(page);
      if (!headers) {
        headers = currentPage.headers;
      } else if (
        headers.length !== currentPage.headers.length ||
        headers.some((header, index) => header !== currentPage.headers[index])
      ) {
        throw new Error('Shipment table columns changed while paging.');
      }

      worksheetData.push(...currentPage.rows);
      console.log(
        `Read page ${pageNumber}: ${currentPage.rows.length} shipment rows ` +
          `(${worksheetData.length} total).`,
      );
      pageNumber += 1;
    } while (await goToNextPage(page, pager));

    if (!headers) {
      throw new Error('No shipment table headings were found.');
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Shipments');
    worksheet.addRow(['Search criteria']);
    worksheet.getRow(1).font = { bold: true, size: 14 };
    worksheet.addRow(['Criteria', 'Value']);
    worksheet.getRow(2).font = { bold: true };
    for (const criterion of searchCriteria) {
      worksheet.addRow([criterion.name, criterion.value]);
    }
    worksheet.addRow([]);
    const headingRow = worksheet.addRow(headers);
    headingRow.font = { bold: true };
    worksheet.views = [{ state: 'frozen', ySplit: headingRow.number }];
    worksheet.autoFilter = {
      from: { row: headingRow.number, column: 1 },
      to: { row: headingRow.number, column: headers.length },
    };

    for (const row of worksheetData) {
      worksheet.addRow(row);
    }

    worksheet.columns = headers.map((header, index) => {
      const maxLength = Math.max(
        header.length,
        ...worksheetData.map((row) => String(row[index] || '').length),
      );
      return { width: Math.min(Math.max(maxLength + 2, 12), 60) };
    });
    worksheet.getColumn(1).width = Math.max(
      worksheet.getColumn(1).width || 0,
      24,
    );
    worksheet.getColumn(2).width = Math.max(
      worksheet.getColumn(2).width || 0,
      40,
    );

    const outputPath = createOutputPath(searchCriteria);
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Saved ${worksheetData.length} shipments to ${outputPath}`);
  } finally {
    if (page) {
      await page.close();
    }
    if (browser) {
      await browser.close();
    }
  }
}

main().catch((error) => {
  console.error(
    `Export failed: ${error instanceof Error ? error.message : String(error)}`,
  );
  process.exitCode = 1;
});
