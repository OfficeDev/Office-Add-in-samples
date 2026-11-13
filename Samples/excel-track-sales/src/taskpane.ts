/* global console, document, Office, OfficeExtension, HTMLInputElement, HTMLSelectElement, HTMLFormElement, HTMLFieldSetElement, HTMLElement, HTMLButtonElement */

import {
  CHART_THEMES,
  CHART_TYPES,
  ChartOptions,
  formatChart,
  formatTable,
  hasSalesTable,
  importSales,
  TABLE_STYLES,
} from "./sales-workbook";
import {
  END_DATE,
  getSales,
  HEADERS,
  loadSalesData,
  METRICS,
  PRODUCTS,
  START_DATE,
  validateDateRange,
} from "./sales-data";

const form = document.querySelector<HTMLFormElement>("#sales-form");
const product = document.querySelector<HTMLSelectElement>("#product");
const date = document.querySelector<HTMLInputElement>("#start-date");
const endDate = document.querySelector<HTMLInputElement>("#end-date");
const metric = document.querySelector<HTMLSelectElement>("#chart-metric");
const chartType = document.querySelector<HTMLSelectElement>("#chart-type");
const theme = document.querySelector<HTMLSelectElement>("#chart-theme");
const tableStyle = document.querySelector<HTMLSelectElement>("#table-style");
const settings = document.querySelector<HTMLFieldSetElement>("#settings");
const status = document.querySelector<HTMLElement>("#status");
let dataLoaded = false;
let busy = false;

function fillOptions(select: HTMLSelectElement, values: readonly string[]): void {
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
}

function selectedKey<T extends object>(select: HTMLSelectElement, choices: T): keyof T {
  const value = select.value;
  if (!Object.prototype.hasOwnProperty.call(choices, value)) {
    throw new Error("Choose a valid setting.");
  }
  return value as keyof T;
}

function chartOptions(): ChartOptions {
  return {
    metric: metric.value,
    type: selectedKey(chartType, CHART_TYPES),
    theme: selectedKey(theme, CHART_THEMES),
  };
}

async function perform(action: () => Promise<void>, message: string): Promise<void> {
  if (busy) {
    return;
  }
  busy = true;
  form.querySelector<HTMLButtonElement>("button").disabled = true;
  settings.disabled = true;
  status.textContent = "Working...";
  status.className = "";
  try {
    await action();
    status.textContent = message;
  } catch (error) {
    console.error(error);
    if (error instanceof OfficeExtension.Error) {
      const location = error.debugInfo.errorLocation;
      status.textContent = `${error.message} (${error.code}${location ? ` at ${location}` : ""})`;
    } else {
      status.textContent = error instanceof Error ? error.message : String(error);
    }
    status.className = "error";
  } finally {
    busy = false;
    form.querySelector<HTMLButtonElement>("button").disabled = false;
    settings.disabled = !dataLoaded;
  }
}

fillOptions(product, PRODUCTS);
fillOptions(metric, METRICS);
metric.value = "ProductSales";
date.value = START_DATE;
endDate.value = END_DATE;
function updateDateRangeValidity(): void {
  endDate.setCustomValidity(
    date.value && endDate.value && endDate.value < date.value
      ? "End date must be on or after the start date."
      : ""
  );
}
[date, endDate].forEach((input) => {
  input.min = START_DATE;
  input.max = END_DATE;
  if (input.type !== "date") {
    input.pattern = "\\d{4}-\\d{2}-\\d{2}";
    input.placeholder = "YYYY-MM-DD";
  }
  input.addEventListener("input", updateDateRangeValidity);
});
HEADERS.forEach((header) => {
  const label = document.createElement("label");
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.value = header;
  checkbox.checked = true;
  label.appendChild(checkbox);
  label.appendChild(document.createTextNode(` ${header}`));
  document.getElementById("columns").appendChild(label);
});

form.addEventListener("submit", (event) => {
  event.preventDefault();
  updateDateRangeValidity();
  if (!form.reportValidity()) {
    return;
  }
  const selectedProduct = product.value;
  const selectedStartDate = date.value;
  const selectedEndDate = endDate.value;
  void perform(async () => {
    validateDateRange(selectedStartDate, selectedEndDate);
    const sales = getSales(
      selectedProduct,
      selectedStartDate,
      selectedEndDate,
      await loadSalesData()
    );
    await importSales(sales, chartOptions());
    dataLoaded = true;
    tableStyle.value = "blue";
    document.querySelectorAll<HTMLInputElement>("#columns input").forEach((input) => {
      input.checked = true;
    });
  }, `Imported sales for ${selectedProduct} from ${selectedStartDate} through ${selectedEndDate}.`);
});

document.getElementById("apply-table").addEventListener("click", () => {
  void perform(async () => {
    const columns = Array.from(
      document.querySelectorAll<HTMLInputElement>("#columns input:checked")
    ).map((input) => input.value);
    await formatTable(selectedKey(tableStyle, TABLE_STYLES), columns);
  }, "Table settings applied.");
});

document.getElementById("apply-chart").addEventListener("click", () => {
  void perform(() => formatChart(chartOptions()), "Chart settings applied.");
});

Office.onReady((info) => {
  document.getElementById("sideload-msg").hidden = true;
  if (info.host !== Office.HostType.Excel) {
    status.textContent = "Open this add-in in Excel to track sales.";
    return;
  }
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    status.textContent =
      "This sample requires ExcelApi 1.9 or later. Update Excel before continuing.";
    return;
  }
  document.getElementById("app-body").hidden = false;
  void perform(async () => {
    dataLoaded = await hasSalesTable();
  }, "Ready. Import demo sales, or apply settings to an existing sales table.");
});
