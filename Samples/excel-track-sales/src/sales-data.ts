/* global fetch */

export const PRODUCTS = ["Keyboard", "Mouse", "Monitor", "Laptop"] as const;
export type Product = (typeof PRODUCTS)[number];
export const START_DATE = "2026-11-01";
export const END_DATE = "2026-11-12";
export const HEADERS = [
  "Date",
  "Units",
  "AvgUnitPrice",
  "Tax",
  "Discount",
  "ProductSales",
  "ServiceSales",
];
export const METRICS = HEADERS.slice(1);

export interface Sale {
  date: string;
  units: number;
  averageUnitPrice: number;
  tax: number;
  discount: number;
  productSales: number;
  serviceSales: number;
}

export interface ProductSale extends Sale {
  product: Product;
}

export function parseSalesData(value: unknown): ProductSale[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error("Sales data must be a nonempty array of records.");
  }
  return value.map((record: unknown) => {
    if (typeof record !== "object" || record === null) {
      throw new Error("Sales data contains an invalid record.");
    }
    const sale = record as Record<string, unknown>;
    const product = PRODUCTS.find((product) => product === sale.product);
    if (typeof sale.date !== "string" || !product) {
      throw new Error("Sales data contains an invalid product or date.");
    }
    validateDate(sale.date);
    const number = (field: string): number => {
      const value = sale[field];
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
        throw new Error(`Sales data contains an invalid numeric value for ${field}.`);
      }
      return value;
    };
    const units = number("units");
    if (!Number.isInteger(units)) {
      throw new Error("Sales units must be whole numbers.");
    }
    return {
      product,
      date: sale.date,
      units,
      averageUnitPrice: number("averageUnitPrice"),
      tax: number("tax"),
      discount: number("discount"),
      productSales: number("productSales"),
      serviceSales: number("serviceSales"),
    };
  });
}

export async function loadSalesData(): Promise<ProductSale[]> {
  const response = await fetch("./sales-data.json");
  if (!response.ok) {
    throw new Error(`Could not load demo sales data (HTTP ${response.status}).`);
  }
  return parseSalesData(await response.json());
}

export function validateDate(value: string): void {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw new Error("Enter a date in YYYY-MM-DD format.");
  }
  const timestamp = Date.parse(`${value}T00:00:00Z`);
  if (!Number.isFinite(timestamp) || new Date(timestamp).toISOString().slice(0, 10) !== value) {
    throw new Error("Enter a valid calendar date.");
  }
}

export function validateDateRange(startDate: string, endDate: string): void {
  validateDate(startDate);
  validateDate(endDate);
  if (startDate < START_DATE || startDate > END_DATE) {
    throw new Error(`Choose a start date between ${START_DATE} and ${END_DATE}.`);
  }
  if (endDate < START_DATE || endDate > END_DATE) {
    throw new Error(`Choose an end date between ${START_DATE} and ${END_DATE}.`);
  }
  if (endDate < startDate) {
    throw new Error("End date must be on or after the start date.");
  }
}

export function getSales(
  product: string,
  startDate: string,
  endDate: string,
  transactions: readonly ProductSale[]
): Sale[] {
  const productIndex = PRODUCTS.findIndex((item) => item === product);
  if (productIndex < 0) {
    throw new Error("Choose Keyboard, Mouse, Monitor, or Laptop.");
  }
  validateDateRange(startDate, endDate);
  return transactions
    .filter((sale) => sale.product === product && sale.date >= startDate && sale.date <= endDate)
    .slice()
    .sort((left, right) => left.date.localeCompare(right.date));
}

export function toTableRows(sales: Sale[], dateSerials: readonly number[]): number[][] {
  if (dateSerials.length !== sales.length || dateSerials.some((value) => !Number.isFinite(value))) {
    throw new Error("Each sales record must have a valid Excel date value.");
  }
  return sales.map((sale, index) => {
    validateDate(sale.date);
    return [
      dateSerials[index],
      sale.units,
      sale.averageUnitPrice,
      sale.tax,
      sale.discount,
      sale.productSales,
      sale.serviceSales,
    ];
  });
}
