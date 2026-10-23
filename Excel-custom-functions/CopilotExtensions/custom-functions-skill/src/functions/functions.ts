/* global CustomFunctions, fetch, Response */

const FUNCTION_CURRENCY_CODE_PATTERN = /^[A-Z]{3}$/;
const FRANKFURTER_API = "https://api.frankfurter.dev/v2";

interface FrankfurterRate {
  rate?: unknown;
}

/**
 * Returns the latest exchange rate between two currencies.
 * @customfunction FXRATE
 * @param sourceCurrency Three-letter source currency code.
 * @param reportingCurrency Three-letter reporting currency code.
 * @returns The latest exchange rate.
 */
export function fxrate(sourceCurrency: string, reportingCurrency: string): Promise<number> {
  const source = normalizeFunctionCurrency(sourceCurrency, "source");
  const reporting = normalizeFunctionCurrency(reportingCurrency, "reporting");

  if (source === reporting) {
    return Promise.resolve(1);
  }

  return fetch(`${FRANKFURTER_API}/rate/${source.toLowerCase()}/${reporting.toLowerCase()}`).then(
    (response: Response) => {
      if (!response.ok) {
        const errorCode =
          response.status === 400 || response.status === 404 || response.status === 422
            ? CustomFunctions.ErrorCode.invalidValue
            : CustomFunctions.ErrorCode.notAvailable;
        throw new CustomFunctions.Error(errorCode, "No exchange rate is available for this currency pair.");
      }

      return response.json().then(
        (payload: FrankfurterRate) => {
          if (typeof payload.rate !== "number" || !Number.isFinite(payload.rate) || payload.rate <= 0) {
            throw new CustomFunctions.Error(
              CustomFunctions.ErrorCode.notAvailable,
              "The exchange-rate service did not return a valid rate."
            );
          }
          return payload.rate;
        },
        () => {
          throw new CustomFunctions.Error(
            CustomFunctions.ErrorCode.notAvailable,
            "The exchange-rate service returned an invalid response."
          );
        }
      );
    },
    () => {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        "The exchange-rate service could not be reached."
      );
    }
  );
}

/**
 * Classifies a converted quote relative to its budget.
 * @customfunction BUDGETSTATUS
 * @param convertedQuote Converted quote amount in the reporting currency.
 * @param budgetLimit Positive budget limit in the reporting currency.
 * @param warningThreshold Warning threshold from zero through one.
 * @returns Within budget, Near limit, or Over budget.
 */
export function budgetStatus(convertedQuote: number, budgetLimit: number, warningThreshold: number): string {
  if (!Number.isFinite(convertedQuote) || convertedQuote < 0) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "The converted quote must be a nonnegative number."
    );
  }
  if (!Number.isFinite(budgetLimit) || budgetLimit <= 0) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "The budget limit must be a positive number."
    );
  }
  if (!Number.isFinite(warningThreshold) || warningThreshold < 0 || warningThreshold > 1) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "The warning threshold must be from zero through one."
    );
  }

  if (convertedQuote > budgetLimit) {
    return "Over budget";
  }
  if (convertedQuote >= budgetLimit * (1 - warningThreshold)) {
    return "Near limit";
  }
  return "Within budget";
}

function normalizeFunctionCurrency(value: string, label: string): string {
  if (typeof value !== "string") {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      `The ${label} currency must be a three-letter code.`
    );
  }

  const currency = value.trim().toUpperCase();
  if (!FUNCTION_CURRENCY_CODE_PATTERN.test(currency)) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      `The ${label} currency must be a three-letter code.`
    );
  }
  return currency;
}
