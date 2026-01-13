import ExcelJS from "exceljs";
import { NextResponse } from "next/server";
import yahooFinance from "yahoo-finance2";

export const runtime = "nodejs";

const MODULES = [
  "assetProfile",
  "incomeStatementHistory",
  "balanceSheetHistory",
  "cashflowStatementHistory",
  "defaultKeyStatistics",
  "price"
] as const;

type RawValue = { raw?: number | string } | number | string | null | undefined;

type StatementRow = Record<string, RawValue> & {
  endDate?: { raw?: number; fmt?: string } | number;
};

const numberFrom = (value: RawValue): number | null => {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return value;
  if (typeof value === "string") return Number(value);
  if (typeof value === "object" && "raw" in value) {
    const rawValue = value.raw;
    if (rawValue === null || rawValue === undefined) return null;
    return typeof rawValue === "number" ? rawValue : Number(rawValue);
  }
  return null;
};

const yearFrom = (row?: StatementRow): string => {
  if (!row?.endDate) return "N/A";
  if (typeof row.endDate === "number") return new Date(row.endDate * 1000).getFullYear().toString();
  if (row.endDate.raw) return new Date(row.endDate.raw * 1000).getFullYear().toString();
  return row.endDate.fmt ?? "N/A";
};

const applyHeaderStyle = (row: ExcelJS.Row) => {
  row.font = { bold: true };
};

const applyCurrency = (cell: ExcelJS.Cell) => {
  cell.numFmt = "$#,##0;($#,##0)";
};

const applyNumber = (cell: ExcelJS.Cell) => {
  cell.numFmt = "#,##0;(#,##0)";
};

const applyPercent = (cell: ExcelJS.Cell) => {
  cell.numFmt = "0.0%";
};

const applyInputFill = (cell: ExcelJS.Cell) => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD6E9FF" }
  };
};

const setSheetDefaults = (sheet: ExcelJS.Worksheet) => {
  sheet.views = [{ state: "frozen", ySplit: 1 }];
};

const addStatementSheet = (
  sheet: ExcelJS.Worksheet,
  statement: StatementRow[] | undefined,
  rows: {
    label: string;
    key?: string;
    formatter?: (cell: ExcelJS.Cell) => void;
    value?: (row: StatementRow) => RawValue;
  }[]
) => {
  const statementRows = statement?.slice(0, 4) ?? [];
  const years = statementRows.map((row) => yearFrom(row));
  const headerRow = sheet.addRow(["Metric", ...years]);
  applyHeaderStyle(headerRow);

  rows.forEach((item) => {
    const values = statementRows.map((row) => {
      if (item.value) return numberFrom(item.value(row));
      if (!item.key) return null;
      return numberFrom(row?.[item.key]);
    });
    const row = sheet.addRow([item.label, ...values]);
    row.eachCell((cell, colNumber) => {
      if (colNumber === 1) return;
      if (typeof cell.value === "number") {
        (item.formatter ?? applyNumber)(cell);
      }
    });
  });

  sheet.columns.forEach((col, index) => {
    col.width = index === 0 ? 28 : 16;
  });

  setSheetDefaults(sheet);
};

const average = (values: number[]): number | null => {
  if (!values.length) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
};

export async function POST(request: Request) {
  const { ticker } = await request.json();
  if (!ticker || typeof ticker !== "string") {
    return NextResponse.json({ error: "Please provide a valid ticker." }, { status: 400 });
  }

  let data;
  try {
    data = await yahooFinance.quoteSummary(ticker, { modules: [...MODULES] });
  } catch (error) {
    return NextResponse.json({ error: "Ticker not found or unavailable." }, { status: 404 });
  }

  if (!data?.price?.shortName && !data?.price?.longName) {
    return NextResponse.json({ error: "Ticker not found or unavailable." }, { status: 404 });
  }

  const assetProfile = data.assetProfile;
  const incomeHistory = data.incomeStatementHistory?.incomeStatementHistory as
    | StatementRow[]
    | undefined;
  const balanceHistory = data.balanceSheetHistory?.balanceSheetStatements as
    | StatementRow[]
    | undefined;
  const cashflowHistory = data.cashflowStatementHistory?.cashflowStatements as
    | StatementRow[]
    | undefined;
  const keyStats = data.defaultKeyStatistics;
  const price = data.price;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Financial Model Generator";
  workbook.created = new Date();

  const summarySheet = workbook.addWorksheet("Summary");
  summarySheet.addRow(["Company", price?.longName || price?.shortName || "N/A"]);
  summarySheet.addRow(["Ticker", price?.symbol || ticker.toUpperCase()]);
  summarySheet.addRow(["Sector", assetProfile?.sector || "N/A"]);
  summarySheet.addRow(["Industry", assetProfile?.industry || "N/A"]);
  summarySheet.addRow([]);
  summarySheet.addRow(["Current Price", numberFrom(price?.regularMarketPrice)]);
  summarySheet.addRow(["Shares Outstanding", numberFrom(keyStats?.sharesOutstanding)]);
  summarySheet.addRow(["Market Cap", numberFrom(price?.marketCap)]);
  summarySheet.addRow([]);
  summarySheet.addRow(["P/E", numberFrom(keyStats?.trailingPE)]);
  summarySheet.addRow(["EV/EBITDA", numberFrom(keyStats?.enterpriseToEbitda)]);
  summarySheet.addRow(["P/B", numberFrom(keyStats?.priceToBook)]);

  summarySheet.getColumn(1).width = 24;
  summarySheet.getColumn(2).width = 32;
  summarySheet.getRow(1).font = { bold: true };
  summarySheet.getRow(6).font = { bold: true };
  summarySheet.getRow(10).font = { bold: true };
  summarySheet.eachRow((row) => {
    row.getCell(1).font = { bold: true };
  });

  applyCurrency(summarySheet.getCell(\"B6\"));
  applyNumber(summarySheet.getCell(\"B7\"));
  applyCurrency(summarySheet.getCell(\"B8\"));
  [10, 11, 12].forEach((rowNumber) => {
    applyNumber(summarySheet.getCell(`B${rowNumber}`));
  });

  setSheetDefaults(summarySheet);

  const incomeSheet = workbook.addWorksheet("Income Statement");
  addStatementSheet(incomeSheet, incomeHistory, [
    { label: "Revenue", key: "totalRevenue", formatter: applyCurrency },
    { label: "Gross Profit", key: "grossProfit", formatter: applyCurrency },
    { label: "Operating Income", key: "operatingIncome", formatter: applyCurrency },
    { label: "Net Income", key: "netIncome", formatter: applyCurrency },
    { label: "EPS", key: "dilutedEPS", formatter: applyNumber }
  ]);

  const balanceSheet = workbook.addWorksheet("Balance Sheet");
  addStatementSheet(balanceSheet, balanceHistory, [
    { label: "Total Assets", key: "totalAssets", formatter: applyCurrency },
    { label: "Total Liabilities", key: "totalLiab", formatter: applyCurrency },
    { label: "Total Equity", key: "totalStockholderEquity", formatter: applyCurrency },
    { label: "Cash", key: "cash", formatter: applyCurrency },
    { label: "Debt", key: "shortLongTermDebt", formatter: applyCurrency },
    { label: "Retained Earnings", key: "retainedEarnings", formatter: applyCurrency }
  ]);

  const cashflowSheet = workbook.addWorksheet("Cash Flow Statement");
  addStatementSheet(cashflowSheet, cashflowHistory, [
    { label: "Operating Cash Flow", key: "totalCashFromOperatingActivities", formatter: applyCurrency },
    { label: "CapEx", key: "capitalExpenditures", formatter: applyCurrency },
    {
      label: "Free Cash Flow",
      formatter: applyCurrency,
      value: (row) => {
        const cfo = numberFrom(row.totalCashFromOperatingActivities);
        const capex = numberFrom(row.capitalExpenditures);
        if (cfo === null || capex === null) return null;
        return cfo - capex;
      }
    }
  ]);

  const ratiosSheet = workbook.addWorksheet("Ratios & Metrics");
  const latestIncome = incomeHistory?.[0];
  const latestBalance = balanceHistory?.[0];

  const revenue = numberFrom(latestIncome?.totalRevenue);
  const grossProfit = numberFrom(latestIncome?.grossProfit);
  const operatingIncome = numberFrom(latestIncome?.operatingIncome);
  const netIncome = numberFrom(latestIncome?.netIncome);
  const totalAssets = numberFrom(latestBalance?.totalAssets);
  const totalEquity = numberFrom(latestBalance?.totalStockholderEquity);
  const totalLiabilities = numberFrom(latestBalance?.totalLiab);
  const currentAssets = numberFrom(latestBalance?.totalCurrentAssets);
  const currentLiabilities = numberFrom(latestBalance?.totalCurrentLiabilities);
  const debt = numberFrom(latestBalance?.shortLongTermDebt);

  const revenuePrev = numberFrom(incomeHistory?.[1]?.totalRevenue);
  const epsCurrent = numberFrom(latestIncome?.dilutedEPS ?? latestIncome?.basicEPS);
  const epsPrev = numberFrom(incomeHistory?.[1]?.dilutedEPS ?? incomeHistory?.[1]?.basicEPS);

  const ratios = [
    {
      label: "Gross margin %",
      value: revenue && grossProfit ? grossProfit / revenue : null
    },
    {
      label: "Operating margin %",
      value: revenue && operatingIncome ? operatingIncome / revenue : null
    },
    {
      label: "Net margin %",
      value: revenue && netIncome ? netIncome / revenue : null
    },
    {
      label: "ROE",
      value: totalEquity && netIncome ? netIncome / totalEquity : null
    },
    {
      label: "ROA",
      value: totalAssets && netIncome ? netIncome / totalAssets : null
    },
    {
      label: "Debt / Equity",
      value: totalEquity && debt ? debt / totalEquity : null
    },
    {
      label: "Current Ratio",
      value: currentLiabilities && currentAssets ? currentAssets / currentLiabilities : null
    },
    {
      label: "Revenue growth %",
      value: revenuePrev && revenue ? (revenue - revenuePrev) / revenuePrev : null
    },
    {
      label: "EPS growth %",
      value: epsPrev && epsCurrent ? (epsCurrent - epsPrev) / epsPrev : null
    }
  ];

  const ratiosHeader = ratiosSheet.addRow(["Metric", "Value"]);
  applyHeaderStyle(ratiosHeader);

  ratios.forEach((ratio) => {
    const row = ratiosSheet.addRow([ratio.label, ratio.value]);
    const valueCell = row.getCell(2);
    if (typeof ratio.value === "number") {
      applyPercent(valueCell);
    }
  });

  ratiosSheet.getColumn(1).width = 26;
  ratiosSheet.getColumn(2).width = 18;
  setSheetDefaults(ratiosSheet);

  const assumptionsSheet = workbook.addWorksheet("Driver Assumptions");
  const revenueGrowthRates = (incomeHistory ?? [])
    .slice(0, 4)
    .map((entry, index, array) => {
      const current = numberFrom(entry.totalRevenue);
      const next = numberFrom(array[index + 1]?.totalRevenue);
      if (!current || !next) return null;
      return (current - next) / next;
    })
    .filter((value): value is number => value !== null);

  const avgRevenueGrowth = average(revenueGrowthRates);
  const grossMargin = revenue && grossProfit ? grossProfit / revenue : null;
  const operatingMargin = revenue && operatingIncome ? operatingIncome / revenue : null;
  const taxRate = (() => {
    const taxExpense = numberFrom(latestIncome?.incomeTaxExpense);
    const preTaxIncome = numberFrom(latestIncome?.incomeBeforeTax);
    if (taxExpense && preTaxIncome) return taxExpense / preTaxIncome;
    return 0.21;
  })();

  const assumptions = [
    { label: "Revenue growth %", value: avgRevenueGrowth },
    { label: "Gross margin %", value: grossMargin },
    { label: "Operating margin %", value: operatingMargin },
    { label: "Tax rate %", value: taxRate }
  ];

  const assumptionsHeader = assumptionsSheet.addRow(["Assumption", "Value"]);
  applyHeaderStyle(assumptionsHeader);

  assumptions.forEach((item) => {
    const row = assumptionsSheet.addRow([item.label, item.value]);
    const valueCell = row.getCell(2);
    if (typeof item.value === "number") {
      applyPercent(valueCell);
      applyInputFill(valueCell);
    }
  });

  assumptionsSheet.getColumn(1).width = 26;
  assumptionsSheet.getColumn(2).width = 18;
  setSheetDefaults(assumptionsSheet);

  const buffer = await workbook.xlsx.writeBuffer();

  return new NextResponse(buffer, {
    status: 200,
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename=${ticker.toUpperCase()}-financial-model.xlsx`
    }
  });
}
