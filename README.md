## Financial Model Generator

A web app that generates downloadable 3-statement financial models from live market data.

## Features

- Enter any stock ticker and get a complete financial model
- Historical financial statements (Income Statement, Balance Sheet, Cash Flow)
- Automatically calculated ratios and metrics
- Editable driver assumptions
- Professional Excel formatting
- No API key required

## Getting Started

### Prerequisites

- Node.js (download from https://nodejs.org)

### Installation

1. Clone or download this repository
2. Open a terminal in the project folder
3. Run:

```bash
npm install --legacy-peer-deps
```

4. Start the app:

```bash
npm run dev
```

5. Open your browser and go to `http://localhost:3000`

### Usage

1. Enter a stock ticker (e.g., AAPL, MSFT, GOOGL)
2. Click "Generate Model"
3. Wait for the Excel file to download
4. Open the file and review/edit as needed

## Excel Workbook Structure

| Sheet | Contents |
|-------|----------|
| Summary | Company overview, valuation multiples |
| Income Statement | Historical P&L data |
| Balance Sheet | Historical assets, liabilities, equity |
| Cash Flow Statement | Historical CFO, CapEx, FCF |
| Ratios & Metrics | Profitability, liquidity, leverage ratios |
| Driver Assumptions | Editable inputs (highlighted in blue) |

## Data Source

All financial data is sourced from Yahoo Finance. No API key required.

## Tech Stack

- Next.js 14
- ExcelJS
- yahoo-finance2
- Tailwind CSS

## License

MIT
