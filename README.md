# NiftyLens

A market intelligence platform for the Nifty 50 index of the National Stock Exchange of India. NiftyLens pulls daily stock data via yfinance, engineers 30+ market health metrics in Python, stores everything in SQL Server, and surfaces it as a live Power BI dashboard embedded in a React web application.

---

## What it does

**Page 1 — Market Health Overview**
Tracks the overall health of the Nifty 50 on any given trading day through five composite scores: Health Score, Breadth Score, Trend Score, Volatility Score, and % Above MA20. Includes a time-series of health over the past three years, a market participation donut (gainers vs losers), market concentration, and a breadth indicators overlay.

**Page 2 — Individual Stock Overview**
Lets users drill into any of the 50 constituent stocks. Shows OCHL candlestick chart, latest price, health score, valuation label (overvalued or not), price trend with MA20 and MA50, Risk vs Reward by quarter, and RSI gauge.

---

## Tech stack

| Layer | Technology |
|---|---|
| Data source | Yahoo Finance via `yfinance` Python library |
| Data processing | Python (pandas, numpy) |
| Storage | Microsoft SQL Server |
| Analytics | Microsoft Power BI (DirectQuery) |
| Hosting | Microsoft Azure (Power BI workspace + Azure AD) |
| Authentication | MSAL.js (Azure AD SPA flow, PKCE) |
| Frontend | React (Create React App) |
| Deployment | Vercel |
| AI assistant | Google Gemini 2.5 Flash Lite |

---

## Project structure

```
powerbi-demo/
  public/
    team/              # Team member photos (jpeg)
  src/
    App.js             # Entire frontend — components, styles, data
  .env                 # Environment variables (never commit this)
  package.json
  README.md
```

---

## How the data pipeline works

```
Yahoo Finance (yfinance)
        |
        v
Python ETL script
  - Downloads OHLCV + sector metadata for all 50 NSE tickers
  - Computes derived metrics: returns, MA20/50, RSI, breadth,
    volatility, health score, overvalued flag, driver score, etc.
  - Runs daily after NSE market close (15:30 IST)
        |
        v
SQL Server
  - stock_prices table: one row per ticker per day
  - market_health table: one row per day (index-level aggregates)
  - Indexed on (ticker, date) for fast DirectQuery reads
        |
        v
Power BI (Azure, DirectQuery)
  - DAX measures for final composite scores and labels
  - Two report pages: Market Health Overview, Stock Level Analysis
  - Embedded via AAD token (no backend server needed)
        |
        v
React app (Vercel)
  - User signs in with Microsoft account via MSAL loginRedirect
  - AAD token acquired silently and passed to powerbi-client-react
  - Gemini AI panel for plain-English explanations of every visual
```

---

## Derived variables

All metrics are computed in Python before reaching Power BI.

**Direction**
- `return` = (Close - Close_prev) / Close_prev * 100
- `reward` = same as return (used in Risk vs Reward chart)

**Volatility**
- `range` = (High - Low) / Close * 100
- `volatility` = rolling_std(return, 20)
- `risk` = same as volatility

**Trend**
- `MA20` = rolling_mean(Close, 20)
- `MA50` = rolling_mean(Close, 50)
- `dev_MA20` = (Close - MA20) / MA20 * 100
- `dev_MA50` = (Close - MA50) / MA50 * 100

**Overvaluation**
- `RS` = avg_gain / avg_loss (14-period)
- `RSI` = 100 - (100 / (1 + RS))
- `zscore` = (Close - mean_20) / std_20
- `overvalued` = dev_MA50 > 10 AND RSI > 70

**Volume**
- `vol_avg20` = rolling_mean(Volume, 20)
- `vol_ratio` = Volume / vol_avg20

**Market Breadth**
- `is_gainer` = return > 0
- `gainers` = count(is_gainer == True)
- `A/D ratio` = gainers / losers
- `above_MA20` = Close > MA20
- `% above MA` = mean(above_MA20) * 100

**Health Score (final output)**
```
breadth_score = mean(is_gainer)
vol_score     = 1 - mean(min_max(volatility))
trend_score   = mean(min_max(dev_MA20))

Health Score = (breadth_score * 0.4) + (vol_score * 0.3) + (trend_score * 0.3)
```

Breadth carries the highest weight (40%) because it is the most direct signal of market participation. Volatility and trend each carry 30% as confirming signals.

---

## Setup and installation

### Prerequisites
- Node.js 18+
- A Microsoft account with Power BI Pro or Premium Per User license
- An Azure AD tenant with an app registration (SPA platform)
- A Google Gemini API key (free tier at aistudio.google.com)

### 1. Clone and install

```bash
git clone https://github.com/your-org/niftylens.git
cd niftylens
npm install
```

### 2. Azure AD app registration

1. Go to portal.azure.com and search for App Registrations
2. Click New Registration
3. Name it anything (e.g. `niftylens-app`)
4. Under Supported account types, choose Single tenant
5. Under Redirect URI, select Single-page application and enter `http://localhost:3000`
6. Click Register
7. Copy the Application (client) ID and Directory (tenant) ID
8. Go to API Permissions, click Add a permission, choose Power BI Service
9. Select Delegated permissions and add `Report.Read.All`
10. Click Grant admin consent

### 3. Update App.js

In `src/App.js`, replace the client and tenant IDs:

```js
const msalConfig = {
  auth: {
    clientId: "your-client-id-here",
    authority: "https://login.microsoftonline.com/your-tenant-id-here",
    redirectUri: window.location.origin,
  },
  ...
};
```

Also update the Power BI report ID and embed URL if you are using a different report:

```js
id: "your-report-id-here",
embedUrl: "https://app.powerbi.com/reportEmbed?reportId=your-report-id-here",
```

### 4. Add team photos

Place team photos in `public/team/`:
```
public/team/hriman.jpeg
public/team/divy.jpeg
public/team/boopalan.jpeg
public/team/yug.jpeg
public/team/avni.jpeg
```

### 5. Run locally

```bash
npm start
```

Open http://localhost:3000, click Sign In, and authenticate with your Microsoft account.

---

## Deployment on Vercel

1. Push the repo to GitHub
2. Import the project at vercel.com
3. No build configuration needed — Vercel detects Create React App automatically
4. After deployment, copy your production URL (e.g. `https://niftylens.vercel.app`)
5. Go back to your Azure AD app registration, open Authentication, and add the production URL as a second redirect URI under the SPA platform
6. Save

Every subsequent push to the main branch triggers an automatic redeploy.

---

## Using the AI assistant

The Gemini panel is accessible via the Ask AI button in the navbar or the Explain this Dashboard button next to the report. It requires a Gemini API key entered once per session (stored in memory only, never sent anywhere other than the Gemini API).

The assistant has full context of both dashboard pages including every KPI value, chart type, and derived metric. It is configured to answer in plain English for non-finance users.

To get a free key: go to aistudio.google.com, click Get API key, and create a key in a new project. The model used is `gemini-2.5-flash-lite` which has the highest free tier rate limits.

---

## NSE market hours

The navbar shows a live indicator of whether NSE is currently open or closed. The check runs client-side every 30 seconds.

- Market open: Monday to Friday, 9:15 AM to 3:30 PM IST
- Pre-open session: 9:00 AM to 9:15 AM IST (not reflected in the indicator)
- The indicator does not account for NSE holidays

---

## Academic references

The indicators and methodology in NiftyLens are grounded in published research:

1. Fama, E. F. (1970). Efficient Capital Markets. *Journal of Finance*, 25(2).
2. Krishnan, R. & Mishra, V. (2013). Stock Market Efficiency in Indian Equity Markets. *International Review of Financial Analysis*, 28.
3. Sehgal, S. & Garhyan, A. (2002). Technical Analysis and the Nifty 50. *Finance India*, 16(3).
4. Kumar, S. S. S. (2006). Volatility Modelling of Nifty 50. *IIMB Management Review*, 18(4).
5. Patel, S. A. (2010). Market Breadth and the Business Cycle. *Vikalpa*, 35(3).
6. Tripathi, V. & Gupta, S. (2018). RSI as a Momentum Indicator on NSE Stocks. *Theoretical Economics Letters*, 8(9).
7. Bhattacharya et al. (2019). Candlestick Analysis in India. *Global Business Review*, 22(4).

---

## License

This project was built for academic purposes at TISS. Not intended for commercial use or financial advice.
