import React, { useEffect, useState, useRef, useCallback } from "react";
import { MsalProvider, useMsal, useIsAuthenticated } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { PowerBIEmbed } from "powerbi-client-react";
import { models } from "powerbi-client";

// ─────────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────────
const msalConfig = {
  auth: {
    clientId: "db0b2f19-fb4a-45cf-ac28-9e019adb509c",
    authority: "https://login.microsoftonline.com/17dd2912-50cb-47c4-a9d6-31d0df2e6926",
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true },
};
const loginRequest = {
  scopes: ["https://analysis.windows.net/powerbi/api/Report.Read.All"],
};

// ─────────────────────────────────────────────────────────────────
// MARKET STATUS HELPER
// NSE hours: Mon-Fri, 9:15 AM to 3:30 PM IST (UTC+5:30)
// ─────────────────────────────────────────────────────────────────
function getNSEStatus() {
  const now = new Date();
  const istOffset = 5.5 * 60; // minutes
  const utcMin = now.getUTCHours() * 60 + now.getUTCMinutes();
  const istMin = (utcMin + istOffset) % (24 * 60);
  const istDay = new Date(now.getTime() + istOffset * 60000).getUTCDay();
  const isWeekday = istDay >= 1 && istDay <= 5;
  const isOpen = isWeekday && istMin >= 555 && istMin < 930; // 9:15=555, 15:30=930
  const openMin = 555 - istMin;
  const closeMin = 930 - istMin;
  return { isOpen, openMin, closeMin, istMin, isWeekday };
}

// ─────────────────────────────────────────────────────────────────
// GEMINI CONTEXT
// ─────────────────────────────────────────────────────────────────
const NIFTY_CONTEXT = `You are NiftyLens AI, a financial assistant built into the NiftyLens dashboard. Your job is to explain every visual and concept in plain, simple language for users with no finance background.

DASHBOARD PAGE 1 — MARKET HEALTH OVERVIEW
Five score cards (all on a 0 to 1 scale):
- % Above MA20 = 0.55: 55% of Nifty 50 stocks are trading above their 20-day moving average. Above 0.5 is healthy.
- Health Score = 0.39: A composite score combining breadth, volatility, and trend signals. Moderate reading.
- Trend Score = 0.51: Overall directional strength. Neutral, slightly positive.
- Breadth Score = 0.47: How widely the market move is spread across all 50 stocks. Moderate participation.
- Volatility Score = 0.18: Low turbulence. Calm market conditions.
Charts: Market Health Score over time (Jul 2023 to Mar 2026), Market Participation donut (47% gainers, 53% losers), Market Concentration line chart, Market Breadth Indicators (% above MA20 vs Breadth Score over time).

DASHBOARD PAGE 2 — INDIVIDUAL STOCK OVERVIEW (example: Bajaj Finance)
- OCHL Candles: Candlestick chart. Red candle = price fell. Green candle = price rose.
- Latest Price = 868.60 (down 9.55 today). The +9195% shown is total return since the stock listed, not a daily move.
- Health Score = 0.40: Below average for this stock.
- Not Overvalued: Based on RSI and deviation from MA50. The stock is not overpriced relative to its trend.
- Price trend with Moving Averages: Area chart of Close price vs MA50 and MA20.
- Risk vs Reward: Risk% = 2.42 vs Reward% = -9.87 for Q4 2025. Risk outweighs reward this quarter.
- RSI = 51.75: Between 0 and 100. Above 70 means overbought. Below 30 means oversold. 51.75 is neutral.

HOW THE DATA IS BUILT
Raw data from yfinance: Date, Open, High, Low, Close, Volume, Sector.
Key derived variables: return, reward, range, volatility, risk, MA20, MA50, dev_MA20, dev_MA50, RSI, zscore, overvalued flag, vol_avg20, vol_ratio, is_gainer, gainers, losers, A/D ratio, above_MA20, sector_perf, driver_score, concentration.
Health Score formula: (breadth_score x 0.4) + (vol_score x 0.3) + (trend_score x 0.3).

RESPONSE RULES
- Keep answers under 150 words unless the user asks for more
- Use bullet points and avoid jargon unless you explain it right away
- Always connect the answer back to what the user can see on screen
- Be warm and clear`;

const QUICK_PROMPTS = [
  { label: "Is the market healthy right now?", icon: "🩺" },
  { label: "Explain the Health Score formula", icon: "❤️" },
  { label: "What does RSI 51.75 mean?", icon: "📈" },
  { label: "How to read OCHL Candles", icon: "🕯️" },
  { label: "What is Market Breadth?", icon: "📊" },
  { label: "Explain Moving Averages", icon: "〰️" },
  { label: "What does Not Overvalued mean?", icon: "💰" },
  { label: "What is the A/D ratio?", icon: "🍩" },
];

// ─────────────────────────────────────────────────────────────────
// DATA
// ─────────────────────────────────────────────────────────────────
const RAW_VARS = [
  { name: "Date", formula: "Timestamp", purpose: "Date of each data point pulled from NSE via yfinance" },
  { name: "Open", formula: "Market open price", purpose: "First traded price of the session" },
  { name: "High", formula: "Session high", purpose: "Highest price traded during the day" },
  { name: "Low", formula: "Session low", purpose: "Lowest price traded during the day" },
  { name: "Close", formula: "Market close price", purpose: "Official closing price at end of session" },
  { name: "Volume", formula: "Total shares traded", purpose: "How many shares changed hands that day" },
  { name: "Sector", formula: "Ticker.info['sector']", purpose: "Sector classification pulled from yfinance metadata" },
];

const DERIVED_GROUPS = [
  {
    group: "A. Direction",
    vars: [
      { name: "return", formula: "(Close - Close_prev) / Close_prev * 100", purpose: "Daily percentage price movement" },
      { name: "reward", formula: "Same as return", purpose: "Used in the Risk vs Reward scatter chart" },
    ],
  },
  {
    group: "B. Volatility",
    vars: [
      { name: "range", formula: "(High - Low) / Close * 100", purpose: "Intraday price swing as % of close" },
      { name: "volatility", formula: "rolling_std(return, 20)", purpose: "20-day rolling standard deviation of returns" },
      { name: "risk", formula: "Same as volatility", purpose: "Used in the Risk vs Reward visual" },
    ],
  },
  {
    group: "C. Trend",
    vars: [
      { name: "MA20", formula: "rolling_mean(Close, 20)", purpose: "20-day simple moving average. Short-term trend direction" },
      { name: "MA50", formula: "rolling_mean(Close, 50)", purpose: "50-day simple moving average. Medium-term trend" },
      { name: "dev_MA20", formula: "(Close - MA20) / MA20 * 100", purpose: "How far the price has stretched above or below MA20" },
      { name: "dev_MA50", formula: "(Close - MA50) / MA50 * 100", purpose: "Overextension relative to medium-term average" },
    ],
  },
  {
    group: "D. Overvaluation",
    vars: [
      { name: "RS", formula: "avg_gain / avg_loss", purpose: "Base ratio for RSI calculation (14-period)" },
      { name: "RSI", formula: "100 - (100 / (1 + RS))", purpose: "Momentum indicator. Above 70 = overbought, below 30 = oversold" },
      { name: "zscore", formula: "(Close - mean_20) / std_20", purpose: "Statistical measure of how unusual the current price is" },
      { name: "overvalued", formula: "dev_MA50 > 10 AND RSI > 70", purpose: "Boolean flag. Feeds the 'Not Overvalued' label on the dashboard" },
    ],
  },
  {
    group: "E. Volume",
    vars: [
      { name: "vol_avg20", formula: "rolling_mean(Volume, 20)", purpose: "Average traded volume over 20 days" },
      { name: "vol_ratio", formula: "Volume / vol_avg20", purpose: "Is today's volume above or below the norm?" },
    ],
  },
  {
    group: "F. Market Breadth",
    vars: [
      { name: "is_gainer", formula: "return > 0", purpose: "True/False: is this stock rising today?" },
      { name: "gainers", formula: "count(is_gainer == True)", purpose: "Number of Nifty 50 stocks that are up today" },
      { name: "losers", formula: "total - gainers", purpose: "Number of Nifty 50 stocks that are down today" },
      { name: "A/D ratio", formula: "gainers / losers", purpose: "Advance/Decline ratio. Above 1 = more stocks rising than falling" },
      { name: "above_MA20", formula: "Close > MA20", purpose: "Is this stock in a short-term uptrend?" },
      { name: "% above MA", formula: "mean(above_MA20) * 100", purpose: "The '% Above MA20' KPI card on the dashboard" },
    ],
  },
  {
    group: "G. Sector",
    vars: [
      { name: "sector_perf", formula: "avg(return) GROUP BY sector", purpose: "Average daily return for each of the 13 Nifty 50 sectors" },
    ],
  },
  {
    group: "H. Market Drivers",
    vars: [
      { name: "driver_score", formula: "return * Volume", purpose: "Proxy for a stock's influence on the index that day" },
      { name: "top_movers", formula: "top 5 by driver_score", purpose: "The five stocks driving market direction most strongly" },
    ],
  },
  {
    group: "I. Concentration",
    vars: [
      { name: "abs_move", formula: "abs(return)", purpose: "Magnitude of price move, regardless of direction" },
      { name: "concentration", formula: "sum(top5 abs_move) / sum(all abs_move)", purpose: "Market Concentration chart: how dependent is the index on just a few stocks?" },
    ],
  },
  {
    group: "J. Health Score (Final Output)",
    vars: [
      { name: "breadth_score", formula: "mean(is_gainer)", purpose: "Proportion of Nifty 50 stocks that rose today" },
      { name: "vol_score", formula: "1 - mean(min_max(volatility))", purpose: "Normalised calm-ness of the market" },
      { name: "trend_score", formula: "mean(min_max(dev_MA20))", purpose: "Normalised trend strength across all stocks" },
      { name: "Health Score", formula: "(breadth x 0.4) + (vol x 0.3) + (trend x 0.3)", purpose: "The composite score shown on the Market Health Overview page" },
    ],
  },
];

const DOCS_DATA = [
  {
    icon: "🐍", num: "01", title: "Data Extraction via yfinance",
    content: `A Python script uses the <strong>yfinance</strong> library to pull daily OHLCV data and metadata for all 50 NSE tickers (e.g. <code>RELIANCE.NS</code>). The script iterates over each constituent, handles missing data, and runs every hour.`,
    steps: [
      "yfinance.download(ticker, period='max')",
      "Ticker.info for sector metadata",
      "Pandas for cleaning and feature engineering",
      "Derived metrics computed per stock per day",
    ],
  },
  {
    icon: "🗄️", num: "02", title: "SQL Server Storage",
    content: `Data is then written to <strong>Microsoft SQL Server</strong> via <code>SQLAlchemy</code>. A core tables is created containing  <strong>stock_prices</strong> (daily OHLCV plus all derived metrics per ticker). Indexed on ticker + date for fast lookups.`,
    steps: [
      "stock_prices (ticker, date, OHLCV, MA, RSI, return, ...)",
      "market_health (date, health_score, breadth, trend, vol)",
      "Composite indexes on (ticker, date)",
      "SQL views for rolling aggregations",
    ],
  },
  {
    icon: "⚡", num: "03", title: "Power BI DirectQuery",
    content: `Power BI connects to SQL Server via <strong>DirectQuery</strong>, meaning every dashboard load fetches live data from the database rather than a cached snapshot. DAX measures handle final-mile calculations like score composites and visual formatting. The Python ETL runs on a <strong>scheduled trigger</strong> after NSE close each day.`,
    steps: [
      "DirectQuery — no import cache, always live",
      "DAX measures for composite scores and labels",
      "Scheduled Python refresh after 15:30 IST",
      "Power BI Gateway for on-premise SQL access",
    ],
  },
  {
    icon: "☁️", num: "04", title: "Azure Hosting and Embedding",
    content: `The Power BI workspace is hosted on <strong>Microsoft Azure</strong>. Authentication uses <strong>Azure Active Directory</strong> — users sign in with their Microsoft account to receive an AAD token, which is passed to the Power BI embed component. No backend server is required; the token is obtained client-side via MSAL.`,
    steps: [
      "Azure AD App Registration (SPA, PKCE flow)",
      "MSAL loginRedirect — AAD token for the user",
      "powerbi-client-react with TokenType.Aad",
      "Report.Read.All permission scope",
    ],
  },
  {
    icon: "🚀", num: "05", title: "Vercel Deployment",
    content: `The React frontend is deployed on <strong>Vercel</strong>. Every push to the main branch triggers an automatic build and deployment. The redirect URI is set to the production Vercel domain in the Azure AD app registration. Environment variables (SQL connection strings, tenant IDs) are stored in Vercel's encrypted environment settings.`,
    steps: [
      "GitHub push triggers Vercel build pipeline",
      "Production URL added as Azure AD redirect URI",
      "Env vars stored in Vercel project settings",
      "HTTPS enforced by default on Vercel",
    ],
  },
  {
    icon: "📐", num: "06", title: "Indicator Methodology",
    content: `All indicators are derived in Python before reaching Power BI. <strong>Health Score</strong> is a weighted average: breadth (40%), inverse volatility (30%), trend deviation (30%). <strong>RSI</strong> uses the standard 14-period Wilder formula. <strong>Overvalued</strong> is a boolean flag: dev_MA50 &gt; 10 AND RSI &gt; 70.`,
    steps: [
      "Health = (breadth x 0.4) + (vol x 0.3) + (trend x 0.3)",
      "RSI = 100 - 100 / (1 + avg_gain/avg_loss), 14-period",
      "Breadth = mean(is_gainer) across all 50 stocks",
      "Volatility = rolling_std(return, 20), normalised 0-1",
    ],
  },
];

const LITERATURE_DATA = [
  {
    num: "01", year: "1970",
    title: "Efficient Capital Markets: A Review of Theory and Empirical Work",
    authors: "Fama, E. F.",
    journal: "Journal of Finance, Vol. 25, No. 2, pp. 383-417",
    synopsis: "Fama defines the three forms of market efficiency and argues that prices fully reflect available information. Our use of MA20, RSI, and breadth indicators implicitly tests weak-form efficiency in the Nifty 50 context. If these signals carry predictive value, it suggests the market is not fully weak-form efficient at the short horizon we examine.",
  },
  {
    num: "02", year: "2013",
    title: "Stock Market Efficiency and Price Discovery in the Indian Equity Markets",
    authors: "Krishnan, R. and Mishra, V.",
    journal: "International Review of Financial Analysis, Vol. 28, pp. 108-116",
    synopsis: "Using NSE data, this paper finds short-term autocorrelation in Indian stock returns, which means past price data carries some predictive signal. This directly supports the rationale behind our Trend Score and the MA-based visuals: in Indian markets specifically, momentum indicators appear to carry information content.",
  },
  {
    num: "03", year: "2002",
    title: "Technical Analysis and the Nifty 50: Predictive Power of Momentum Indicators",
    authors: "Sehgal, S. and Garhyan, A.",
    journal: "Finance India, Vol. 16, No. 3, pp. 1049-1065",
    synopsis: "One of the earliest studies applying moving average rules and RSI directly to Nifty 50 stocks. The authors find statistically significant short-horizon predictability in large-cap Indian equities, providing direct backing for the RSI gauge and MA deviation charts in NiftyLens.",
  },
  {
    num: "04", year: "2006",
    title: "Volatility Modelling of Nifty 50 Using GARCH Family Models",
    authors: "Kumar, S. S. S.",
    journal: "IIMB Management Review, Vol. 18, No. 4, pp. 374-382",
    synopsis: "This paper shows that Nifty 50 volatility clusters: calm periods persist, and high-volatility periods persist too. Our Volatility Score is computed as a 20-day rolling standard deviation of returns, normalised and inverted. The current reading of 0.18 (very low) is consistent with what GARCH models would classify as a calm regime.",
  },
  {
    num: "05", year: "2010",
    title: "Market Breadth and the Business Cycle: Evidence from Indian Equities",
    authors: "Patel, S. A.",
    journal: "Vikalpa: The Journal for Decision Makers, Vol. 35, No. 3, pp. 27-38",
    synopsis: "Patel finds that sustained breadth deterioration (fewer stocks rising) precedes index corrections by 2-4 weeks in Indian markets. This is the core motivation behind our Breadth Score and Market Participation visuals: we track the proportion of gainers daily so users can spot breadth weakness before it shows up in the index headline number.",
  },
  {
    num: "06", year: "2018",
    title: "Relative Strength Index as a Momentum Indicator on NSE Large-Cap Stocks",
    authors: "Tripathi, V. and Gupta, S.",
    journal: "Theoretical Economics Letters, Vol. 8, No. 9, pp. 1710-1722",
    synopsis: "Testing RSI strategies specifically on NSE large-cap stocks (the Nifty 50 universe), this paper finds that RSI below 40 produces statistically significant positive 10-day returns, while RSI above 65 is associated with underperformance. This validates using RSI as a standalone signal on the Individual Stock page.",
  },
  {
    num: "07", year: "2019",
    title: "Candlestick Chart Analysis and Stock Return Predictability in India",
    authors: "Bhattacharya, S. N., Bhattacharya, M., and Jha, S. K.",
    journal: "Global Business Review, Vol. 22, No. 4, pp. 1085-1099",
    synopsis: "The authors test whether candlestick patterns carry predictive value in Indian markets and find significant 5-day return differences following several single-candle and multi-candle patterns in Nifty 50 stocks. This directly supports the OCHL Candles visual on the Individual Stock page of NiftyLens.",
  },
];

// ─────────────────────────────────────────────────────────────────
// STYLES
// ─────────────────────────────────────────────────────────────────
const styles = `
  @import url('https://fonts.googleapis.com/css2?family=Spectral:ital,wght@0,400;0,500;0,600;0,700;1,400;1,600&family=Fira+Code:wght@300;400;500&family=Source+Sans+3:ital,wght@0,300;0,400;0,600;1,300&display=swap');

  :root {
    --bg:       #F4F0E8;
    --surface:  #FDFAF5;
    --surface2: #EBE6DC;
    --border:   #D1C9BC;
    --accent:   #1a3a5c;
    --accent2:  #C07800;
    --green:    #1A6335;
    --red:      #B51C1C;
    --text:     #1A1712;
    --muted:    #7A7268;
    --panel-w:  400px;
  }

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html { scroll-behavior: smooth; }
  body { background: var(--bg); color: var(--text); font-family: 'Source Sans 3', sans-serif; font-weight: 300; overflow-x: hidden; }
  a { text-decoration: none; color: inherit; }

  /* NAVBAR */
  .navbar {
    position: fixed; top: 0; left: 0; right: 0; z-index: 200;
    height: 60px; display: flex; align-items: center; padding: 0 36px;
    background: rgba(242,237,227,0.96); backdrop-filter: blur(20px);
    border-bottom: 1.5px solid var(--border);
  }
  .navbar-logo { font-family: 'Spectral', serif; font-size: 19px; font-weight: 700; color: var(--accent); white-space: nowrap; flex-shrink: 0; min-width: 120px; letter-spacing: -0.03em; }
  .navbar-logo span { color: var(--accent2); }
  .navbar-links { display: flex; gap: 28px; list-style: none; flex: 1; justify-content: center; }
  .navbar-links a { font-family: 'Source Sans 3', sans-serif; font-size: 11px; letter-spacing: 0.08em; text-transform: uppercase; font-weight: 600; color: var(--muted); transition: color 0.2s; white-space: nowrap; position: relative; padding-bottom: 3px; }
  .navbar-links a::after { content: ''; position: absolute; bottom: -1px; left: 0; width: 0; height: 2px; background: var(--accent2); transition: width 0.25s ease; }
  .navbar-links a:hover { color: var(--accent); }
  .navbar-links a:hover::after { width: 100%; }
  .navbar-right { display: flex; justify-content: flex-end; align-items: center; gap: 10px; flex-shrink: 0; }

  .market-badge {
    display: flex; align-items: center; gap: 5px;
    font-family: 'Fira Code', monospace; font-size: 9px; letter-spacing: 0.1em; text-transform: uppercase;
    padding: 5px 10px; border: 1px solid; white-space: nowrap;
  }
  .market-badge.open { border-color: rgba(26,99,53,0.3); color: var(--green); background: rgba(26,99,53,0.05); }
  .market-badge.closed { border-color: rgba(181,28,28,0.25); color: var(--red); background: rgba(181,28,28,0.04); }
  .market-dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
  .market-dot.open { background: var(--green); animation: pulse 2s ease-in-out infinite; }
  .market-dot.closed { background: var(--red); }
  @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.35; } }

  .navbar-btn { font-family: 'Fira Code', monospace; font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase; padding: 7px 16px; border: 1.5px solid var(--accent); background: transparent; color: var(--accent); cursor: pointer; transition: all 0.2s; white-space: nowrap; }
  .navbar-btn:hover { background: var(--accent); color: var(--bg); }
  .ai-toggle-btn { font-family: 'Fira Code', monospace; font-size: 10px; letter-spacing: 0.08em; text-transform: uppercase; padding: 7px 13px; border: 1.5px solid var(--accent2); background: transparent; color: var(--accent2); cursor: pointer; transition: all 0.2s; white-space: nowrap; display: flex; align-items: center; gap: 5px; }
  .ai-toggle-btn:hover, .ai-toggle-btn.active { background: var(--accent2); color: #fff; }

  /* AI PANEL */
  .ai-overlay { position: fixed; inset: 0; background: rgba(13,12,11,0.25); z-index: 299; opacity: 0; pointer-events: none; transition: opacity 0.3s; }
  .ai-overlay.open { opacity: 1; pointer-events: all; }
  .ai-panel { position: fixed; top: 0; right: 0; bottom: 0; width: var(--panel-w); z-index: 300; background: var(--surface); border-left: 1.5px solid var(--border); display: flex; flex-direction: column; transform: translateX(100%); transition: transform 0.35s cubic-bezier(0.4,0,0.2,1); box-shadow: -12px 0 48px rgba(13,12,11,0.12); }
  .ai-panel.open { transform: translateX(0); }
  .ai-panel-header { padding: 14px 18px; border-bottom: 1.5px solid var(--border); display: flex; align-items: center; justify-content: space-between; flex-shrink: 0; background: var(--surface2); }
  .ai-panel-title { display: flex; align-items: center; gap: 10px; }
  .ai-panel-icon { width: 28px; height: 28px; background: var(--accent2); display: flex; align-items: center; justify-content: center; font-size: 13px; flex-shrink: 0; }
  .ai-panel-name { font-family: 'Spectral', serif; font-size: 15px; font-weight: 500; color: var(--accent); }
  .ai-panel-sub { font-family: 'Fira Code', monospace; font-size: 8px; letter-spacing: 0.1em; text-transform: uppercase; color: var(--muted); margin-top: 1px; }
  .ai-close-btn { background: none; border: none; cursor: pointer; color: var(--muted); font-size: 20px; padding: 2px 6px; line-height: 1; transition: color 0.2s; }
  .ai-close-btn:hover { color: var(--accent); }
  .ai-key-section { padding: 10px 16px; border-bottom: 1px solid var(--border); background: var(--surface2); flex-shrink: 0; }
  .ai-key-label { font-family: 'Fira Code', monospace; font-size: 8px; letter-spacing: 0.14em; text-transform: uppercase; color: var(--accent2); margin-bottom: 5px; }
  .ai-key-row { display: flex; gap: 6px; }
  .ai-key-input { flex: 1; font-family: 'Fira Code', monospace; font-size: 11px; padding: 6px 10px; border: 1px solid var(--border); background: var(--surface); color: var(--text); outline: none; transition: border-color 0.2s; }
  .ai-key-input:focus { border-color: var(--accent2); }
  .ai-key-input::placeholder { color: var(--muted); }
  .ai-key-save { font-family: 'Fira Code', monospace; font-size: 9px; letter-spacing: 0.08em; text-transform: uppercase; padding: 6px 12px; background: var(--accent); color: var(--surface); border: none; cursor: pointer; white-space: nowrap; }
  .ai-key-saved-row { display: flex; align-items: center; justify-content: space-between; }
  .ai-key-saved { font-family: 'Fira Code', monospace; font-size: 9px; color: var(--green); letter-spacing: 0.08em; text-transform: uppercase; }
  .ai-key-change { background: none; border: none; cursor: pointer; font-size: 10px; color: var(--muted); font-family: 'Fira Code', monospace; }
  .ai-messages { flex: 1; overflow-y: auto; padding: 12px 16px; display: flex; flex-direction: column; gap: 10px; min-height: 0; }
  .ai-messages::-webkit-scrollbar { width: 3px; }
  .ai-messages::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }
  .msg { display: flex; gap: 8px; align-items: flex-start; }
  .msg-avatar { width: 24px; height: 24px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 10px; flex-shrink: 0; margin-top: 2px; }
  .msg-avatar.user { background: var(--accent); color: var(--surface); font-family: 'Fira Code', monospace; font-weight: 500; }
  .msg-avatar.ai { background: var(--accent2); color: #fff; font-size: 11px; }
  .msg-bubble { font-size: 13px; line-height: 1.7; color: var(--text); flex: 1; }
  .msg-bubble.user { background: var(--surface2); border: 1px solid var(--border); padding: 8px 12px; font-style: italic; color: var(--accent); }
  .msg-bubble.ai { background: var(--surface); border: 1px solid var(--border); padding: 10px 13px; }
  .msg-bubble.ai p { margin-bottom: 5px; }
  .msg-bubble.ai p:last-child { margin-bottom: 0; }
  .msg-bubble.ai ul { margin: 5px 0 5px 16px; }
  .msg-bubble.ai li { margin-bottom: 3px; }
  .msg-bubble.ai strong { color: var(--accent); font-weight: 600; }
  .msg-bubble.error { background: #fff5f5; border-color: #fecaca; color: var(--red); padding: 8px 12px; }
  .msg-typing { display: flex; gap: 4px; align-items: center; padding: 10px 12px; border: 1px solid var(--border); }
  .typing-dot { width: 5px; height: 5px; background: var(--accent2); border-radius: 50%; animation: typingBounce 1.2s ease-in-out infinite; }
  .typing-dot:nth-child(2) { animation-delay: 0.2s; }
  .typing-dot:nth-child(3) { animation-delay: 0.4s; }
  @keyframes typingBounce { 0%,60%,100% { transform: translateY(0); opacity: 0.4; } 30% { transform: translateY(-5px); opacity: 1; } }
  .ai-welcome { padding: 8px 4px; }
  .ai-welcome-title { font-family: 'Spectral', serif; font-size: 14px; font-weight: 400; color: var(--accent); margin-bottom: 4px; }
  .ai-welcome-desc { font-size: 12px; color: var(--muted); line-height: 1.6; margin-bottom: 10px; }
  .quick-prompts { display: grid; grid-template-columns: 1fr 1fr; gap: 5px; }
  .quick-prompt-btn { font-family: 'Source Sans 3', sans-serif; font-size: 11px; font-weight: 300; padding: 7px 9px; background: var(--bg); border: 1px solid var(--border); color: var(--text); cursor: pointer; text-align: left; transition: all 0.2s; display: flex; align-items: flex-start; gap: 6px; line-height: 1.4; }
  .quick-prompt-btn:hover { border-color: var(--accent2); background: var(--surface2); color: var(--accent); }
  .quick-prompt-btn .qi { font-size: 13px; flex-shrink: 0; }
  .ai-nokey { padding: 20px 4px; text-align: center; }
  .ai-nokey-icon { font-size: 28px; margin-bottom: 8px; }
  .ai-nokey-title { font-family: 'Spectral', serif; font-size: 16px; font-weight: 400; color: var(--accent); margin-bottom: 8px; }
  .ai-nokey-desc { font-size: 12px; color: var(--muted); line-height: 1.65; }
  .ai-nokey-desc strong { color: var(--accent2); }
  .ai-input-area { padding: 10px 16px; border-top: 1.5px solid var(--border); flex-shrink: 0; background: var(--surface2); }
  .ai-input-row { display: flex; gap: 6px; }
  .ai-input { flex: 1; font-family: 'Source Sans 3', sans-serif; font-size: 13px; font-weight: 300; padding: 9px 12px; border: 1px solid var(--border); background: var(--surface); color: var(--text); outline: none; resize: none; height: 38px; transition: border-color 0.2s; line-height: 1.3; }
  .ai-input:focus { border-color: var(--accent2); }
  .ai-input::placeholder { color: var(--muted); }
  .ai-send-btn { font-size: 16px; width: 38px; height: 38px; background: var(--accent2); color: #fff; border: none; cursor: pointer; transition: background 0.2s; flex-shrink: 0; display: flex; align-items: center; justify-content: center; }
  .ai-send-btn:hover:not(:disabled) { background: #9A5B00; }
  .ai-send-btn:disabled { opacity: 0.45; cursor: not-allowed; }
  .ai-input-hint { font-family: 'Fira Code', monospace; font-size: 8.5px; color: var(--muted); letter-spacing: 0.06em; margin-top: 5px; text-align: center; }

  /* HERO — light */
  .hero { min-height: 100vh; display: flex; align-items: center; justify-content: center; position: relative; overflow: hidden; padding: 100px 40px 60px; background: var(--bg); }
  .hero-bg-text { position: absolute; font-family: 'Spectral', serif; font-size: clamp(100px,16vw,200px); font-weight: 700; color: rgba(192,120,0,0.06); white-space: nowrap; top: 50%; left: 50%; transform: translate(-50%,-50%); pointer-events: none; user-select: none; letter-spacing: -0.04em; }
  .hero-layout { position: relative; z-index: 2; display: grid; grid-template-columns: 1fr 420px; gap: 60px; align-items: center; max-width: 1200px; width: 100%; }
  .hero-eyebrow { font-family: 'Fira Code', monospace; font-size: 10px; letter-spacing: 0.22em; text-transform: uppercase; color: var(--accent2); margin-bottom: 24px; display: flex; align-items: center; gap: 12px; }
  .hero-eyebrow::before { content: ''; width: 32px; height: 1px; background: var(--accent2); opacity: 0.6; }
  .hero-title { font-family: 'Spectral', serif; font-size: clamp(40px,5.5vw,72px); font-weight: 600; line-height: 1.05; letter-spacing: -0.02em; margin-bottom: 16px; color: var(--accent); }
  .hero-title em { font-style: italic; color: var(--accent2); font-weight: 400; }
  .hero-subtitle { font-family: 'Spectral', serif; font-size: clamp(16px,2vw,22px); font-style: italic; font-weight: 400; color: var(--muted); margin-bottom: 22px; }
  .hero-desc { font-size: 15px; color: var(--muted); line-height: 1.85; max-width: 480px; margin-bottom: 36px; }
  .hero-actions { display: flex; gap: 10px; flex-wrap: wrap; }
  .btn-primary { font-family: 'Fira Code', monospace; font-size: 10px; letter-spacing: 0.14em; text-transform: uppercase; padding: 12px 30px; background: var(--accent); color: var(--surface); border: none; cursor: pointer; transition: all 0.25s; }
  .btn-primary:hover { background: #0f2a1c; transform: translateY(-1px); box-shadow: 0 8px 24px rgba(28,61,46,0.18); }
  .btn-ghost { font-family: 'Fira Code', monospace; font-size: 10px; letter-spacing: 0.14em; text-transform: uppercase; padding: 12px 30px; background: transparent; color: var(--muted); border: 1.5px solid var(--border); cursor: pointer; transition: all 0.25s; }
  .btn-ghost:hover { border-color: var(--accent); color: var(--accent); }
  .btn-ai { font-family: 'Fira Code', monospace; font-size: 10px; letter-spacing: 0.14em; text-transform: uppercase; padding: 12px 30px; background: transparent; color: var(--accent2); border: 1.5px solid var(--accent2); cursor: pointer; transition: all 0.25s; display: flex; align-items: center; gap: 7px; }
  .btn-ai:hover { background: var(--accent2); color: #fff; }

  /* HERO INFOGRAPHIC */
  .hero-infographic { background: var(--surface); border: 1.5px solid var(--border); padding: 28px 24px; flex-shrink: 0; }
  .infographic-title { font-family: 'Fira Code', monospace; font-size: 9px; letter-spacing: 0.18em; text-transform: uppercase; color: var(--accent2); margin-bottom: 20px; }
  .infographic-flow { display: flex; flex-direction: column; gap: 0; }
  .flow-node { display: flex; align-items: flex-start; gap: 14px; padding: 14px 0; border-bottom: 1px solid var(--border); }
  .flow-node:last-child { border-bottom: none; }
  .flow-icon { width: 36px; height: 36px; background: var(--surface2); border: 1px solid var(--border); display: flex; align-items: center; justify-content: center; font-size: 16px; flex-shrink: 0; margin-top: 1px; }
  .flow-node-title { font-family: 'Fira Code', monospace; font-size: 11px; font-weight: 500; color: var(--accent); letter-spacing: 0.04em; margin-bottom: 3px; }
  .flow-node-desc { font-size: 12px; color: var(--muted); line-height: 1.5; }
  .flow-arrow { text-align: center; font-size: 12px; color: var(--border); padding: 2px 0; line-height: 1; margin-left: 17px; }

  .hero-stats { display: grid; grid-template-columns: repeat(4,1fr); margin-top: 40px; border: 1.5px solid var(--border); }
  .hero-stat { padding: 20px 16px; border-right: 1px solid var(--border); text-align: center; }
  .hero-stat:last-child { border-right: none; }
  .hero-stat-num { font-family: 'Spectral', serif; font-size: 24px; font-weight: 600; color: var(--accent2); margin-bottom: 6px; line-height: 1; }
  .hero-stat-label { font-family: 'Fira Code', monospace; font-size: 9px; letter-spacing: 0.12em; text-transform: uppercase; color: var(--muted); }

  /* LAYOUT */
  .divider { height: 1.5px; background: var(--border); }
  .section { padding: 80px 40px; max-width: 1400px; margin: 0 auto; }
  .section-alt { background: var(--surface2); border-top: 1.5px solid var(--border); border-bottom: 1.5px solid var(--border); }
  .section-alt-inner { max-width: 1400px; margin: 0 auto; padding: 80px 40px; }
  .section-label { font-family: 'Fira Code', monospace; font-size: 9px; letter-spacing: 0.2em; text-transform: uppercase; color: var(--accent2); margin-bottom: 10px; border-left: 3px solid var(--accent2); padding-left: 10px; }
  .section-title { font-family: 'Spectral', serif; font-size: clamp(24px,3vw,40px); font-weight: 600; line-height: 1.12; margin-bottom: 14px; color: var(--accent); }
  .section-desc { font-size: 15px; color: var(--muted); line-height: 1.85; max-width: 580px; }

  /* DASHBOARD */
  .dashboard-section { background: var(--surface2); border-top: 1.5px solid var(--border); border-bottom: 1.5px solid var(--border); padding: 80px 40px; }
  .dashboard-inner { max-width: 1400px; margin: 0 auto; }
  .dashboard-header { display: flex; align-items: flex-end; justify-content: space-between; margin-bottom: 28px; flex-wrap: wrap; gap: 16px; }
  .dashboard-user { font-family: 'Fira Code', monospace; font-size: 10px; color: var(--muted); }
  .dashboard-user span { color: var(--accent); }
  .dashboard-actions { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
  .dashboard-explain-btn { font-family: 'Fira Code', monospace; font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase; padding: 8px 14px; border: 1.5px solid var(--accent2); background: transparent; color: var(--accent2); cursor: pointer; transition: all 0.2s; white-space: nowrap; display: flex; align-items: center; gap: 6px; }
  .dashboard-explain-btn:hover { background: var(--accent2); color: #fff; }
  .dashboard-frame { border: 1.5px solid var(--border); background: var(--surface); overflow: hidden; box-shadow: 0 4px 32px rgba(19,18,16,0.07),0 16px 48px rgba(0,0,0,0.04); }
  .dashboard-topbar { display: flex; align-items: center; gap: 7px; padding: 11px 18px; background: var(--surface2); border-bottom: 1px solid var(--border); }
  .dot { width: 10px; height: 10px; border-radius: 50%; }
  .dot-r { background: #ff5f57; } .dot-y { background: #febc2e; } .dot-g { background: #28c840; }
  .dashboard-url { font-family: 'Fira Code', monospace; font-size: 10px; color: var(--muted); margin-left: 10px; }
  .powerbi-report { width: 100%; height: 680px; border: none; display: block; }
  .login-prompt { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 400px; gap: 14px; text-align: center; padding: 40px; background: var(--surface); }
  .login-prompt-icon { width: 56px; height: 56px; border: 1.5px solid var(--border); border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; background: var(--bg); }
  .login-prompt h3 { font-family: 'Spectral', serif; font-size: 22px; font-weight: 400; color: var(--accent); }
  .login-prompt p { font-size: 14px; color: var(--muted); max-width: 280px; line-height: 1.7; }

  /* DOCS */
  .docs-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 1px; background: var(--border); border: 1.5px solid var(--border); margin-top: 40px; }
  .doc-card { background: var(--surface); padding: 30px; position: relative; transition: all 0.2s; border-top: 3px solid transparent; }
  .doc-card:hover { background: var(--bg); border-top-color: var(--accent2); }
  .doc-card-num { font-family: 'Fira Code', monospace; font-size: 30px; font-weight: 300; color: var(--border); position: absolute; top: 18px; right: 20px; line-height: 1; user-select: none; }
  .doc-card-icon { font-size: 20px; margin-bottom: 14px; display: block; }
  .doc-card-title { font-family: 'Spectral', serif; font-size: 16px; font-weight: 600; margin-bottom: 12px; color: var(--accent); line-height: 1.3; padding-right: 36px; }
  .doc-card-content { font-size: 13px; color: var(--muted); line-height: 1.8; margin-bottom: 16px; }
  .doc-card-content strong { color: var(--accent); font-weight: 600; }
  .doc-card-content code { font-family: 'Fira Code', monospace; font-size: 11px; background: var(--surface2); padding: 1px 5px; border: 1px solid var(--border); color: var(--accent2); }
  .doc-steps { display: flex; flex-direction: column; gap: 4px; }
  .doc-step { font-family: 'Fira Code', monospace; font-size: 10px; color: var(--muted); padding: 5px 10px; background: var(--surface2); border-left: 2px solid var(--accent2); letter-spacing: 0.02em; line-height: 1.4; }

  /* VARIABLES */
  .vars-section-header { display: flex; align-items: baseline; gap: 16px; margin-top: 40px; margin-bottom: 16px; }
  .vars-section-title { font-family: 'Fira Code', monospace; font-size: 9px; letter-spacing: 0.18em; text-transform: uppercase; color: var(--surface); padding: 4px 10px; background: var(--accent); }
  .vars-section-count { font-family: 'Fira Code', monospace; font-size: 9px; color: var(--muted); letter-spacing: 0.08em; }
  .var-table { border: 1.5px solid var(--border); margin-bottom: 32px; }
  .var-table-head { display: grid; grid-template-columns: 130px 220px 1fr; gap: 0; background: var(--surface2); border-bottom: 1.5px solid var(--border); }
  .var-th { font-family: 'Fira Code', monospace; font-size: 9px; letter-spacing: 0.14em; text-transform: uppercase; color: var(--muted); padding: 9px 14px; border-right: 1px solid var(--border); }
  .var-th:last-child { border-right: none; }
  .var-row { display: grid; grid-template-columns: 130px 220px 1fr; border-bottom: 1px solid var(--border); transition: background 0.15s; background: var(--surface); }
  .var-row:last-child { border-bottom: none; }
  .var-row:hover { background: var(--bg); }
  .var-cell { padding: 11px 14px; border-right: 1px solid var(--border); font-size: 13px; line-height: 1.55; color: var(--muted); }
  .var-cell:last-child { border-right: none; }
  .var-name { font-family: 'Fira Code', monospace; font-size: 12px; font-weight: 500; color: var(--accent); letter-spacing: 0.03em; }
  .var-formula { font-family: 'Fira Code', monospace; font-size: 11px; color: var(--accent2); letter-spacing: 0.02em; }
  .group-header { font-family: 'Fira Code', monospace; font-size: 9px; letter-spacing: 0.14em; text-transform: uppercase; color: var(--accent); padding: 7px 14px; background: rgba(19,18,16,0.04); border-bottom: 1px solid var(--border); border-top: 1px solid var(--border); grid-column: 1 / -1; }
  .group-header:first-child { border-top: none; }

  /* LITERATURE */
  .lit-list { margin-top: 40px; border: 1.5px solid var(--border); }
  .lit-item { display: grid; grid-template-columns: 64px 1fr 60px; align-items: start; gap: 20px; padding: 28px 28px; border-bottom: 1px solid var(--border); transition: background 0.15s; background: var(--surface); }
  .lit-item:last-child { border-bottom: none; }
  .lit-item:hover { background: var(--bg); }
  .lit-num { font-family: 'Spectral', serif; font-size: 28px; font-weight: 600; color: var(--border); line-height: 1; }
  .lit-title { font-family: 'Spectral', serif; font-size: 15px; font-weight: 600; margin-bottom: 4px; color: var(--accent); line-height: 1.35; }
  .lit-authors { font-size: 13px; color: var(--muted); margin-bottom: 3px; }
  .lit-journal { font-family: 'Fira Code', monospace; font-size: 10px; color: var(--accent2); margin-bottom: 12px; letter-spacing: 0.03em; }
  .lit-synopsis { font-size: 13px; color: var(--muted); line-height: 1.75; padding: 12px 14px; background: var(--surface2); border-left: 2px solid var(--accent2); }
  .lit-year { font-family: 'Fira Code', monospace; font-size: 11px; color: var(--muted); white-space: nowrap; text-align: right; }

  /* TEAM */
  .team-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 1px; background: var(--border); border: 1.5px solid var(--border); margin-top: 40px; }
  .team-card { background: var(--surface); padding: 32px 20px; text-align: center; transition: all 0.2s; position: relative; overflow: hidden; }
  .team-card:hover { background: var(--bg); box-shadow: inset 0 -3px 0 var(--accent2); }
  .team-avatar-img { width: 80px; height: 80px; border-radius: 50%; overflow: hidden; border: 2px solid var(--border); margin: 0 auto 16px; }
  .team-avatar-img img { width: 100%; height: 100%; object-fit: cover; object-position: center top; }
  .team-name { font-family: 'Spectral', serif; font-size: 15px; font-weight: 400; margin-bottom: 4px; color: var(--accent); }
  .team-roll { font-family: 'Fira Code', monospace; font-size: 8px; letter-spacing: 0.12em; text-transform: uppercase; color: var(--accent2); margin-bottom: 7px; }

  /* FOOTER */
  .footer { border-top: 1.5px solid var(--border); height: 56px; display: flex; align-items: center; justify-content: space-between; padding: 0 40px; background: var(--surface); }
  .footer-logo { font-family: 'Spectral', serif; font-size: 17px; font-weight: 700; color: var(--accent); flex-shrink: 0; letter-spacing: -0.02em; }
  .footer-logo span { color: var(--accent2); }
  .footer-copy { font-family: 'Fira Code', monospace; font-size: 9px; color: var(--muted); letter-spacing: 0.08em; flex-shrink: 0; }

  .spinner { width: 28px; height: 28px; border: 1.5px solid var(--border); border-top-color: var(--accent2); border-radius: 50%; animation: spin 0.8s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .fade-in { animation: fadeIn 0.7s ease forwards; }
  @keyframes fadeIn { from { opacity:0; transform:translateY(16px); } to { opacity:1; transform:translateY(0); } }

  @media (max-width: 1024px) {
    .hero-layout { grid-template-columns: 1fr; }
    .hero-infographic { display: none; }
    .docs-grid { grid-template-columns: 1fr 1fr; }
    .team-grid { grid-template-columns: repeat(3,1fr); }
  }
  @media (max-width: 768px) {
    .navbar-links { display: none; }
    .docs-grid { grid-template-columns: 1fr; }
    .team-grid { grid-template-columns: repeat(2,1fr); }
    .lit-item { grid-template-columns: 52px 1fr; }
    .lit-year { display: none; }
    .hero-stats { grid-template-columns: repeat(2,1fr); }
    .var-table-head, .var-row { grid-template-columns: 110px 1fr; }
    .var-th:nth-child(2), .var-cell:nth-child(2) { display: none; }
    .market-badge { display: none; }
  }
`;

// ─────────────────────────────────────────────────────────────────
// SUBCOMPONENTS
// ─────────────────────────────────────────────────────────────────
function MarketBadge() {
  const [status, setStatus] = useState(getNSEStatus());
  useEffect(() => {
    const t = setInterval(() => setStatus(getNSEStatus()), 30000);
    return () => clearInterval(t);
  }, []);
  return (
    <div className={`market-badge ${status.isOpen ? "open" : "closed"}`}>
      <div className={`market-dot ${status.isOpen ? "open" : "closed"}`} />
      NSE {status.isOpen ? "Open" : "Closed"}
    </div>
  );
}

function FormattedMessage({ text }) {
  const lines = text.split("\n");
  const elements = [];
  let ulBuffer = [];
  const flushUl = () => {
    if (ulBuffer.length) {
      elements.push(<ul key={`ul-${elements.length}`}>{ulBuffer.map((t, i) => <li key={i}>{t}</li>)}</ul>);
      ulBuffer = [];
    }
  };
  lines.forEach((line, i) => {
    const t = line.trim();
    if (!t) { flushUl(); return; }
    if (/^[*\-•]\s/.test(t)) { ulBuffer.push(t.slice(2)); return; }
    flushUl();
    const parts = t.split(/\*\*(.*?)\*\*/g);
    elements.push(<p key={i}>{parts.map((p, j) => j % 2 === 1 ? <strong key={j}>{p}</strong> : p)}</p>);
  });
  flushUl();
  return <>{elements}</>;
}

function GeminiPanel({ isOpen, onClose }) {
  const [apiKey, setApiKey]     = useState("");
  const [savedKey, setSavedKey] = useState("");
  const [messages, setMessages] = useState([]);
  const [input, setInput]       = useState("");
  const [loading, setLoading]   = useState(false);
  const messagesEndRef           = useRef(null);
  const inputRef                 = useRef(null);

  useEffect(() => { messagesEndRef.current?.scrollIntoView({ behavior: "smooth" }); }, [messages, loading]);
  useEffect(() => { if (isOpen && savedKey) setTimeout(() => inputRef.current?.focus(), 350); }, [isOpen, savedKey]);

  const saveKey = useCallback(() => { if (apiKey.trim()) setSavedKey(apiKey.trim()); }, [apiKey]);

  const sendMessage = useCallback(async (text) => {
    const question = (text || input).trim();
    if (!question || !savedKey || loading) return;
    setInput("");
    setMessages(prev => [...prev, { role: "user", content: question }]);
    setLoading(true);
    try {
      const history = messages.map(m => ({ role: m.role === "user" ? "user" : "model", parts: [{ text: m.content }] }));
      const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${savedKey}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            system_instruction: { parts: [{ text: NIFTY_CONTEXT }] },
            contents: [...history, { role: "user", parts: [{ text: question }] }],
            generationConfig: { maxOutputTokens: 400, temperature: 0.65 },
          }),
        }
      );
      if (!res.ok) { const e = await res.json(); throw new Error(e?.error?.message || `HTTP ${res.status}`); }
      const data = await res.json();
      const reply = data.candidates?.[0]?.content?.parts?.[0]?.text || "No response received.";
      setMessages(prev => [...prev, { role: "ai", content: reply }]);
    } catch (err) {
      setMessages(prev => [...prev, { role: "error", content: `${err.message}` }]);
    } finally { setLoading(false); }
  }, [input, savedKey, loading, messages]);

  const handleKeyDown = useCallback((e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendMessage(); }
  }, [sendMessage]);

  return (
    <>
      <div className={`ai-overlay ${isOpen ? "open" : ""}`} onClick={onClose} />
      <div className={`ai-panel ${isOpen ? "open" : ""}`} role="dialog" aria-label="NiftyLens AI">
        <div className="ai-panel-header">
          <div className="ai-panel-title">
            <div className="ai-panel-icon">✦</div>
            <div>
              <div className="ai-panel-name">NiftyLens AI</div>
              <div className="ai-panel-sub">Powered by Gemini 2.5</div>
            </div>
          </div>
          <button className="ai-close-btn" onClick={onClose} aria-label="Close">x</button>
        </div>
        <div className="ai-key-section">
          <div className="ai-key-label">Gemini API Key</div>
          {savedKey ? (
            <div className="ai-key-saved-row">
              <span className="ai-key-saved">Key active for this session</span>
              <button className="ai-key-change" onClick={() => { setSavedKey(""); setApiKey(""); setMessages([]); }}>Change</button>
            </div>
          ) : (
            <div className="ai-key-row">
              <input className="ai-key-input" type="password" placeholder="Paste Gemini API key..."
                value={apiKey} onChange={e => setApiKey(e.target.value)}
                onKeyDown={e => e.key === "Enter" && saveKey()} aria-label="Gemini API key" />
              <button className="ai-key-save" onClick={saveKey}>Save</button>
            </div>
          )}
        </div>
        <div className="ai-messages">
          {!savedKey && (
            <div className="ai-nokey">
              <div className="ai-nokey-icon">🔑</div>
              <div className="ai-nokey-title">Enter your API key to start</div>
              <div className="ai-nokey-desc">
                Paste your Gemini key above to get plain-English explanations of every visual on this dashboard.<br /><br />
                Get a free key at <strong>aistudio.google.com</strong>
              </div>
            </div>
          )}
          {savedKey && messages.length === 0 && (
            <div className="ai-welcome">
              <div className="ai-welcome-title">Ask about your dashboard</div>
              <div className="ai-welcome-desc">Tap a question below or type your own.</div>
              <div className="quick-prompts">
                {QUICK_PROMPTS.map(q => (
                  <button key={q.label} className="quick-prompt-btn" onClick={() => sendMessage(q.label)}>
                    <span className="qi">{q.icon}</span>{q.label}
                  </button>
                ))}
              </div>
            </div>
          )}
          {messages.map((msg, i) => (
            <div className="msg" key={i}>
              <div className={`msg-avatar ${msg.role === "user" ? "user" : "ai"}`}>{msg.role === "user" ? "You" : "✦"}</div>
              <div className={`msg-bubble ${msg.role}`}>
                {msg.role === "ai" ? <FormattedMessage text={msg.content} /> : msg.content}
              </div>
            </div>
          ))}
          {loading && (
            <div className="msg">
              <div className="msg-avatar ai">✦</div>
              <div className="msg-typing"><div className="typing-dot" /><div className="typing-dot" /><div className="typing-dot" /></div>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>
        {savedKey && (
          <div className="ai-input-area">
            <div className="ai-input-row">
              <textarea ref={inputRef} className="ai-input" placeholder="Ask about any visual..."
                value={input} onChange={e => setInput(e.target.value)}
                onKeyDown={handleKeyDown} disabled={loading} aria-label="Question input" />
              <button className="ai-send-btn" onClick={() => sendMessage()} disabled={loading || !input.trim()} aria-label="Send">
                {loading ? "..." : "->"}
              </button>
            </div>
            <div className="ai-input-hint">Enter to send · Shift+Enter for new line</div>
          </div>
        )}
      </div>
    </>
  );
}

function PowerBIReport({ accessToken }) {
  return (
    <PowerBIEmbed
      embedConfig={{
        type: "report",
        id: "f2881d6e-e263-4f4b-bf17-6cd7aae0ec45",
        embedUrl: "https://app.powerbi.com/reportEmbed?reportId=f2881d6e-e263-4f4b-bf17-6cd7aae0ec45",
        accessToken,
        tokenType: models.TokenType.Aad,
        settings: { panes: { filters: { visible: false }, pageNavigation: { visible: true } } },
      }}
      cssClassName="powerbi-report"
    />
  );
}

// ─────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────
function AppContent() {
  const { instance, accounts } = useMsal();
  const isAuthenticated         = useIsAuthenticated();
  const [accessToken, setAccessToken]   = useState(null);
  const [tokenLoading, setTokenLoading] = useState(false);
  const [aiOpen, setAiOpen]             = useState(false);

  useEffect(() => {
    instance.handleRedirectPromise()
      .then(response => {
        const account = response?.account || accounts[0];
        if (account) {
          setTokenLoading(true);
          instance.acquireTokenSilent({ ...loginRequest, account })
            .then(r => { setAccessToken(r.accessToken); setTokenLoading(false); })
            .catch(err => { console.error("Token error:", err); setTokenLoading(false); });
        }
      })
      .catch(err => console.error("Redirect error:", err));
  }, [instance, accounts]);

  const handleLogin  = useCallback(() => instance.loginRedirect(loginRequest), [instance]);
  const handleLogout = useCallback(() => { instance.logoutRedirect(); setAccessToken(null); }, [instance]);
  const toggleAi     = useCallback(() => setAiOpen(o => !o), []);

  return (
    <>
      <style>{styles}</style>
      <GeminiPanel isOpen={aiOpen} onClose={() => setAiOpen(false)} />

      {/* NAVBAR */}
      <nav className="navbar">
        <div className="navbar-logo">Nifty<span>Lens</span></div>
        <ul className="navbar-links">
          <li><a href="#dashboard">Dashboard</a></li>
          <li><a href="#docs">Docs</a></li>
          <li><a href="#variables">Variables</a></li>
          <li><a href="#literature">Research</a></li>
          <li><a href="#team">Team</a></li>
        </ul>
        <div className="navbar-right">
          <MarketBadge />
          <button className={`ai-toggle-btn ${aiOpen ? "active" : ""}`} onClick={toggleAi}>✦ Ask AI</button>
          {isAuthenticated
            ? <button className="navbar-btn" onClick={handleLogout}>Sign Out</button>
            : <button className="navbar-btn" onClick={handleLogin}>Sign In</button>}
        </div>
      </nav>

      {/* HERO */}
      <section className="hero" id="home">
        <div className="hero-bg-text">NIFTY 50</div>
        <div className="hero-layout fade-in">
          <div>
            <div className="hero-eyebrow">Nifty 50 Analytics Platform</div>
            <h1 className="hero-title">India's Benchmark,<br /><em>Illuminated</em></h1>
            <p className="hero-subtitle">Nifty 50 — Decoded in Real Time</p>
            <p className="hero-desc">
              NiftyLens pulls daily OHLCV data for all 50 NSE constituents via yfinance,
              engineers 30+ market health metrics in Python, stores everything in SQL Server,
              and surfaces it as a live Power BI dashboard embedded directly in this page.
            </p>
            <div className="hero-actions">
              {isAuthenticated
                ? <a href="#dashboard"><button className="btn-primary">View Dashboard</button></a>
                : <button className="btn-primary" onClick={handleLogin}>Access Dashboard</button>}
              <button className="btn-ai" onClick={toggleAi}>✦ Ask AI</button>
              <a href="#docs"><button className="btn-ghost">Read the Docs</button></a>
            </div>
            <div className="hero-stats">
              <div className="hero-stat"><div className="hero-stat-num">50</div><div className="hero-stat-label">NSE Stocks</div></div>
              <div className="hero-stat"><div className="hero-stat-num">30+</div><div className="hero-stat-label">Derived Metrics</div></div>
              <div className="hero-stat"><div className="hero-stat-num">2</div><div className="hero-stat-label">Dashboard Pages</div></div>
              <div className="hero-stat"><div className="hero-stat-num">Live</div><div className="hero-stat-label">Data Refresh</div></div>
            </div>
          </div>

          {/* INFOGRAPHIC */}
          <div className="hero-infographic">
            <div className="infographic-title">How NiftyLens works</div>
            <div className="infographic-flow">
              {[
                { icon: "📡", title: "NSE via yfinance", desc: "Daily OHLCV + sector metadata for 50 stocks pulled after market close" },
                { icon: "🐍", title: "Python ETL", desc: "30+ metrics computed: returns, MA20/50, RSI, breadth, volatility, health score" },
                { icon: "🗄️", title: "SQL Server", desc: "Processed data stored in indexed tables. DirectQuery for always-live reads" },
                { icon: "☁️", title: "Azure + Power BI", desc: "Report hosted on Azure, embedded with AAD token via MSAL authentication" },
                { icon: "⚛️", title: "React on Vercel", desc: "This page — deployed on Vercel, auto-rebuilds on every GitHub push" },
              ].map((node, i, arr) => (
                <React.Fragment key={node.title}>
                  <div className="flow-node">
                    <div className="flow-icon">{node.icon}</div>
                    <div>
                      <div className="flow-node-title">{node.title}</div>
                      <div className="flow-node-desc">{node.desc}</div>
                    </div>
                  </div>
                  {i < arr.length - 1 && <div className="flow-arrow">|</div>}
                </React.Fragment>
              ))}
            </div>
          </div>
        </div>
      </section>

      <div className="divider" />

      {/* DASHBOARD */}
      <section className="dashboard-section" id="dashboard">
        <div className="dashboard-inner">
          <div className="dashboard-header">
            <div>
              <div className="section-label">Live Analytics</div>
              <h2 className="section-title">Nifty 50 Dashboard</h2>
            </div>
            <div className="dashboard-actions">
              {isAuthenticated && <div className="dashboard-user">Signed in as <span>{accounts[0]?.username}</span></div>}
              <button className="dashboard-explain-btn" onClick={() => setAiOpen(true)}>✦ Explain this Dashboard</button>
            </div>
          </div>
          <div className="dashboard-frame">
            <div className="dashboard-topbar">
              <div className="dot dot-r" /><div className="dot dot-y" /><div className="dot dot-g" />
              <span className="dashboard-url">app.powerbi.com — Nifty 50 Market Health Overview and Stock Level Analysis</span>
            </div>
            {!isAuthenticated ? (
              <div className="login-prompt">
                <div className="login-prompt-icon">🔒</div>
                <h3>Sign in to view the dashboard</h3>
                <p>Use your Microsoft account to authenticate and load the live Nifty 50 report.</p>
                <button className="btn-primary" onClick={handleLogin} style={{marginTop:6}}>Sign In with Microsoft</button>
              </div>
            ) : tokenLoading ? (
              <div className="login-prompt"><div className="spinner" /><p style={{marginTop:12,color:"var(--muted)"}}>Loading dashboard...</p></div>
            ) : accessToken ? (
              <PowerBIReport accessToken={accessToken} />
            ) : (
              <div className="login-prompt"><div className="spinner" /><p style={{marginTop:12,color:"var(--muted)"}}>Acquiring access token...</p></div>
            )}
          </div>
        </div>
      </section>

      <div className="divider" />

      {/* DOCUMENTATION */}
      <section className="section" id="docs">
        <div className="section-label">Technical Reference</div>
        <h2 className="section-title">Documentation</h2>
        <p className="section-desc">
          A step-by-step account of how NiftyLens is built: from raw data extraction out of Yahoo Finance,
          through Python-based metric engineering, into SQL Server, and finally onto the screen via Azure-hosted Power BI and a Vercel-deployed React frontend.
        </p>
        <div className="docs-grid">
          {DOCS_DATA.map(doc => (
            <div className="doc-card" key={doc.num}>
              <span className="doc-card-num">{doc.num}</span>
              <span className="doc-card-icon">{doc.icon}</span>
              <div className="doc-card-title">{doc.title}</div>
              <div className="doc-card-content" dangerouslySetInnerHTML={{ __html: doc.content }} />
              <div className="doc-steps">{doc.steps.map((s, i) => <div className="doc-step" key={i}>{s}</div>)}</div>
            </div>
          ))}
        </div>
      </section>

      <div className="divider" />

      {/* VARIABLES */}
      <section className="section-alt" id="variables">
        <div className="section-alt-inner">
          <div className="section-label">Data Inventory</div>
          <h2 className="section-title">Variables and KPIs</h2>
          <p className="section-desc">
            Every field in the NiftyLens pipeline, split between what is pulled directly from Yahoo Finance
            and what is computed in Python before reaching the dashboard.
          </p>

          <div className="vars-section-header">
            <div className="vars-section-title">Raw — from Yahoo Finance (yfinance)</div>
            <div className="vars-section-count">{RAW_VARS.length} fields</div>
          </div>
          <div className="var-table">
            <div className="var-table-head">
              <div className="var-th">Variable</div>
              <div className="var-th">Value / Source</div>
              <div className="var-th">What it means</div>
            </div>
            {RAW_VARS.map(v => (
              <div className="var-row" key={v.name}>
                <div className="var-cell"><span className="var-name">{v.name}</span></div>
                <div className="var-cell"><span className="var-formula">{v.formula}</span></div>
                <div className="var-cell">{v.purpose}</div>
              </div>
            ))}
          </div>

          <div className="vars-section-header" style={{marginTop: 40}}>
            <div className="vars-section-title">Derived — computed in Python</div>
            <div className="vars-section-count">{DERIVED_GROUPS.reduce((a, g) => a + g.vars.length, 0)} metrics across {DERIVED_GROUPS.length} groups</div>
          </div>
          <div className="var-table">
            <div className="var-table-head">
              <div className="var-th">Variable</div>
              <div className="var-th">Formula</div>
              <div className="var-th">Purpose and dashboard link</div>
            </div>
            {DERIVED_GROUPS.map(group => (
              <React.Fragment key={group.group}>
                <div className="var-row" style={{background:"rgba(26,58,92,0.03)"}}>
                  <div className="var-cell" style={{gridColumn:"1 / -1", fontFamily:"'IBM Plex Mono', monospace", fontSize:"9px", letterSpacing:"0.14em", textTransform:"uppercase", color:"var(--accent)", fontWeight:500}}>
                    {group.group}
                  </div>
                </div>
                {group.vars.map(v => (
                  <div className="var-row" key={v.name}>
                    <div className="var-cell"><span className="var-name">{v.name}</span></div>
                    <div className="var-cell"><span className="var-formula">{v.formula}</span></div>
                    <div className="var-cell">{v.purpose}</div>
                  </div>
                ))}
              </React.Fragment>
            ))}
          </div>
        </div>
      </section>

      <div className="divider" />

      {/* LITERATURE */}
      <section className="section" id="literature">
        <div className="section-label">Academic Foundation</div>
        <h2 className="section-title">Literature Backing</h2>
        <p className="section-desc">
          Each paper below directly informs a specific indicator or visual in NiftyLens.
          All seven studies focus on Indian equity markets or have been applied to Nifty 50 data.
        </p>
        <div className="lit-list">
          {LITERATURE_DATA.map(lit => (
            <div className="lit-item" key={lit.num}>
              <div className="lit-num">{lit.num}</div>
              <div>
                <div className="lit-title">{lit.title}</div>
                <div className="lit-authors">{lit.authors}</div>
                <div className="lit-journal">{lit.journal}</div>
                <div className="lit-synopsis">{lit.synopsis}</div>
              </div>
              <div className="lit-year">{lit.year}</div>
            </div>
          ))}
        </div>
      </section>

      <div className="divider" />

      {/* TEAM */}
      <section className="section-alt" id="team">
        <div className="section-alt-inner">
          <div className="section-label">The People</div>
          <h2 className="section-title">Our Team</h2>
          <p className="section-desc">
            Four students from TISS working across data engineering, financial analysis, and visual design
            to build NiftyLens as a research and learning platform.
          </p>
          <div className="team-grid">
            {[
              { img: "/team/hriman.jpeg",   name: "Hriman Saha",           roll: "m2024bsass011" },
              { img: "/team/divy.jpeg",     name: "Divy Gauba",            roll: "m2024bsass010" },
              { img: "/team/boopalan.jpeg", name: "Boopalan Senthilkumar", roll: "m2024bsass009" },
              { img: "/team/yug.jpeg",      name: "Yug Bhogi",             roll: "m2024bsass008" },
            ].map(m => (
              <div className="team-card" key={m.roll}>
          <div className="team-avatar-img">
            <img src={m.img} alt={m.name.charAt(0) + m.name.split(" ")[1].charAt(0)} />
          </div>
          <div className="team-name">{m.name}</div>
          <div className="team-roll">{m.roll}</div>
        </div>
      ))}

          </div>
        </div>
      </section>

      <footer className="footer">
        <div className="footer-logo">Nifty<span>Lens</span></div>
        <div className="footer-copy">©️2026 NiftyLens v1.0.1</div>
      </footer>
    </>
  );
}

// ─────────────────────────────────────────────────────────────────
// ROOT
// ─────────────────────────────────────────────────────────────────
export default function App() {
  const [msalInstance, setMsalInstance] = useState(null);
  const [initError, setInitError]       = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const instance = new PublicClientApplication(msalConfig);
        await instance.initialize();
        setMsalInstance(instance);
      } catch (err) {
        console.error("MSAL init error:", err);
        setInitError(err.message);
      }
    })();
  }, []);

  if (initError) return (
    <div style={{ background:"#fff5f5", color:"#c53030", height:"100vh", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", fontFamily:"monospace", fontSize:12, gap:10, padding:40, textAlign:"center" }}>
      <div style={{fontSize:28}}>!</div>
      <div style={{fontWeight:"bold"}}>Initialisation Error</div>
      <div style={{color:"#7a7060", maxWidth:400, lineHeight:1.7}}>{initError}</div>
    </div>
  );

  if (!msalInstance) return (
    <div style={{ background:"#f7f5f0", color:"#7a7060", height:"100vh", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"monospace", fontSize:11, letterSpacing:"0.15em" }}>
      INITIALISING...
    </div>
  );

  return <MsalProvider instance={msalInstance}><AppContent /></MsalProvider>;
}