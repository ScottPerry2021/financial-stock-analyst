import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

def get_stock_data(ticker, period="1y", interval="1d"):
    stock = yf.Ticker(ticker)
    hist = stock.history(period=period, interval=interval)
    return stock, hist

def analyze_stock(stock, hist):
    info = stock.info
    pe_ratio = info.get('trailingPE')
    eps = info.get('trailingEps')
    market_cap = info.get('marketCap')
    beta = info.get('beta')

    # Technical indicators
    hist["50_MA"] = hist["Close"].rolling(window=50).mean()
    hist["200_MA"] = hist["Close"].rolling(window=200).mean()
    hist["Daily_Return"] = hist["Close"].pct_change()
    hist["Volatility"] = hist["Daily_Return"].rolling(window=20).std() * (252 ** 0.5)

    analysis = {
        "Ticker": stock.ticker,
        "Current Price": hist["Close"][-1],
        "Market Cap": market_cap,
        "P/E Ratio": pe_ratio,
        "EPS": eps,
        "Beta": beta,
        "50-Day MA": hist["50_MA"][-1],
        "200-Day MA": hist["200_MA"][-1],
        "Volatility (20-day)": hist["Volatility"][-1]
    }

    return analysis, hist

def export_to_excel(ticker, analysis, hist):
    output_file = f"{ticker}_financial_analysis.xlsx"
    with pd.ExcelWriter(output_file) as writer:
        pd.DataFrame([analysis]).to_excel(writer, sheet_name="Summary", index=False)
        hist.to_excel(writer, sheet_name="Historical Data")
    print(f"Report exported to {output_file}")

def main():
    ticker = input("Enter the stock ticker symbol (e.g., AAPL): ").upper()
    stock, hist = get_stock_data(ticker)
    analysis, hist = analyze_stock(stock, hist)
    export_to_excel(ticker, analysis, hist)

if __name__ == "__main__":
    main()
