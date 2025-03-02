import requests
import pandas as pd
from datetime import datetime
import os
import matplotlib.pyplot as plt
import numpy as np
import base64
from io import BytesIO


COINGECKO_API_URL = "https://api.coingecko.com/api/v3"
REPORT_FILE_PATH = "Crypto_Analysis_Report.html"

def fetch_top_50_cryptos():
    """Fetch the top 50 cryptocurrencies by market capitalization from CoinGecko API."""
    try:
        endpoint = f"{COINGECKO_API_URL}/coins/markets"
        params = {
            "vs_currency": "usd",
            "order": "market_cap_desc",
            "per_page": 50,
            "page": 1,
            "sparkline": False,
            "price_change_percentage": "24h"
        }
        
        response = requests.get(endpoint, params=params)
        response.raise_for_status()  
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data from CoinGecko API: {e}")
        return None

def process_crypto_data(data):
    """Process and transform the raw cryptocurrency data."""
    if not data:
        return None
    
  
    processed_data = []
    for coin in data:
        processed_data.append({
            "Name": coin["name"],
            "Symbol": coin["symbol"].upper(),
            "Current Price (USD)": coin["current_price"],
            "Market Cap (USD)": coin["market_cap"],
            "24h Trading Volume (USD)": coin["total_volume"],
            "24h Price Change (%)": coin["price_change_percentage_24h"] if coin["price_change_percentage_24h"] else 0
        })
    
    return processed_data

def analyze_crypto_data(data):
    """Perform analysis on the cryptocurrency data."""
    if not data:
        return None
    
    df = pd.DataFrame(data)
    
   
    analysis = {
        "top_5_by_market_cap": df.nlargest(5, "Market Cap (USD)")[["Name", "Symbol", "Market Cap (USD)", "Current Price (USD)"]].to_dict("records"),
        "top_5_by_volume": df.nlargest(5, "24h Trading Volume (USD)")[["Name", "Symbol", "24h Trading Volume (USD)"]].to_dict("records"),
        "top_5_gainers": df.nlargest(5, "24h Price Change (%)")[["Name", "Symbol", "24h Price Change (%)"]].to_dict("records"),
        "top_5_losers": df.nsmallest(5, "24h Price Change (%)")[["Name", "Symbol", "24h Price Change (%)"]].to_dict("records"),
        "average_price": df["Current Price (USD)"].mean(),
        "median_price": df["Current Price (USD)"].median(),
        "highest_price_change": df.loc[df["24h Price Change (%)"].idxmax()],
        "lowest_price_change": df.loc[df["24h Price Change (%)"].idxmin()],
        "total_market_cap": df["Market Cap (USD)"].sum(),
        "total_trading_volume": df["24h Trading Volume (USD)"].sum(),
        "bitcoin_dominance": (df.loc[df["Name"] == "Bitcoin", "Market Cap (USD)"].values[0] / df["Market Cap (USD)"].sum()) * 100 if "Bitcoin" in df["Name"].values else 0,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "data": df
    }
    
    return analysis

def create_market_cap_chart(analysis):
    """Create a pie chart of market cap distribution for top 5 cryptocurrencies."""
    plt.figure(figsize=(8, 6))
    
   
    top5 = analysis["top_5_by_market_cap"]
    
    
    top5_sum = sum(coin["Market Cap (USD)"] for coin in top5)
    others = analysis["total_market_cap"] - top5_sum
    
    
    labels = [f"{coin['Symbol']} (${coin['Market Cap (USD)'] / 1e9:.2f}B)" for coin in top5]
    labels.append(f"Others (${others / 1e9:.2f}B)")
    
    sizes = [coin["Market Cap (USD)"] for coin in top5]
    sizes.append(others)
    
    
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, shadow=False)
    plt.axis('equal')
    plt.title('Market Cap Distribution (in Billions USD)')
    
    
    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    plt.close()
    

    buffer.seek(0)
    image_png = buffer.getvalue()
    buffer.close()
    
    encoded = base64.b64encode(image_png).decode('utf-8')
    return encoded

def create_price_change_chart(analysis):
    """Create a bar chart of top 5 gainers and losers."""
    plt.figure(figsize=(10, 6))
    
    
    gainers = analysis["top_5_gainers"]
    losers = analysis["top_5_losers"]
    
    
    labels = [coin["Symbol"] for coin in gainers] + [coin["Symbol"] for coin in losers]
    values = [coin["24h Price Change (%)"] for coin in gainers] + [coin["24h Price Change (%)"] for coin in losers]
    colors = ['green' if val >= 0 else 'red' for val in values]
    
    
    plt.bar(range(len(values)), values, tick_label=labels, color=colors)
    plt.axhline(y=0, color='black', linestyle='-', alpha=0.3)
    plt.title('Top 5 Gainers and Losers (24h Price Change %)')
    plt.ylabel('Price Change (%)')
    plt.xticks(rotation=45)
    
    
    for i, v in enumerate(values):
        plt.text(i, v + (1 if v >= 0 else -1), f"{v:.2f}%", ha='center', va='bottom' if v >= 0 else 'top')
    
    plt.tight_layout()
    
   
    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    plt.close()
    
    
    buffer.seek(0)
    image_png = buffer.getvalue()
    buffer.close()
    
    encoded = base64.b64encode(image_png).decode('utf-8')
    return encoded

def generate_html_report(analysis):
    """Generate an HTML report that can be opened in Word."""
    if not analysis:
        return False
    
    
    market_cap_chart = create_market_cap_chart(analysis)
    price_change_chart = create_price_change_chart(analysis)
    
   
    formatted_market_cap = "${:,.2f}".format(analysis["total_market_cap"])
    formatted_volume = "${:,.2f}".format(analysis["total_trading_volume"])
    formatted_avg_price = "${:,.2f}".format(analysis["average_price"])
    formatted_median_price = "${:,.2f}".format(analysis["median_price"])
    
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Cryptocurrency Market Analysis Report</title>
        <style>
            body {{ font-family: 'Calibri', Arial, sans-serif; margin: 20px; }}
            h1, h2, h3 {{ color: #2F5597; }}
            table {{ border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 20px; }}
            th, td {{ border: 1px solid #DDDDDD; text-align: left; padding: 8px; }}
            th {{ background-color: #2F5597; color: white; }}
            tr:nth-child(even) {{ background-color: #F2F2F2; }}
            .container {{ margin-bottom: 30px; }}
            .chart {{ text-align: center; margin: 20px 0; }}
            .positive {{ color: green; }}
            .negative {{ color: red; }}
            .timestamp {{ font-style: italic; color: #666666; text-align: right; }}
        </style>
    </head>
    <body>
        <h1>Cryptocurrency Market Analysis Report</h1>
        <p class="timestamp">Generated on: {analysis["timestamp"]}</p>
        
        <div class="container">
            <h2>Market Overview</h2>
            <table>
                <tr>
                    <th>Metric</th>
                    <th>Value</th>
                </tr>
                <tr>
                    <td>Total Market Cap of Top 50</td>
                    <td>{formatted_market_cap}</td>
                </tr>
                <tr>
                    <td>Total 24h Trading Volume</td>
                    <td>{formatted_volume}</td>
                </tr>
                <tr>
                    <td>Average Price of Top 50</td>
                    <td>{formatted_avg_price}</td>
                </tr>
                <tr>
                    <td>Median Price of Top 50</td>
                    <td>{formatted_median_price}</td>
                </tr>
                <tr>
                    <td>Bitcoin Dominance</td>
                    <td>{analysis["bitcoin_dominance"]:.2f}%</td>
                </tr>
            </table>
        </div>
        
        <div class="chart">
            <h2>Market Cap Distribution</h2>
            <img src="data:image/png;base64,{market_cap_chart}" alt="Market Cap Distribution" />
        </div>
        
        <div class="container">
            <h2>Top 5 Cryptocurrencies by Market Cap</h2>
            <table>
                <tr>
                    <th>Rank</th>
                    <th>Name</th>
                    <th>Symbol</th>
                    <th>Market Cap (USD)</th>
                    <th>Current Price (USD)</th>
                </tr>
    """
    
    
    for i, coin in enumerate(analysis["top_5_by_market_cap"], 1):
        html_content += f"""
                <tr>
                    <td>{i}</td>
                    <td>{coin["Name"]}</td>
                    <td>{coin["Symbol"]}</td>
                    <td>${coin["Market Cap (USD)"]:,.2f}</td>
                    <td>${coin["Current Price (USD)"]:.6f}</td>
                </tr>
        """
    
    html_content += """
            </table>
        </div>
        
        <div class="chart">
            <h2>24-Hour Price Change: Top Gainers and Losers</h2>
            <img src="data:image/png;base64,""" + price_change_chart + """" alt="Price Change Chart" />
        </div>
        
        <div class="container">
            <h2>Top 5 Gainers (24h)</h2>
            <table>
                <tr>
                    <th>Rank</th>
                    <th>Name</th>
                    <th>Symbol</th>
                    <th>24h Price Change (%)</th>
                </tr>
    """
    
    
    for i, coin in enumerate(analysis["top_5_gainers"], 1):
        html_content += f"""
                <tr>
                    <td>{i}</td>
                    <td>{coin["Name"]}</td>
                    <td>{coin["Symbol"]}</td>
                    <td class="positive">+{coin["24h Price Change (%)"]:.2f}%</td>
                </tr>
        """
    
    html_content += """
            </table>
        </div>
        
        <div class="container">
            <h2>Top 5 Losers (24h)</h2>
            <table>
                <tr>
                    <th>Rank</th>
                    <th>Name</th>
                    <th>Symbol</th>
                    <th>24h Price Change (%)</th>
                </tr>
    """
    
    
    for i, coin in enumerate(analysis["top_5_losers"], 1):
        html_content += f"""
                <tr>
                    <td>{i}</td>
                    <td>{coin["Name"]}</td>
                    <td>{coin["Symbol"]}</td>
                    <td class="negative">{coin["24h Price Change (%)"]:.2f}%</td>
                </tr>
        """
    
    html_content += """
            </table>
        </div>
        
        <div class="container">
            <h2>Conclusion</h2>
            <p>
                This report provides a snapshot of the cryptocurrency market, highlighting key metrics,
                trends, and notable performers. The data is sourced from the CoinGecko API and represents
                the market state as of the timestamp indicated above.
            </p>
            <p>
                For more detailed information and live updates, please refer to the Excel file that
                accompanies this report. The Excel file is automatically updated every 5 minutes to
                reflect the most current market conditions.
            </p>
        </div>
        
        <p class="timestamp">
            Note: This report was generated automatically as part of the cryptocurrency tracking system.
            The data is subject to market volatility and should be considered a point-in-time analysis.
        </p>
    </body>
    </html>
    """
    
    
    with open(REPORT_FILE_PATH, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    print(f"HTML report generated at {REPORT_FILE_PATH}")
    return True

def main():
    """Main function to generate the analysis report."""
    print("Generating Cryptocurrency Analysis Report...")
    
    
    crypto_data = fetch_top_50_cryptos()
    
    if crypto_data:
        
        processed_data = process_crypto_data(crypto_data)
        
        
        analysis_results = analyze_crypto_data(processed_data)
        
        
        if generate_html_report(analysis_results):
            print(f"Report successfully generated at {REPORT_FILE_PATH}")
            print("You can open this HTML file in Word to save it as a Word document or PDF.")
        else:
            print("Failed to generate the report.")
    else:
        print("Failed to fetch cryptocurrency data.")

if __name__ == "__main__":
    main() 