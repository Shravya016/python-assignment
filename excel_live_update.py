import requests
import pandas as pd
import time
import os
from datetime import datetime
import win32com.client
import pythoncom
import sys


COINGECKO_API_URL = "https://api.coingecko.com/api/v3"
EXCEL_FILE_PATH = os.path.abspath("crypto_data_live.xlsx")
REFRESH_INTERVAL = 300  

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
    
    # Analysis results
    analysis = {
        "top_5_by_market_cap": df.nlargest(5, "Market Cap (USD)")[["Name", "Symbol", "Market Cap (USD)"]].to_dict("records"),
        "average_price": df["Current Price (USD)"].mean(),
        "highest_price_change": df.loc[df["24h Price Change (%)"].idxmax()],
        "lowest_price_change": df.loc[df["24h Price Change (%)"].idxmin()],
        "total_market_cap": df["Market Cap (USD)"].sum(),
        "total_trading_volume": df["24h Trading Volume (USD)"].sum(),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    return analysis

def create_excel_template():
    """Create the initial Excel template if it doesn't exist."""
    if os.path.exists(EXCEL_FILE_PATH):
        return
    
    
    columns = [
        "Name", "Symbol", "Current Price (USD)", 
        "Market Cap (USD)", "24h Trading Volume (USD)", 
        "24h Price Change (%)"
    ]
    df = pd.DataFrame(columns=columns)
    
    with pd.ExcelWriter(EXCEL_FILE_PATH, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Live Crypto Data', index=False)
        analysis_df = pd.DataFrame(columns=["Metric", "Value"])
        analysis_df.to_excel(writer, sheet_name='Analysis', index=False)
    
    print(f"Created Excel template at {EXCEL_FILE_PATH}")

def update_excel_with_com(data, analysis):
    
    if not data or not analysis:
        return False
    
    try:
        
        pythoncom.CoInitialize()
        df = pd.DataFrame(data)
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  
        

        try:
            wb = excel.Workbooks(os.path.basename(EXCEL_FILE_PATH))
            print("Excel file is already open, updating...")
        except:
            
            if not os.path.exists(EXCEL_FILE_PATH):
                create_excel_template()
            wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
            print("Opened Excel file for updating...")
        
        
        ws_data = wb.Sheets("Live Crypto Data")
        
        
        if ws_data.UsedRange.Rows.Count > 1:
            ws_data.Range(f"A2:F{ws_data.UsedRange.Rows.Count}").Clear()
        
        
        for r, row in enumerate(df.values, start=2):
            for c, value in enumerate(row, start=1):
                ws_data.Cells(r, c).Value = value
        
        
        data_range = ws_data.Range(f"A1:F{len(df) + 1}")
        data_range.Columns.AutoFit()
        
        
        ws_data.Cells(1, 8).Value = "Last Updated:"
        ws_data.Cells(1, 9).Value = analysis["timestamp"]
        
        
        ws_analysis = wb.Sheets("Analysis")
        
        
        ws_analysis.UsedRange.Clear()
        
        
        ws_analysis.Cells(1, 1).Value = "Cryptocurrency Market Analysis"
        title_range = ws_analysis.Range("A1:D1")
        title_range.Merge()
        title_range.Font.Bold = True
        title_range.Font.Size = 14
        
        
        ws_analysis.Cells(3, 1).Value = "Summary Metrics"
        ws_analysis.Range("A3:B3").Merge()
        ws_analysis.Range("A3:B3").Font.Bold = True
        
        ws_analysis.Cells(4, 1).Value = "Average Price (USD)"
        ws_analysis.Cells(4, 2).Value = analysis["average_price"]
        
        ws_analysis.Cells(5, 1).Value = "Total Market Cap (USD)"
        ws_analysis.Cells(5, 2).Value = analysis["total_market_cap"]
        
        ws_analysis.Cells(6, 1).Value = "Total 24h Trading Volume (USD)"
        ws_analysis.Cells(6, 2).Value = analysis["total_trading_volume"]
        
        ws_analysis.Cells(7, 1).Value = "Last Updated"
        ws_analysis.Cells(7, 2).Value = analysis["timestamp"]
        
        
        ws_analysis.Cells(9, 1).Value = "Top 5 Cryptocurrencies by Market Cap"
        ws_analysis.Range("A9:D9").Merge()
        ws_analysis.Range("A9:D9").Font.Bold = True
        
        ws_analysis.Cells(10, 1).Value = "Rank"
        ws_analysis.Cells(10, 2).Value = "Name"
        ws_analysis.Cells(10, 3).Value = "Symbol"
        ws_analysis.Cells(10, 4).Value = "Market Cap (USD)"
        
        for i, coin in enumerate(analysis["top_5_by_market_cap"], 1):
            ws_analysis.Cells(10 + i, 1).Value = i
            ws_analysis.Cells(10 + i, 2).Value = coin["Name"]
            ws_analysis.Cells(10 + i, 3).Value = coin["Symbol"]
            ws_analysis.Cells(10 + i, 4).Value = coin["Market Cap (USD)"]
        
        
        ws_analysis.Cells(17, 1).Value = "24-Hour Price Change Extremes"
        ws_analysis.Range("A17:D17").Merge()
        ws_analysis.Range("A17:D17").Font.Bold = True
        
        ws_analysis.Cells(18, 1).Value = "Type"
        ws_analysis.Cells(18, 2).Value = "Name"
        ws_analysis.Cells(18, 3).Value = "Symbol"
        ws_analysis.Cells(18, 4).Value = "Price Change (%)"
        
        
        ws_analysis.Cells(19, 1).Value = "Highest"
        ws_analysis.Cells(19, 2).Value = analysis["highest_price_change"]["Name"]
        ws_analysis.Cells(19, 3).Value = analysis["highest_price_change"]["Symbol"]
        ws_analysis.Cells(19, 4).Value = analysis["highest_price_change"]["24h Price Change (%)"]
        
        
        ws_analysis.Cells(20, 1).Value = "Lowest"
        ws_analysis.Cells(20, 2).Value = analysis["lowest_price_change"]["Name"]
        ws_analysis.Cells(20, 3).Value = analysis["lowest_price_change"]["Symbol"]
        ws_analysis.Cells(20, 4).Value = analysis["lowest_price_change"]["24h Price Change (%)"]
        
        
        ws_analysis.Columns.AutoFit()
        
        
        wb.Save()
        print(f"Excel file updated at {analysis['timestamp']}")
        
        return True
    
    except Exception as e:
        print(f"Error updating Excel: {e}")
        return False
    finally:
        
        try:
            wb = None
            excel.Quit()
            excel = None
            pythoncom.CoUninitialize()
        except:
            pass

def main():
    """Main function to run the cryptocurrency tracker with live Excel updates."""
    print("Cryptocurrency Live Excel Tracker Started")
    print("=======================================")
    print("Fetching and analyzing top 50 cryptocurrencies...")
    print("Updates will be sent directly to Excel every 5 minutes.")
    print("Excel file will be opened automatically.")
    print("Press Ctrl+C to stop the program.")
    
    
    if not os.path.exists(EXCEL_FILE_PATH):
        create_excel_template()
    
    run_count = 0
    
    try:
        while True:
            run_count += 1
            print(f"\nUpdate #{run_count} - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            
            crypto_data = fetch_top_50_cryptos()
            
            if crypto_data:
                
                processed_data = process_crypto_data(crypto_data)
                
                
                analysis_results = analyze_crypto_data(processed_data)
                
                
                success = update_excel_with_com(processed_data, analysis_results)
                
                if success:
                    print(f"Excel updated successfully. Next update in {REFRESH_INTERVAL // 60} minutes.")
                else:
                    print("Failed to update Excel. Will retry in the next cycle.")
            else:
                print("Failed to fetch data. Will retry in the next cycle.")
            
            
            time.sleep(REFRESH_INTERVAL)
            
    except KeyboardInterrupt:
        print("\nProgram terminated by user.")
    except Exception as e:
        print(f"\nAn error occurred: {e}")
    finally:
        print("\nCryptocurrency Live Excel Tracker Stopped")

if __name__ == "__main__":
    main() 