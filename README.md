# Cryptocurrency Live Tracker
This project fetches and analyzes live data for the top 50 cryptocurrencies by market capitalization and updates an Excel file with the latest information every 5 minutes.

## Features
- **Real-time Data Fetching**: Uses the CoinGecko API to retrieve live cryptocurrency data.
- **Basic Analysis**:
  - Identifies the top 5 cryptocurrencies by market cap.
  - Calculates the average price of the top 50 cryptocurrencies.
  - Determines the highest and lowest 24-hour price changes.
- **Live Updates**: Continuously updates an Excel file every 5 minutes.
- **Analysis Report**: Generates a text-based summary report.

## Requirements
- **Python Version**: 3.7 or higher
- **Required Python Packages** (listed in `requirements.txt`):
  - `requests`
  - `pandas`
  - `openpyxl`

## Installation
1. Clone the repository:
   ```sh
   git clone https://github.com/Shravya016/python-assignment.git
   cd python-assignment
   ```
2. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```

## Usage
Run the script:
```sh
python crypto_tracker.py
```
This will:
- Fetch data from the CoinGecko API.
- Create and update an Excel file (`crypto_data_live.xlsx`).
- Generate an analysis report (`crypto_analysis_report.txt`).
- Continue updating data every 5 minutes until manually stopped.

To stop the script, press `Ctrl+C` in the terminal.

## Excel File Structure
The generated Excel file includes:
1. **Live Crypto Data**: Complete dataset for the top 50 cryptocurrencies.
2. **Analysis Sheet**:
   - Average price, total market cap, and total trading volume.
   - Top 5 cryptocurrencies by market cap.
   - Highest and lowest 24-hour price changes.

## Notes
- The **CoinGecko API** has rate limits. If errors occur, you might have exceeded these limits.
- The refresh interval is set to **5 minutes**. Modify the `REFRESH_INTERVAL` constant to change it.

## Analysis Report
The script automatically generates a text-based analysis report. For a more professional report:
1. Open `crypto_analysis_report.txt`.
2. Copy the content into a Word document.
3. Format as needed and save as **PDF** or **Word format**.

## Live Excel Updates
The Excel file updates automatically. To view live changes:
1. Open `crypto_data_live.xlsx` in Microsoft Excel.
2. Enable automatic refresh or manually reload the file periodically.
