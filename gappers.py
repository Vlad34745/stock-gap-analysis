import yfinance as yf
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. CONFIGURATION & DATA INGESTION
# ==========================================

# Change the ticker to any asset (e.g., "TSLA", "NVDA", "BTC-USD")
ticker = "AAPL"  

print(f"Fetching historical data for {ticker} from Yahoo Finance...")
data = yf.download(ticker, start="2026-01-01", end="2026-05-15", auto_adjust=False, actions=False)

# Fix for MultiIndex column structure returned by modern yfinance API
if isinstance(data.columns, pd.MultiIndex):
    data.columns = data.columns.droplevel(1)

if data.empty:
    print(f"Error: Failed to fetch data for ticker '{ticker}'. Please check the symbol or network connection.")
    exit()

# ==========================================
# 2. QUANTITATIVE STRATEGY CALCULATIONS
# ==========================================

# Calculate the previous day's closing price
data['Prev_Close'] = data['Close'].shift(1)

# Calculate the opening gap percentage (stored as decimal for Excel formatting: 0.015 = 1.5%)
data['Gap_Pct'] = (data['Open'] - data['Prev_Close']) / data['Prev_Close']

# Calculate Day 2 and Day 3 price action relative to Day 1 Close
data['Day2_Move_Pct'] = (data['Close'].shift(-1) - data['Close']) / data['Close']
data['Day3_Move_Pct'] = (data['Close'].shift(-2) - data['Close']) / data['Close']

# Filter: Find opening gaps greater than 1.5% and drop incomplete data rows
gappers = data[data['Gap_Pct'] > 0.015].dropna(subset=['Prev_Close'])

# ==========================================
# 3. EXCEL WORKBOOK GENERATION & STYLING
# ==========================================

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Gap Analysis Summary"

# Ensure grid lines are visible in the spreadsheet
ws.views.sheetView[0].showGridLines = True  

# Typography definitions (clean Segoe UI family)
font_title = Font(name="Segoe UI", size=16, bold=True, color="004D40")
font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
font_data = Font(name="Segoe UI", size=11)
font_bold = Font(name="Segoe UI", size=11, bold=True)

# Color palettes (Emerald Corporate Theme)
fill_header = PatternFill(start_color="004D40", end_color="004D40", fill_type="solid")
fill_zebra = PatternFill(start_color="F0F7F4", end_color="F0F7F4", fill_type="solid")
fill_green = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid") # Soft pastel green for profits
fill_red = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")     # Soft pastel red for losses

# Muted gray borders for data cells
thin_border = Border(
    left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC')
)

# Main sheet title block
ws["A1"] = f"Historical Stock Gap Analysis: {ticker}"
ws["A1"].font = font_title
ws.merge_cells("A1:G1")
ws.row_dimensions[1].height = 30

# Table headers alignment
ws.cell(row=4, column=1, value="Date")
headers = ["Open Price", "Prev Close", "Gap %", "Day 1 Close", "Day 2 Move", "Day 3 Move"]
for idx, h in enumerate(headers, start=2):
    ws.cell(row=4, column=idx, value=h)

# Apply styling properties to the header row
for col in range(1, 8):
    cell = ws.cell(row=4, column=col)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border
ws.row_dimensions[4].height = 25

# Populate spreadsheet rows with calculated dataframe results
current_row = 5
for date, row in gappers.iterrows():
    is_zebra = (current_row % 2 == 0)
    row_fill = fill_zebra if is_zebra else PatternFill(fill_type=None)
    
    # 1. Date column formatting
    c_date = ws.cell(row=current_row, column=1, value=date.strftime('%Y-%m-%d'))
    c_date.alignment = Alignment(horizontal="center", vertical="center")
    
    # 2. Currency values formatting (Open, Previous Close, Day 1 Close)
    ws.cell(row=current_row, column=2, value=float(row['Open'])).number_format = "$#,##0.00"
    ws.cell(row=current_row, column=3, value=float(row['Prev_Close'])).number_format = "$#,##0.00"
    ws.cell(row=current_row, column=5, value=float(row['Close'])).number_format = "$#,##0.00"
    
    # 3. Gap percentage formatting (bold for visibility)
    c_gap = ws.cell(row=current_row, column=4, value=float(row['Gap_Pct']))
    c_gap.number_format = "0.00%"
    c_gap.font = font_bold
    
    # 4. Day 2 and Day 3 metrics with smart performance-based color coding
    for col_idx, col_name in [(6, 'Day2_Move_Pct'), (7, 'Day3_Move_Pct')]:
        val = float(row[col_name])
        c_move = ws.cell(row=current_row, column=col_idx, value=val)
        c_move.number_format = "0.00%"
        
        if val > 0:
            c_move.fill = fill_green
            c_move.font = Font(name="Segoe UI", size=11, color="1B5E20", bold=True)  # Dark green text
        else:
            c_move.fill = fill_red
            c_move.font = Font(name="Segoe UI", size=11, color="B71C1C")            # Dark red text

    # Apply general alignment, borders, and zebra patterns to the current row
    for col in range(1, 8):
        cell = ws.cell(row=current_row, column=col)
        cell.border = thin_border
        if col != 1:
            cell.alignment = Alignment(horizontal="right", vertical="center")
        if col not in [6, 7] and is_zebra:  # Protect conditional performance fills from zebra style overrides
            cell.fill = row_fill
            
    ws.row_dimensions[current_row].height = 20
    current_row += 1

# ==========================================
# 4. SUMMARY / ACCOUNTING TOTALS ROW
# ==========================================
summary_row = current_row + 1

# Insert label row for statistical summaries
ws.cell(row=summary_row, column=1, value="Average").font = font_bold
ws.cell(row=summary_row, column=1).alignment = Alignment(horizontal="left", vertical="center")

# Inject Excel native formulas for aggregate metrics
for col_idx, letter in [(4, 'D'), (6, 'F'), (7, 'G')]:
    c_avg = ws.cell(row=summary_row, column=col_idx, value=f"=AVERAGE({letter}5:{letter}{summary_row-2})")
    c_avg.font = font_bold
    c_avg.number_format = "0.00%"
    c_avg.alignment = Alignment(horizontal="right", vertical="center")
    # Standard accounting border rule: single thin top border, double thin bottom border
    c_avg.border = Border(top=Side(style='thin', color="000000"), bottom=Side(style='double', color="000000"))

# ==========================================
# 5. DYNAMIC SHEET WIDTH FIT & EXPORT
# ==========================================
for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    # Add defensive character padding to guarantee clean presentation without text truncation
    ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

# Asset-specific output filename generation
output_filename = f"{ticker}_Gap_Analysis_Project.xlsx"
wb.save(output_filename)

print("-" * 50)
print(f"Success! Analytical report for {ticker} has been fully generated.")
print(f"File exported as: {output_filename}")
print("-" * 50)