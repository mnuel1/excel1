import openpyxl
from openpyxl.styles import numbers, Font, PatternFill, Border, Side, Alignment
import pandas as pd
import os
import datetime

# === Load the Excel files ===
masterlist_path = 'masterlist.xlsx'
report_path = 'report.xlsx'

master_df = pd.read_excel(masterlist_path)
report_df = pd.read_excel(report_path)

# Normalize column names to uppercase and strip spaces
report_df.columns = report_df.columns.str.strip().str.upper()

# Fill down necessary columns
report_df[['DATE', 'CODE', 'CUSTOMER']] = report_df[['DATE', 'CODE', 'CUSTOMER']].fillna(method='ffill')

# Remove rows where any of DATE, CODE, or CUSTOMER are still missing
report_df = report_df.dropna(subset=['DOC. NO.'])

# Strip any extra spaces
report_df['CODE'] = report_df['CODE'].astype(str).str.strip()
report_df['CUSTOMER'] = report_df['CUSTOMER'].astype(str).str.strip()


# === Output folder for reports ===
output_dir = 'project_reports'
os.makedirs(output_dir, exist_ok=True)

# === Get unique projects from masterlist ===
projects = master_df['Project'].unique()

for project in projects:
    # Get all customers for the current project
    customers = master_df[master_df['Project'] == project]['Name'].unique()

    # Filter report rows for those customers
    filtered_report = report_df[
        (report_df['CODE'].astype(str).str.strip() + ' - ' + report_df['CUSTOMER'].astype(str).str.strip()).isin(customers)
    ]

    if filtered_report.empty:
        continue  # Skip if there's no data for this project
        
    # Construct the output dataframe
    output_data = []

    # Ensure unit price is numeric
    report_df['UNIT PRICE'] = pd.to_numeric(report_df['UNIT PRICE'], errors='coerce').fillna(0)

    grand_total = 0

    # Group by CODE - CUSTOMER
    for cust_code, group in filtered_report.groupby(filtered_report['CODE'] + ' - ' + filtered_report['CUSTOMER']):
        cust_total = group['UNIT PRICE'].sum()
        first_row = True

        for _, row in group.iterrows():
            line = {
                '110204010104 - Other Receivables': cust_code if first_row else '',
                'Balance': cust_total if first_row else '',
                'Particulars': row['DOC. NO.'],
                # 'Date': row['DATE'],
                'Amount': row['UNIT PRICE'],
                # 'Remarks': '',
                # 'Status': "OVERPAYMENT" if cust_total < 0 else "OUTSTANDING" 
            }
            output_data.append(line)
            first_row = False
            
        # Add subtotal row for this customer
        output_data.append({
            '110204010104 - Other Receivables': '',
            'Balance': '',
            'Particulars': 'TOTAL',
            # 'Date': '',
            'Amount': cust_total,
            # 'Remarks': '',
            # 'Status': ''
        })
    # Add grand total row
    grand_total = sum([row['Amount'] for row in output_data if row['Particulars'] == 'TOTAL'])

    output_data.append({
        '110204010104 - Other Receivables': 'TOTAL',
        'Balance': grand_total,
        'Particulars': '',
        # 'Date': '',
        'Amount': '',
        # 'Remarks': '',
        # 'Status': ''
    })        

    
    output_df = pd.DataFrame(output_data)
    
    # Save to Excel
    output_file = os.path.join(output_dir, f'{project}_report.xlsx')
    output_df.to_excel(output_file, index=False)

    wb = openpyxl.load_workbook(output_file)
    ws = wb.active

    # styles
    default_font = Font(name='Arial Narrow', size=10, bold=False)
    bold_font = Font(name='Arial Narrow', size=10, bold=True)
    white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    thin_border = Border(bottom=Side(style='thin', color='000000'))

    # Header alignment rules
    alignment_left = Alignment(horizontal='left', vertical='center')
    alignment_right = Alignment(horizontal='right', vertical='center')
    
    # title = ws["B2"]
    # title.value = "OTHER RECEIVABLES"
    # title.font = bold_font
    # title.fill = white_fill

    # date = ws['B3']
    # date.value = "04/01/2000"
    # date.font = bold_font
    # date.fill = white_fill

    header = [cell.value for cell in ws[1]]
    balance_col = header.index('Balance') + 1
    amount_col = header.index('Amount') + 1
        
    accounting_format = '_(#,##0.00_);_((#,##0.00);_("-"??_);_(@_)'
    
    for row in ws.iter_rows(min_row=2, min_col=balance_col, max_col=balance_col):
        for cell in row:
            cell.number_format = accounting_format

    for row in ws.iter_rows(min_row=2, min_col=amount_col, max_col=amount_col):
        for cell in row:
            cell.number_format = accounting_format

    wb.save(output_file)
print("All project reports have been generated.")
