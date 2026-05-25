import pandas as pd
import shutil
import os
from datetime import datetime

def format_to_billsync(extracted_items, template_path, output_path):
    # Strict column adherence as per user request (Columns A to S from Sample BillSync.xlsx)
    # Added 3 new columns at start: Invoice Number, Company, User
    columns = [
        'Invoice Number',       # New Column A
        'Company',              # New Column B
        'User',                 # New Column C
        'Invoice Date',         # Column D
        'Working Timekeeper',   # Column E
        'Billing Timekeeper',   # Column F
        'Description',          # Column G
        'Date of Item',         # Column H
        'Last Date to add Attorney Information', # Column I
        'Appeal Status',        # Column J
        'Matter Number',        # Column K
        'Task ID',              # Column L
        'Item Type',            # Column M
        'UNITS',                # Column N
        'RATE',                 # Column O
        'AMOUNT',               # Column P
        'Reduced Amount',       # Column Q
        'Total Invoice Amount', # Column R
        'Narrative',            # Column S
        'Attorney Comment',     # Column T
        'Attachment',           # Column U
        'Attachment : URL'      # Column V
    ]

    # Create DataFrame from items
    rows = []
    for item in extracted_items:
        # Rule 1) Ignore line items having ZERO reductions.
        red_amt = item.get('reduced_amount', 0.0)
        
        # User Request: "include only reduced line items"
        # So we strictly filter out anything with 0 reduction.
        if abs(red_amt) < 0.001:
            continue
            
        row = {col: "" for col in columns}
        
        # New Columns
        row['Invoice Number'] = item.get('invoice_number', '')
        row['Company'] = "" # Placeholder as per user request to add header
        row['User'] = ""    # Placeholder as per user request to add header
        
        # Extracted Fields mapping
        # Rule 2) Working Timekeeper name is required (Handled in parser upstream or here)
        # Parser now puts Name in 'timekeeper' field if found.
        row['Working Timekeeper'] = item.get('timekeeper', '')
        
        # Rule 3) No need to fill Billing Timekeeper. Keep it blank.
        row['Billing Timekeeper'] = "" 
        
        # Date of line items -> Date of Item
        # (Invoice Date column might refer to the header Invoice Date)
        row['Date of Item'] = item.get('date', '')
        row['Invoice Date'] = item.get('invoice_date', '') # Extracted from header if available
        
        # Matter Number -> Matter Number
        row['Matter Number'] = item.get('matter_number', '')
        
        # Item Type (Fee/Expense) -> Item Type
        row['Item Type'] = item.get('item_type', '')
        
        # Unit -> UNITS
        row['UNITS'] = item.get('units', 0.0)
        
        # Rate -> RATE
        row['RATE'] = item.get('rate', 0.0)
        
        # Amount -> AMOUNT
        row['AMOUNT'] = item.get('amount', 0.0)
        
        # Reduced Amount -> Reduced Amount
        row['Reduced Amount'] = red_amt
        
        # Rule 4) Merge Description and Narrative with [Description + “Line Break” + “Audit Reason:” + Narrative]
        full_desc = item.get('description', '').strip()
        reason = item.get('reason', '').strip()
        notes = item.get('notes', '').strip()
        
        # Narrative logic (Audit Reason)
        audit_reason = f"{reason} {notes}".strip()
        amount_billed = row['AMOUNT']
        
        # User Requirement: skip those audit reason where amt billed (AMOUNT) in ZERO.
        if audit_reason and amount_billed != 0:
            # Using \n for Line Break in Excel cells
            combined_desc = f"{full_desc}\nAudit Reason: {audit_reason}"
        else:
            combined_desc = full_desc
            
        # User Request: Put whole text in Narrative. Keep Description Blank.
        row['Description'] = ""
        row['Narrative'] = combined_desc
            
        rows.append(row)
        
    df = pd.DataFrame(rows, columns=columns)
    
    # Save
    try:
        df.to_excel(output_path, index=False)
        print(f"Saved BillSync report to {output_path}")
    except Exception as e:
        print(f"Error saving Excel: {e}")

if __name__ == "__main__":
    pass
