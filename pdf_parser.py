import pdfplumber
import re
import os

def parse_line_item(line):
    # Regex to find the trailing numbers (Rate, Units, Amount)
    # They usually appear at the end of the line.
    # Pattern: Space, Float, Space, Float, Space, Float (End of line)
    # Note: Amount can be negative.
    
    # Example: 255.00 1.30 331.50
    # Example: 0.00 0.00 0.10
    
    pattern = r'\s+(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})$'
    match = re.search(pattern, line)
    
    if match:
        rate_str, units_str, amount_str = match.groups()
        
        # Remove these from the line to get the prefix
        prefix = line[:match.start()].strip()
        
        # Now parse the prefix: LineNo Date Timekeeper? Description
        # Regex: Start with digits (LineNo), then Date
        prefix_match = re.match(r'^(\d+)\s+(\d{1,2}/\d{1,2}/\d{4})\s+(.*)$', prefix)
        if prefix_match:
            line_no, date, remainder = prefix_match.groups()
            
            # Check for Timekeeper in remainder
            # Timekeeper seems to be the first word in remainder if it exists
            # Example Fee: "CT289876 Draft letter..."
            # Example Exp: "PHOTOCOPIES - CORRESPONDENCE" (No ID, or ID is the description?)
            
            # Heuristic: If first word is alphanumeric/uppercase and looks like an ID?
            # Or simplified: Start description after first space?
            tokens = remainder.split(' ', 1)
            first_token = tokens[0]
            
            timekeeper = ""
            description = ""
            item_type = ""
            
            rate_val = float(rate_str.replace(',', ''))
            
            if rate_val > 0:
                item_type = "Fee"
                timekeeper = first_token
                description = tokens[1] if len(tokens) > 1 else ""
            else:
                item_type = "Expense"
                timekeeper = "Expense" # User request: Working Timekeeper should be 'Expense' for expenses
                description = remainder
                
            return {
                "line_no": line_no,
                "date": date,
                "timekeeper": timekeeper,
                "description": description,
                "rate": float(rate_str.replace(',', '')),
                "units": float(units_str.replace(',', '')),
                "amount": float(amount_str.replace(',', '')),
                "item_type": item_type,
                "reduced_amount": 0.0,
                "reason": "",
                "notes": "",
                "adjustment_active": False # Track if we are in an adjustment explanation
            }
            
    return None

def extract_adjustment_summary(pdf):
    """
    Builds a mapping of {Code: Comment} from the 'Adjustment Summary' table.
    """
    summary_map = {}
    # Updated regex to handle both "Amount Adjusted" and "Amount Reversal" (to correctly split entries)
    entry_pattern = r'([A-Z0-9\-]+)\s+(?:Hours|Amount)\s+(?:Adjusted|Reversed|Reversal):\s*(-?[\d\.,]+)\s+(?:[A-Za-z]?\s*)?(?:Amount|Hours)\s+(?:Adjusted|Reversal|Reversed):\s*(-?[\d,\.]+)'
    
    for page in pdf.pages:
        text = page.extract_text()
        if not text: continue
        
        matches = list(re.finditer(entry_pattern, text))
        if not matches: continue

        for i, match in enumerate(matches):
            code = match.group(1)
            start_index = match.end()
            
            if i < len(matches) - 1:
                end_index = matches[i+1].start()
            else:
                end_index = len(text)
                total_match = re.search(r'(Total.*Hours Adjusted|Time Keeper Summary|Payment Summary)', text[start_index:])
                if total_match:
                     end_index = start_index + total_match.start()
            
            raw_comment = text[start_index:end_index].strip()
            
            # Skip common footer noise at the end of the report
            noise_footers = [
                "Adjustments on this invoice may be appealed",
                "Appeals may be filed at",
                "Time Keeper Summary",
                "For Control Number:",
                "PRIVILEGED AND CONFIDENTIAL",
                "The information contained in this document",
                "intended only for the use of the individual",
                "If the reader of this document",
                "strictly prohibited",
                "immediately notify us",
                "return the original document",
                "Galleria Circle",
                "Postal Service",
                "Thank you.",
                "Copyright 2026 Quovant",
                "Page "
            ]
            
            clean_lines = []
            for line in raw_comment.split('\n'):
                line = line.strip()
                if not line: continue
                if any(k in line for k in ["Percent of Total", "Adjustment Summary"]):
                    continue
                if any(k in line for k in noise_footers):
                    continue
                if line in ["P", "A", "E"]: continue 
                clean_lines.append(line)
                
            summary_map[code] = " ".join(clean_lines).strip()
    return summary_map

def extract_invoice_data(pdf_path):
    items = []
    current_item = None
    
    # Metadata to extract
    invoice_date = ""
    matter_number = ""
    invoice_number = ""
    finalized_date = ""
    timekeeper_map = {}
    
    total_submitted = 0.0
    total_payable = 0.0
    
    with pdfplumber.open(pdf_path) as pdf:
        # Pre-extract adjustment summary comments
        adj_summary_comments = extract_adjustment_summary(pdf)
        
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            
            lines = text.split('\n')
            
            # Extract header info (usually on first page)
            if not invoice_date:
                header_text = "\n".join(lines[:30]) # Check top 30 lines
                
                inv_match = re.search(r'Invoice Date:\s+(\d{1,2}/\d{1,2}/\d{4})', header_text)
                if inv_match:
                    invoice_date = inv_match.group(1)
                    
                inv_num_match = re.search(r'Invoice Number:\s+(\S+)', header_text)
                if inv_num_match:
                    invoice_number = inv_num_match.group(1)
            
            # Check for Finalized Date (Authorized On) separately to ensure we catch it
            # It might appear later or earlier
            if not finalized_date:
                # "Authorized On: 12/4/2025 12:23:41 PM"
                auth_match = re.search(r'Authorized On:\s+(\d{1,2}/\d{1,2}/\d{4})', text)
                if auth_match:
                    finalized_date = auth_match.group(1)
            
            # Matter Number extraction
            if not matter_number:
                mat_match = re.search(r'Matter Number:\s+(\d+)', text)
                if mat_match:
                    matter_number = mat_match.group(1)

            # Payment Summary Extraction
            if "Payment Summary" in text:
                # Look for TOTAL and PAY TO FIRM lines in this page's text
                for line in lines:
                    line = line.strip()
                    if line.startswith("TOTAL") and "Fees" not in line:
                         # Last token is the Grand Total
                         parts = line.split()
                         if parts:
                             try:
                                 total_submitted = float(parts[-1].replace(',', ''))
                             except: pass
                    
                    if line.startswith("PAY TO FIRM"):
                        parts = line.split()
                        if parts:
                            try:
                                total_payable = float(parts[-1].replace(',', ''))
                            except: pass

            # Build Timekeeper Mapping
            if "Time Keeper Summary" in text:
                 tk_lines = text.split("Time Keeper Summary")[1].split('\n')
                 for tk_line in tk_lines:
                     tk_line = tk_line.strip()
                     tk_match = re.search(r'^(.*)\s+([A-Za-z0-9]+)\s+[A-Za-z]+\s+[\d,]+\.\d{2}', tk_line)
                     if tk_match:
                         name = tk_match.group(1).strip()
                         tk_id = tk_match.group(2).strip()
                         if "Identifier" not in tk_id and "Time" not in name:
                             timekeeper_map[tk_id] = name

            for line in lines:
                line = line.strip()
                
                # Check for new line item
                if re.match(r'^\d+\s+\d{1,2}/\d{1,2}/\d{4}', line):
                    if current_item:
                        # Map Timekeeper ID to Name
                        tk_id = current_item.get('timekeeper', '')
                        if tk_id in timekeeper_map:
                            current_item['timekeeper'] = timekeeper_map[tk_id]
                            
                        current_item['invoice_date'] = invoice_date
                        current_item['matter_number'] = matter_number
                        current_item['invoice_number'] = invoice_number
                        items.append(current_item)
                    
                    # Start new item
                    parsed = parse_line_item(line)
                    if parsed:
                        current_item = parsed
                        current_item['notes_active'] = False
                    else:
                        pass
                
                # Check for new line adjustment (handles leading status P or A)
                elif current_item:
                    adj_match = re.search(r'(?:^|[A-Z]?\s+)([A-Z][A-Z0-9\-]{1,})\s+(.*?)\s+(-?[\d,]+\.\d{2})(?:\s+(-?[\d,]+\.\d{2}))?$', line)
                    if adj_match:
                        code, local_reason, val1, val2 = adj_match.groups()
                        amt_adj_str = val2 if val2 else val1
                        val_adj = float(amt_adj_str.replace(',', ''))
                        current_item['reduced_amount'] += val_adj
                        
                        if abs(val_adj) > 0.001:
                            current_item['last_code'] = code
                            if "0.00000" not in local_reason:
                                # Start with local reason if meaningful
                                current_item['reason'] += f"{local_reason} ".strip() + " "
                            
                            # Add boilerplate comment from the Adjustment Summary
                            if code in adj_summary_comments:
                                comment = adj_summary_comments[code]
                                if comment not in current_item['reason']:
                                    current_item['reason'] += comment + " "
                            
                            current_item['adjustment_active'] = True # Found a reduction, succeeding text is likely explanation
                        current_item['notes_active'] = False
                        continue
                        
                    if "0.00000" in line:
                        continue

                    if "Client Notes:" in line or "Quovant Notes:" in line:
                         current_item['notes'] += line + " "
                         current_item['notes_active'] = True
                         current_item['adjustment_active'] = False
                         continue
                         
                    if line.startswith("Payment Summary") or line.startswith("LCN:"):
                         current_item['notes_active'] = False
                         continue
                         
                    noise_keywords = [
                        "Copyright", "Page", "PRIVILEGED", "Originally Printed:", 
                        "Reprinted:", "PO Box", "Dallas, TX", "Phone:", 
                        "Compliance Report", "Control Number:", "Invoice Details",
                        "Line Date Atty", "Rate Billed", "Originally Printed",
                        "Time Keeper Summary", "For Control Number:", "Identifier Position",
                        "Time Keeper Identifier", "Authorized On:", "Payment Summary"
                    ]
                    
                    found_noise = False
                    for k in noise_keywords:
                        if k in line:
                            # Truncate the line before the noise if possible
                            line = line.split(k)[0].strip()
                            found_noise = True
                            break
                    
                    if found_noise:
                        if line and current_item.get('adjustment_active'):
                             current_item['reason'] += line + " "
                        current_item['notes_active'] = False
                        current_item['adjustment_active'] = False
                        continue
                        
                    if line.startswith("The information contained") or line.startswith("document is not the intended"):
                        current_item['adjustment_active'] = False
                        continue
                    if line.startswith("immediately notify us") or line.startswith("submitted in writing to Quovant"):
                        current_item['adjustment_active'] = False
                        continue
                        
                    if "UNAUTHORIZED TIMEKEEPER" in line and "A70" not in line:
                         if "0.00000" not in line:
                            current_item['reason'] += line + " "
                         current_item['notes_active'] = False
                         current_item['adjustment_active'] = True
                    elif current_item.get('notes_active'):
                         current_item['notes'] += line + " "
                    elif current_item.get('adjustment_active'):
                         # Avoid appending the boilerplate again if it was already added from summary
                         clean_line = line.strip()
                         
                         # Check if this line is already covered by the boilerplate comment
                         is_boilerplate_part = False
                         code = current_item.get('last_code', '') # We need to store the code
                         if code in adj_summary_comments:
                             if clean_line in adj_summary_comments[code]:
                                 is_boilerplate_part = True
                         
                         if clean_line and not is_boilerplate_part and clean_line not in current_item['reason']:
                             # Special case: line references like "See Line 1100" 
                             # User requested "only highlighted part", which usually excludes these refs.
                             if "See Line" in clean_line:
                                 # If it's just a line ref, we skip it as per "highlighted part" request
                                 pass
                             else:
                                 current_item['reason'] += clean_line + " "
                    else:
                         current_item['description'] += " " + line

    # Add last item
    if current_item:
        tk_id = current_item.get('timekeeper', '')
        if tk_id in timekeeper_map:
            current_item['timekeeper'] = timekeeper_map[tk_id]
            
        current_item['invoice_date'] = invoice_date
        current_item['matter_number'] = matter_number
        current_item['invoice_number'] = invoice_number
        items.append(current_item)
        
    total_reduction = total_submitted - total_payable
    
    metadata = {
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "finalized_date": finalized_date,
        "total_submitted": total_submitted,
        "total_reduction": total_reduction,
        "total_payable": total_payable
    }
        
    return items, metadata

if __name__ == "__main__":
    # Test on the sample PDF
    import sys
    test_pdf = "12039854.pdf" # Assumes in CWD
    if len(sys.argv) > 1:
        test_pdf = sys.argv[1]
        
    data, meta = extract_invoice_data(test_pdf)
    print("Metadata:", meta)
    print(f"Extracted {len(data)} items")
