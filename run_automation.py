import os
import glob
import shutil
import email_processor
import pdf_parser
import billsync_formatter
import tracker_manager
from datetime import datetime

def main():
    # Configuration - Using OneDrive paths as requested
    base_dir = r"e:\Antigravity Workspace\QUOVANT"
    template_path = os.path.join(base_dir, "Sample BillSync.xlsx")
    
    # 1. Incoming Folder (For ZIP attachments and Extraction)
    incoming_dir = r"C:\Users\EliteAdm\OneDrive - GMD\General - Appeals\QUOVANT\Invoices_Incoming"
    # 2. Processed Folder (For final storage of parsed PDFs)
    processed_dir = r"C:\Users\EliteAdm\OneDrive - GMD\General - Appeals\QUOVANT\Invoices_Processed"
    # 3. Output Folder (For final Excel)
    output_dir = r"C:\Users\EliteAdm\OneDrive - GMD\General - Appeals\QUOVANT\Output_Excel"
    
    # Ensure all directories exist
    for d in [incoming_dir, processed_dir, output_dir]:
        if not os.path.exists(d):
            os.makedirs(d)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel = os.path.join(output_dir, f"Files_BillSync_Output_{timestamp}.xlsx")
    
    print("--- Step 1: Processing Emails (Downloading and Extracting to Incoming) ---")
    try:
        # Extract everything into incoming_dir first (Staging)
        email_processor.process_emails(incoming_dir, incoming_dir)
    except Exception as e:
        print(f"Warning: Email processing failed. Error: {e}")
        
    print("\n--- Step 2: Parsing PDFs (From Incoming) ---")
    all_extracted_items = []
    
    # Find all PDFs in the incoming directory (extracted from ZIPs)
    pdf_files = glob.glob(os.path.join(incoming_dir, "*.pdf"))
    
    if not pdf_files:
        print("No new PDF files found to process.")
        return
        
    for pdf_file in pdf_files:
        print(f"Processing: {os.path.basename(pdf_file)}")
        try:
            items, metadata = pdf_parser.extract_invoice_data(pdf_file)
            print(f"  Extracted {len(items)} line items. Meta: {metadata['invoice_number']}")
            all_extracted_items.extend(items)
            
            # --- Tracker Update ---
            # Update Tracker immediately per file or batch? Per file is safer.
            tracker_path = os.path.join(output_dir, "Master_Tracker.xlsx")
            # We use the DESTINATION path for the link in tracker, because the file will be moved there.
            # But the move happens later.
            # Let's predict the destination path.
            final_pdf_path = os.path.join(processed_dir, os.path.basename(pdf_file))
            tracker_manager.update_tracker(tracker_path, metadata, final_pdf_path)
            
        except Exception as e:
            print(f"  Error parsing {pdf_file}: {e}")
            
    print(f"\nTotal Line Items Extracted: {len(all_extracted_items)}")
    
    print("\n--- Step 3: Generating BillSync Report ---")
    if all_extracted_items:
        billsync_formatter.format_to_billsync(all_extracted_items, template_path, output_excel)
        
        # --- Step 4: Cleanup and Organization ---
        print("\n--- Step 4: Cleanup and Organization ---")
        
        # Move processed PDFs to Invoices_Processed
        for pdf_file in pdf_files:
            try:
                dest = os.path.join(processed_dir, os.path.basename(pdf_file))
                # If destination exists, add timestamp or overwrite? Assuming move/overwrite for now
                if os.path.exists(dest):
                    os.remove(dest)
                shutil.move(pdf_file, dest)
                print(f"  Moved {os.path.basename(pdf_file)} to Invoices_Processed")
            except Exception as e:
                print(f"  Warning: Could not move {pdf_file}: {e}")
                
        # Delete ZIP files in Incoming
        zip_files = glob.glob(os.path.join(incoming_dir, "*.zip"))
        for zip_file in zip_files:
            try:
                os.remove(zip_file)
                print(f"  Deleted {os.path.basename(zip_file)}")
            except Exception as e:
                print(f"  Warning: Could not delete {zip_file}: {e}")
                
        # Cleanup temporary logs/clutter
        clutter_files = [
            "extraction_results.txt", 
            "extraction_results_v2.txt", 
            "pdf_debug_dump.txt", 
            "pdf_dump_new.txt"
        ]
        for cf in clutter_files:
            cf_path = os.path.join(base_dir, cf)
            if os.path.exists(cf_path):
                try:
                    os.remove(cf_path)
                    print(f"  Removed clutter: {cf}")
                except: pass
                
    else:
        print("No data extracted. Skipping report generation and cleanup.")
        
    print("\n--- Automation Complete ---")

if __name__ == "__main__":
    main()
