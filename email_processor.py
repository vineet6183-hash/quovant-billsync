import win32com.client
import os
import zipfile
import pathlib

def process_emails(download_folder, extract_folder):
    """
    Connects to Outlook, finds unread emails with subject 'MBR Compliance Reports',
    downloads ZIPs to download_folder, and extracts PDFs to extract_folder.
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) # 6 is Inbox
    
    # User's request: "Subject contains 'Mitratech MBR Compliance Reports'"
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    
    print("Checking for the latest unread email with subject 'Mitratech MBR Compliance Reports'...")
    
    found_any = False
    for message in messages:
        try:
            # Filter by Subject as requested
            if "Mitratech MBR Compliance Reports" in message.Subject:
                if message.UnRead: 
                    print(f"Found latest unread email: {message.Subject} received {message.ReceivedTime}")
                    
                    for attachment in message.Attachments:
                        if attachment.FileName.lower().endswith('.zip'):
                            save_path = os.path.join(download_folder, attachment.FileName)
                            print(f"Downloading {attachment.FileName} to {save_path}")
                            attachment.SaveAsFile(save_path)
                            
                            # Extract to the processed folder
                            with zipfile.ZipFile(save_path, 'r') as zip_ref:
                                zip_ref.extractall(extract_folder)
                                print(f"  Extracted {attachment.FileName} to {extract_folder}")
                            
                    # Mark as read after successful attachment processing
                    message.UnRead = False
                    found_any = True
                    # Break after the first (latest) matching email
                    break
        except Exception as e:
            print(f"Error processing message: {e}")
            continue
    
    if not found_any:
        print("No new unread emails found matching criteria.")

if __name__ == "__main__":
    # Test paths
    inc = r"C:\Users\EliteAdm\OneDrive - GMD\General - Appeals\QUOVANT\Invoices_Incoming"
    proc = r"C:\Users\EliteAdm\OneDrive - GMD\General - Appeals\QUOVANT\Invoices_Processed"
    for d in [inc, proc]:
        if not os.path.exists(d): os.makedirs(d)
        
    process_emails(inc, proc)
