import streamlit as st
import pandas as pd
import mailbox
import io
import os
import zipfile
from datetime import datetime
import email.utils
import re
from email.header import decode_header
import tempfile

def clean_filename(filename):
    """Clean filename to remove invalid characters"""
    # Remove or replace invalid characters for Windows filenames
    invalid_chars = r'[<>:"/\\|?*]'
    cleaned = re.sub(invalid_chars, '_', filename)
    # Limit length to 200 characters
    if len(cleaned) > 200:
        cleaned = cleaned[:200]
    return cleaned

def decode_mime_words(s):
    """Decode MIME encoded words in headers"""
    if not s:
        return ""
    
    decoded_parts = []
    for part, encoding in decode_header(s):
        if isinstance(part, bytes):
            if encoding:
                try:
                    part = part.decode(encoding)
                except (LookupError, UnicodeDecodeError):
                    part = part.decode('utf-8', errors='ignore')
            else:
                part = part.decode('utf-8', errors='ignore')
        decoded_parts.append(part)
    
    return ''.join(decoded_parts)

def extract_email_body(message):
    """Extract the plain text body from an email message"""
    body = ""
    
    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            
            # Skip attachments
            if "attachment" in content_disposition:
                continue
                
            if content_type == "text/plain":
                charset = part.get_content_charset()
                if charset is None:
                    charset = 'utf-8'
                
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body = payload.decode(charset, errors='ignore')
                        break
                except (UnicodeDecodeError, LookupError):
                    try:
                        body = payload.decode('utf-8', errors='ignore')
                        break
                    except:
                        continue
    else:
        charset = message.get_content_charset()
        if charset is None:
            charset = 'utf-8'
        
        try:
            payload = message.get_payload(decode=True)
            if payload:
                body = payload.decode(charset, errors='ignore')
        except (UnicodeDecodeError, LookupError):
            try:
                body = payload.decode('utf-8', errors='ignore')
            except:
                body = str(message.get_payload())
    
    return body.strip()

def process_mbox_file(mbox_file):
    """Process mbox file and extract email data"""
    emails_data = []
    
    # Create temporary file to work with mbox
    with tempfile.NamedTemporaryFile(delete=False, suffix='.mbox') as temp_file:
        temp_file.write(mbox_file.read())
        temp_file_path = temp_file.name
    
    try:
        # Open the mbox file
        mbox = mailbox.mbox(temp_file_path)
        
        for i, message in enumerate(mbox):
            try:
                # Extract basic information
                subject = decode_mime_words(message.get('Subject', ''))
                sender = decode_mime_words(message.get('From', ''))
                recipient = decode_mime_words(message.get('To', ''))
                cc = decode_mime_words(message.get('Cc', ''))
                bcc = decode_mime_words(message.get('Bcc', ''))
                date_str = message.get('Date', '')
                
                # Parse date
                parsed_date = None
                if date_str:
                    try:
                        parsed_date = email.utils.parsedate_to_datetime(date_str)
                    except:
                        parsed_date = None
                
                # Extract body
                body = extract_email_body(message)
                
                # Get message ID
                message_id = message.get('Message-ID', '')
                
                # Get other headers
                reply_to = decode_mime_words(message.get('Reply-To', ''))
                in_reply_to = message.get('In-Reply-To', '')
                references = message.get('References', '')
                
                email_data = {
                    'Index': i + 1,
                    'Message-ID': message_id,
                    'Subject': subject,
                    'From': sender,
                    'To': recipient,
                    'Cc': cc,
                    'Bcc': bcc,
                    'Reply-To': reply_to,
                    'Date': parsed_date.isoformat() if parsed_date else date_str,
                    'Date_Raw': date_str,
                    'In-Reply-To': in_reply_to,
                    'References': references,
                    'Body': body,
                    'Body_Length': len(body)
                }
                
                emails_data.append(email_data)
                
            except Exception as e:
                st.warning(f"Error processing email {i+1}: {str(e)}")
                continue
    
    finally:
        # Clean up temporary file
        try:
            os.unlink(temp_file_path)
        except:
            pass
    
    return emails_data

def convert_to_csv(emails_data):
    """Convert emails data to CSV format"""
    df = pd.DataFrame(emails_data)
    return df.to_csv(index=False)

def convert_to_excel(emails_data):
    """Convert emails data to Excel format"""
    df = pd.DataFrame(emails_data)
    
    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Emails', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Emails']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            # Set column width (with some padding)
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()

def convert_to_txt(emails_data):
    """Convert emails data to readable text format"""
    txt_content = []
    
    for email in emails_data:
        txt_content.append("=" * 80)
        txt_content.append(f"Email #{email['Index']}")
        txt_content.append("=" * 80)
        txt_content.append(f"Message-ID: {email['Message-ID']}")
        txt_content.append(f"Subject: {email['Subject']}")
        txt_content.append(f"From: {email['From']}")
        txt_content.append(f"To: {email['To']}")
        if email['Cc']:
            txt_content.append(f"Cc: {email['Cc']}")
        if email['Bcc']:
            txt_content.append(f"Bcc: {email['Bcc']}")
        if email['Reply-To']:
            txt_content.append(f"Reply-To: {email['Reply-To']}")
        txt_content.append(f"Date: {email['Date']}")
        txt_content.append("-" * 80)
        txt_content.append(email['Body'])
        txt_content.append("\n")
    
    return "\n".join(txt_content)

def main():
    st.set_page_config(
        page_title="DCI Mbox Converter",
        page_icon="📧",
        layout="wide"
    )
    
    st.title("📧 DCI Mbox Converter")
    st.markdown("Convert your mbox files to Excel, CSV, or TXT format with **no limits**!")
    
    # File upload
    st.header("Upload Mbox File")
    uploaded_file = st.file_uploader(
        "Choose an mbox file",
        type=['mbox'],
        help="Upload your mbox file to convert it to Excel, CSV, or TXT format"
    )
    
    if uploaded_file is not None:
        st.success(f"File uploaded: {uploaded_file.name}")
        
        # Processing options
        st.header("Conversion Options")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            convert_to_excel_btn = st.checkbox("Convert to Excel (.xlsx)", value=True)
        with col2:
            convert_to_csv_btn = st.checkbox("Convert to CSV (.csv)", value=True)
        with col3:
            convert_to_txt_btn = st.checkbox("Convert to TXT (.txt)", value=True)
        
        if st.button("🚀 Start Conversion", type="primary"):
            try:
                # Show progress
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                status_text.text("Processing mbox file...")
                progress_bar.progress(20)
                
                # Process the mbox file
                emails_data = process_mbox_file(uploaded_file)
                
                if not emails_data:
                    st.error("No emails found in the mbox file or file could not be processed.")
                    return
                
                progress_bar.progress(60)
                status_text.text(f"Found {len(emails_data)} emails. Generating downloads...")
                
                # Display summary
                st.success(f"✅ Successfully processed {len(emails_data)} emails!")
                
                # Create download buttons
                st.header("Download Converted Files")
                
                # Prepare filename base
                base_filename = clean_filename(uploaded_file.name.replace('.mbox', ''))
                
                col1, col2, col3 = st.columns(3)
                
                if convert_to_excel_btn:
                    with col1:
                        excel_data = convert_to_excel(emails_data)
                        st.download_button(
                            label="📊 Download Excel",
                            data=excel_data,
                            file_name=f"{base_filename}_emails.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                if convert_to_csv_btn:
                    with col2:
                        csv_data = convert_to_csv(emails_data)
                        st.download_button(
                            label="📄 Download CSV",
                            data=csv_data,
                            file_name=f"{base_filename}_emails.csv",
                            mime="text/csv"
                        )
                
                if convert_to_txt_btn:
                    with col3:
                        txt_data = convert_to_txt(emails_data)
                        st.download_button(
                            label="📝 Download TXT",
                            data=txt_data,
                            file_name=f"{base_filename}_emails.txt",
                            mime="text/plain"
                        )
                
                progress_bar.progress(100)
                status_text.text("Conversion completed successfully! 🎉")
                
                # Show preview
                st.header("Preview")
                if emails_data:
                    df_preview = pd.DataFrame(emails_data)
                    # Show first few columns for preview
                    preview_cols = ['Index', 'Subject', 'From', 'To', 'Date', 'Body_Length']
                    available_cols = [col for col in preview_cols if col in df_preview.columns]
                    st.dataframe(df_preview[available_cols].head(10), use_container_width=True)
                    
                    if len(emails_data) > 10:
                        st.info(f"Showing first 10 of {len(emails_data)} emails")
                
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")
                st.error("Please ensure you've uploaded a valid mbox file.")
    
    # Information section
    st.header("About DCI Mbox Converter")
    st.markdown("""
    This tool converts mbox (mailbox) files to Excel, CSV, or TXT formats with **no conversion limits**.
    
    **Features:**
    - ✅ Convert unlimited number of emails
    - 📊 Export to Excel with auto-formatted columns
    - 📄 Export to CSV for data analysis
    - 📝 Export to readable TXT format
    - 🔧 Handles MIME encoding and various character sets
    - 🛡️ Secure processing (files are processed locally)
    
    **Supported Data:**
    - Email headers (Subject, From, To, Cc, Bcc, etc.)
    - Email body content
    - Timestamps and Message IDs
    - Reply chains and references
    
    **File Formats:**
    - Input: `.mbox` files
    - Output: `.xlsx`, `.csv`, `.txt`
    """)

if __name__ == "__main__":
    main()