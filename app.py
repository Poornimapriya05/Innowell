import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import imaplib
import email
from email.header import decode_header
import pdfplumber
import pandas as pd
from pymongo import MongoClient
import numpy as np
from fpdf import FPDF
import re
import os

# Email credentials
smtp_host = "smtp.gmail.com"
imap_host = "imap.gmail.com"
email_address = "poornimavadivel190@gmail.com"
password = "rhkyqcmpxnymjctc"

# Directory to save attachments
attachments_dir = "C:/Users/prade/OneDrive/Desktop/inno"
os.makedirs(attachments_dir, exist_ok=True)

# Specify the attachment to download
specific_filename = "input_doc.pdf"
specific_filetype = ".pdf"

def send_email_with_attachment(receiver_email, subject, body, file_path, csv_path=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = email_address
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        def attach_file(filepath, filename):
            with open(filepath, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={filename}")
                msg.attach(part)

        attach_file(file_path, os.path.basename(file_path))
        if csv_path:
            attach_file(csv_path, os.path.basename(csv_path))

        with smtplib.SMTP(smtp_host, 587) as server:
            server.starttls()
            server.login(email_address, password)
            server.sendmail(email_address, receiver_email, msg.as_string())
        print("Email sent successfully!")
    except Exception as e:
        print("Failed to send email:", e)

def download_specific_attachments():
    try:
        with imaplib.IMAP4_SSL(imap_host) as mail:
            mail.login(email_address, password)
            mail.select("inbox")

            status, messages = mail.search(None, 'ALL')
            email_ids = messages[0].split()

            if email_ids:
                latest_email_id = email_ids[-1]
                status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject = decode_header(msg["Subject"])[0][0]
                        if isinstance(subject, bytes):
                            subject = subject.decode()
                        print("Processing latest email with subject:", subject)

                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_disposition() == "attachment":
                                    filename = part.get_filename()
                                    if filename:
                                        if filename == specific_filename or filename.endswith(specific_filetype):
                                            filepath = os.path.join(attachments_dir, filename)
                                            with open(filepath, "wb") as f:
                                                f.write(part.get_payload(decode=True))
                                            print(f"Attachment {filename} saved to {filepath}")
                                        else:
                                            print(f"Skipping attachment: {filename}")
            else:
                print("No emails found in the inbox.")
    except Exception as e:
        print("Failed to download attachments:", e)

def clean_text(text):
    return text.replace('–', '-')

def extract_field(text, field_name):
    pattern = rf"{field_name}:\s*(.*?)\n"
    match = re.search(pattern, text, re.IGNORECASE)
    return clean_text(match.group(1).strip()) if match else "Not Found"

def process_pdf():
    pdf_path = "C:/Users/prade/OneDrive/Desktop/inno/attachments/input_doc.pdf"
    csv_path = 'extracted_table.csv'
    final_pdf_path = "final_data.pdf"
    image_path = "extracted_image.png"

    try:
        with pdfplumber.open(pdf_path) as pdf:
            table = pdf.pages[3].extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                df.to_csv(csv_path, index=False)
    except Exception as e:
        print(f"Error extracting table from PDF: {e}")

    try:
        client = MongoClient('mongodb://localhost:27017/')
        db = client['costume']
        collection = db['spec']
        df = pd.read_csv(csv_path)
        collection.insert_many(df.to_dict('records'))
    except Exception as e:
        print(f"Error storing data in MongoDB: {e}")

    try:
        data = list(collection.find())
        df_retrieved = pd.DataFrame(data)
    except Exception as e:
        print(f"Error retrieving data from MongoDB: {e}")

    try:
        df_retrieved['per_rate'] = np.random.uniform(0, 2, size=len(df_retrieved))
        if 'Qty' in df_retrieved.columns:
            df_retrieved['Total'] = df_retrieved['Qty'].astype(float) * df_retrieved['per_rate']
        else:
            print("Column 'Qty' not found in the DataFrame.")
            df_retrieved['Total'] = np.nan
        df_retrieved['per_rate'] = df_retrieved['per_rate'].astype(float).fillna(0).round(2)
        df_retrieved['Total'] = df_retrieved['Total'].astype(float).fillna(0).round(2)
        df_retrieved = df_retrieved.fillna('')
        if '_id' in df_retrieved.columns:
            df_retrieved.drop(columns=['_id'], inplace=True)
        df_limited = df_retrieved[['Placement', 'Composition', 'Qty', 'per_rate', 'Total']].head(8)
        total_cost = df_limited['Total'].astype(float).sum()
    except Exception as e:
        print(f"Error processing data: {e}")


    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()

            # Extract fields
            style = extract_field(text, "Style")
            style_number = extract_field(text, "Style number")
            brand = extract_field(text, "Brand")
            sizes = extract_field(text, "Sizes")
            commodity = extract_field(text, "Commodity")
            email = extract_field(text, "E-mail")
            care_address = extract_field(text, "Care Address")

            # Extract image (if exists)
            images = page.images
            if images:
                image = images[0]
                x0, y0, x1, y1 = image['x0'], image['top'], image['x1'], image['bottom']
                img = page.within_bbox((x0, y0, x1, y1)).to_image()
                img.save(image_path)
    except Exception as e:
        print(f"Error extracting additional information: {e}")

    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial",style='B', size=10)

        if os.path.exists(image_path):
            pdf.image(image_path, x=150, y=10, w=40)

        pdf.cell(200, 10, txt="Costing Sheet", ln=True, align='C')

        # Use a set to track added fields
        already_added = set()

        # Define fields with labels
        fields_to_add = [
            ("Style", style),
            ("Style Number", style_number),
            ("Brand", brand),
            ("Sizes", sizes),
            ("Commodity", commodity),
            ("E-mail", email),
            ("Care Address", care_address),
        ]

        # Add fields to PDF
        added_fields = set()  # Track added label-value pairs
        for label, value in fields_to_add:
            field_pair = f"{label}:{value}"  # Combine label and value to ensure uniqueness
            if field_pair not in added_fields and value != "Not Found":
                pdf.set_font("Arial", style='B', size=10)  # Bold for the label
                pdf.cell(40, 10, txt=f"{label}:", ln=False)  # Keep the label and value on the same line
                pdf.set_font("Arial", size=10)  # Normal for the value
                pdf.cell(0, 10, txt=value, ln=True)  # Add the value right after the label without extra spaces
                added_fields.add(field_pair) 

        pdf.set_font("Arial",style='B',size=10)
        pdf.cell(200, 10, txt="Spec Sheet:", ln=True, align='L')
        pdf.cell(0, 10, txt="", ln=True)
        

        # Add table header
        column_widths = {"Placement": 40, "Composition": 70, "Qty": 25, "per_rate": 25, "Total": 25}
        pdf.set_font("Arial", style='B', size=10)  # Bold for headers
        for column in ["Placement", "Composition", "Qty", "per_rate", "Total"]:
            pdf.cell(column_widths[column], 10, column, border=1, align='C')
        pdf.ln()

        # Add table rows
        pdf.set_font("Arial", size=10)  # Normal for data rows
        for _, row in df_limited.iterrows():
            for column in ["Placement", "Composition", "Qty", "per_rate", "Total"]:
                pdf.cell(column_widths[column], 10, str(row[column]), border=1, align='C')
            pdf.ln()

        # Add total cost
        pdf.cell(column_widths['Placement'], 10, '', border=1, align='C')
        pdf.cell(column_widths['Composition'], 10, '', border=1, align='C')
        pdf.cell(column_widths['Qty'], 10, '', border=1, align='C')
        pdf.cell(column_widths['per_rate'], 10, 'Total', border=1, align='C')
        pdf.cell(column_widths['Total'], 10, f"{total_cost:.2f}", border=1, align='C')
        pdf.ln()

        pdf.output(final_pdf_path)

        if os.path.exists(image_path):
            os.remove(image_path)

        print(f"Final data saved to '{final_pdf_path}'.")
    except Exception as e:
        print(f"Error generating PDF: {e}")
    
def send_generated_pdf():
    try:
        receiver_email = "poornimavadivel190@gmail.com"
        subject = "Costing Sheet"
        body = "Costumes PDF"
        send_email_with_attachment(receiver_email, subject, body, "final_data.pdf", "output.csv")
    except Exception as e:
        print(f"Error sending email with attachment: {e}")

if __name__ == "__main__":
    download_specific_attachments()
    process_pdf()
    send_generated_pdf()