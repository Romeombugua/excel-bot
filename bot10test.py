import pytesseract
from pdf2image import convert_from_path
import pdfplumber
import re
from datetime import datetime
import xlwings as xw

# Extract free-form text from page 1 using OCR
def extract_text_from_first_page(pdf_path):
    images = convert_from_path(pdf_path, first_page=1, last_page=1)
    ocr_text = pytesseract.image_to_string(images[0])
    return ocr_text

# Function to extract key data points from the OCR text (first page)
def extract_data_from_text(text):
    dba_name_match = re.search(r"Trading as\s+(.+)\s+Customer reference", text)
    dba_name = dba_name_match.group(1).strip() if dba_name_match else ""

    registered_name_match = re.search(r"^(.+)\s+Trading as", text, re.MULTILINE)
    registered_name = registered_name_match.group(1).strip() if registered_name_match else ""

    address_match = re.search(r"Customer reference\s+\w+\n(.+?)\s+(Merchant ID|Invoice number)", text)
    address = address_match.group(1).strip() if address_match else ""
    postcode_match = re.search(r"\b[A-Z]{1,2}[0-9R][0-9A-Z]? ?[0-9][A-Z]{2}\b", text)
    postcode = postcode_match.group(0).strip() if postcode_match else ""

    invoice_date_match = re.search(r"Invoice date\s+(\d{2} \w+ \d{4})", text)
    invoice_date = invoice_date_match.group(1) if invoice_date_match else ""
    original_statement_date = re.search(r"Invoice period\s+\d{2} \w+ to \d{2} \w+ (\d{4})", text)
    original_statement = f"{invoice_date.split()[1]}-{original_statement_date.group(1)}" if invoice_date and original_statement_date else ""

    current_provider = "Dojo"

    data = {
        "DBA Name": f"{dba_name} : {address}",
        "DBA Address": address,
        "DBA POST CODE": postcode,
        "Analsys Date": datetime.today().strftime('%d-%b-%y'),
        "Date of Original Statement": original_statement,
        "Name of Owner": "",
        "Title": "Owner",
        "Date of Quote Generated": "",
        "Current Provider": current_provider,
        "Registered Business Name": registered_name,
        "Address": address,
        "Post Code": postcode
    }
    print("Extracted and populated Data:")
    print(data)

    return data

# Function to extract specific financial fields from the OCR text (first page)
def extract_page1_fields(text):
    page1_fields = {
        "Card transactions": None,
        "Card machine services": None,
        "Net amount": None,
        "VAT": None,
        "Total due": None
    }

    for field in page1_fields.keys():
        match = re.search(rf"{field}\s+(\£[0-9,.]+)", text)
        if match:
            page1_fields[field] = match.group(1)

    return page1_fields

# Function to extract card type data from the OCR text (subsequent pages)
def extract_tables_from_other_pages(pdf_path):
    all_table_data = []
    service_type_seen = False
    secure_transaction_fees_sum = {
        'Fee type': 'Sum Secure Transaction Fee',
        'Quantity': 0,
        'Fee per transaction': 0.0,
        'Total': 0.0,
        'VAT code': 'S'
    }

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages[1:], start=2):
            tables = page.extract_tables()

            if tables:
                for table in tables:
                    headers = table[0]
                    for row in table[1:]:
                        row_data = {headers[col]: row[col].strip() if row[col] else None for col in range(len(headers))}

                        if 'Service type' in row_data:
                            if service_type_seen:
                                continue
                            else:
                                service_type_seen = True
                            service_data = row_data['Service type'].split()
                            if len(service_data) >= 4:
                                row_data['Service type'] = ' '.join(service_data[:-3])
                                row_data['Quantity'] = service_data[-3]
                                row_data['Total'] = service_data[-2].replace('£', '')
                                row_data['VAT code'] = service_data[-1]

                        if 'Card type' in row_data and row_data['Card type']:
                            card_type_data = row_data['Card type'].split()
                            if len(card_type_data) >= 6:
                                num_trans_index = next((i for i, item in enumerate(card_type_data) if item.replace(',', '').isdigit()), -1)
                                if num_trans_index != -1:
                                    row_data['Card type'] = ' '.join(card_type_data[:num_trans_index]).strip()
                                    row_data['Number of transactions'] = card_type_data[num_trans_index]
                                    row_data['Total value of transactions'] = card_type_data[num_trans_index + 1]
                                    rate_parts = card_type_data[num_trans_index + 2:]
                                    rate_info = ' '.join(rate_parts)
                                    percentage_rate = next((part for part in rate_parts if '%' in part), None)
                                    row_data['Rate per transaction'] = percentage_rate if percentage_rate else ''
                                    if '+' in rate_info:
                                        fee_parts = rate_info.split('+')
                                        if len(fee_parts) > 1:
                                            fee_per_trx = fee_parts[1].strip().split()[0]
                                            row_data['Fee per TRX'] = fee_per_trx
                                    remaining = [part for part in rate_parts if part not in [percentage_rate, fee_per_trx] and '£' in part]
                                    if len(remaining) >= 2:
                                        row_data['Total'] = remaining[0]
                                        row_data['VAT code'] = remaining[1]
                                    elif len(remaining) == 1:
                                        row_data['Total'] = remaining[0]

                        if 'Fee type' in row_data and row_data['Fee type']:
                            fee_data = row_data['Fee type'].split()
                            if len(fee_data) >= 5:
                                row_data['Fee type'] = ' '.join(fee_data[:-4])
                                row_data['Quantity'] = fee_data[-4]
                                row_data['Fee per transaction'] = fee_data[-3].replace('£', '')
                                row_data['Total'] = fee_data[-2].replace('£', '')
                                row_data['VAT code'] = fee_data[-1]

                        all_table_data.append(row_data)

                        if 'Fee type' in row_data and 'Secure Transaction Fee' in row_data['Fee type']:
                            quantity = int(row_data['Quantity']) if row_data['Quantity'] else 0
                            fee_per_transaction = float(row_data['Fee per transaction']) if row_data['Fee per transaction'] else 0.0
                            total = float(row_data['Total']) if row_data['Total'] else 0.0
                            secure_transaction_fees_sum['Quantity'] += quantity
                            secure_transaction_fees_sum['Fee per transaction'] += fee_per_transaction
                            secure_transaction_fees_sum['Total'] += total

    return all_table_data, secure_transaction_fees_sum

def extract_total_card_volume(table_data):
    for entry in table_data:
        if 'Card type' in entry and entry['Card type'] and 'Subtotal' in entry['Card type']:
            match = re.search(r'£([\d,]+\.\d+)', entry['Card type'])
            if match:
                return float(match.group(1).replace(',', ''))
    return None

# Function to process a PDF and extract all necessary data
def process_pdf(pdf_path):
    first_page_text = extract_text_from_first_page(pdf_path)
    extracted_data = extract_data_from_text(first_page_text)
    page1_fields = extract_page1_fields(first_page_text)
    table_data, secure_transaction_fees_sum = extract_tables_from_other_pages(pdf_path)
    total_card_volume = extract_total_card_volume(table_data)

    services_for_dojo_go = next((entry for entry in table_data if entry.get('Service type') == 'Services for Dojo Go'), None)

    return extracted_data, page1_fields, secure_transaction_fees_sum, services_for_dojo_go, total_card_volume, table_data

# Function to update Excel with the extracted data using xlwings
def update_excel_with_data(excel_path, extracted_data, sum_secure_transaction_fee, services_for_dojo_go, page1_fields, total_card_volume, table_data):
    app = xw.App(visible=False)
    try:
        wb = app.books.open(excel_path)
        sheet = wb.sheets[0]  # Assuming we're working with the first sheet

        # Update cells
        sheet.range("B6").value = extracted_data["DBA Name"]
        sheet.range("B7").value = extracted_data["DBA Address"]
        sheet.range("B8").value = extracted_data["DBA POST CODE"]
        sheet.range("B9").value = extracted_data["Analsys Date"]
        sheet.range("B10").value = extracted_data["Date of Original Statement"]
        sheet.range("B15").value = extracted_data["Registered Business Name"]
        sheet.range("B16").value = extracted_data["Address"]
        sheet.range("B17").value = extracted_data["Post Code"]
        sheet.range("B22").value = sum_secure_transaction_fee["Quantity"]
        sheet.range("C22").value = sum_secure_transaction_fee["Total"]
        sheet.range("B24").value = services_for_dojo_go["Quantity"] if services_for_dojo_go else ""
        sheet.range("C24").value = services_for_dojo_go["Total"] if services_for_dojo_go else ""
        sheet.range("C21").value = page1_fields.get("Card transactions", "")
        sheet.range("C25").value = page1_fields.get("Net amount", "")
        sheet.range("C26").value = page1_fields.get("VAT", "")
        sheet.range("C27").value = page1_fields.get("Total due", "")

        if total_card_volume:
            sheet.range("C28").value = f"£{total_card_volume:,.2f}"

        # Update the card rates breakdown table
        for record in table_data:
            if 'Card type' in record and record['Card type']:
                card_type = record['Card type']
                for row in sheet.range("A2:A100"):  # Adjust range as needed
                    if row.value == card_type:
                        row.offset(column=1).value = record.get('Number of transactions', None)
                        row.offset(column=2).value = record.get('Total value of transactions', None)
                        row.offset(column=3).value = record.get('Rate per transaction', None)
                        if 'Fee per TRX' in record:
                            row.offset(column=4).value = record['Fee per TRX']
                        row.offset(column=5).value = record.get('Total', None)
                        break

        wb.save()
        print("Excel file updated successfully")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        wb.close()
        app.quit()

# Main execution
if __name__ == "__main__":
    pdf_path = "/path/to/your/Dojo invoice.pdf"
    excel_path = "/path/to/your/main.xlsx"

    extracted_data, page1_fields, sum_secure_transaction_fee, services_for_dojo_go, total_card_volume, table_data = process_pdf(pdf_path)
    update_excel_with_data(excel_path, extracted_data, sum_secure_transaction_fee, services_for_dojo_go, page1_fields, total_card_volume, table_data)
