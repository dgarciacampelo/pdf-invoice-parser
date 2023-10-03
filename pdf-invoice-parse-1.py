import os
import pdfplumber
import re
from openpyxl import Workbook

# Global dictionary to store data for each CUPS
global_cups_data = {}

# Threshold of rows of data to be saved in a single XLSX file
PENDING_ROWS_THRESHOLD: int = 100


# Function to process a PDF file and extract data
def process_pdf(pdf_file_path):
    text = ""
    # Open the PDF file with pdfplumber
    with pdfplumber.open(pdf_file_path) as pdf:
        num_pages = len(pdf.pages)
        print(f"Processing {pdf_file_path} ({num_pages} pages)...")

        for page in pdf.pages:
            text += page.extract_text()

    # Initialize a dictionary to store extracted data
    extracted_data = dict()

    # Extract Datos del Titular using regex
    datos_titular_match = re.search(
        r"DATOS DEL TITULAR\s*(.*?)\s*Nº Factura", text, re.DOTALL
    )
    if datos_titular_match:
        extracted_data["Datos del Titular"] = datos_titular_match.group(1).strip()

    # Extract Nº Factura using regex
    invoice_number_match = re.search(r"Nº Factura:\s*(\w+)", text)
    if invoice_number_match:
        extracted_data["Nº Factura"] = invoice_number_match.group(1)

    # Extract Período de facturación using regex
    billing_period_match = re.search(
        r"Período de facturación:\s*([\d/]+ - [\d/]+)", text
    )
    if billing_period_match:
        extracted_data["Período de facturación"] = billing_period_match.group(1)

    # Extract CUPS and Ref. Contrato Acceso using regex
    cups_match = re.search(
        r"CUPS:\s*(.*?)\s*Ref. Contrato Acceso:\s*(\d+)", text, re.DOTALL
    )
    if cups_match:
        extracted_data["CUPS"] = cups_match.group(1).strip()
        extracted_data["Ref. Contrato Acceso"] = cups_match.group(2).strip()

    # Extract the line starting with "P1. Energía activa"
    p1_energia_activa_line = re.search(r"P1\. Energía activa[^\n]*", text)

    if p1_energia_activa_line:
        line_text = p1_energia_activa_line.group(0)
        # Split the line by whitespace to obtain the values
        values = re.findall(r"[\d,.]+", line_text)

        if len(values) > 3:
            # values[0] should be 1 from P1.
            extracted_data["P1. Energía activa (Cantidad)"] = values[1]
            extracted_data["P1. Energía activa (Precio/u)"] = values[2]
            extracted_data["P1. Energía activa (Total)"] = values[3]

    return extracted_data


# Function to save data to an XLSX file
def save_to_xlsx(data, cups_code):
    if cups_code not in global_cups_data:
        global_cups_data[cups_code] = []

    global_cups_data[cups_code].append(data)

    # Check if the number of rows for this CUPS reaches a threshold
    if len(global_cups_data[cups_code]) >= PENDING_ROWS_THRESHOLD:
        save_data_to_xlsx(cups_code)


# Function to save pending data for a CUPS to an XLSX file
def save_data_to_xlsx(cups_code):
    if cups_code in global_cups_data:
        # Create a new workbook and select the active worksheet
        workbook = Workbook()
        worksheet = workbook.active

        # Write column headers to the worksheet
        column_headers = [
            "Datos del Titular",
            "Nº Factura",
            "Período de facturación",
            "CUPS",
            "Ref. Contrato Acceso",
            "P1. Energía activa (Cantidad)",
            "P1. Energía activa (Precio/u)",
            "P1. Energía activa (Total)",
        ]
        worksheet.append(column_headers)

        # Append the data for this CUPS to the worksheet
        for data in global_cups_data[cups_code]:
            data_row = [data.get(key, "") for key in column_headers]
            worksheet.append(data_row)

        # Define the XLSX file path based on the CUPS
        xlsx_file_path = os.path.join(cwd, f"{cups_code}.xlsx")

        # Save the workbook as an XLSX file
        workbook.save(xlsx_file_path)
        print(f"XLSX file '{xlsx_file_path}' created successfully.")

        # Clear the saved data for this CUPS
        del global_cups_data[cups_code]


if __name__ == "__main__":
    # Get the current working directory
    cwd = os.getcwd()

    # List all PDF files in the current directory
    pdf_files = [file for file in os.listdir(cwd) if file.endswith(".pdf")]

    # Process each PDF file and store the data
    for pdf_file in pdf_files:
        pdf_file_path = os.path.join(cwd, pdf_file)

        try:
            extracted_data = process_pdf(pdf_file_path)
            cups_code = extracted_data.get("CUPS", "Unknown")
            save_to_xlsx(extracted_data, cups_code)
        except Exception as e:
            print(f"Error processing {pdf_file}: {str(e)}")

    # Save any remaining pending data to XLSX files
    """
    By creating a list from the keys before iterating, you are iterating over a
    static list of keys, and you can safely modify the original global_cups_data
    dictionary inside the loop without encountering a RuntimeError:
    # ! dictionary changed size during iteration
    """
    for cups_code in list(global_cups_data.keys()):
        save_data_to_xlsx(cups_code)
