# Invoice Automation
Mass invoice generation and automation tool

## Table of Contents
- [Project Background](#project-background)
- [Project Goal](#project-goal)
- [File Structure](#file-structure)
- [Instructions](#instructions)
  - [1. Packages Used](#1-packages-used)
  - [2. Invoice Template Word and Invoice Info Excel](#2-invoice-template-word-and-invoice-info-excel)
  - [3. Build Invoice Data Dictionary](#3-build-invoice-data-dictionary)
  - [4. Word Invoice Generation with Placeholder Replacement](#4-word-invoice-generation-with-placeholder-replacement)
  - [5. Convert Word Files to PDF Files](#5-convert-word-files-to-pdf-files)
- [Future Improvements](#future-improvements)
- [License](#license)

## Project Background
In many businesses, invoice generation is a critical but often time-consuming task, especially when handling a large volume of customer transactions. Manual processing not only increases the risk of errors but also leads to significant inefficiencies in operations.

To address these challenges, this project was developed to automate the invoice generation process. By streamlining data integration and automating invoice creation, companies can reduce manual workload, minimize errors, and improve overall operational efficiency.

## Project Goal
This project aims to automate the generation of company invoices in PDF by using **Python Jupyter Notebook**, **Microsoft Excel** and **Microsoft Word**. It is designed to process large-scale customer sales data efficiently, ensuring accurate, consistent, and scalable invoice creation with minimal manual intervention.

## File Structure
- `README.md` – project overview
- `LICENSE.txt` – license information
- `.gitignore` – git ignore config
- `.gitattributes` – git attributes config
- `.gitmodules` – git submodules config
- [`utils/`](https://github.com/leopengningchuan/personal_utils) – submodule used
- `invoice_automation.ipynb` – notebook for invoice generation  
- `inv_info_sample.xlsx`– sample invoice input data  
- `inv_template.docx` – word invoice template
- `SAMCO_1–4.pdf` – sample output

## Instructions

### 1. Packages Used
- `pandas, datetime, re`: for data manipulation
- [`personal_utils.docx_manipulate`](https://github.com/leopengningchuan/personal_utils): for modifying Word files

### 2. Invoice Template Word and Invoice Info Excel
This project uses two key supporting files to generate customized invoices:
- `inv_template.docx`
This is the base Word template used for invoice generation. It contains placeholders (e.g., `CUSTOMER`, `DOC_DATE`, `AMOUNT1`, `TOTAL_AMOUNT`, etc.) that will be replaced with actual data from the Excel sheet.
The visual layout of the invoice—such as the company logo, table formatting, and footer—is pre-defined in this file. For the company issuing the invoice, information like the company name, address, and banking details can be edited directly within the Word file.

- `inv_info.xlsx`
This is the base Excel file contains structured invoice data. Each row represents one invoice, with columns corresponding to invoice fields:

| Invoice No. |  Customer  | Customer Address1 |    Customer Address2   | Invoice Date | Payment Terms |    Item   |       Detail      |               Unit Price               | Quantity |
|:-----------:|:----------:|:-----------------:|:----------------------:|:------------:|:-------------:|:---------:|:-----------------:|:--------------------------------------:|:--------:|
|   SAMCO_1   | Customer A |    123 ABC St.    |  Town_A, State_I 10001 |  2024/12/31  |       30      | Product 1 | Product 1 details |                                  9.00  |    170   |
|   SAMCO_2   | Customer B |  Unit 1, 456 Ave. | City_B, State_II 20002 |   2025/1/11  |       90      | Product 7 | Product 7 details |                               69.00    |    302   |

The program reads this data row by row and fills the template accordingly to generate one PDF file per invoice.

### 3. Build Invoice Data Dictionary
The program processes each row from the Excel file and converts it into a structured Python dictionary. Each key in the dictionary corresponds to a placeholder in the Word template (e.g., `CUSTOMER`, `DOC_DATE`, `AMOUNT1`, `TOTAL_AMOUNT`, etc.). This step ensures the data is clean, properly formatted (e.g., currency, dates), and ready for insertion into the template.

### 4. Word Invoice Generation with Placeholder Replacement
The invoice automation process uses `populate_docx_table()` from the `utils.docx_manipulate` module to dynamically replace placeholders in a Word docx template using values from an invoice dictionary.

For each invoice entry, the script fills in matching placeholders within table cells and generates a Word invoice. The output preserves the formatting and structure defined in the original template, which remains unchanged to ensure consistent formatting across all invoices.

The `utils/` folder is included as a Git submodule and contains a reusable function library maintained in the [`personal_utils`](https://github.com/leopengningchuan/personal_utils). You can refer to that repository for detailed function documentation and personal notes.

### 5. Convert Word Files to PDF Files
After each invoice is generated as a Word file, it is immediately converted into a PDF file for final output. To keep the workspace clean and prevent duplication, the intermediate Word file is automatically deleted after the PDF is successfully created.

## Future Improvements
- **Duplicate Invoice Detection**: Add a check to flag or prevent generation of duplicate invoices based on invoice number or client name.
- **Web Interface**: Build a lightweight front-end interface to allow non-technical users to upload invoice data and download results without using Jupyter or the command line.
- **Email Integration**: Add functionality to automatically send generated invoices via email to clients, with customizable messages and attachments.

## License
This project is licensed under the MIT License - see the [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/leopengningchuan/invoice-automation?tab=MIT-1-ov-file) file for details.
