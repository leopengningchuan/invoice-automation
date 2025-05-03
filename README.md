# Invoice Automation
Mass invoice generation and automation tool

## Project Background
In many businesses, invoice generation is a critical but often time-consuming task, especially when handling a large volume of customer transactions. Manual processing not only increases the risk of errors but also leads to significant inefficiencies in operations.

To address these challenges, this project was developed to automate the invoice generation process. By streamlining data integration and automating invoice creation, companies can reduce manual workload, minimize errors, and improve overall operational efficiency.

## Project Goal
This project aims to automate the generation of company invoices in PDF by using **Python Jupyter Notebook**, **Microsoft Excel** and **Microsoft Word**. It is designed to process large-scale customer sales data efficiently, ensuring accurate, consistent, and scalable invoice creation with minimal manual intervention.

## Table of Contents
- README.md
- LICENSE.txt
- Invoice Automation.ipynb
- inv_info_sample.xlsx
- inv_template.docx
- SAMCO_1–4.pdf (sample invoices generated)

## Instructions

### 1. Package Used
- `pandas, datetime, re`: for data manipulation
- `os`: for file management
- `docx`: for using Microsoft Word as a template
- `docx2pdf`: for converting a Microsoft Word file to a PDF file

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

### 3. Parse the Information





### 4. Convert Word Files to PDF
After each invoice is generated as a Word document, it is immediately converted into a PDF file for final output. To keep the workspace clean and prevent duplication, the intermediate Word file is automatically deleted after the PDF is successfully created.



## License
This project is licensed under the MIT License - see the [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) file for details.

## Acknowledgements
- Thanks to [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) for web scraping.
- Thanks to [Douban](https://www.douban.com) for providing the platform.
