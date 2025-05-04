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
- `os`: for handling file and directory operations, such as saving and deleting file
- `python-docx`: for enabling reading and editing Word files to populate invoice templates
- `docx2pdf`: for converting the generated Word files into PDF files for final delivery

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

### 4. Populate Word Template with Invoice Data
Using the invoice dictionary, the script replaces all matching placeholders in the Word template with actual data. This dynamic substitution generates a personalized invoice Word file for each entry, preserving the original layout and formatting defined in the template.
The generated Word invoice is saved as a new file, while the original template remains unchanged. This ensures that each new invoice starts from a clean, unaltered template for consistent formatting and accurate substitutions.

> [!NOTE]  
> Handling Placeholder Substitution Issues in Word Templates

> When using Word templates for invoice generation, placeholders like UNITPRICE1 may not always be stored as a single contiguous string. Instead, Word can split them into multiple runs (e.g., UNIT, PRICE, 1), especially if the placeholder is manually typed letter-by-letter or if formatting changes occur mid-text. This makes accurate substitution difficult.

> To address this, there are two possible solutions:

> Best Practice: Always paste the full placeholder (e.g., UNITPRICE1) into the Word template instead of typing it character by character. This helps Word treat it as a single run.
Programmatic Workaround: Merge all runs in a paragraph into one string, perform substitutions on the combined text, and then rewrite the paragraph with the updated content. However, this method overwrites the original formatting of the paragraph.
Here’s the code implementation of the workaround:

```python
# define the function of replacing the template placeholder with invoice information
def replace_text_in_doc_table(item_dict, docx_template_name):

    # open the template docx
    doc = Document(docx_template_name)

    # replace the placeholder in the docx for all the invoices information
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:

                    # combine all the run as a full text
                    full_text = ''.join([run.text for run in para.runs])

                    if full_text in item_dict.keys():
                        full_text = item_dict[full_text]

                    # clear the para
                    para.clear()
                    para.add_run(full_text)

    # add the docx path and docx name
    new_doc_path = item_dict['INV_NO'] + '.docx'

    # save the docx to the docx path
    doc.save(new_doc_path)

replace_text_in_doc_table(item_dict, 'inv_template.docx')
```

### 5. Convert Word Files to PDF Files
After each invoice is generated as a Word file, it is immediately converted into a PDF file for final output. To keep the workspace clean and prevent duplication, the intermediate Word file is automatically deleted after the PDF is successfully created.

## License
This project is licensed under the MIT License - see the [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/leopengningchuan/invoice-automation?tab=MIT-1-ov-file) file for details.

## Acknowledgements
- Thanks to Microsoft Word for providing a flexible document format that allows for easy templating.
- Thanks to the Python community for the powerful libraries that made this project possible, including [`python-docx`](https://pypi.org/project/docx2pdf/) and [`openpyxl`](https://pypi.org/project/docx2pdf/).
