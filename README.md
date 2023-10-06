# Google Apps Script Invoice Automation

This Google Apps Script project automates invoice creation, management, and email notifications. It includes the following functions:

- **Submit Invoice**: This function submits invoices to the system. If the invoice ID (cell O2) is blank, it creates a new invoice. If not, it updates an existing invoice.

- **Make PDF**: This function generates a PDF document for the invoice and saves it in a specified folder on Google Drive.

- **Make PDF Order**: Similar to "Make PDF," this function generates a PDF document for the order and saves it in the same folder on Google Drive.

- **Make History**: This function logs the invoice details in a "Tagihan" sheet.

- **Copy Rows**: Copies invoice data to a "Produk Tagihan" sheet for tracking sold products.

- **Clear Invoice Fields**: Clears invoice-related fields to prepare for new entries.

- **Search**: Searches for invoice records based on the invoice ID (cell O2).

- **Dropdown List**: Populates dropdown lists for product sizes based on the selected product.

- **Dropdown Return List**: Populates dropdown lists for product sizes in the "Return" sheet.

- **Dropdown Label List**: Populates dropdown lists for product sizes in the "Pelabelan" sheet.

- **Dropdown Packing List**: Populates dropdown lists for product sizes in the "Packing" sheet.

- **Delete Row With Empty Values**: Deletes rows in the "Produk Tagihan" sheet if they contain empty values.

- **Create Trigger**: Creates a trigger to run the "Delete Row With Empty Values" function on edits.

- **On Edit**: Runs various functions on spreadsheet edits, including deleting rows with empty values and updating dropdown lists.

- **Get Gmail Emails**: Retrieves emails related to invoices and logs them in an "Email List" sheet.

- **GDrive Files**: Lists all files in a specified Google Drive folder and logs their details in the "Link" sheet.

- **Move Files To Folder**: Moves files to the corresponding folders based on the information in the "Link" sheet.

## Usage

1. Open your Google Sheets document containing the "Invoice," "Tagihan," "Produk Tagihan," "Return," "Pelabelan," "Packing," and "Link" sheets.

2. Ensure that the script is correctly associated with your Google Sheets document.

3. Run the `submitInvoice()` function to submit or update invoices.

4. Use other functions as needed for generating PDFs, managing history, and more.

## Important Notes

- This script is designed for specific use cases and may require adjustments for your specific needs.

- Be cautious when running the script, as it may modify your data.

## Author

- Musnida Ulya

