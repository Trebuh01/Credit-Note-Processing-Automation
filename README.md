# Credit Note Processing Automation

## Project Overview
This project consists of two Python programs designed to automate the process of downloading and analyzing credit note PDFs from an email account. The goal is to filter and process credit notes based on specific criteria and automatically saving or discarding PDFs based on the results.

The project uses the following technologies:
- **Python**: The core programming language used.
- **IMAP**: For fetching emails from a WP email server.
- **PyMuPDF (fitz)**: For extracting text from PDFs.
- **Geopy**: For geocoding and route planning in the first program.
- **Shutil**: For file operations, such as moving and deleting files.

---

## Program 1: Credit Note Processing with Geolocation and Route Validation

### Purpose
This program automates the process of downloading credit note PDFs, extracting key information (such as cities mentioned after a `500` code), and comparing extracted routes with the actual route information fetched from the Maps API.

### Key Features:
1. **Download Attachments**: Downloads PDFs with the subject containing `CREDITNOTE` from a WP email inbox.
2. **Extract Data**: Extracts the `Total Amount`, starting city, and ending city from the PDF. It also handles cases like cities with slashes (`/`) or hyphens (`-`).
3. **Route Comparison**: Uses the Maps API to check if the route in the credit note corresponds to the actual driving route between the starting and ending cities.
4. **Error Handling**: If there are missing data points (e.g., missing fuel surcharge, road tax, or incorrect route), the PDF is moved to a folder with an error suffix added to the filename.
5. **File Management**: PDFs with issues are moved to a `potencjalne_bledy` folder, with errors added to the file name.

### How to Use:
1. Update your **email credentials** and **Maps API key** in the code.
2. Run the program to download PDFs, extract the required data, and validate routes using the Maps API.
3. The program will categorize PDFs based on validation results and move them to appropriate folders.

## Program 2: Negative Total Amount Credit Note Finder
### Purpose
This program downloads all credit note PDFs from an email account for a specified date range, checks for credit notes with negative Total Amount values, and separates them into a different folder for further processing or archiving.

### Key Features:
1. **Download All Credit Notess**: Downloads all credit note PDFs from a WP email account for a user-specified date range.
2. **Identify Negative Amounts**: Scans each PDF to determine if it contains a negative Total Amount value.
3. **Organize PDFs**: Moves PDFs with negative Total Amount values to a separate folder (negative_creditnotes) for review or archiving.
4. **Clean-up**: Deletes all files from the all_creditnotes folder after processing.
