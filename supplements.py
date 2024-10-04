import imaplib
import email
import os
import fitz  # PyMuPDF
import re
import shutil

DATE_SINCE = "01-Aug-2024"  
DATE_BEFORE = "01-Sep-2024"  


def clean_filename(filename):
    filename = re.sub(r'[\\/*?:"<>|]', "", filename)
    filename = filename.replace("\r", "").replace("\n", "")
    return filename


def download_all_creditnotes(username, password, download_folder='all_creditnotes', subject_filter="CREDITNOTE"):
    mail = imaplib.IMAP4_SSL('imap.wp.pl')
    mail.login(username, password)
    mail.select('inbox')

    search_query = f'(SINCE "{DATE_SINCE}" BEFORE "{DATE_BEFORE}")'
    typ, dane = mail.search(None, search_query)

    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    for num in dane[0].split():
        typ, dane = mail.fetch(num, '(RFC822)')
        email_msg = email.message_from_bytes(dane[0][1])
        email_subject = email_msg['subject']

        if subject_filter.upper() in email_subject.upper():
            print(f"Pobieranie załączników z maila o temacie: {email_subject}")

            typ, uid_data = mail.fetch(num, '(UID)')
            uid = uid_data[0].split()[2].decode()  

            for part in email_msg.walk():
                if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                    continue

                if part.get_content_type() == "application/pdf":
                    filename = part.get_filename()
                    filename = clean_filename(filename)

                    
                    if filename:
                        filename = f"{uid}_{filename}"

                    
                    filepath = os.path.join(download_folder, filename)

                    
                    print(f"Zapisywanie pliku: {filepath}")
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))

    mail.logout()


def contains_negative_total_amount(pdf_path):
    doc = fitz.open(pdf_path)
    text = ''

    for page in doc:
        text += page.get_text("text")

    lines = text.splitlines()

    for i, line in enumerate(lines):
        if "Total Amount" in line:
            if i + 1 < len(lines):
                amount_line = lines[i + 1].strip()
                try:
                    
                    amount = float(amount_line.replace(",", "."))
                    if amount < 0:
                        print(f"Znaleziono ujemną wartość: {amount_line}")
                        return True  
                except ValueError:
                    pass
    return False  


def copy_negative_creditnotes(source_folder='all_creditnotes', target_folder='negative_creditnotes'):
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    for filename in os.listdir(source_folder):
        file_path = os.path.join(source_folder, filename)

        if contains_negative_total_amount(file_path):
            target_path = os.path.join(target_folder, filename)
            shutil.copy(file_path, target_path)
            print(f"Plik {file_path} skopiowany do {target_path} z ujemną wartością 'Total Amount'.")


def clean_all_creditnotes_folder(folder='all_creditnotes'):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
                print(f"Plik {file_path} usunięty.")
        except Exception as e:
            print(f"Nie udało się usunąć pliku {file_path}. Błąd: {e}")


if __name__ == "__main__":
    username = ""  
    password = ""


    all_creditnotes_folder = 'all_creditnotes'
    negative_creditnotes_folder = 'negative_creditnotes'

    download_all_creditnotes(username, password, all_creditnotes_folder, "CREDITNOTE")

    copy_negative_creditnotes(all_creditnotes_folder, negative_creditnotes_folder)

    clean_all_creditnotes_folder(all_creditnotes_folder)