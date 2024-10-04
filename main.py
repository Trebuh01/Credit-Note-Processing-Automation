import imaplib
import email
import time
import fitz  # PyMuPDF
from geopy.geocoders import Nominatim
import os
import openrouteservice
from openrouteservice import convert
import shutil
route_cache = {}
def clean_filename(filename):
    import re
    
    cleaned_filename = re.sub(r'[\\/*?:"<>|]', "", filename)
    cleaned_filename = cleaned_filename.replace("\r", "").replace("\n", "")
    return cleaned_filename

def download_attachments_from_wp(username, password, download_folder='pdf_files', subject_filter="CREDITNOTE"):
    
    mail = imaplib.IMAP4_SSL('imap.wp.pl')
    mail.login(username, password)
    mail.select('inbox')

    
    search_query = f'(SINCE "01-Apr-2024" BEFORE "01-May-2024")'
    typ, dane = mail.search(None, search_query)

    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

   
    for num in dane[0].split():
        typ, dane = mail.fetch(num, '(RFC822)')
        email_msg = email.message_from_bytes(dane[0][1])
        email_subject = email_msg['subject']

        
        if "CREDITNOTE" in email_subject.upper():  
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
                    print(f"Zapisywanie pliku do: {filepath}")  
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    print(f"Pobrano załącznik: {filename}")
        else:
            print(f"Temat nie zawiera słowa 'CREDITNOTE': {email_subject}")

    mail.logout()



def extract_data_from_pdf(pdf_path):
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
                        print(f"Ignorowanie pliku {pdf_path} z ujemną wartością 'Total Amount': {amount_line}")
                        return None  
                except ValueError:
                    pass
    
    start_city = None
    end_city = None
    cities = []

    
    for line in lines:
        if line.startswith("500") and not line.startswith("500,"):  
            
            city = line.split("500", 1)[-1].strip()  

            
            if "/" in city:
                city = city.split("/")[0].strip()  

            
            if "-" in city:
                city = city.split("-")[0].strip()  

            cities.append(city)

    
    if cities:
        start_city = cities[0]  
        end_city = cities[-1]  
        if start_city == end_city and len(cities) > 2:
            start_city = cities[1]

   
    road_tax_de = any("Road tax DE" in line for line in lines)
    road_tax_be = any("Road tax BE" in line for line in lines)
    road_tax_fr = any("Road tax FR" in line for line in lines)
    fuel_surcharge = any("Fuel surcharge" in line for line in lines)

    
    return {
        "start_city": start_city,
        "end_city": end_city,
        "road_tax_de": road_tax_de,
        "road_tax_be": road_tax_be,
        "road_tax_fr": road_tax_fr,
        "fuel_surcharge": fuel_surcharge
    }


import openrouteservice
from openrouteservice import convert
from geopy.geocoders import Nominatim



country_id_map = {
    74: 'Germany',
    53: 'Denmark',
    70: 'France',
    17: 'Belgium',
    
}


def find_countries_by_route(start_city, end_city, api_key):
    global route_cache
    route_key = (start_city, end_city)

    
    if route_key in route_cache:
        print(f"Trasa dla {start_city} -> {end_city} znaleziona w cache.")
        return route_cache[route_key]
    client = openrouteservice.Client(key=api_key)  

    
    geolocator = Nominatim(user_agent="route_checker")
    try:
        start_location = geolocator.geocode(start_city)
        end_location = geolocator.geocode(end_city)
    except Exception as e:
        print(f"Problem z geolokalizacją: {e}")
        return []

    if not start_location or not end_location:
        print(f"Nie udało się znaleźć lokalizacji dla {start_city} lub {end_city}")
        return []

    
    coords = [
        [start_location.longitude, start_location.latitude],
        [end_location.longitude, end_location.latitude]
    ]

    try:
        
        route = client.directions(
            coordinates=coords,
            profile='driving-car',
            format='geojson',
            extra_info=['countryinfo']  
        )
    except Exception as e:
        print(f"Problem z wyznaczaniem trasy: {e}")
        return []

    
    country_ids = set()
    extras = route['features'][0]['properties']['extras']

    
    if 'countryinfo' in extras:
        country_segments = extras['countryinfo']['values']
        for segment in country_segments:
            country_id = segment[2]  
            country_ids.add(country_id)

    
    decoded_countries = [country_id_map.get(country_id, f"Unknown Country ID {country_id}") for country_id in country_ids]

    route_cache[route_key] = decoded_countries
    return decoded_countries


def compare_data(fuel_surcharge, road_tax_de, road_tax_be, road_tax_fr, route_countries):
    errors = []

    
    if not fuel_surcharge:
        errors.append("NO_FUEL_SURCHARGE")

    
    if "Germany" in route_countries and not road_tax_de:
        errors.append("NO_ROAD_TAX_DE")

    if "Belgium" in route_countries and not road_tax_be:
        errors.append("NO_ROAD_TAX_BE")

    if "France" in route_countries and not road_tax_fr:
        errors.append("NO_ROAD_TAX_FR")

    return errors



import shutil  


def move_file_to_error_folder(pdf_path, error_folder='potencjalne_bledy', error_reasons=None):
    if not os.path.exists(error_folder):
        os.makedirs(error_folder)  

    file_name, file_extension = os.path.splitext(os.path.basename(pdf_path))  

    
    if error_reasons:
        error_suffix = "_" + "_".join(error_reasons)
        file_name = f"{error_suffix}{file_name}"

    new_file_name = f"{file_name}{file_extension}"
    new_path = os.path.join(error_folder, new_file_name)  

    shutil.move(pdf_path, new_path)  
    print(f"Przeniesiono plik {pdf_path} do folderu {error_folder} z nową nazwą: {new_file_name}")


def main(username, password, api_key):
    
    download_attachments_from_wp(username, password)

    
    folder_path = 'pdf_files'
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"Przetwarzanie pliku: {pdf_file}")

        
        extracted_data = extract_data_from_pdf(pdf_path)
        
        if extracted_data is None:
            print(f"Ignorowanie pliku {pdf_file} z powodu ujemnej wartości 'Total Amount'.")
            continue
        
        if extracted_data['start_city'] == "BIESHEIM" or extracted_data['end_city'] == "BIESHEIM":
            print(f"Trasa z miastem BIESHEIM. Ignorowanie pliku: {pdf_file}")
            continue  
        
        print(f"Wyniki dla pliku {pdf_file}:")
        print(f"Miasto startowe: {extracted_data['start_city']}")
        print(f"Miasto końcowe: {extracted_data['end_city']}")
        print(f"Opłata drogowa DE: {'Tak' if extracted_data['road_tax_de'] else 'Nie'}")
        print(f"Opłata drogowa BE: {'Tak' if extracted_data['road_tax_be'] else 'Nie'}")
        print(f"Opłata drogowa FR: {'Tak' if extracted_data['road_tax_fr'] else 'Nie'}")
        print(f"Dopłata paliwowa: {'Tak' if extracted_data['fuel_surcharge'] else 'Nie'}")
        print("-" * 50)

       
        route_countries = find_countries_by_route(
            extracted_data['start_city'],
            extracted_data['end_city'],
            api_key
        )
        print(f"Trasa prowadzi przez: {route_countries}")
        
        if not route_countries:
            print(f"Nie udało się znaleźć trasy dla pliku {pdf_file}. Przenoszenie do folderu z błędami.")
            move_file_to_error_folder(pdf_path)
            continue  
        
        errors = compare_data(
            extracted_data['fuel_surcharge'],
            extracted_data['road_tax_de'],
            extracted_data['road_tax_be'],
            extracted_data['road_tax_fr'],
            route_countries
        )

        if errors:
            print(f"Błędy w pliku {pdf_file}:")
            for error in errors:
                print(f"  - {error}")
            
            move_file_to_error_folder(pdf_path, error_reasons=errors)
        else:
            print(f"Wszystko OK w pliku: {pdf_file}")
        print("=" * 50)


if __name__ == "__main__":
    username = ""
    password = ""
    api_key = ""
    subject_filter = "CREDITNOTE"
    main(username, password, api_key)