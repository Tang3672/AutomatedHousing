import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin# Load the Excel spreadsheet
excel_file = 'client_info.xlsx'
df = pd.read_excel(excel_file)

# Function to generate the URL for each client
def generate_url(client):
    url = "https://www.theunitedeffort.org/housing/affordable-housing/"
    url += f"?city={client['Preferred Location']}"
    url += f"&unitType={client['Type of Unit']}"
    
    # Handling for Type of Unit column
    unit_type = client['Type of Unit']
    if isinstance(unit_type, float):  # Check if the value is a float
       unit_type = str(unit_type)  # Convert float to string
    url += f"&unitType={unit_type}"
   

    return url

# Function to download webpage and save it as HTML file
def download_webpage(client_name, url):
    response = requests.get(url)
    if response.status_code == 200:
        with open(f"{client_name}.html", 'wb') as f:
            f.write(response.content)
        print(f"{client_name}.html downloaded successfully.")
    else:
        print(f"Failed to download {client_name}.html")
    

# Loop through each client
for index, client in df.iterrows():
    print(f"Processing {client['Client Name']}...")
    url = generate_url(client)
    
    # Open the website
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Click checkboxes based on client information
        
        # City
        for city_checkbox in soup.find_all('input', {'name': 'city'}):
            if not isinstance(client['PreferredLocation'], float): 
                if city_checkbox['value'].lower() == client['Preferred Location'].lower():
                    city_checkbox['checked'] = ''
        
        # Type of Unit
        for unit_checkbox in soup.find_all('input', {'name': 'unitType'}):
            value = unit_checkbox['value']
            if not isinstance(client['Type of Unit'], float):  # Check if the value is not a float
                if value.lower() in [u.lower() for u in client['Type of Unit'].split(',')]:
                    unit_checkbox['checked'] =""
        
        # Availability
        availability_checkbox = soup.find('input', {'name': 'availability', 'value': 'Available'})
        availability_checkbox['checked'] = ''
        
        # Populations Served
        populations_served_checkbox = soup.find('input', {'name': 'populationsServed', 'value': 'General Population'})
        populations_served_checkbox['checked'] = ''
        
        # Rent
        rent_max_input = soup.find('input', {'name': 'rentMax'})
        rent_max_input['value'] = client['Rent']
        
        # Income
        income_input = soup.find('input', {'name': 'income'})
        income_input['value'] = client['Income']
        
        # Click update results
        update_button = soup.find('input', {'id': 'filter-submit', 'type': 'submit', 'value': 'Update results'})
        form_data = {update_button['value']: ""}
        response = requests.post(url, data=form_data)
        
        # Click show printable summary
        show_printable_summary_button = soup.find('a', {'class': 'btn btn_secondary noprint btn_print'})
        full_url = urljoin("https://theunitedeffort.org/", show_printable_summary_button['href'])
        response = requests.get(full_url)
        
        # Download webpage
        download_webpage(client['Client Name'], response.url)
        
    else:
        print(f"Failed to open {url}")

print("All clients processed.")
