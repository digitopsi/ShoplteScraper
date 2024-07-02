import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import sys
import traceback
import time

def generate_url(sku):
    return f"https://shoplet.pl/szukaj?controller=search&orderby=position&orderway=desc&searchInDescriptions=0&search_query={sku}"

def compare_stocks(supplier_file, output_file):
    try:
        supplier_df = pd.read_excel(supplier_file)
        supplier_df['Identyfikator'] = supplier_df['Identyfikator'].astype(str)
        new_skus_df = pd.DataFrame({
            'New_SKUs': supplier_df['Identyfikator'],
            'New_SKUs_URLs': [generate_url(sku) for sku in supplier_df['Identyfikator']]
        })
        with pd.ExcelWriter(output_file) as writer:
            new_skus_df.to_excel(writer, sheet_name='New_SKUs', index=False)
        print(f"Stock comparison completed. Output saved to {output_file}", flush=True)
    except Exception as e:
        print(f"Exception in compare_stocks: {str(e)}", flush=True)

def extract_between_first_slashes_from_end(description):
    parts = description.split('/')
    if len(parts) >= 2:
        return parts[-2].strip()
    return ''

def generate_extract_with_info():
    print("Starting generate_extract_with_info function.", flush=True)
    try:
        df = pd.read_excel('edited_universal_stock.xlsx')
        print(f"Read {len(df)} rows from edited_universal_stock.xlsx", flush=True)
        
        filtered_df = df[df['Identyfikator'].apply(lambda x: str(x).startswith('L1'))].reset_index(drop=True)
        print(f"Filtered to {len(filtered_df)} rows where Identyfikator starts with 'L1'", flush=True)
        
        final_rows = []

        for index, row in filtered_df.iterrows():
            sku = row['Identyfikator']
            url = generate_url(sku)
            response = requests.get(url)
            soup = BeautifulSoup(response.text, 'html.parser')
            if soup.find_all('div', {'class': 'product-inner'}):
                final_rows.append(row)
                print(f"Product found for SKU {sku} at {url}", flush=True)

        final_df = pd.DataFrame(final_rows).reset_index(drop=True)
        print(f"{len(final_df)} rows retained after checking for product-inner", flush=True)
        
        if 'Opis' not in final_df.columns:
            print("Column 'Opis' not found in final_df", flush=True)
            final_df['Correct_Extracted_Info'] = ''
        else:
            final_df['Correct_Extracted_Info'] = final_df['Opis'].apply(extract_between_first_slashes_from_end)
            print(f"Extracted Correct_Extracted_Info: {final_df['Correct_Extracted_Info'].unique()}", flush=True)

        values_to_keep = ["A KL", "A- KL", "A1 KL", "A2 KL", "AM1 KL", "AM2 KL", "B KL", "B1 KL", "B2 KL", "KL A1", "KL A2", "KL AM1", "KL AM2", "KL B1", "KL B2", "V1 KL"]
        print(f"Values to keep: {values_to_keep}", flush=True)

        final_filtered_df = final_df[final_df['Correct_Extracted_Info'].isin(values_to_keep)]
        print(f"Correct_Extracted_Info values to be filtered: {final_df['Correct_Extracted_Info'].tolist()}", flush=True)
        print(f"{len(final_filtered_df)} rows retained after filtering Correct_Extracted_Info", flush=True)
        
        final_filtered_df.to_excel('extract_with_info.xlsx', index=False)
        print("Extract with info generated successfully.", flush=True)
    except Exception as e:
        print(f"Exception in generate_extract_with_info: {str(e)}", flush=True)

def first_script():
    try:
        urls = pd.read_excel("SKU_Comparison.xlsx", sheet_name="New_SKUs")
        
        results = pd.DataFrame(columns=["URL", "Link"])
        
        total_urls = len(urls['New_SKUs_URLs'])
        
        for idx, url in enumerate(urls['New_SKUs_URLs']):
            print(f"Processing URL: {url}", flush=True)
            
            r = requests.get(url)
            
            soup = BeautifulSoup(r.text, 'html.parser')
            
            links = soup.find_all('a', class_='back-image')

            for link in links:
                href = link.get('href')
                results.loc[len(results)] = [url, href]
                print(f"Links found: {href}", flush=True)

        results.to_excel("scrapedurls_img.xlsx", index=False)
        print("URL scraping completed.", flush=True)
    except Exception as e:
        print(f"Exception in first_script: {str(e)}", flush=True)

def process_files():
    print("Starting process_files function.", flush=True)
    Tk().withdraw()

    supplier_csv = askopenfilename(title="Select the universal_stock.csv file", filetypes=[("CSV files", "*.csv")])

    if not supplier_csv:
        print("You must select the file to proceed.", flush=True)
        exit()

    supplier_df = pd.read_csv(supplier_csv, delimiter=';')
    print(f"Read {len(supplier_df)} rows from {supplier_csv}", flush=True)
    
    supplier_df = supplier_df[supplier_df['Symbol'].str.startswith('L1', na=False)]
    print(f"Filtered to {len(supplier_df)} rows where Symbol starts with 'L1'", flush=True)
    
    supplier_df.rename(columns={'Symbol': 'SKU'}, inplace=True)
    supplier_xlsx = 'edited_universal_stock.xlsx'
    supplier_df.to_excel(supplier_xlsx, index=False)
    print(f"Processed {supplier_csv} to {supplier_xlsx}", flush=True)

    return supplier_xlsx

def second_script():
    print("Starting second_script function.", flush=True)
    try:
        df = pd.read_excel("scrapedurls_img.xlsx")
        row_list = []
        total_links = len(df['Link'])

        for idx, link in enumerate(df['Link']):
            print(f"Processing URL: {link}", flush=True)
            response = requests.get(link)
            soup = BeautifulSoup(response.text, 'html.parser')
            ul_element = soup.find('ul', {'id': 'thumbs_list_frame'})
            image_hrefs = []
            if ul_element:
                for li in ul_element.find_all('li'):
                    a = li.find('a')
                    if a and 'href' in a.attrs:
                        image_hrefs.append(a['href'])
            row_dict = {"Link": link}
            for i, href in enumerate(image_hrefs):
                col_name = f"Image_Href_{i+1}"
                row_dict[col_name] = href
            row_list.append(row_dict)
        image_results = pd.DataFrame(row_list)
        image_results.to_excel("scraped_image_hrefs.xlsx", index=False)
        print("Image href scraping completed.", flush=True)
    except Exception as e:
        print(f"Exception in second_script: {str(e)}", flush=True)

def third_script():
    print("Starting third_script function.", flush=True)
    try:
        df = pd.read_excel("scrapedurls_img.xlsx")
        chrome_options = Options()
        chrome_profile_path = "C:\\Users\\roksc\\AppData\\Local\\Google\\Chrome\\User Data"
        chrome_options.add_argument(f"user-data-dir={chrome_profile_path}")
        chrome_options.add_argument("profile-directory=Profile 1")
        driver = webdriver.Chrome(options=chrome_options)
        attributes_to_extract = [
            "Mark", "Model", "Razred izdelka", "Model procesorja", "velikost RAM-a", "Kapaciteta diska",
            "Diagonala zaslona", "Zaslon na dotik", "Konektorji", "Garancija",
            "Operacijski sistem", "Komunikacija", "Multimedija", "Model grafične kartice"
        ]
        extracted_data_df = pd.DataFrame()
        total_links = len(df['Link'])

        def slow_scroll(driver):
            last_height = driver.execute_script("return window.pageYOffset;")
            while True:
                body = driver.find_element(By.TAG_NAME, 'body')
                body.send_keys(Keys.PAGE_DOWN)
                time.sleep(1)
                new_height = driver.execute_script("return window.pageYOffset;")
                if new_height == last_height:
                    break
                last_height = new_height

        for idx, link in enumerate(df['Link']):
            print(f"Processing URL: {link}", flush=True)
            driver.get(link)
            time.sleep(1)
            slow_scroll(driver)
            
            # Handle fancybox overlay
            try:
                if driver.find_element(By.CLASS_NAME, 'fancybox-overlay'):
                    print("Found fancybox overlay, attempting to close it...", flush=True)
                    close_button = driver.find_element(By.CLASS_NAME, 'fancybox-close')
                    close_button.click()
                    time.sleep(1)
            except:
                print("No fancybox overlay found, continuing...", flush=True)

            try:
                table = driver.find_element(By.CLASS_NAME, 'table-data-sheet')
                extracted_data = {'Link': link}
                for row in table.find_elements(By.TAG_NAME, 'tr'):
                    cells = row.find_elements(By.TAG_NAME, 'td')
                    if len(cells) == 2:
                        attribute = cells[0].text.strip()
                        value = cells[1].text.strip()
                        if attribute in attributes_to_extract:
                            extracted_data[attribute] = value
                new_row = pd.DataFrame([extracted_data])
                extracted_data_df = pd.concat([extracted_data_df, new_row], ignore_index=True)
            except Exception as e:
                print(f"An error occurred: {e}", flush=True)
                print("Table not found.", flush=True)
        driver.quit()
        extracted_data_df.to_excel("extracted_attributes.xlsx", index=False)
        print("Attributes extraction completed.", flush=True)
    except Exception as e:
        print(f"Exception in third_script: {str(e)}", flush=True)

if __name__ == "__main__":
    print("Main execution started.", flush=True)
    try:
        supplier_xlsx = process_files()
        
        generate_extract_with_info()
        compare_stocks('extract_with_info.xlsx', 'SKU_Comparison.xlsx')
        first_script()
        second_script()
        third_script()
    
    except Exception as e:
        print(f"Exception in main: {str(e)}", flush=True)
    print("Main execution completed.", flush=True)

    import pandas as pd36

def update_graphic_card_model(cell):
    if pd.isna(cell) or cell == "":
        return "Integrirana grafična kartica"
    return cell

def filter_mark(cell):
    if pd.isna(cell):
        return None
    brands = ['Lenovo', 'HP', 'Dell']
    for brand in brands:
        if brand in cell:
            return brand
    return None

# Load the original Excel file into a DataFrame
original_file_path = 'extracted_attributes.xlsx'
df_original = pd.read_excel(original_file_path)

# Modify 'Model grafične kartice' column
df_original['Model grafične kartice'] = df_original['Model grafične kartice'].apply(update_graphic_card_model)

# Step 1: Remove Rows Based on Conditions
df_original = df_original[df_original['Operacijski sistem'].str.contains('Windows|Chrome', na=False)]
df_original = df_original[df_original['Razred izdelka'].str.contains('IN|-IN|B|Novo', na=False)]
df_original['Mark'] = df_original['Mark'].apply(filter_mark)
df_original = df_original[df_original['Mark'].notna()]
df_original.dropna(inplace=True)

# Step 2: Modify Data Inside Cells
df_original['Diagonala zaslona'] = df_original['Diagonala zaslona'].astype(str)
df_original['Diagonala zaslona'] = df_original['Diagonala zaslona'].str.replace('.', ',')
df_original['Diagonala zaslona'] = df_original['Diagonala zaslona'].apply(lambda x: f'{x}\"')
df_original['Zaslon na dotik'] = df_original['Zaslon na dotik'].replace({'ja': 'Da', 'št': 'Ne'})

# Extract SKU from Link and handle duplicates
df_original['SKU'] = df_original['Link'].apply(lambda x: x[-17:-10])
df_original['SKU_Counter'] = df_original.groupby('SKU').cumcount() + 1
df_original['SKU'] = df_original.apply(lambda x: f"{x['SKU']}+{x['SKU_Counter']}" if x['SKU_Counter'] > 1 else x['SKU'], axis=1)
df_original.drop(columns='SKU_Counter', inplace=True)

df_original['Razred izdelka'] = df_original['Razred izdelka'].str.replace('IN', 'A', case=False)
df_original['Razred izdelka'] = df_original['Razred izdelka'].str.replace('IN-', 'B', case=False)
df_original['Razred izdelka'] = df_original['Razred izdelka'].str.replace('B', 'C', case=False)
df_original['Razred izdelka'] = df_original['Razred izdelka'].str.replace('Novo', 'A+', case=False)

# Additional modifications to the 'Model' column
df_original['Model'] = df_original['Model'].str.replace('Natančnost', 'Precision', case=False)
df_original['Model'] = df_original['Model'].str.replace('Zemljepisna širina', 'Latitude', case=False)

# Load and filter df_image_hrefs
image_hrefs_file_path = 'scraped_image_hrefs.xlsx'
df_image_hrefs = pd.read_excel(image_hrefs_file_path)
common_links = set(df_original['Link']).intersection(set(df_image_hrefs['Link']))
df_original_common = df_original[df_original['Link'].isin(common_links)].copy()
df_image_hrefs_common = df_image_hrefs[df_image_hrefs['Link'].isin(common_links)].copy()

# Specify columns to merge from df_image_hrefs
cols_to_merge = ['Image_Href_2', 'Image_Href_3', 'Image_Href_4', 'Image_Href_5', 'Image_Href_6']

# Sort and merge DataFrames
df_original_common.sort_values('Link', inplace=True)
df_image_hrefs_common.sort_values('Link', inplace=True)
df_merged = pd.merge(df_original_common, df_image_hrefs_common[['Link'] + cols_to_merge], on='Link', how='left')

# Save the final merged DataFrame
final_output_file_path = 'final_merged_attributes.xlsx'
df_merged.to_excel(final_output_file_path, index=False)

import pandas as pd

# Function to extract the year of the laptop based on the processor model
def extract_year_from_processor(cell):
    if '-' in cell:
        series_number = cell.split('-')[-1]
        digits = ''.join(filter(str.isdigit, series_number))
        if len(digits) >= 2:
            generation_candidate = int(digits[:2])
            if generation_candidate > 20:
                generation = int(digits[0])
            else:
                generation = generation_candidate
            if generation >= 4:
                return 2010 + generation
    return None

# Function to convert disk capacity from "1000 GB" to "1 TB"
def convert_disk_capacity(cell):
    if cell == "1000 GB":
        return "1 TB"
    return cell

# Function to ensure only one quotation mark in 'Diagonala zaslona'
def adjust_screen_size(cell):
    if cell.count('\"') > 1:
        return cell.replace('\"', '', 1)
    return cell

# Load the final merged DataFrame
df_merged = pd.read_excel('final_merged_attributes.xlsx')

# Update the 'Leto' and 'GRADE' columns in the DataFrame
df_merged['Leto'] = df_merged['Model procesorja'].apply(extract_year_from_processor)

# Remove rows where the processor is older than 2014 or where 'Leto'/'GRADE' is None
df_merged = df_merged[df_merged['Leto'].notna() & df_merged['Razred izdelka'].notna()]
df_merged = df_merged[df_merged['Leto'] >= 2014]

# Apply the conversion functions for 'Kapaciteta diska' and 'Diagonala zaslona'
df_merged['Kapaciteta diska'] = df_merged['Kapaciteta diska'].apply(convert_disk_capacity)
df_merged['Diagonala zaslona'] = df_merged['Diagonala zaslona'].apply(adjust_screen_size)

# Save the corrected DataFrame to a new Excel file
df_merged.to_excel('corrected_final_merged_attributes.xlsx', index=False)



import pandas as pd

# Load the data
data = pd.read_excel("corrected_final_merged_attributes.xlsx")

# Lists to store the generated descriptions
descriptions = []

# Iterate through each row in the DataFrame
for index, row in data.iterrows():
    # Retrieve the necessary information
    brand = row["Mark"]
    model = row["Model"]
    processor_model = row["Model procesorja"]
    ram = row["velikost RAM-a"]
    disk = row["Kapaciteta diska"]
    screen_size = row["Diagonala zaslona"]
    os = row["Operacijski sistem"]
    graphics_card = row["Model grafične kartice"]
    touch_screen = row["Zaslon na dotik"]
    ports = row["Konektorji"]
    multimedia = row["Multimedija"]
    wireless_tech = row["Komunikacija"]
    warranty = row["Garancija"]
    year = row["Leto"]
    grade = row["Razred izdelka"]

    
    # Generate the description based on the computer's specifications
    if processor_model.startswith("Intel Core i5") or ("AMD Ryzen 3") and ram == "8 GB" and disk in ["128 GB", "256 GB"]:
        description = f"<p>{brand} {model} je idealen za vsakodnevno uporabo. Z {disk} prostora in {ram} pomnilnika, {model} odlično obvladuje osnovne naloge, kot so brskanje po spletu, urejanje dokumentov in gledanje filmov. {processor_model} zagotavlja gladko delovanje, medtem ko {screen_size} zagotavlja dovolj prostora za udobno delo in zabavo.</p> <br> <h3><strong>Tehnične specifikacije</strong></h3> <br> <ul><li>Znamka: <strong>{brand}</strong></li><li>Leto: <strong>{year}</strong></li><li>Diagonala zaslona: <strong>{screen_size}</strong></li><li>Vrsta procesorja: <strong>{processor_model}</strong></li><li>Operacijski sistem: <strong>{os}</strong></li><li>Velikost pomnilnika: <strong>{ram}</strong></li><li>Velikost diska: <strong>{disk}</strong></li><li>Grafična kartica: <strong>{graphics_card}</strong></li><li>Zaslon na dotik: <strong>{touch_screen}</strong></li><li>Razred: <strong>{grade}</strong></li><li>Priključki: <strong>{ports}</strong></li><li>Multimedija: <strong>{multimedia}</strong></li><li>Brezžične tehnologije: <strong>{wireless_tech}</strong></li><li>Garancija v mesecih: <strong>{warranty}</strong></li></ul><br><p>Oprema ima lahko vidne napake, kot so odrgnine in udrtine na ohišju. Te napake so vidne, vendar nikakor ne motijo ​​vsakodnevne uporabe. Poleg tega ima lahko zaslon nekaj manjših prask in odrgnin. Pri vključenem zaslonu te napake ne bi smele v ničemer ovirati dela.</p>"
    elif processor_model.startswith("Intel Core i5") or ("AMD Ryzen 5") and ram in ["12 GB", "16 GB"] and disk in ["512 GB", "128 GB", "256 GB"]:
        description = f"<p>{brand} {model} se ponaša z {processor_model}, {ram} pomnilnika in {disk} prostora, kar pomeni, da je kos zahtevnejšim nalogam. Učinkovitost tega prenosnika je idealna za študij, domače pisarne ali manjše podjetje. Ponuja odlično ravnovesje med zmogljivostjo in cenovno dostopnostjo, ne da bi pri tem žrtvoval kakovost.</p> <br> <h3><strong>Tehnične specifikacije</strong></h3> <br> <ul><li>Znamka: <strong>{brand}</strong></li><li>Leto: <strong>{year}</strong></li><li>Diagonala zaslona: <strong>{screen_size}</strong></li><li>Vrsta procesorja: <strong>{processor_model}</strong></li><li>Operacijski sistem: <strong>{os}</strong></li><li>Velikost pomnilnika: <strong>{ram}</strong></li><li>Velikost diska: <strong>{disk}</strong></li><li>Grafična kartica: <strong>{graphics_card}</strong></li><li>Zaslon na dotik: <strong>{touch_screen}</strong></li><li>Razred: <strong>{grade}</strong></li><li>Priključki: <strong>{ports}</strong></li><li>Multimedija: <strong>{multimedia}</strong></li><li>Brezžične tehnologije: <strong>{wireless_tech}</strong></li><li>Garancija v mesecih: <strong>{warranty}</strong></li></ul><br><p>Oprema ima lahko vidne napake, kot so odrgnine in udrtine na ohišju. Te napake so vidne, vendar nikakor ne motijo ​​vsakodnevne uporabe. Poleg tega ima lahko zaslon nekaj manjših prask in odrgnin. Pri vključenem zaslonu te napake ne bi smele v ničemer ovirati dela.</p>"
    elif processor_model.startswith("Intel Core i7") and ram in ["16 GB", "32 GB"] and disk in ["512 GB", "128 GB", "256 GB", "1 TB"]:
        description = f"<p>{brand} {model} je zasnovan za tiste, ki potrebujejo moč in hitrost. {processor_model} in {ram} pomnilnik zagotavljata neverjetno hitro in učinkovito delovanje, medtem ko {disk} prostora omogoča shranjevanje velikih datotek in zahtevnih aplikacij. Idealno za zahtevne uporabnike, kot so grafični oblikovalci, programerji ali gamerji.</p> <br> <h3><strong>Tehnične specifikacije</strong></h3> <br> <ul><li>Znamka: <strong>{brand}</strong></li><li>Leto: <strong>{year}</strong></li><li>Diagonala zaslona: <strong>{screen_size}</strong></li><li>Vrsta procesorja: <strong>{processor_model}</strong></li><li>Operacijski sistem: <strong>{os}</strong></li><li>Velikost pomnilnika: <strong>{ram}</strong></li><li>Velikost diska: <strong>{disk}</strong></li><li>Grafična kartica: <strong>{graphics_card}</strong></li><li>Zaslon na dotik: <strong>{touch_screen}</strong></li><li>Razred: <strong>{grade}</strong></li><li>Priključki: <strong>{ports}</strong></li><li>Multimedija: <strong>{multimedia}</strong></li><li>Brezžične tehnologije: <strong>{wireless_tech}</strong></li><li>Garancija v mesecih: <strong>{warranty}</strong></li></ul><br><p>Oprema ima lahko vidne napake, kot so odrgnine in udrtine na ohišju. Te napake so vidne, vendar nikakor ne motijo ​​vsakodnevne uporabe. Poleg tega ima lahko zaslon nekaj manjših prask in odrgnin. Pri vključenem zaslonu te napake ne bi smele v ničemer ovirati dela.</p>"
    elif processor_model.startswith("Intel Core i9"):
        description = f"<p>{brand} {model} je namenjen zahtevnim uporabnikom, ki potrebujejo vrhunske specifikacije za svoje delo. Z {processor_model}in {ram} za hitro in učinkovito delovanje, ter z {disk} za obilico prostora za shranjevanje, je ta prenosnik pripravljen na najzahtevnejše naloge. Odlična izbira za strokovnjake kot so video producente, 3D umetnike ali profesionalne gamerje.</p> <br> <h3><strong>Tehnične specifikacije</strong></h3> <br> <ul><li>Znamka: <strong>{brand}</strong></li><li>Leto: <strong>{year}</strong></li><li>Diagonala zaslona: <strong>{screen_size}</strong></li><li>Vrsta procesorja: <strong>{processor_model}</strong></li><li>Operacijski sistem: <strong>{os}</strong></li><li>Velikost pomnilnika: <strong>{ram}</strong></li><li>Velikost diska: <strong>{disk}</strong></li><li>Grafična kartica: <strong>{graphics_card}</strong></li><li>Zaslon na dotik: <strong>{touch_screen}</strong></li><li>Razred: <strong>{grade}</strong></li><li>Priključki: <strong>{ports}</strong></li><li>Multimedija: <strong>{multimedia}</strong></li><li>Brezžične tehnologije: <strong>{wireless_tech}</strong></li><li>Garancija v mesecih: <strong>{warranty}</strong></li></ul><br><p>Oprema ima lahko vidne napake, kot so odrgnine in udrtine na ohišju. Te napake so vidne, vendar nikakor ne motijo ​​vsakodnevne uporabe. Poleg tega ima lahko zaslon nekaj manjših prask in odrgnin. Pri vključenem zaslonu te napake ne bi smele v ničemer ovirati dela.</p>"
    else:
        description = f"<h3><strong>Tehnične specifikacije</strong></h3> <br> <ul><li>Znamka: <strong>{brand}</strong></li><li>Leto: <strong>{year}</strong></li><li>Diagonala zaslona: <strong>{screen_size}</strong></li><li>Vrsta procesorja: <strong>{processor_model}</strong></li><li>Operacijski sistem: <strong>{os}</strong></li><li>Velikost pomnilnika: <strong>{ram}</strong></li><li>Disk: <strong>{disk}</strong></li><li>Grafična kartica: <strong>{graphics_card}</strong></li><li>Zaslon na dotik: <strong>{touch_screen}</strong></li><li>Razred: <strong>{grade}</strong></li><li>Priključki: <strong>{ports}</strong></li><li>Multimedija: <strong>{multimedia}</strong></li><li>Brezžične tehnologije: <strong>{wireless_tech}</strong></li><li>Garancija v mesecih: <strong>{warranty}</strong></li></ul><br><p>Oprema ima lahko vidne napake, kot so odrgnine in udrtine na ohišju. Te napake so vidne, vendar nikakor ne motijo ​​vsakodnevne uporabe. Poleg tega ima lahko zaslon nekaj manjših prask in odrgnin. Pri vključenem zaslonu te napake ne bi smele v ničemer ovirati dela.</p>"
    
    # Add the description to the list
    descriptions.append(description)

if len(descriptions) != len(data):
    raise ValueError(f"Length of descriptions ({len(descriptions)}) does not match length of DataFrame ({len(data)})")

# Add the descriptions to the DataFrame
data["Description"] = descriptions

# Save the DataFrame to a new Excel file
data.to_excel("descriptions.xlsx", index=False)

# ---- Final Export Code Starts Here ----


import pandas as pd



# Create the new export DataFrame with required column headers
new_export_df = pd.DataFrame(columns=['link','handleId', 'fieldType', 'name', 'description', 'productImageUrl', 'collection', 'SKU', 'ribbon', 'price', 'surcharge', 'visible', 'discountMode','discountValue','inventory','weight','cost'])

# Read the updated descriptions Excel file with the 'Nivo' column
descriptions = pd.read_excel('descriptions.xlsx')

new_export_df['link'] = descriptions['Link']

new_export_df['handleId'] = descriptions['SKU'].reset_index(drop=True)

new_export_df['fieldType'] = 'Product'

new_export_df['name'] = descriptions['Mark'] + ' ' + descriptions['Model'] + '/' + descriptions['Model procesorja'].str[:13] + '/' + descriptions['velikost RAM-a'] + '/' + descriptions['Kapaciteta diska'] + '/' + descriptions['Diagonala zaslona'] + '/' + descriptions['Razred izdelka']

# Add the description data
new_export_df['description'] = descriptions['Description'].reset_index(drop=True)

new_export_df['productImageUrl'] = descriptions['Image_Href_2'] + ';' + descriptions['Image_Href_3'] + ';' + descriptions['Image_Href_4'] + ';' + descriptions['Image_Href_5']

new_export_df['collection'] ="SHOPLET;" + "Prenosni računalniki;" + descriptions['Model procesorja'].str[:13] + ';' + descriptions['Leto'].astype(str) + ';' + descriptions['Operacijski sistem'] + ';' + descriptions['Razred izdelka'] + ';' + descriptions['Diagonala zaslona'] + ';' + descriptions['Mark'] + ';' + descriptions['velikost RAM-a'] + ';' + descriptions['Kapaciteta diska'] + ';' + descriptions['Mark'].upper() + " LAPTOP"

# Populate SKU column from the updated descriptions DataFrame
new_export_df['SKU'] = descriptions['SKU'].reset_index(drop=True)

new_export_df['ribbon'] = ''

new_export_df['price'] = ''

new_export_df['surcharge'] = ''
# Add the static columns
new_export_df['visible'] = 'TRUE'
new_export_df['discountMode'] = 'Percent'
new_export_df['discountValue'] = 0
new_export_df['inventory'] = 'InStock'

new_export_df['weight'] = ''

new_export_df['cost'] = ''

#Save the new export DataFrame to an Excel file
new_export_file_path_with_images = 'new_export_with_extra_columns_and_images.xlsx'
new_export_df.to_excel(new_export_file_path_with_images, index=False)

    