import pandas as pd

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

def generate_unique_skus(df):
    unique_skus = []
    current_sku = None
    counter = 1

    for index, row in df.iterrows():
        sku = row['SKU']
        if sku != current_sku:
            current_sku = sku
            counter = 1
            unique_skus.append(sku)
        else:
            unique_skus.append(f"{sku}+{counter}")
            counter += 1
    return unique_skus

# Load the original Excel file into a DataFrame
original_file_path = 'extracted_attributes.xlsx'
df_original = pd.read_excel(original_file_path)

# Print column names to identify the correct column name
print("Columns in df_original:", df_original.columns)

# Assuming the correct column name for the link is identified
correct_link_column_name = 'Link'  # Update this to the correct column name if needed

# Extract SKU from Link
df_original['SKU'] = df_original[correct_link_column_name].apply(lambda x: x[-17:-10])

# Ensure SKUs are unique by processing each SKU group
df_original['SKU'] = generate_unique_skus(df_original)

# Ensure the SKU-Handle DataFrame is unique by 'Link'
df_sku_handle = df_original[[correct_link_column_name, 'SKU']].drop_duplicates(subset=[correct_link_column_name])

# Merge the SKU and Handle IDs back into the original data
df_merged = pd.merge(df_original, df_sku_handle, on=correct_link_column_name, how='left', suffixes=('', '_drop'))
df_merged.drop([col for col in df_merged.columns if 'drop' in col], axis=1, inplace=True)

# Modify 'Model grafične kartice' column
df_merged['Model grafične kartice'] = df_merged['Model grafične kartice'].apply(update_graphic_card_model)

# Step 1: Remove Rows Based on Conditions
df_merged = df_merged[df_merged['Operacijski sistem'].str.contains('Windows|Chrome', na=False)]
df_merged = df_merged[df_merged['Razred izdelka'].str.contains('IN|-IN|B|Novo', na=False)]
df_merged['Mark'] = df_merged['Mark'].apply(filter_mark)
df_merged = df_merged[df_merged['Mark'].notna()]
df_merged.dropna(inplace=True)

# Step 2: Modify Data Inside Cells
df_merged['Diagonala zaslona'] = df_merged['Diagonala zaslona'].astype(str).str.replace('.', ',').apply(lambda x: f'{x}\"')
df_merged['Zaslon na dotik'] = df_merged['Zaslon na dotik'].replace({'ja': 'Da', 'št': 'Ne'})

df_merged['Razred izdelka'] = df_merged['Razred izdelka'].str.replace('IN', 'A', case=False).str.replace('IN-', 'B', case=False).str.replace('B', 'C', case=False).str.replace('Novo', 'A+', case=False)

df_merged['Model'] = df_merged['Model'].str.replace('Natančnost', 'Precision', case=False).str.replace('Zemljepisna širina', 'Latitude', case=False)

# Load the image hrefs data
image_hrefs_file_path = 'scraped_image_hrefs.xlsx'
df_image_hrefs = pd.read_excel(image_hrefs_file_path)

# Filter and sort DataFrames based on common 'Link'
common_links = set(df_merged[correct_link_column_name]).intersection(set(df_image_hrefs['Link']))
df_merged_common = df_merged[df_merged[correct_link_column_name].isin(common_links)].copy()
df_image_hrefs_common = df_image_hrefs[df_image_hrefs['Link'].isin(common_links)].copy()

df_merged_common.sort_values(correct_link_column_name, inplace=True)
df_image_hrefs_common.sort_values('Link', inplace=True)

# Specify columns to merge from df_image_hrefs
cols_to_merge = ['Image_Href_2', 'Image_Href_3', 'Image_Href_4', 'Image_Href_5', 'Image_Href_6']

# Merge DataFrames
df_final_merged = pd.merge(df_merged_common, df_image_hrefs_common[['Link'] + cols_to_merge], left_on=correct_link_column_name, right_on='Link', how='left')

# Remove duplicates
df_final_merged.drop_duplicates(subset=[correct_link_column_name], inplace=True)

# Save the final merged DataFrame
final_output_file_path = 'final_merged_attributes.xlsx'
df_final_merged.to_excel(final_output_file_path, index=False)

# Display the first few rows to verify
print(df_final_merged.head())


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
new_export_df = pd.DataFrame(columns=['link','handleId', 'fieldType', 'name', 'description', 'productImageUrl', 'collection', 'SKU', 'ribbon', 'price', 'surcharge', 'visible', 'discountMode', 'discountValue', 'inventory', 'weight', 'cost'])

# Read the updated descriptions Excel file with the 'Nivo' column
descriptions = pd.read_excel('descriptions.xlsx')

# Check for required columns in the descriptions DataFrame
required_columns = [
    'Link', 'SKU', 'Mark', 'Model', 'Model procesorja', 'velikost RAM-a', 'Kapaciteta diska', 
    'Diagonala zaslona', 'Razred izdelka', 'Description', 'Image_Href_2', 'Image_Href_3', 
    'Image_Href_4', 'Image_Href_5', 'Leto', 'Operacijski sistem'
]

for col in required_columns:
    if col not in descriptions.columns:
        raise ValueError(f"Missing required column: {col}")

new_export_df['link'] = descriptions['Link']

new_export_df['handleId'] = descriptions['SKU'].reset_index(drop=True)

new_export_df['fieldType'] = 'Product'

new_export_df['name'] = (
    descriptions['Mark'] + ' ' + descriptions['Model'] + '/' + descriptions['Model procesorja'].str[:13] + '/' +
    descriptions['velikost RAM-a'] + '/' + descriptions['Kapaciteta diska'] + '/' + descriptions['Diagonala zaslona'] + '/' +
    descriptions['Razred izdelka']
)

new_export_df['description'] = descriptions['Description'].reset_index(drop=True)

new_export_df['productImageUrl'] = (
    descriptions['Image_Href_2'].fillna('') + ';' + 
    descriptions['Image_Href_3'].fillna('') + ';' + 
    descriptions['Image_Href_4'].fillna('') + ';' + 
    descriptions['Image_Href_5'].fillna('')
)

new_export_df['collection'] = (
    "SHOPLET;" + "Prenosni računalniki;" + descriptions['Model procesorja'].str[:13] + ';' +
    descriptions['Leto'].astype(str).fillna('') + ';' + descriptions['Operacijski sistem'].fillna('') + ';' +
    descriptions['Razred izdelka'].fillna('') + ';' + descriptions['Diagonala zaslona'].fillna('') + ';' + 
    descriptions['Mark'].fillna('') + ';' + descriptions['velikost RAM-a'].fillna('') + ';' +
    descriptions['Kapaciteta diska'].fillna('') + ';' + descriptions['Mark'].str.upper().fillna('') + " LAPTOP"
)

new_export_df['SKU'] = descriptions['SKU'].reset_index(drop=True)

new_export_df['ribbon'] = ''
new_export_df['price'] = ''
new_export_df['surcharge'] = ''
new_export_df['visible'] = 'TRUE'
new_export_df['discountMode'] = 'Percent'
new_export_df['discountValue'] = 0
new_export_df['inventory'] = 'InStock'
new_export_df['weight'] = ''
new_export_df['cost'] = ''

#Save the new export DataFrame to an Excel file
new_export_file_path_with_images = 'new_export_with_extra_columns_and_images.xlsx'
new_export_df.to_excel(new_export_file_path_with_images, index=False)

import pandas as pd
import requests
from bs4 import BeautifulSoup
import random

def scrape_price(url):
    print(f"Scraping price from {url}")
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        price_element = soup.find('span', id='our_price_display')
        if price_element and price_element.has_attr('content'):
            price = float(price_element['content'])
            print(f"Price found: {price}")
            return price
        else:
            print("Price element not found.")
            return None
    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return None

def adjust_price(price):
    if price is None or pd.isna(price):
        print("Invalid or missing price data. Cannot adjust.")
        return None
    print(f"Adjusting price for {price:.2f}")
    price_str = f"{price:.2f}".split(".")[0]
    last_digit = int(price_str[-1])
    if last_digit >= 5:
        adjusted_price = int(price_str[:-1] + "9") + 1
    else:
        adjusted_price = int(price_str[:-1] + "9")
    adjusted_price_formatted = f"{adjusted_price}.99"
    print(f"Adjusted price: {adjusted_price_formatted}")
    return adjusted_price_formatted

file_path = 'new_export_with_extra_columns_and_images.xlsx'
data = pd.read_excel(file_path)

for index, row in data.iterrows():
    print(f"Processing row {index+1} with URL {row['link']}")
    scraped_price = scrape_price(row['link'])
    if scraped_price is not None:
        cost = scraped_price * 0.22
        data.at[index, 'cost'] = f"{cost:.2f}"
        print(f"Calculated cost: {cost:.2f}")

        # Determine additional margin based on cost
        if cost < 300:
            additional_margin = 40
        elif 300 <= cost < 400:
            additional_margin = 60
        elif 400 <= cost < 500:
            additional_margin = 80
        else:
            additional_margin = 100  # If cost is 500 or more, apply $100 margin

        final_price = cost + additional_margin
        final_price_with_tax = final_price * 1.025
        adjusted_final_price = adjust_price(final_price_with_tax)
        data.at[index, 'price'] = adjusted_final_price
    else:
        print(f"Failed to scrape price for URL {row['link']}")

# Randomly apply 'Top Ponudba!' ribbon to 50% of the entries
num_rows = len(data)
rows_to_update = random.sample(range(num_rows), num_rows // 2)
for row in rows_to_update:
    print(f"Applying 'Top Ponudba!' to row {row}")
    data.at[row, 'ribbon'] = 'Top Ponudba!'
    data.at[row, 'discountValue'] = 100  # Assume discount value needed
    if data.at[row, 'price'] is not None and not pd.isna(data.at[row, 'price']):
        additional_margin = float(data.at[row, 'price']) + 100
        adjusted_price = adjust_price(additional_margin)
        data.at[row, 'price'] = adjusted_price
    else:
        print(f"Cannot apply additional margin due to missing or invalid price at row {row}")

# Remove rows where 'cost' or 'price' is empty
data.dropna(subset=['cost', 'price'], inplace=True)

# Remove the 'link' column from the DataFrame
data.drop(columns=['link'], inplace=True)

# Update collection field
data['collection'] = data['collection'].astype(str) + ";Top Ponudba"
data['discountMode'] = "AMOUNT"

# Save the updated DataFrame to a new Excel file
updated_file_path = 'updated_export_with_extra_columns_and_images.xlsx'
data.to_excel(updated_file_path, index=False)

print("Updated Excel file is saved with discounts and adjusted prices.")
