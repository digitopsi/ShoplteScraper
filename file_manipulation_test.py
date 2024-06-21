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

log_file = open("logs.txt", "w")
sys.stdout = log_file
sys.stderr = log_file

def log_exception(e):
    """
    Log an exception to the log file.
    """
    exc_type, exc_value, exc_traceback = sys.exc_info()
    traceback_details = {
        'filename': exc_traceback.tb_frame.f_code.co_filename,
        'lineno'  : exc_traceback.tb_lineno,
        'name'    : exc_traceback.tb_frame.f_code.co_name,
        'type'    : exc_type.__name__,
        'message' : exc_value.message,
    }
    log_file.write("Exception occurred: {}\n".format(traceback_details))
    traceback.print_exception(exc_type, exc_value, exc_traceback, file=log_file)

Tk().withdraw()  # to hide the main window

# Ask for 'universal_stock.csv' file
supplier_csv = askopenfilename(title="Select the universal_stock.csv file", filetypes=[("CSV files", "*.csv")])

# Ask for 'catalog_products.csv' file
website_csv = askopenfilename(title="Select the catalog_products.csv file", filetypes=[("CSV files", "*.csv")])

# Check if the user selected files
if not supplier_csv or not website_csv:
    print("You must select both files to proceed.")
    exit()

def generate_url(sku):
    return f"https://shoplet.pl/szukaj?controller=search&orderby=position&orderway=desc&searchInDescriptions=0&search_query={sku}"

def compare_stocks(supplier_file, website_file, output_file):
    supplier_df = pd.read_excel(supplier_file)
    website_df = pd.read_excel(website_file)
    supplier_df['Identyfikator'] = supplier_df['Identyfikator'].astype(str)
    website_df['SKU'] = website_df['SKU'].astype(str)
    new_skus = list(set(supplier_df['Identyfikator']) - set(website_df['SKU']))
    new_skus_df = pd.DataFrame({
        'New_SKUs': new_skus,
        'New_SKUs_URLs': [generate_url(sku) for sku in new_skus]
    })
    with pd.ExcelWriter(output_file) as writer:
        new_skus_df.to_excel(writer, sheet_name='New_SKUs', index=False)

# Generating 'edited_universal_stock.xlsx'
supplier_df = pd.read_csv(supplier_csv, delimiter=';')
supplier_df = supplier_df[supplier_df['Symbol'].str.startswith('L1', na=False)]
supplier_df.rename(columns={'Symbol': 'SKU'}, inplace=True)
supplier_xlsx = 'edited_universal_stock.xlsx'
supplier_df.to_excel(supplier_xlsx, index=False)

# Generating 'edited_catalog_products.xlsx'
website_df = pd.read_csv(website_csv, delimiter=',')
website_df.rename(columns={'sku': 'SKU'}, inplace=True)
website_xlsx = 'edited_catalog_products.xlsx'
website_df.to_excel(website_xlsx, index=False)

# Use 'extract_with_info.xlsx' as the supplier file in compare_stocks function
output_xlsx = 'SKU_Comparison.xlsx'
compare_stocks('edited_universal_stock.xlsx', 'edited_catalog_products.xlsx', output_xlsx)



def generate_extract_with_info():
    df = pd.read_excel('edited_universal_stock.xlsx')  # Read the edited file
    filtered_df = df[df['Identyfikator'].apply(lambda x: str(x).startswith('L1'))].reset_index(drop=True)
    final_rows = []

    for index, row in filtered_df.iterrows():
        sku = row['Identyfikator']
        url = generate_url(sku)
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        if soup.find_all('div', {'class': 'product-inner'}):
            final_rows.append(row)

    final_df = pd.DataFrame(final_rows).reset_index(drop=True)
    final_df['Correct_Extracted_Info'] = final_df['Opis'].apply(extract_between_first_slashes_from_end)
    values_to_keep = ["A KL", "A- KL", "A1 KL", "A2 KL", "AM1 KL", "AM2 KL", "B KL", "B1 KL", "B2 KL", "KL A1", "KL A2", "KL AM1", "KL AM2", "KL B1", "KL B2", "V1 KL"]
    final_filtered_df = final_df[final_df['Correct_Extracted_Info'].isin(values_to_keep)]
    final_filtered_df.to_excel('extract_with_info.xlsx', index=False)



# Generating 'edited_universal_stock.xlsx'
supplier_df = pd.read_csv(supplier_csv, delimiter=';')
supplier_df = supplier_df[supplier_df['Symbol'].str.startswith('L1', na=False)]
supplier_df.rename(columns={'Symbol': 'SKU'}, inplace=True)
supplier_xlsx = 'edited_universal_stock.xlsx'
supplier_df.to_excel(supplier_xlsx, index=False)

# Generating 'edited_catalog_products.xlsx'
website_df = pd.read_csv(website_csv, delimiter=',')
website_df.rename(columns={'sku': 'SKU'}, inplace=True)
website_xlsx = 'edited_catalog_products.xlsx'
website_df.to_excel(website_xlsx, index=False)

# Call generate_extract_with_info here
generate_extract_with_info()

# Now, use 'extract_with_info.xlsx' as the supplier file in compare_stocks function
output_xlsx = 'SKU_Comparison.xlsx' 
compare_stocks('extract_with_info.xlsx', website_xlsx, output_xlsx)



def first_script():
    urls = pd.read_excel("SKU_Comparison.xlsx", sheet_name="New_SKUs")
    
    results = pd.DataFrame(columns=["URL", "Link"])
    
    total_urls = len(urls['New_SKUs_URLs'])
    
    for idx, url in enumerate(urls['New_SKUs_URLs']):
        print(f"Processing URL: {url}")
        
        r = requests.get(url)
        
        soup = BeautifulSoup(r.text, 'html.parser')
        
        links = soup.find_all('a', class_='back-image')

        for link in links:
            href = link.get('href')
            results.loc[len(results)] = [url, href]
            print(f"Links found: {href}")
     
        progress_bar(idx+1, total_urls)

    results.to_excel("scrapedurls_img.xlsx", index=False)


def second_script():
    df = pd.read_excel("scrapedurls_img.xlsx")
    row_list = []
    total_links = len(df['Link'])

    for idx, link in enumerate(df['Link']):
        print(f"Processing URL: {link}")
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
        progress_bar(idx+1, total_links)
    image_results = pd.DataFrame(row_list)
    image_results.to_excel("scraped_image_hrefs.xlsx", index=False)

def third_script():
    df = pd.read_excel("scrapedurls_img.xlsx")
    chrome_options = Options()
    chrome_profile_path = "C:\\Users\\AŽBE\\AppData\\Local\\Google\\Chrome\\User Data"
    chrome_options.add_argument(f"user-data-dir={chrome_profile_path}")
    chrome_options.add_argument("profile-directory=Profile 11")
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
        print(f"Processing URL: {link}")
        driver.get(link)
        time.sleep(1)
        slow_scroll(driver)
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
            print(f"An error occurred: {e}")
            print("Table not found.")
        progress_bar(idx+1, total_links)
    driver.quit()
    extracted_data_df.to_excel("extracted_attributes.xlsx", index=False)

def main():
    try:
        generate_extract_with_info()
        print("Running the second script...")
        print("Running the SKU comparison script...")
        create_SKU_comparison()
        print("SKU comparison script completed.")
        
        print("Running the first script...")
        first_script()
        print("First script completed.")

        print("Running the second script...")
        second_script()
        print("Second script completed.")

        print("Running the third script...")
        third_script()
        print("Third script completed.")
    except Exception as e:
        log_exception(e)

if __name__ == "__main__":
    log_file = open("logs.txt", "w")
    sys.stdout = log_file
    sys.stderr = log_file

    main()

    log_file.close()
