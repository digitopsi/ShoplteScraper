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
