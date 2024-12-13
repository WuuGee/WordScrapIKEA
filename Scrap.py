import os
import requests
from bs4 import BeautifulSoup
import csv
import random
from time import sleep
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys 
import pandas as pd

def read_csv(filename):
    """Read CSV and handle empty rows"""
    products = []
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)  # Skip header
            for row in reader:
                if row and row[0].strip():  # Check if row exists and first column is not empty
                    products.append(row)
    except Exception as e:
        print(f"Error reading CSV file: {e}")
    return products

def write_to_excel(details):
    """Write product details to Excel file"""
    fieldnames = ['Name', 'Color', 'Price', 'Article Num', 'Summary', 'Dimension']
    
    try:
        # Try to read existing Excel file
        if os.path.exists('StoreProduct.xlsx'):
            df = pd.read_excel('StoreProduct.xlsx')
        else:
            # Create new DataFrame if file doesn't exist
            df = pd.DataFrame(columns=fieldnames)
        
        # Append new row
        new_row = pd.DataFrame([details])
        df = pd.concat([df, new_row], ignore_index=True)
        
        # Write back to Excel
        df.to_excel('StoreProduct.xlsx', index=False)
    except Exception as e:
        print(f"Error writing to Excel: {e}")

def process_product_details(driver):
    """Extract all product details from the current page"""
    details = {}
    try:
        details['Name'] = driver.find_element(By.CLASS_NAME, "pip-header-section__title--big").text
        description = driver.find_element(By.CLASS_NAME, "pip-header-section__description-text").text
        # Extract color from description (text between first and second comma)
        try:
            color = description.split(',')[1].strip()
            details['Color'] = color
        except IndexError:
            details['Color'] = description  # Keep original if format doesn't match
        details['Price'] = driver.find_element(By.CLASS_NAME, "pip-temp-price__integer").text
        details['Article Num'] = driver.find_element(By.CLASS_NAME, "pip-product-identifier__value").text
        details['Summary'] = driver.find_element(By.CLASS_NAME, "pip-product-summary__description").text
        try:
            details['Dimension'] = driver.find_element(By.CLASS_NAME, "pip-header-section__description-measurement").text
        except:
            details['Dimension'] = "Not specified"
        
        # Write to CSV
        write_to_excel(details)
    except Exception as e:
        print(f"Error getting product details: {e}")
    return details



def process_color_variation(driver, href):
    """Process a single color variation"""
    try:
        print(f"\nNavigating to color variation: {href}")
        driver.get(href)
        sleep(2)
        return process_product_details(driver)
    except Exception as e:
        print(f"Error processing color variation: {e}")
        return None
    
def normalize_swedish_chars(text):
    """Convert Swedish characters to their regular alphabet counterparts"""
    swedish_chars = {
        'Ä': 'A',
        'Å': 'A',
        'Ö': 'O',
        'É': 'E',
        'Ü': 'U',
        'ä': 'a',
        'å': 'a',
        'ö': 'o',
        'é': 'e',
        'ü': 'u'
    }
    for swedish, regular in swedish_chars.items():
        text = text.replace(swedish, regular)
    return text

def scrape_ikea_malaysia():
    driver = webdriver.Chrome()
    base_url = "https://www.ikea.com/my/en/"
    
    # Read all products from CSV
    products = read_csv('ProductName.csv')
    
    try:
        # Initial setup and cookie handling
        driver.get(base_url)
        wait = WebDriverWait(driver, 10)
        
        try:
            cookie_button = wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler")))
            cookie_button.click()
            sleep(1)
        except:
            print("No cookie consent popup found or already accepted")

        # Process each product from CSV
        for product_row in products:
            current_product = product_row[0]  # Assuming product name is in first column
            print(f"\n{'='*50}")
            print(f"Processing product: {current_product}")
            print(f"{'='*50}")
            
            try:
                # Search for product
                search_bar = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ikea-search-input"]')))
                search_bar.clear()
                search_bar.send_keys(current_product)
                search_bar.send_keys(Keys.RETURN)
                sleep(2)
                
                try:
                    searchButton = driver.find_element(By.XPATH, '/html/body/header/div/div[2]/div/div/form/div[1]/div/span[2]/span[2]/div/button')
                    searchButton.click()
                except:
                    print("Search button not found, continuing with enter key search")

                try:
                    # Wait for search results
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "plp-fragment-wrapper")))
                    sleep(2)

                    # Find matching products
                    fragment_wrappers = driver.find_elements(By.CLASS_NAME, "plp-fragment-wrapper")
                    matching_product_links = []
                    
                    for wrapper in fragment_wrappers:
                        product_div = wrapper.find_element(By.CSS_SELECTOR, 'div[data-product-name]')
                        product_name = product_div.get_attribute('data-product-name')
                        print(f"Found product: {product_name}")

                        normalized_product_name = normalize_swedish_chars(product_name.upper())
                        normalized_current_product = normalize_swedish_chars(current_product.upper())
                        
                        if normalized_product_name == normalized_current_product:
                            product_link = wrapper.find_element(By.CLASS_NAME, 'plp-product__image-link')
                            matching_product_links.append(product_link.get_attribute('href'))

                    print(f"Found {len(matching_product_links)} matching links for {current_product}")
                except:
                    print(f"No results found for {current_product}, moving to next product")
                    driver.get(base_url)
                    sleep(2)
                    continue

                # Process each matching product
                for link in matching_product_links:
                    print(f"\nProcessing link: {link}")
                    driver.get(link)
                    sleep(2)

                    try:
                        # Check for color variations
                        style_items = driver.find_elements(By.CLASS_NAME, "pip-product-styles__items")
                        if style_items:
                            style_container = style_items[0]
                            color_hrefs = []
                            
                            # Collect all color variation links
                            link_items = style_container.find_elements(By.TAG_NAME, "a")
                            for link_item in link_items:
                                try:
                                    href = link_item.get_attribute('href')
                                    if href:
                                        color_hrefs.append(href)
                                except:
                                    continue
                            
                            # Process current color
                            print("\nProcessing current color:")
                            current_color_details = process_product_details(driver)
                            print(current_color_details)
                            
                            # Process each color variation
                            for href in color_hrefs:
                                color_details = process_color_variation(driver, href)
                                if color_details:
                                    print(color_details)
                        
                        else:
                            # Process single product without variations
                            print("\nProcessing single product:")
                            details = process_product_details(driver)
                            print(details)
                            
                    except Exception as e:
                        print(f"Error processing product: {e}")

                # Return to base URL for next product
                driver.get(base_url)
                sleep(2)
                
            except Exception as e:
                print(f"Error processing product {current_product}: {e}")
                driver.get(base_url)  # Return to base URL even if error occurs
                sleep(2)

    except Exception as e:
        print(f"Fatal error in main scraping process: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_ikea_malaysia()