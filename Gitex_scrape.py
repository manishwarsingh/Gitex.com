from selenium import webdriver
import pandas as pd
from pandas import ExcelWriter
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

# Function to check if a value exists in a column of a DataFrame
def is_value_exists(df, column_name, value):
    
    return value in df[column_name].values

# Read the Excel file
file_path = 'testcase1.xlsx'
df = pd.read_excel(file_path)
df1 = df['button6 href']

chrome_options = Options()
# chrome_options.add_argument('--headless')
path = os.getcwd()
# Path to the ChromeDriver executable
print(f"Start chrome driver....")
chrome_driver_path = path +"\\chromedriver.exe"

def scrape_data(link):
    # Implement the function here
    driver = webdriver.Chrome(service=Service(chrome_driver_path), options=chrome_options)
    driver.get(link)
    time.sleep(1)
    print(f"link--> {link}")

    try:

        print(f"link--> {link}")
        
        # Extract the Title from the page
        try:
            title_element = driver.find_element(By.XPATH, "//*[@id='cphContents_Label22']")
            Title = title_element.text if title_element else "Not Record on the Site"
            print(f"Title--- {Title}")
        except Exception as err:
            print(err)
            Title = "Title Not Found"

        # Click on the Company_profile to get the text
        try:
            time.sleep(1)
            company_profile_element = driver.find_element(By.XPATH, "//*[@id='cphContents_lblOnlineProfile']")
            Company_profile = company_profile_element.text if company_profile_element else "Not Record on the Site"
            print(f"Company_profile ------> {Company_profile}")
        except Exception as err:
            print(err)
            Company_profile = "Company_profile Not Found"
        
        # Click on the LinkedIn icon to get the text
        try:
            time.sleep(1)
            linkedin_elements = driver.find_elements(By.CSS_SELECTOR, "#cphContents_ddLinkedln")
            linkedin_links = linkedin_elements[0].get_attribute("href") if linkedin_elements else "Not Record on the Site"
            print(f"linkedin_links ---> {linkedin_links}")
        except Exception as e:
            print(f"e-----\n {e}")
            linkedin_links = "linkedin_links Not Found"

        # Click on the industry to get the text
        try:
            industry_elements = driver.find_elements(By.XPATH, "//*[@id='DivMiddle']/div[1]/div/div/div[2]/div/ul/li[6]/a")
            if industry_elements:
                industry_elements[0].click()
                industry_text_element = driver.find_element(By.XPATH, "//*[@id='INDUSTRY']/div")
                Industry_text = industry_text_element.get_attribute("innerText").split(" ,") if industry_text_element else "Not Record on the Site"
            else:
                Industry_text = "Not Record on the Site"
            print(f"Industry ------ {Industry_text}")
        except Exception as err:
            print(err)
            Industry_text = "Industry_text Not Found"
        
    except Exception as err:
        print(f"Unexpected error {err=}, {type(err)=}")
        Title = "Error"
        Company_profile = "Error"
        linkedin_links = "Error"
        Industry_text = "Error"

    driver.quit()
    
    return Title, Company_profile, Industry_text, linkedin_links

# Loop through each link in the Excel file, scrape data, and insert into the sheet one by one
for index, link in enumerate(df1):
    # Check if the data already exists in the DataFrame
    if is_value_exists(df, 'button6 href', link):
        # Call the function here
        print(df)
        print(link)
        
        Title, Company_profile, Industry_text, linkedin_links = scrape_data(link)
        # Assign the new data to the DataFrame
        df.loc[index, 'Title'] = Title
        df.loc[index, 'Company_profile'] = Company_profile
        df.loc[index, 'Industry_text'] = Industry_text
        df.loc[index, 'linkedin_links'] = linkedin_links

        # Save the updated DataFrame back to the Excel file after each link
        with ExcelWriter(file_path) as writer:
            df.to_excel(writer, index=False)

print(f"Scraped data successfully and updated in Excel sheet.")
