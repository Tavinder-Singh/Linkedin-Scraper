import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import tkinter as tk
from tkinter import filedialog, messagebox

def read_companies_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df['Company Name'].tolist()

def linkedin_login(driver, username, password):
    driver.get("https://www.linkedin.com/login")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password"))).send_keys(password)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[type="submit"]'))).click()
    time .sleep (90 )      #Time to fill the capcha


def search_company_on_google(company_name):
    driver.get("https://www.google.com")
    search_box = driver.find_element(By.NAME, "q")
    search_box.send_keys(company_name + " LinkedIn")
    search_box.send_keys(Keys.RETURN)

    time.sleep(3)  # Wait for search results to load

    try:
        first_result = driver.find_element(By.CSS_SELECTOR, 'a[jsname="UWckNb"]')
        return first_result.get_attribute('href')
    except Exception as e:
        print(f"Error finding LinkedIn link for {company_name}: {e}")
        return None

def get_linkedin_data(linkedin_url):
    driver.get(linkedin_url)

    time.sleep(2)  # Wait for the page to load

    employee_count = "Not Found"
    followers_count = "Not Found"
    industry = "Not Found"
    website = "Not Found"

    try:
        employee_count_elem = driver.find_element(By.CSS_SELECTOR, 'span.t-normal.t-black--light.link-without-visited-state.link-without-hover-state')
        employee_count = employee_count_elem.text
    except Exception as e:
        print(f"Error extracting employee count from LinkedIn page: {e}")

    try:
        followers_count_elems = driver.find_elements(By.CSS_SELECTOR, 'div.org-top-card-summary-info-list__info-item')
        followers_count = followers_count_elems[-1].text.strip()
    except Exception as e:
        print(f"Error extracting followers count from LinkedIn page: {e}")

    try:
        industry_elem = driver.find_element(By.CSS_SELECTOR, 'div.org-top-card-summary-info-list__info-item')
        industry = industry_elem.text.strip()
    except Exception as e:
        print(f"Error extracting industry from LinkedIn page: {e}")

    try:
        driver.get(linkedin_url + "/about")
        time.sleep(2)  # Wait for the page to load

        website_elem = driver.find_element(By.CSS_SELECTOR, 'span.link-without-visited-state[dir="ltr"]')
        website = website_elem.text.strip()
    except Exception as e:
        print(f"Error extracting website from LinkedIn about page: {e}")

    return employee_count, followers_count, industry, website

def write_to_excel(file_path, data):
    df = pd.DataFrame(data, columns=['Company', 'LinkedIn URL', 'Employee Count', 'Followers Count', 'Industry', 'Website'])
    df.to_excel(file_path, index=False)

def start_scraping():
    global driver

    input_file_path = input_file_entry.get()
    output_file_path = output_file_entry.get()
    linkedin_username = email_entry.get()
    linkedin_password = password_entry.get()
    chromedriver_path = chromedriver_path_entry.get()

    if not input_file_path or not output_file_path or not linkedin_username or not linkedin_password or not chromedriver_path:
        messagebox.showerror("Error", "All fields are required!")
        return

    company_names = read_companies_from_excel(input_file_path)

    driver = webdriver.Chrome(service=Service(chromedriver_path))

    try:
        linkedin_login(driver, linkedin_username, linkedin_password)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to log in to LinkedIn: {e}")
        driver.quit()
        return

    linkedin_data = []
    for company in company_names:
        linkedin_url = search_company_on_google(company)
        if linkedin_url:
            employee_count, followers_count, industry, website = get_linkedin_data(linkedin_url)
            linkedin_data.append([company, linkedin_url, employee_count, followers_count, industry, website])
        else:
            linkedin_data.append([company, "Not Found", "Not Found", "Not Found", "Not Found", "Not Found"])

        write_to_excel(output_file_path, linkedin_data)

    driver.quit()
    messagebox.showinfo("Success", "Scraping completed and data saved to Excel!")

# GUI setup
root = tk.Tk()
root.title("LinkedIn Scraper")

tk.Label(root, text="Make sure the column name in the Input file should be 'Company Name'", fg="red").grid(row=0, columnspan=3, padx=10, pady=5)

tk.Label(root, text="Input Excel File:").grid(row=1, column=0, padx=10, pady=5)
input_file_entry = tk.Entry(root, width=50)
input_file_entry.grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: input_file_entry.insert(0, filedialog.askopenfilename())).grid(row=1, column=2, padx=10, pady=5)

tk.Label(root, text="Output Excel File:").grid(row=2, column=0, padx=10, pady=5)
output_file_entry = tk.Entry(root, width=50)
output_file_entry.grid(row=2, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: output_file_entry.insert(0, filedialog.asksaveasfilename(defaultextension=".xlsx"))).grid(row=2, column=2, padx=10, pady=5)

tk.Label(root, text="LinkedIn Email:").grid(row=3, column=0, padx=10, pady=5)
email_entry = tk.Entry(root, width=50)
email_entry.grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="LinkedIn Password:").grid(row=4, column=0, padx=10, pady=5)
password_entry = tk.Entry(root, show='*', width=50)
password_entry.grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="Chromedriver Path:").grid(row=5, column=0, padx=10, pady=5)
chromedriver_path_entry = tk.Entry(root, width=50)
chromedriver_path_entry.grid(row=5, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=lambda: chromedriver_path_entry.insert(0, filedialog.askopenfilename())).grid(row=5, column=2, padx=10, pady=5)

tk.Button(root, text="Start Scraping", command=start_scraping).grid(row=6, column=1, pady=20)

tk.Label(root, text="Made by Tavinder Singh Sahni", fg="blue").grid(row=7, columnspan=3, pady=5)

root.mainloop()
