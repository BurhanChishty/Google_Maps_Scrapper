import time
import pandas as pd
import threading
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

stop_scraping = False

def get_business_details(driver):
    global stop_scraping
    
    action = ActionChains(driver)
    data_list = []
    seen_names = set()
    
    while not stop_scraping:
        business_elements = driver.find_elements(By.CLASS_NAME, "hfpxzc")

        if not business_elements:
            print("No business elements found.")
            break

        for business in business_elements:
            if stop_scraping:
                break

            try:
                action.move_to_element(business).perform()
                business.click()
                time.sleep(3)  # Wait for details panel to load

                source = driver.page_source
                soup = BeautifulSoup(source, 'html.parser')

                name_element = soup.find('h1', {"class": "DUwDvf lfPIob"})
                name = name_element.text.strip() if name_element else "N/A"

                if name in seen_names:
                    continue
                seen_names.add(name)

                address_element = soup.find(attrs={"data-tooltip": "Copy address"})
                address = address_element.get_text(strip=True) if address_element else "N/A"

                phone_element = soup.find(attrs={"data-tooltip": "Copy phone number"})
                phone = phone_element.get_text(strip=True) if phone_element else "N/A"

                website_element = soup.find(attrs={"data-tooltip": "Open website"})
                website = website_element.get_text(strip=True) if website_element else "Not available"

                data_list.append([name, address, phone, website])
                
                # Update GUI Table in real-time
                tree.insert("", tk.END, values=(name, address, phone, website))

                print(f"Extracted: {name}, {address}, {phone}, {website}")  # Debugging line
                
                time.sleep(2)  # Small delay before clicking the next business

            except Exception as e:
                print(f"Error extracting data: {e}")
                continue

        # Scroll down to load more businesses if available
        action.scroll_by_amount(0, 1000).perform()
        time.sleep(3)  # Allow new businesses to load

    return data_list

def save_data_to_excel(data):
    if not data:
        messagebox.showwarning("Warning", "No data was scraped.")
        return
    
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[["Excel Files", "*.xlsx"]], title="Save Data As")
    if file_path:
        df = pd.DataFrame(data, columns=["Business Name", "Address", "Phone Number", "Website"])
        df.to_excel(file_path, index=False, engine='openpyxl')
        messagebox.showinfo("Success", f"Data saved to {file_path}")
    else:
        messagebox.showwarning("Warning", "No file was selected. Data was not saved.")

def scrape_google_maps(search_query, status_label):
    global stop_scraping
    
    options = Options()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    driver.maximize_window()  # Maximize Chrome window
    
    google_maps_url = "https://www.google.com/maps"
    driver.get(google_maps_url)
    time.sleep(5)
    
    search_box = driver.find_element(By.ID, "searchboxinput")
    search_box.send_keys(search_query)
    search_box.send_keys(Keys.RETURN)
    time.sleep(7)  # Increased time to ensure page loads properly
    
    all_data = get_business_details(driver)
    driver.quit()
    
    save_data_to_excel(all_data)
    status_label.config(text="Scraping completed.", fg="green")

def start_scraping():
    global stop_scraping
    stop_scraping = False
    search_query = entry.get()
    if not search_query:
        messagebox.showerror("Error", "Please enter a search query.")
        return

    # Clear previous data in Treeview
    for row in tree.get_children():
        tree.delete(row)

    threading.Thread(target=scrape_google_maps, args=(search_query, status_label), daemon=True).start()

def stop_scraping_func():
    global stop_scraping
    stop_scraping = True

# GUI Setup
root = tk.Tk()
root.title("Google Maps Scraper")
root.geometry("700x400")

tk.Label(root, text="Enter search query:").pack(pady=5)
entry = tk.Entry(root, width=40)
entry.pack(pady=5)

start_button = tk.Button(root, text="Start Scraping", command=start_scraping)
start_button.pack(pady=5)

stop_button = tk.Button(root, text="Stop Scraping", command=stop_scraping_func, fg="red")
stop_button.pack(pady=5)

status_label = tk.Label(root, text="", fg="black")
status_label.pack(pady=10)

# Add Treeview for displaying extracted data
tree = ttk.Treeview(root, columns=("Business Name", "Address", "Phone Number", "Website"), show="headings")
tree.heading("Business Name", text="Business Name")
tree.heading("Address", text="Address")
tree.heading("Phone Number", text="Phone Number")
tree.heading("Website", text="Website")

tree.pack(pady=10, fill=tk.BOTH, expand=True)

root.mainloop()
