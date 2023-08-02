# Mini Project on Web Scraping and generating dataset
import pandas as pd
import openpyxl
from bs4 import BeautifulSoup
import requests

product_name = []
product_prices = []
product_desc = []
ratings = []

for i in range(2, 40):

    url = "https://www.flipkart.com/search?q=t+shirts+for+men&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=on&as=off&page=" + str(i)
    response = requests.get(url) # sends a GET request to the url asking the server to return the content of the web page
    
    
    soup = BeautifulSoup(response.text, "lxml")# creates an object soup by parsing the content of HTML stored in response

    next_page = soup.find('a', class_="_1LKTO3")

    if next_page:# checks if next_page variable is not None
        next_page = next_page.get("href")#extracts the URL to the next page from the anchor tag

    box = soup.find("div", class_="_1YokD2 _3Mn1Gg")

    names = box.find_all("div", class_="_4rR01T")
    prices = box.find_all("div", class_="_30jeq3 _1_WHN1")
    descriptions = box.find_all("ul", class_="_1xgFaf")
    reviews = box.find_all("div", class_="_3LWZlK")

    for idx in range(max(len(names), len(prices), len(descriptions))):
        name = names[idx].text if idx < len(names) else None
        product_name.append(name)

        price = prices[idx].text if idx < len(prices) else None
        product_prices.append(price)

        desc = descriptions[idx].text if idx < len(descriptions) else None
        product_desc.append(desc)

        # the following was done as length of reviews list was different than length of others
        if idx < len(reviews):
            rating = reviews[idx].text
        else:
            rating = "NA"
        ratings.append(rating)

data_frame = pd.DataFrame({
    "Name of Mobile": product_name,
    "Price": product_prices,
    "Product Description": product_desc,
    "Ratings": ratings
})

# print(data_frame)
# Write DataFrame to Excel file with wrapping enabled for the 'Product Description' column
file_path = "C:/Users/Adarsh Joshi/windows important stuff/OneDrive/Desktop/web scraper/T shirts.xlsx"
data_frame.to_excel(file_path, index=False, engine='openpyxl')

# Open the workbook and get the worksheet
wkbook = openpyxl.load_workbook(file_path)
wksheet = wkbook.active

# Set the column widths for better readability
wksheet.column_dimensions['A'].width = 40  # Name of Mobile
wksheet.column_dimensions['B'].width = 15  # Price
wksheet.column_dimensions['C'].width = 80  # Product Description
wksheet.column_dimensions['D'].width = 15  # Ratings

# Apply the wrapping format to the 'Product Description' column
for row in wksheet.iter_rows(min_row=2, min_col=3, max_col=3):  # Start from row 2 to skip header
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(wrapText=True)

# Save the Excel file
wkbook.save(file_path)
