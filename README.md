# Amazon Product Scraper

This project is an Amazon product scraper built using Python with **Selenium** and **BeautifulSoup**. It scrapes product details such as titles, prices, ratings, and links from Amazon and saves them in an Excel file with clickable links.

## Features:
- Scrapes product details like **Title**, **Original Price**, **Featured Price**, **Rating**, and **Product Link**.
- Solves **CAPTCHA** using `amazoncaptcha`.
- Handles pagination to scrape multiple pages.
- Saves the scraped data to an **Excel** file.
- Makes product links in the Excel file clickable.

## Setup Instructions:

1. **Install the required dependencies:**
   Before running the project, ensure that all necessary libraries are installed. You can install them by running:
    
   ```bash
   pip install -r requirements.txt
    ```

    Alternatively, you can install the libraries manually using:
   ```bash
   pip install selenium webdriver-manager amazoncaptcha beautifulsoup4 lxml openpyxl pandas
    ```

2. **Download ChromeDriver:**
    Ensure that ChromeDriver is installed and matches your Chrome browser version. This project uses the webdriver-manager library to handle it automatically.

## Running the Script:

**Define the Product Search:**
    Modify the search_element and num_pages variables in the script to choose the product category and the number of pages to scrape:<br>
    For example:

<mark>search_element = "Laptops"</mark>

<mark>num_pages = 10</mark>

**Run the Script:**
    Run the script with your Python interpreter:
   ```bash
   python amazon_scraper.py
   ```

**The script will:**<br>
    > Solve CAPTCHAs automatically if encountered.<br>
    > Scrape product data and save it to an Excel file named amazon_<search_element>_data.xlsx.<br>
    > Add clickable product links to the Excel file.<br>

**Excel File Output:**
    The output Excel file will be named according to the search element, e.g., amazon_Laptops_data.xlsx.

## Script Workflow:
**Initialization:**
    ChromeDriver is set up in headless mode for silent execution.
    Product search element and number of pages are defined.

**CAPTCHA Handling:**   
    If a CAPTCHA is detected, the script uses amazoncaptcha to solve it and proceed with the scraping process.

**Scraping Details:**
    Title, Price (original and featured), Rating, and Product Link are extracted using BeautifulSoup.
    The script handles cases where some products may not have certain details available (e.g., missing price or rating).

**Pagination:**
    The script automatically clicks the "Next" button and scrapes data from subsequent pages until the defined num_pages are scraped.

**Excel Operations:**
    After scraping, the data is saved to an Excel file with clickable product links using openpyxl.

**Error Handling:**
    The script includes error handling for:<br>
        *CAPTCHA Solving:* Attempts to solve CAPTCHA up to 3 times.<br>
        *Timeout and NoSuchElementException:* Retries to load elements when they are not found.<br>
