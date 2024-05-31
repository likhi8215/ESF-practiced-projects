from robocorp.tasks import task
from robocorp import browser
from RPA.Browser.Selenium import Selenium
from RPA.Tables import Tables
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive  # Import Archive library for ZIP archiving
import time
import os

pdf = PDF()
browser = Selenium()

@task
def order_the_robots_from_robotsparebin(): 
    browser.configure(
        slowmo=100,
    )
    open_the_browser()
    close_annoying_modal()
    orders = get_orders()
    process_orders(orders) 
    archive_receipts()  # Call the function to archive receipt files

def open_the_browser():
    browser.goto("https://robotsparebinindustries.com/")

def get_orders():
    http = HTTP()
    tables = Tables()
    url = "https://robotsparebinindustries.com/orders.csv"
    filename = "orders.csv"
    http.download(url, filename, overwrite=True)
    return tables.read_table_from_csv(filename)

def process_orders(orders):
    for order in orders:
        fill_the_form(order)

def fill_the_form(order):
    for attempt in range(3):
        try:
            # Fill the form fields with the order data
            browser.input_text('id:orderNumber', order['Order number'])
            browser.input_text('id:address', order['Address'])
            browser.input_text('id:headline', order['Robot model'])

            # For the leg number input, we need a different approach since the id changes
            leg_input_xpath = '//input[contains(@placeholder, "number of legs")]'
            browser.input_text(leg_input_xpath, order['Number of legs'])

            # Submit the form
            browser.click('css:button.btn-primary')

            # Check for error message
            time.sleep(2)  # Wait a bit for any error message to appear
            error_message = browser.get_text("css:.error-message")

            if not error_message:
                print(f"Successfully submitted order: {order['Order number']}")
                receipt_path = store_receipt_as_pdf(order['Order number'])
                screenshot_path = take_screenshot(order['Order number'])
                embed_screenshot_to_receipt(screenshot_path, receipt_path)  # Embed screenshot to PDF
                break
            else:
                print(f"Error: {error_message}")
                print("Retrying order...")
        except Exception as e:
            print(f"Attempt {attempt + 1} failed for order {order['Order number']}: {e}")
            if attempt == 2:  # On the last attempt, log failure
                print(f"Failed to submit order {order['Order number']} after 3 attempts")
        time.sleep(2)  # Add a delay before retrying

def close_annoying_modal():
    try:
        browser.click('css:button.close')
        print("Modal closed successfully")
    except Exception as e:
        print("Failed to close modal:", e)

def store_receipt_as_pdf(order_number):
    pdf = PDF()  # Create an instance of the PDF class
    receipt_dir = 'output/receipts'  # Directory to store receipts

    # Check if the directory exists, create if not
    if not os.path.exists(receipt_dir):
        os.makedirs(receipt_dir)

    # Define the path for the PDF file
    receipt_path = os.path.join(receipt_dir, f"order_receipt_{order_number}.pdf")

    # Convert the current HTML to a PDF
    pdf.html_to_pdf(browser.get_html(), receipt_path)

    # Print a message indicating the PDF has been stored
    print(f"Stored receipt as PDF: {receipt_path}")

    # Return the path to the stored PDF
    return receipt_path

def take_screenshot(order_number):
    # Take a screenshot of the robot
    screenshot_path = f"output/screenshots/robot_{order_number}.png"
    browser.screenshot(screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()  # Create an instance of the PDF class
    
    # Embed the screenshot to the PDF
    pdf.append_images_to_pdf([screenshot], pdf_file)

    # Print a message indicating the screenshot has been embedded
    print(f"Embedded screenshot to receipt PDF: {pdf_file}")

from RPA.Archive import Archive

def archive_receipts():
    archive = Archive()
    receipt_dir = 'output/receipts'
    zip_file = 'output/receipts.zip'

    # Create a ZIP archive of receipt files
    archive.create_zip_archive(zip_file, receipt_dir)

    # Print a message indicating the ZIP file creation
    print(f"Receipts archived successfully: {zip_file}")

    # Clean up: remove individual receipt files after archiving
    for root, _, files in os.walk(receipt_dir):
        for file in files:
            os.remove(os.path.join(root, file))





















# =======================================level 2 itself=============================================


# from robocorp.tasks import task
# from robocorp import browser
# from RPA.Tables import Tables
# from RPA.HTTP import HTTP
# import csv

# @task
# def order_the_robots_from_robotsparebin(): 
#     """
#     1.Orders robots from RobotSpareBin Industries Inc.
#     Saves the order HTML receipt as a PDF file.
#     Saves the screenshot of the ordered robot.
#     Embeds the screenshot of the robot to the PDF receipt.
#     Creates ZIP archive of the receipts and the images.
#     """
#     open_the_browser()
#     close_annoying_modal()
#     orders=get_orders()
#     # print(orders)
#     process_orders(orders)  
#     for order in orders:
#         fill_the_form(order)
    


# def open_the_browser():
#     """opening the browser to integrate with intranet"""
#     browser.goto("https://robotsparebinindustries.com/")



# def get_orders():
#     http=HTTP()
#     tables=Tables()

#     # define the URL of CSV file
#     url="https://robotsparebinindustries.com/orders.csv"
#     filename="orders.csv"

#     # downloading the CSV file
#     http.download(url,filename,overwrite=True)#======================================download point

#     # read the CSV file into a table
#     table = tables.read_table_from_csv(filename)
#     # returning table
#     return table


# #continous looping of orders
# def process_orders(orders):
#     for order in orders:
#         print("processing orders:{order}")
    

# """closing annoying modal"""
# def close_annoying_modal():
#     try:
#         click_on_it='css:button.close'
#         browser.click(click_on_it)
#         print("modal closed successfully")
#     except Exception as e:
#         print("Failed to close modal:{e}")

# def fill_the_form(row):
#     #  browser.wait_until_page_contains_element("css:#order_form", timeout=10)
#     browser.input_text("css:#order_number", row['Order number'])
#     browser.input_text("css:#customer_name", row['Customer name'])
#     browser.select_from_list_by_value("css:#product_select", row['Product'])
#     leg_number_input = browser.find_element("css:[id*=leg_number]")
#     leg_number_input.input_text(row['Leg number'])
#     browser.click("css:#submit_button")
#     print("Form filled successfully")















# ========================================Level 1 with XPATH===========================================================

# from robocorp.tasks import task
# from robocorp import browser
# import time
# from lxml import etree
# from RPA.HTTP import HTTP
# from RPA.Excel.Files import Files

# @task
# def robot_spare_bin_python():
#     """Insert the sales data for the week and export it as a PDF"""
#     browser.configure(
#         slowmo=100,
#     )
#     open_the_intranet_website()
#     log_in()
#     download_excel_file()
#     fill_form_with_excel_data()

# def open_the_intranet_website():
#     """Navigates to the given URL"""
#     browser.goto("https://robotsparebinindustries.com/")
#     print("open browser")

# def log_in():
#     """Fills in the login form and clicks the 'Log in' button"""
#     page = browser.page()
#     page.fill("//input[@id='username']", "maria")
#     time.sleep(2)
#     page.fill("//input[@id='password']", "thoushallnotpass")
#     page.click("//button[contains(text(), 'Log in')]")
#     time.sleep(2)

# def download_excel_file():
#     """Downloads excel file from the given URL"""
#     http = HTTP()
#     http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True) 

# def fill_form_with_excel_data():
#     """Read data from excel and fill in the sales form"""
#     excel = Files()
#     excel.open_workbook("SalesData.xlsx")
#     worksheet = excel.read_worksheet_as_table("data", header=True)
#     excel.close_workbook()

#     for row in worksheet:
#         fill_and_submit_sales_form(row)      

# def fill_and_submit_sales_form(sales_rep):
#     """Fills in the sales data and click the 'Submit' button"""
#     page = browser.page()

#     page.fill("//input[@id='firstname']", sales_rep["First Name"])
#     page.fill("//input[@id='lastname']", sales_rep["Last Name"])
#     page.select_option("//select[@id='salestarget']", str(sales_rep["Sales Target"]))
#     page.fill("//input[@id='salesresult']", str(sales_rep["Sales"]))
#     page.click("text=Submit")















# ============================================Level 1====================================================================






# from robocorp.tasks import task
# from robocorp import browser
# import time
# from RPA.HTTP import HTTP
# from RPA.Excel.Files import Files
# from RPA.PDF import PDF

# @task
# def robot_spare_bin_python():
#     """Insert the sales data for the week and export it as a PDF"""
#     browser.configure(
#         slowmo=100,
#     )
#     open_the_intranet_website()
#     log_in()
#     download_excel_file()
#     fill_form_with_excel_data()
#     collect_results()
#     export_as_pdf()
#     log_out()
   

# def open_the_intranet_website():
#     """Navigates to the given URL"""
#     browser.goto("https://robotsparebinindustries.com/")
#     print("open browser")

# def log_in():
#     """Fills in the login form and clicks the 'Log in' button"""
#     page = browser.page()
#     page.fill("#username", "maria")
#     time.sleep(2)
#     page.fill("#password", "thoushallnotpass")
#     page.click("button:text('Log in')")
#     time.sleep(2)

# def download_excel_file():
#     """Downloads excel file from the given URL"""
#     http = HTTP()
#     http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)  

# def fill_form_with_excel_data():
#     """Read data from excel and fill in the sales form"""
#     excel = Files()
#     excel.open_workbook("SalesData.xlsx")
#     worksheet = excel.read_worksheet_as_table("data", header=True)
#     excel.close_workbook()

#     for row in worksheet:
#         fill_and_submit_sales_form(row)      

# def fill_and_submit_sales_form(sales_rep):
#     """Fills in the sales data and click the 'Submit' button"""
#     page = browser.page()

#     page.fill("#firstname", sales_rep["First Name"])
#     page.fill("#lastname", sales_rep["Last Name"])
#     page.select_option("#salestarget", str(sales_rep["Sales Target"]))
#     page.fill("#salesresult", str(sales_rep["Sales"]))
#     page.click("text=Submit")


# def collect_results():
#     """Take a screenshot of the page"""
#     page=browser.page()
#     page.screenshot(path="output/sales_summary.png")

# def export_as_pdf():
#     """Export the data to a pdf file"""
#     page=browser.page()
#     sales_results_html=page.locator("#sales-results").inner_html()
#     pdf = PDF()
#     pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

# def log_out():
#     """Log out from browser"""
#     page=browser.page()
#     page.click("text=log out") 

