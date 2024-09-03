from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import time

# Initialize the browser and HTTP libraries
browser = Selenium()
http = HTTP()

def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.open_available_browser("https://robotsparebinindustries.com/")

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    browser.input_text("id=username", "maria")
    browser.input_text("id=password", "thoushallnotpass")
    browser.click_button("css:button[type='submit']")
    time.sleep(5)  # Ensure the page has loaded

def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    browser.input_text("id=firstname", sales_rep["First Name"])
    browser.input_text("id=lastname", sales_rep["Last Name"])
    browser.input_text("id=salesresult", str(sales_rep["Sales"]))
    browser.select_from_list_by_value("id=salestarget", str(sales_rep["Sales Target"]))
    browser.click_button("css:button[type='submit']")
    time.sleep(2)  # Wait for the submission to process

def download_excel_file():
    """Downloads the Excel file from the given URL"""
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def fill_form_with_excel_data():
    """Read data from Excel and fill in the sales form"""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_and_submit_sales_form(row)

def collect_results():
    """Take a screenshot of the page"""
    browser.capture_page_screenshot("output/sales_summary.png")

def export_as_pdf():
    """Export the data to a pdf file"""
    # Use XPath or a different selector if ID is not reliable
    browser.wait_until_element_is_visible("//div[@id='sales-results']", timeout=20)
    sales_results_html = browser.get_element_attribute("//div[@id='sales-results']", "innerHTML")
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")


def log_out():
    """Presses the 'Log out' button"""
    browser.wait_until_element_is_visible("//button[contains(text(), 'Log out')]", timeout=10)
    browser.click_button("//button[contains(text(), 'Log out')]")



# Run the main task
robot_spare_bin_python()










