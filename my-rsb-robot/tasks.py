from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

@task
def minimal_task():
   """Inserts the sales data for the week and expot it as a pdf."""
   browser.configure(
       slowmo = 100,
    )
   open_the_intranet_website()
   log_in()
   download_excel_file()
   fill_form_with_excel_data()
   collect_results()
   export_as_pdf()
   log_out()
   
def open_the_intranet_website():
    """Opening the website URL."""
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Logging in to the website."""
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def download_excel_file():
    """Downlod the Escel file form the giving URL."""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/SalesData.xlsx", overwrite="True")

def fill_form_and_submit(sales_rep):
    """Filling nad submit slaes form."""
    page = browser.page();
    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")

def fill_form_with_excel_data():
    """Reads data from excel file and fill the form."""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_form_and_submit(row)

def collect_results():
    """Take screenshot of the result page."""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def export_as_pdf():
    """Export the result page as pdf."""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

def log_out():
    """Press the log out button."""
    page = browser.page()
    page.click = ("text=Log out")
