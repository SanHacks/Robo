"""Template robot with Python."""

import os

from Browser import Browser
from Browser.utils.data_types import SelectAttribute
from RPA.Excel.Files import Files
from RPA.HTTP import HTTP
from RPA.PDF import PDF

browser = Browser()


def open_the_intranet_website():
    """Open the intranet website."""
    browser.new_page("https://robotsparebinindustries.com/")


def login_in():
    """Login in."""
    browser.input_text("css:#username", "maria")
    browser.input_text("css:#password", "thoushallNotpass")
    browser.click("text=Login")


def download_the_excel_file():
    """Download the excel file."""
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)


def fill_and_submit_the_form_for_one_person(sales_rep):
    """Fill and submit the form for one person."""
    browser.type_text("css=#firstname", sales_rep["First Name"])
    browser.type_text("css=#lastname", sales_rep["Last Name"])
    browser.type_text("css=#salesresult", sales_rep["Sales"])
    browser.select_options_by(
        "css=#salestarget", SelectAttribute.VALUE, sales_rep["Target"])
    browser.click("text=Submit")


def collect_the_results():
    """Collect the results."""
    browser.take_screenshot(
        filename=f"{os.getcwd()}/results.png", selecotr="css=div.sales-summary")


def export_the_table_as_a_pdf():
    """Export the table as a pdf."""
    sales_results_html = browser.get_property(
        "css=#sales-results", "outerHTML")
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")


def log_out():
    """Log out."""
    browser.click("text=Logout")


def main():
    try:
        open_the_intranet_website()
        login_in()
        download_the_excel_file()
        excel = Files()
        excel.open_workbook("SalesData.xlsx")
        sales_reps = excel.read_worksheet_as_table(header=True)
        excel.close_workbook()
        for sales_rep in sales_reps:
            fill_and_submit_the_form_for_one_person(sales_rep)
        collect_the_results()
        export_the_table_as_a_pdf()
        log_out()
    finally:
        log_out()
        browser.playwright.close()


if __name__ == "__main__":
    main()
