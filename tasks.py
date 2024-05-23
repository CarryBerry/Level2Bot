from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    screenshots_folder, receipts_folder = create_output_folders()
    open_robot_order_website()
    download_orders_file()
    process_orders(screenshots_folder, receipts_folder)

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def create_output_folders():
    """Creates output folders."""
    os.makedirs("output/screenshots", exist_ok=True)
    os.makedirs("output/receipts", exist_ok=True)
    return "output/screenshots", "output/receipts"

def download_orders_file():
    """Downloads orders.csv file from the given URL"""
    HTTP().download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Gets the orders from the CSV table"""
    return Tables().read_table_from_csv("orders.csv", columns=["Order number","Head","Body","Legs","Address"])

def close_annoying_modal():
    """Closes the annoying modal if needed"""
    page = browser.page()
    okButton_selector = "text=OK"
    if page.is_visible(okButton_selector): page.click(okButton_selector)

def fill_the_form(order):
    """Fill the order form"""
    page = browser.page()

    page.select_option('#head', order['Head'])
    page.click('#id-body-'+order["Body"])
    page.fill('input[placeholder="Enter the part number for the legs"]', order['Legs'])
    page.fill('#address', order['Address'])
    page.click('#preview')
    page.click('#order')

def handle_alert():
    """Unstable part - Exception handling if needed"""
    page = browser.page()
    while page.query_selector('.alert.alert-danger') is not None:
        page.click('#order')

def open_another_order():
    browser.page().click('#order-another')

def process_orders(screenshots_folder, receipts_folder):
    orders = get_orders()

    for order in orders:
        close_annoying_modal()
        fill_the_form(order)
        handle_alert()

        receipt = store_receipt_as_pdf(order['Order number'], receipts_folder)
        screenshot = screenshot_robot_preview(order['Order number'], screenshots_folder)
        embed_screenshot_to_receipt(screenshot, receipt)

        open_another_order()
    
    archive_receipts(receipts_folder)

def store_receipt_as_pdf(order_number, receipts_folder):
    """Store each receipt as a PDF"""
    receipt_html = browser.page().locator("#receipt").inner_html()
    receipt_path = receipts_folder+"/order_{0}.pdf".format(order_number)
    PDF().html_to_pdf(receipt_html, receipt_path)
    return receipt_path

def screenshot_robot_preview(order_number, screenshots_folder):
    """Screenshot the robot preview"""
    screenshot_path = screenshots_folder+"/order_"+order_number+".png"
    browser.page().locator(selector="#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, receipt):
    """Add screenshot to each PDF receipt"""
    PDF().add_files_to_pdf(files=[screenshot], target_document=receipt, append=True)

def archive_receipts(receipts_folder):
    """Create Zip Archive of all reciepts"""
    Archive().archive_folder_with_zip(receipts_folder, receipts_folder+".zip", recursive=True)