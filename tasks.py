from robocorp.tasks import task
from robocorp import browser
from RPA.Tables import Tables
from RPA.HTTP import HTTP
from RPA.PDF import PDF
import zipfile
import os
from time import sleep
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    browser.configure(
        slowmo=1,
    )
    open_robot_order_website()
    give_up_rights()
    orders = get_orders()

    all_receipts = []

    for order in orders:
        order_number = order["Order number"]

        fill_the_form(order)
        preview_order()
        robot_screenshot(order_number)   # Screenshot BEFORE submitting (preview is visible)
        submit_order(retries=10)          # Submit the order
        receipt_html = receipt(order_number)  # Read receipt AFTER successful submission
        screenshot_to_receipt(order_number)
        all_receipts.append(f"output/receipt_{order_number}.pdf")
        next_order()

    zip_receipts(all_receipts)


def open_robot_order_website():
    # Navigates to the given URL
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def give_up_rights():
    # Agree to the terms of service
    page = browser.page()
    page.wait_for_selector("button:text('OK')", timeout=5000)
    page.click("button:text('OK')")


def get_orders():
    # Access orders CSV
    tables = Tables()
    http = HTTP()
    url = "https://robotsparebinindustries.com/orders.csv"
    file_path = "orders.csv"

    # Download specified file and overwrite if needed
    http.download(url, target_file=file_path, overwrite=True)

    # CSV into a table
    orders = tables.read_table_from_csv(file_path)
    return orders


def fill_the_form(order):
    # Filling the order form
    page = browser.page()

    # Select the head
    page.select_option("#head", order["Head"])

    # Select body
    page.click(f"#id-body-{order['Body']}")

    # Enter part number for legs
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])

    # Enter shipping address
    page.fill("#address", order["Address"])


def preview_order():
    page = browser.page()
    page.click("#preview")


def robot_screenshot(order_number):
    # Take screenshot of robot preview BEFORE submitting
    page = browser.page()
    page.locator("#robot-preview-image").screenshot(path=f"output/robot_{order_number}.png")


def submit_order(retries=10):
    page = browser.page()
    for attempt in range(1, retries + 1):
        try:
            page.click("#order")                             # Use the exact button ID
            page.wait_for_selector("#receipt", timeout=3000)
            return                                           # Success — exit the function
        except Exception as e:
            print(f"Order did not go through, retrying {attempt}/{retries}... ({e})")
            sleep(1)
    raise Exception(f"Order failed after {retries} retries — #receipt never appeared.")


def receipt(order_number):
    # Read receipt AFTER order has been successfully submitted
    page = browser.page()
    pdf = PDF()

    receipt_html = page.locator("#receipt").inner_html()

    pdf.html_to_pdf(
        receipt_html,
        f"output/receipt_{order_number}.pdf"
    )

    return receipt_html


def screenshot_to_receipt(order_number):
    # Embed the robot screenshot into the receipt PDF
    pdf = PDF()

    pdf.add_files_to_pdf(
        files=[
            f"output/receipt_{order_number}.pdf",
            f"output/robot_{order_number}.png"
        ],
        target_document=f"output/receipt_{order_number}.pdf"
    )


def next_order():
    page = browser.page()
    page.click("#order-another")
    # Dismiss the modal that appears after clicking "Order another robot"
    page.wait_for_selector("button:text('OK')", timeout=5000)
    page.click("button:text('OK')")


def zip_receipts(all_receipts):
    zip_path = "output/all_receipts.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for receipt_file in all_receipts:
            if os.path.exists(receipt_file):
                zf.write(receipt_file, arcname=os.path.basename(receipt_file))
            else:
                print(f"Warning: file not found, skipping: {receipt_file}")
    print(f"All receipts have been zipped into all_receipts.zip ({len(all_receipts)} files)")
