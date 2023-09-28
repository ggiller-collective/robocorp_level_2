from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

http = HTTP()
tables = Tables()
pdf = PDF()
archive = Archive()

browser.configure(
    slowmo=100,
)


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    get_orders()
    archive_receipts()


def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def close_annoying_modal():
    page = browser.page()
    page.click(".alert-buttons button:nth-child(1)")


def read_csv_into_table(file_path):
    table = tables.read_table_from_csv(file_path)
    return table


def fill_the_form(row):
    close_annoying_modal()
    page = browser.page()
    page.select_option("#head", row["Head"])
    page.click(f".form-group:nth-child(2) input[value='{row['Body']}']")
    page.fill(".form-group:nth-child(3) input", row["Legs"])
    page.fill("#address", row["Address"])
    page.click("#preview")
    robot_preview_image = page.wait_for_selector("#robot-preview-image")
    screenshot_path = f"output/robot_images/{row['Order number']}.png"
    robot_preview_image.screenshot(path=screenshot_path)
    max_retries = 5
    for attempt in range(max_retries):
        try:
            page.click("#order")
            # If the click is successful, break the loop
            page.wait_for_selector("#receipt", timeout=500)
            break
        except Exception:
            print(f"Attempt {attempt+1} failed. Retrying...")
            if attempt + 1 == max_retries:
                print("Max retries reached. Raising the exception.")
                raise
    pdf_file_path = store_receipt_as_pdf(row["Order number"])
    embed_screenshot_to_receipt(screenshot_path, pdf_file_path)
    page.click("#order-another")


def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf_file_path = f"output/receipts/{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, pdf_file_path)
    return pdf_file_path


def embed_screenshot_to_receipt(screenshot_path, pdf_file_path):
    list_of_files = [screenshot_path]
    pdf.add_files_to_pdf(
        files=list_of_files, target_document=pdf_file_path, append=True
    )


def get_orders():
    csv_url = "https://robotsparebinindustries.com/orders.csv"
    http.download(csv_url, "orders.csv", overwrite=True)
    table = read_csv_into_table("orders.csv")
    for row in table:
        fill_the_form(row)


def archive_receipts():
    archive.archive_folder_with_zip("output/receipts", "output/receipts.zip")
