from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Email import Email
from RPA.Robocorp.WorkItems import WorkItems
import os
import shutil
from datetime import datetime

@task
def robot_data_fetcher():
    """Get the electricity prices and insert it to excel"""
    browser.configure(slowmo=100)

    try:
        open_porssisahko_website()
        halvin_hinta, kallein_hinta = fetch_sahko_prices()
        excel_filename = save_to_excel(halvin_hinta, kallein_hinta)
        saasto = calculate_savings(halvin_hinta, kallein_hinta)
        backup_excel(excel_filename)
        pdf_filename = convert_excel_to_pdf(excel_filename)
        send_pdf_by_email(pdf_filename, saasto)
        print(f"Robotin ajo valmis! Säästö: {saasto:.2f} senttiä/kWh.")

    finally:
        browser.close_browser()
        excel = Files()
        excel.close_workbook()


def open_porssisahko_website():
    """Open the porssisahko website"""
    browser.goto("https://www.porssisahko.fi/")

def fetch_sahko_prices():
    """Fetch the cheapest and most expensive prices for the day"""
    page = browser.page()

### Nämä eivät ole vielä oikein
    halvin_hinta = float(page.locator(".price-min").inner_text().split()[0])
    kallein_hinta = float(page.locator(".price-max").inner_text().split()[0])

    return halvin_hinta, kallein_hinta

def save_to_excel(halvin_hinta, kallein_hinta):
    """Save prices to Excel by date"""
    excel = Files()
    today = datetime.now().strftime("%Y-%m-%d")
    excel_filename = f"sahko_hinnat_{today}.xlsx"

### Testausta tälle
    data = [
        {"Päivämäärä": today, "Halvin hinta (snt/kWh)": halvin_hinta, "Kallein hinta (snt/kWh)": kallein_hinta}
    ]
    excel.create_workbook()
    excel.append_rows_to_worksheet(data, header=True)
    excel.save_workbook(excel_filename)

    return excel_filename

def calculate_savings(halvin_hinta, kallein_hinta):
    """Calculate savings with this device (3 kWh/h)."""
    return (kallein_hinta - halvin_hinta) * 3

def backup_excel(excel_filename):
    """Make a backup of the excel-file"""
    backup_folder = "varmuuskopiot"
    os.makedirs(backup_folder, exist_ok=True)
    shutil.copy(excel_filename, os.path.join(backup_folder, excel_filename))

def convert_excel_to_pdf(excel_filename):
    """Convert the excel to a pdf"""
    pdf = PDF()
    pdf_filename = excel_filename.replace(".xlsx", ".pdf")
    pdf.excel_to_pdf(excel_filename, pdf_filename)
    return pdf_filename

def send_pdf_by_email(pdf_filename, saasto):
    """Send the pdf to email addresses"""
    email = Email()
    email.smtp_server = "smtp.gmail.com"  # Muuta tarvittaessa
    email.smtp_port = 587
    email.smtp_username = "lähettäjän_email@gmail.com"
    email.smtp_password = "salasana"  # Käytä turvallista tapaa tallentaa salasana
    email.recipients = ["vastaanottaja1@example.com", "vastaanottaja2@example.com"]
    email.subject = f"Sähkön hinnat ja säästöraportti - {datetime.now().strftime('%Y-%m-%d')}"
    email.body = f"Liitteenä päivän sähkön hinnat. Säästö: {saasto:.2f} senttiä/kWh."
    email.attach_file(pdf_filename)
    email.send_message()
