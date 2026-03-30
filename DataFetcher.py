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

def fetch_hourly_prices():
    """Hakee tunnittain sähkön hinnat sivulta."""
    page = browser.page()
    tunnit_hinnat = []

    # Oletetaan, että hinnat ovat taulukossa, jossa jokaisella rivillä on tunti ja hinta
    for i in range(24):
        tunti = f"{i:02d}-{i+1:02d}"
        hinta_elementti = page.locator(f"tr:nth-child({i+1}) td:nth-child(2)")
        hinta = float(hinta_elementti.inner_text().split()[0])
        tunnit_hinnat.append({"Tunti": tunti, "Hinta (snt/kWh)": hinta})

    return tunnit_hinnat

def save_to_excel(tunnit_hinnat, excel_filename):
    """Tallentaa tunnittain hinnat staattiseen Excel-tiedostoon."""
    excel = Files()

    # Avaa olemassa oleva tiedosto tai luo uusi
    if os.path.exists(excel_filename):
        excel.open_workbook(excel_filename)
    else:
        excel.create_workbook()

    # Lisää uusi rivi päivämäärällä
    today = datetime.now().strftime("%Y-%m-%d")
    data = [{"Päivämäärä": today, **tunti} for tunti in tunnit_hinnat]
    excel.append_rows_to_worksheet(data, header=True)
    excel.save_workbook(excel_filename)

def calculate_prices_and_savings(excel_filename):
    """Laskee halvimman, kalleimman hinnan ja säästön."""
    excel = Files()
    excel.open_workbook(excel_filename)
    worksheet = excel.read_worksheet_as_table(header=True)

    # Etsi viimeisimmän päivän hinnat
    today = datetime.now().strftime("%Y-%m-%d")
    today_data = [row for row in worksheet if row["Päivämäärä"] == today]

    hinnat = [float(row["Hinta (snt/kWh)"]) for row in today_data]
    halvin_hinta = min(hinnat)
    kallein_hinta = max(hinnat)
    saasto = (kallein_hinta - halvin_hinta) * 3  # 3 kWh/h

    return halvin_hinta, kallein_hinta, saasto

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
