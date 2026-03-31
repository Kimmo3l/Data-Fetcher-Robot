from robocorp.tasks import task
from robocorp import browser
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import os
import shutil
from datetime import datetime

# Määritellään tiedostonimi vakiona, jotta se on helppo vaihtaa
EXCEL_FILE = "sahko_hinnat.xlsx"

@task
def robot_data_fetcher():
    """Hakee sähkön hinnat, tallentaa Exceliin ja tekee raportin."""
    browser.configure(
        headless=True,
        slowmo=100,
        )

    try:
        open_porssisahko_website()
        # 1. Haetaan hinnat listana
        hinnat_lista = fetch_hourly_prices()
        
        # 2. Tallennetaan Exceliin
        save_to_excel(hinnat_lista, EXCEL_FILE)
        
        # 3. Lasketaan tilastot (halvin, kallein, säästö)
        halvin, kallein, saasto = calculate_prices_and_savings(EXCEL_FILE)
        
        # 4. Arkistointi ja raportointi
        backup_excel(EXCEL_FILE)
        pdf_file = convert_excel_to_pdf(EXCEL_FILE)
        
        print(f"Ajo valmis! Halvin: {halvin} snt, Kallein: {kallein} snt. Säästö: {saasto:.2f} snt.")

    finally:
        # robocorp.browser sulkee selaimen automaattisesti, 
        # joten emme kutsu tässä close_browser()-metodia.
        try:
            excel = Files()
            excel.close_workbook()
        except:
            # Jos työkirjaa ei ollut edes avattu, ei tehdä mitään
            pass

def open_porssisahko_website():
    browser.goto("https://www.porssisahko.fi/")

def fetch_hourly_prices():
    """Hakee sähkön hinnat sivulta varmistaen, että luetaan vain yksi taulukko."""
    page = browser.page()
    tunnit_hinnat = []

    # Odotetaan, että sivu on latautunut
    page.wait_for_selector("table")

    # .first varmistaa, että otamme vain ensimmäisen vastaantulevan taulukon
    # (Sivulla on kaksi identtistä taulukkoa eri näyttökooille)
    rows = page.locator("table").first.locator("tr").all()

    # Käydään läpi rivit, range(25) varmistaa tasan 24 tuntia otsikon jälkeen
    for i in range(25):
        try:
            # Käytetään tarkkaa järjestysnumeroa (nth) ensimmäisen taulukon sisällä
            tunti = rows[i].locator("td:nth-child(1)").inner_text().strip()
            hinta_teksti = rows[i].locator("td:nth-child(2)").inner_text().strip()
            
            # Puhdistetaan hinta numeroksi
            hinta = float(hinta_teksti.replace(",", ".").split()[0])
            
            tunnit_hinnat.append({
                "Tunti": tunti, 
                "Hinta": hinta
            })
        except Exception as e:
            print(f"Virhe rivillä {i}: {e}")
            continue

    return tunnit_hinnat

def save_to_excel(data_list, filename):
    excel = Files()
    
    # POISTETAAN vanha tiedosto, jos se on olemassa, jotta rivejä ei kerry liikaa
    if os.path.exists(filename):
        try:
            os.remove(filename)
        except PermissionError:
            print("Sulje Excel-tiedosto ennen ajoa!")
            return

    # Luodaan täysin uusi työkirja
    excel.create_workbook(filename)
    
    today = datetime.now().strftime("%Y-%m-%d")
    final_data = [{"Päivämäärä": today, **rivi} for rivi in data_list]
    
    # Nyt tässä on tasan 24 riviä
    excel.append_rows_to_worksheet(final_data, header=True)
    excel.save_workbook()
    excel.close_workbook()

def calculate_prices_and_savings(filename):
    excel = Files()
    excel.open_workbook(filename)
    table = excel.read_worksheet_as_table(header=True)
    excel.close_workbook()

    today = datetime.now().strftime("%Y-%m-%d")
    # Suodatetaan vain tämän päivän rivit
    today_rows = [row for row in table if row["Päivämäärä"] == today]
    
    hinnat = [float(row["Hinta"]) for row in today_rows]
    halvin = min(hinnat)
    kallein = max(hinnat)
    # Laskennallinen säästö (esimerkki)
    saasto = (kallein - halvin) * 3 
    
    return halvin, kallein, saasto

def backup_excel(filename):
    os.makedirs("varmuuskopiot", exist_ok=True)
    shutil.copy(filename, f"varmuuskopiot/backup_{datetime.now().strftime('%Y%m%d')}.xlsx")

def convert_excel_to_pdf(filename):
    pdf = PDF()
    pdf_name = filename.replace(".xlsx", ".pdf")
    # Huom: RPA.PDF:n excel_to_pdf vaatii yleensä Office-asennuksen tai kikkailua. 
    # Usein helpompaa on tallentaa HTML-raportti ja muuttaa se PDF:ksi.
    return pdf_name