from robocorp.tasks import task
from robocorp import browser
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Email.ImapSmtp import ImapSmtp
from dotenv import load_dotenv
import urllib.parse

import os
import shutil
from datetime import datetime

load_dotenv()

# Määritellään tiedostonimi vakiona, jotta se on helppo vaihtaa
EXCEL_FILE = "sahko_hinnat.xlsx"

@task
def robot_data_fetcher():
    """Hakee sähkön hinnat, tallentaa Exceliin ja tekee raportin."""
    browser.configure(
        headless=False,
        slowmo=10,
        )

    try:
        open_porssisahko_website()
        hinnat_lista = fetch_hourly_prices() # Hakee kaikki 24h
        save_to_excel(hinnat_lista, EXCEL_FILE)
        
        # Tilastot koko päivästä (PDF:ää ja popupia varten)
        halvin_pva, kallein_pva, saasto = calculate_prices_and_savings(EXCEL_FILE)
        
        # --- UUSI LOGIIKKA FUNKTIONA ---
        ehdotus = etsi_paras_tuleva_tunti(hinnat_lista)

        if ehdotus:
            # Lähetetään sähköposti, jossa ehdotetaan tulevaa tuntia, 
            # mutta kerrotaan vertailun vuoksi koko päivän kallein hinta.
            laheta_sahkoposti_ilmoitus(ehdotus['Hinta'], kallein_pva, ehdotus['Tunti'])
            print(f"Ehdotus lähetetty: {ehdotus['Tunti']} ({ehdotus['Hinta']} snt)")
        else:
            print("Ei enää tulevia tunteja tälle päivälle.")
        # -------------------------------

        backup_excel(EXCEL_FILE)
        pdf_file = convert_excel_to_pdf(EXCEL_FILE)
        nayta_tulos_selaimessa(halvin_pva, kallein_pva, saasto)
        
    finally:
        try:
            excel = Files()
            excel.close_workbook()
        except:
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

def convert_excel_to_pdf(excel_filename):
    pdf = PDF()
    pdf_name = excel_filename.replace(".xlsx", ".pdf")
    
    # Luetaan Excel-tiedon sisältö raporttia varten
    excel = Files()
    excel.open_workbook(excel_filename)
    data = excel.read_worksheet_as_table(header=True)
    excel.close_workbook()

    # Luodaan HTML-sisältö
    html_content = "<h1>Sähkön hintaraportti</h1><table border='1'><tr><th>Aika</th><th>Hinta (snt)</th></tr>"
    for row in data:
        html_content += f"<tr><td>{row['Tunti']}</td><td>{row['Hinta']}</td></tr>"
    html_content += "</table>"

    # Tallennetaan PDF:ksi
    pdf.html_to_pdf(html_content, pdf_name)
    return pdf_name

def nayta_tulos_selaimessa(halvin, kallein, saasto):
    page = browser.page()
    page.set_viewport_size({"width": 400, "height": 400})
    # Kirjoitetaan tulokset suoraan valkoiselle sivulle
    sisalto = f"<h1>Tiedot haettu ja raportti valmis!</h1><br><h2>Sähkoraportti:</h2><p>Halvin: {halvin}</p><p>Kallein: {kallein}</p><p>Säästö: {saasto:.2f}</p>"
    page.set_content(sisalto)
    
    # Odotetaan, että käyttäjä näkee sivun
    import time
    time.sleep(8)

from datetime import datetime, timedelta
import urllib.parse
from RPA.Email.ImapSmtp import ImapSmtp
import os

def laheta_sahkoposti_ilmoitus(halvin, kallein, tunti_teksti):
    """Lähettää yhteenvedon ja dynaamisen kalenterilinkin Gmaililla."""
    gmail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)

    # Luetaan arvot ympäristömuuttujista
    KAYTTAJA = os.getenv("GMAIL_USER")
    SALASANA = os.getenv("GMAIL_PASSWORD")
    VASTAANOTTAJA = KAYTTAJA

    # 1. PUHDISTETAAN AIKAVÄLI (Kestää eri viivatyypit: –, -, —)
    siisti_tunti_teksti = tunti_teksti.replace('–', '-').replace('—', '-')
    aikaväli_osat = siisti_tunti_teksti.split('-')

    # Muotoillaan ajat muotoon HHMM00 (esim. 03:00 -> 030000)
    alku_tunti = aikaväli_osat[0].strip().replace(':', '') + "00"

    # Tarkistetaan, onko lopputunti "00" (eli yli puolenyön)
    loppu_tunti_teksti = aikaväli_osat[1].strip()
    if loppu_tunti_teksti == "00:00":
        # Jos lopputunti on 00:00, se kuuluu seuraavalle päivälle
        paiva_str = datetime.now().strftime("%Y%m%d")
        seuraava_paiva = (datetime.now() + timedelta(days=1)).strftime("%Y%m%d")
        loppu_tunti = loppu_tunti_teksti.replace(':', '') + "00"
        dates = f"{paiva_str}T{alku_tunti}/{seuraava_paiva}T{loppu_tunti}"
    else:
        # Normaalitapaus: sama päivä
        paiva_str = datetime.now().strftime("%Y%m%d")
        loppu_tunti = loppu_tunti_teksti.replace(':', '') + "00"
        dates = f"{paiva_str}T{alku_tunti}/{paiva_str}T{loppu_tunti}"

    # 2. LUODAAN KALENTERILINKKI
    # Enkoodataan otsikko URL-muotoon (välilyönnit -> %20 jne.)
    otsikko = urllib.parse.quote(f"Halpaa sähköä: {halvin} snt")
    linkki = f"https://www.google.com/calendar/render?action=TEMPLATE&text={otsikko}&dates={dates}"

    # 3. LÄHETETÄÄN VIESTI
    try:
        gmail.authorize(account=KAYTTAJA, password=SALASANA)

        # Käytetään viestissä "a" ja "o" -kirjaimia varmuuden vuoksi
        viesti = (
            f"Pörssisähkorobotti on suorittanut haun.\n\n"
            f"Päivän halvin hinta: {halvin} snt/kWh (klo {tunti_teksti})\n"
            f"Päivän kallein hinta: {kallein} snt/kWh\n\n"
            f"LISÄÄ HALVIN TUNTI KALENTERIIN KLIKKAAMALLA TÄSTÄ:\n"
            f"{linkki}\n\n"
            f"Terveisin, Robottisi"
        )

        gmail.send_message(
            sender=KAYTTAJA,
            recipients=VASTAANOTTAJA,
            subject=f"Sähköraportti: Halvin {halvin} snt",
            body=viesti
        )
        print("Sähköposti lähetetty onnistuneesti kalenterilinkin kera!")

    except Exception as e:
        print(f"Sähköpostin lähetys epäonnistui: {e}")

def etsi_paras_tuleva_tunti(hinnat_lista):
    """Suodattaa listasta vain tulevat tunnit ja palauttaa halvimman niistä."""
    nykyinen_tunti = datetime.now().hour
    tulevat_hinnat = []

    for rivi in hinnat_lista:
        # Puhdistetaan viiva ja otetaan alkutunnin tunti (esim. "16:00 – 17:00" -> 16)
        alku_tunti = int(rivi['Tunti'].replace('–', '-').split(':')[0].strip())
        
        if alku_tunti >= nykyinen_tunti:
            tulevat_hinnat.append(rivi)

    if not tulevat_hinnat:
        return None
    
    # Palauttaa koko sanakirjan (Tunti ja Hinta) halvimmasta tulevasta hetkestä
    return min(tulevat_hinnat, key=lambda x: x['Hinta'])
