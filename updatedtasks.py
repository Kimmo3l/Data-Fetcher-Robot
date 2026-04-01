from robocorp.tasks import task
from robocorp import browser
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import os
import shutil
from datetime import datetime
from RPA.Outlook.Application import Application

@task
def datafetcher_robot():
    browser.configure(
        slowmo=200,
        headless=False,
    )
    open_browser()
    hinnat = get_prices()
    copy_to_excel(hinnat)
    save_as_pdf(hinnat)
    make_backup()
    halvin = min(hinnat, key=lambda x: x["Hinta"])
    tee_kalenterimerkinta(halvin)

def open_browser():
    # Avataan selain ja kuitataan evästeet
    page = browser.goto("https://www.porssisahkoa.fi/")
    try:
        page.get_by_role("button", name="Suostun").click(timeout=5000)
    except:
        pass

def get_prices():
    # Haetaan hinnat, kellonaika ja niiden tallennus
    page = browser.page()
    taulukko = page.locator("table").first
    taulukko.scroll_into_view_if_needed()
    
    rows = taulukko.locator("tr").all()
    hinnat = []
    
    for row in rows:
        solut = row.locator("td").all()
        if len(solut) >= 2:
            aika = solut[0].inner_text().strip() 
            teksti = solut[1].inner_text().strip()
            if "," in teksti:
                hinta = float(teksti.replace(",", ".").split()[0])
                hinnat.append({"Tunti": aika, "Hinta": hinta}) 
    return hinnat

def copy_to_excel(hinnat):
    # Tallentaa päivän sähkön hintatiedot tunneittain Excel-tiedostoon
    excel = Files()
    tiedosto = "sahko.xlsx"
    
    if os.path.exists(tiedosto):
        try:
            os.remove(tiedosto)
        except:
            excel.close_workbook()

    excel.create_workbook(tiedosto)
    
    rivit_exceliin = []
    pvm_tanaan = datetime.now().strftime("%Y-%m-%d")
    
    for h in hinnat:
        rivit_exceliin.append({
            "Päivämäärä": pvm_tanaan,
            "Tunti": h["Tunti"],
            "Hinta": h["Hinta"]
        })
    
    excel.append_rows_to_worksheet(rivit_exceliin, header=True)
    excel.save_workbook()
    excel.close_workbook()

def save_as_pdf(hinnat):
    # Etsitään halvin ja kallein listasta
    h = min(hinnat, key=lambda x: x["Hinta"])
    k = max(hinnat, key=lambda x: x["Hinta"])
    
    teksti = f"""<h1>Sähköraportti valmis</h1>
    <p>Tänään sähkön alin hinta on {h['Hinta']} snt kellon aikana {h['Tunti']} 
    ja kallein hinta on {k['Hinta']} snt ja kellon aikana {k['Tunti']}.</p>
    <p>Excel taulukkoon on tallennettu tarkemmat tiedot.</p>"""
    
    PDF().html_to_pdf(teksti, "raportti.pdf")

def make_backup():
    os.makedirs("backups", exist_ok=True)
    shutil.copy("sahko.xlsx", f"backups/sahko_backup_{datetime.now().strftime('%d%m%y')}.xlsx")

def tee_kalenterimerkinta(h):
    # Luodaan uusi merkintä vanhan Outlookin kalenteriin, käyttäjä itse tallentaa
    outlook = Application()
    outlook.open_application()
    
    try:
        item = outlook.app.CreateItem(1)
        item.Subject = f"Halpa sähkö: {h['Hinta']} snt ({h['Tunti']})"
        tunti_numero = int(h['Tunti'].split(':')[0])
        aloitus = datetime.now().replace(hour=tunti_numero, minute=0, second=0, microsecond=0)
        item.Start = aloitus.strftime("%Y-%m-%d %H:%M")
        
        item.Duration = 60
        item.Body = f"Sähkö on halvimmillaan klo {h['Tunti']}."
        
        item.Save() 
        print("Merkintä tallennettu onnistuneesti!")
    except Exception as e:
        print(f"Virhe merkinnän luonnissa: {e}")
        try: item.Display()
        except: pass
