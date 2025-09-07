"""
AUTOMATYCZNY MONITOR SCHOWKA
Wersja: 2025-07-31  (poprawione wiązanie numeru z promocją + filtr nieaktywnych kont)
Autor: (twoje imię)
"""
import sys, time, os, logging, random, re, requests
from datetime import datetime

# --- zależności ----------------------------------------------------------------
missing = []
for lib in ("pyperclip", "keyboard", "requests", "openpyxl"):
    try:
        __import__(lib)
    except ModuleNotFoundError:
        missing.append(lib)
if missing:
    print("❌ Brakuje bibliotek:", ", ".join(missing))
    print("➡️  pip install", " ".join(missing)); sys.exit(1)

import pyperclip, keyboard
from openpyxl import load_workbook, Workbook

# --- ścieżki --------------------------------------------------------------------
BASE_DIR   = os.path.dirname(os.path.abspath(sys.argv[0]))
excel_path = os.path.join(BASE_DIR, "zielone.xlsx")
log_path   = os.path.join(BASE_DIR, "log.txt")
debug_path = os.path.join(BASE_DIR, "clipboard_debug.txt")

# --- logger ---------------------------------------------------------------------
logging.basicConfig(
    filename=log_path,
    filemode="a",
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

# --- pomocnicze -----------------------------------------------------------------
def ensure_excel_file(path: str) -> None:
    """Tworzy plik Excel z nagłówkami, jeśli nie istnieje."""
    if not os.path.exists(path):
        wb = Workbook()
        wb.active.append(["NIP", "Numer", "Promocja", "Data", "Uslugi"])
        wb.save(path)
        print("🆕 Utworzono nowy plik:", os.path.basename(path))

# --- API GUS --------------------------------------------------------------------
GUS_API_KEY = "bf96e683d9a9449b8958"

class GUSConnector:
    def __init__(self, api_key: str = GUS_API_KEY):
        self.api_key = api_key
        self.session = requests.Session()
        self.url = "https://wyszukiwarkaregon.stat.gov.pl/wsBIR/UslugaBIRzewnPubl.svc"
        self.session.headers.update({"Content-Type": "application/soap+xml;charset=UTF-8"})
        self.sid = None

    def login(self) -> bool:
        envelope = (
            '<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" '
            'xmlns:ns="http://CIS/BIR/PUBL/2014/07" '
            'xmlns:wsa="http://www.w3.org/2005/08/addressing">'
            '<soap:Header>'
            '<wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzewnPubl/Zaloguj</wsa:Action>'
            f'<wsa:To>{self.url}</wsa:To>'
            '</soap:Header>'
            '<soap:Body>'
            '<ns:Zaloguj>'
            f'<ns:pKluczUzytkownika>{self.api_key}</ns:pKluczUzytkownika>'
            '</ns:Zaloguj>'
            '</soap:Body>'
            '</soap:Envelope>'
        )
        try:
            resp = self.session.post(self.url, data=envelope, timeout=5)
            if resp.status_code == 200:
                m = re.search(r"<ZalogujResult>([^<]+)</ZalogujResult>", resp.text)
                if m:
                    self.sid = m.group(1)
                    self.session.headers.update({"sid": self.sid})
                    return True
            logger.error("API GUS: nieudane logowanie (status %s)", resp.status_code)
        except Exception as e:
            logger.error("API GUS: %s", e)
        return False

    def search_by_nip(self, nip: str):
        envelope = (
            '<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" '
            'xmlns:ns="http://CIS/BIR/PUBL/2014/07" '
            'xmlns:wsa="http://www.w3.org/2005/08/addressing" '
            'xmlns:dat="http://CIS/BIR/PUBL/2014/07/DataContract">'
            '<soap:Header>'
            '<wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzewnPubl/DaneSzukajPodmioty</wsa:Action>'
            f'<wsa:To>{self.url}</wsa:To>'
            '</soap:Header>'
            '<soap:Body>'
            '<ns:DaneSzukajPodmioty>'
            '<ns:pParametryWyszukiwania>'
            f'<dat:Nip>{nip}</dat:Nip>'
            '</ns:pParametryWyszukiwania>'
            '</ns:DaneSzukajPodmioty>'
            '</soap:Body>'
            '</soap:Envelope>'
        )
        try:
            resp = self.session.post(self.url, data=envelope, timeout=5)
            text = resp.text if resp.status_code == 200 else ""
            if "<Fault" in text or "<faultcode>" in text:
                if not self.login():
                    return None
                resp = self.session.post(self.url, data=envelope, timeout=5)
                text = resp.text if resp.status_code == 200 else ""
            if not text:
                return None

            m = re.search(r"<DaneSzukajPodmiotyResult>(.*?)</DaneSzukajPodmiotyResult>",
                          text, flags=re.DOTALL)
            if not m:
                return None

            import html, xml.etree.ElementTree as ET
            inner_xml = html.unescape(m.group(1)).strip()
            if not inner_xml:
                return None
            root = ET.fromstring(f"<root>{inner_xml}</root>")
            dane = root.find(".//dane")
            if dane is None:
                return None
            name = dane.findtext("Nazwa") or ""
            inactive = bool(dane.findtext("DataZakonczeniaDzialalnosci"))
            return {"Nazwa": name, "Nieaktywna": inactive}
        except Exception as e:
            logger.error("API GUS: %s", e)
            return None

# --- EKSTRAKCJA DANYCH ---------------------------------------------------------
def _promo_from_details(details_line: str) -> (str, str):
    """Zwraca (promocja, data_do) z linii szczegółów usługi."""
    # data "do DD-MM-YYYY"
    m_do = re.search(r"do (\d{2}-\d{2}-\d{4})", details_line)
    data_do = m_do.group(1) if m_do else ""
    # ogon po ostatniej dacie
    dates = list(re.finditer(r"\d{2}-\d{2}-\d{4}", details_line))
    tail = details_line[dates[-1].end():].strip() if dates else details_line.strip()
    # pierwsze pole z ogona (oddzielone >=2 spacjami)
    promo = re.split(r"\s{2,}", tail)[0].strip()
    return promo, data_do

def extract_client_data(text: str, gus: GUSConnector):
    """
    Zwraca tuple: (lista_aktywnych, lista_pominiętych_nip)
    Każdy wpis ma postać [NIP, Numer, Promocja, Data_do, Uslugi]
    """
    clients, skipped = [], []
    blocks = re.split(r"Wszystkich kont:\s*\d+", text, flags=re.I)
    nip_re = re.compile(r"(?<=Pokaż[\s\t])\d{10}")

    for blk in blocks:
        if not blk.strip():
            continue
        nip_match = nip_re.search(blk)
        if not nip_match:
            continue
        nip = nip_match.group()
        # --- DB logic: skip whole block if "DB" found ---
        if "DB" in blk:
            skipped.append(f"{nip} (DB)")
            logger.info(f"Pominięto NIP {nip} (DB)")
            continue
        # -----------------------------------------------

        gus_data = gus.search_by_nip(nip)
        if not gus_data or gus_data["Nieaktywna"]:
            skipped.append(nip)
            continue

        lines = blk.splitlines()
        active_account = True  # domyślnie aktywne, dopóki nie napotkamy info o 0 z X
        numbers_info = []      # (nr_pełny, promo, data_do)

        for idx, line in enumerate(lines):
            # wykryj przełączenie aktywności konta
            m_acc = re.search(r"(\d+)\s+z\s+(\d+)", line)
            if m_acc:
                active_account = m_acc.group(1) != "0"
                continue

            # linia z numerem usługi (11 cyfr zaczynających się od 48)
            m_num = re.match(r"\s*48\d{9}\s*$", line)
            if m_num and active_account:
                full_num = m_num.group().strip()
                details = lines[idx + 1] if idx + 1 < len(lines) else ""
                promo, data_do = _promo_from_details(details)
                numbers_info.append((full_num, promo, data_do))

        if not numbers_info:
            skipped.append(nip)
            continue

        # preferuj numer z prefiksem 485
        chosen = next((n for n in numbers_info if n[0].startswith("485")), numbers_info[0])
        numer_9cyfr = chosen[0][2:]  # obcięcie prefiksu 48
        promo_val, data_do = chosen[1], chosen[2]

        # policz aktywne usługi tylko z aktywnych kont
        uslugi = sum(int(m.group(1))
                     for m in re.finditer(r"(\d+)\s+z\s+(\d+)", blk)
                     if m.group(1) != "0")

        clients.append([nip, numer_9cyfr, promo_val, data_do, uslugi])
        logger.debug("✅ Dodano NIP=%s, Numer=%s, Promocja='%s'", nip, numer_9cyfr, promo_val)

    return clients, skipped

# --- zapis do Excela -----------------------------------------------------------
def insert_data_to_excel(rows):
    try:
        wb = load_workbook(excel_path)
        sh = wb.active
        for r in rows:
            sh.append(r)
        wb.save(excel_path)
        print(f"✅ Dopisano {len(rows)} rekord(y) do Excela.")
    except Exception as e:
        logger.error("Excel: %s", e)
        print("❌ Błąd zapisu:", e)

# --- monitor schowka -----------------------------------------------------------
def monitor_clipboard(gus: GUSConnector):
    prev = ""
    print("📋 Czekam na Ctrl+C…")
    while True:
        try:
            if keyboard.is_pressed("ctrl+c"):
                time.sleep(0.4)
                clip = pyperclip.paste()
                if clip == prev or len(clip.strip()) < 20:
                    continue
                prev = clip
                with open(debug_path, "w", encoding="utf-8") as f:
                    f.write(clip)

                active_rows, skipped_nips = extract_client_data(clip, gus)
                if active_rows:
                    insert_data_to_excel(active_rows)
                if skipped_nips:
                    print("⛔  Pominięto NIP-y:", ", ".join(skipped_nips))
                    logger.info("Pominięte NIP-y: %s", ", ".join(skipped_nips))
                elif not active_rows:
                    print("⚠️  Brak aktywnych firm lub brak promocji.")
            time.sleep(0.1)
        except Exception as e:
            logger.critical("Monitor: %s", e)
            print("❌ Błąd monitora:", e)

# --- START ---------------------------------------------------------------------
if __name__ == "__main__":
    ensure_excel_file(excel_path)
    gus = GUSConnector()
    if gus.login():
        monitor_clipboard(gus)
    else:
        print("❌ Nie udało się połączyć z API GUS.")
