from __future__ import annotations

import argparse
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional, Tuple, Set

import pandas as pd

# Zależności runtime: selenium, openpyxl, chromedriver zgodny z Chrome
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import NoSuchElementException, WebDriverException
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
except Exception:
    webdriver = None
    ChromeOptions = None
    By = None
    Keys = None
    NoSuchElementException = Exception
    WebDriverException = Exception
    WebDriverWait = None
    EC = None


def _base_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def _normalize_pl(s: str) -> str:
    if s is None:
        return ""
    s = s.strip()
    s = s.lower()
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def parse_args() -> argparse.Namespace:
    base = _base_dir()
    p = argparse.ArgumentParser(description="Agent: pobieranie klientów z CRM (Chrome devtools).")
    # Pliki
    p.add_argument("--towns-excel", default=str(base / "Miejscowosci.xlsx"))
    p.add_argument("--out-excel", default=str(base / "Klienci.xlsx"))
    p.add_argument("--sheet-towns", default=None)  # domyślnie: 'Lokalizacje' niżej
    p.add_argument("--sheet-out", default=None)    # domyślnie: 'Arkusz1'

    # Jednorazowe nadpisanie
    p.add_argument("--woj", default=None)
    p.add_argument("--miasto", default=None)

    # Chrome
    p.add_argument("--chrome-debug-port", type=int, default=9222)
    p.add_argument("--implicit-wait", type=float, default=3.0)
    p.add_argument("--page-wait", type=float, default=0.8)

    # Selektory pól (pierwotne – nie zmieniamy)
    p.add_argument("--xpath-woj-input", default="//th[contains(., 'Województwo')]//input")
    p.add_argument("--xpath-miasto-input", default="//th[contains(., 'Miasto')]//input")

    # Tabela wyników
    p.add_argument("--rows-selector", default="//table//tbody/tr")
    p.add_argument("--name-selector", default="td.oneLineWithEllipsis")
    p.add_argument("--nip-selector", default=None)

    # Czekanie / paginacja
    p.add_argument("--spinner-xpath", default=None)
    p.add_argument("--wait-timeout", type=float, default=12.0)
    p.add_argument("--next-selector-x", default="//a[contains(@class,'next') or contains(., '›') or contains(., 'Następna')]")
    p.add_argument("--max-pages", type=int, default=200)

    # Dodatki (bez zmian logiki)
    p.add_argument("--press-enter-after-city", action="store_true",
                   help="(Ignorowane — Enter jest wysyłany zawsze po mieście.)")
    p.add_argument("--retype-retries", type=int, default=3)
    return p.parse_args()


@dataclass
class ClientRow:
    woj: str
    miasto: str
    nip: str
    nazwa: str


def read_towns(towns_excel: Path, sheet: Optional[str], cli_woj: Optional[str], cli_miasto: Optional[str]) -> List[Tuple[str, str]]:
    if cli_woj or cli_miasto:
        return [(cli_woj or "", cli_miasto or "")]
    if not towns_excel.exists():
        raise FileNotFoundError(f"Brak pliku: {towns_excel}")
    df = pd.read_excel(towns_excel, sheet_name=(sheet or "Lokalizacje"))
    cols = [c.strip().lower() for c in df.columns]
    woj_col = None
    miasto_col = None
    for i, c in enumerate(cols):
        if c in {"woj", "województwo", "wojewodztwo"}:
            woj_col = df.columns[i]
        if c in {"miasto", "miejscowosc", "miejscowość"}:
            miasto_col = df.columns[i]
    if miasto_col is None and "miasto" in df.columns:
        miasto_col = "miasto"
    if miasto_col is None:
        miasto_col = df.columns[0]
    pairs = []
    for _, row in df.iterrows():
        woj = str(row[woj_col]).strip() if woj_col else ""
        miasto = str(row[miasto_col]).strip()
        if miasto and miasto.lower() != "nan":
            pairs.append((woj, miasto))
    return pairs


def attach_chrome(debug_port: int, implicit_wait: float):
    if webdriver is None:
        raise RuntimeError("Brak zależności Selenium/openpyxl/chromedriver.")
    opts = ChromeOptions()
    opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
    try:
        driver = webdriver.Chrome(options=opts)
    except WebDriverException as e:
        raise RuntimeError(
            "Nie połączono z działającym Chrome.\n"
            "Uruchom: chrome.exe --remote-debugging-port=9222\n"
            "Upewnij się, że chromedriver pasuje do wersji Chrome."
        ) from e
    driver.implicitly_wait(implicit_wait)
    return driver


def _safe_text(el) -> str:
    if el is None:
        return ""
    try:
        return el.text.strip()
    except Exception:
        return ""


def _by_xpath_first(drv, xpath: str):
    try:
        return drv.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return None


def _by_xpath_all(drv, xpath: str):
    try:
        return drv.find_elements(By.XPATH, xpath)
    except NoSuchElementException:
        return []


def _by_css_first(ctx, css: Optional[str]):
    if not css:
        return None
    try:
        return ctx.find_element(By.CSS_SELECTOR, css)
    except NoSuchElementException:
        return None


def _value_of(inp):
    try:
        return (inp.get_attribute("value") or "").strip()
    except Exception:
        return ""


def _set_value_js(drv, inp, value: str) -> bool:
    """Ustawienie wartości poprzez JS + zdarzenia input/change; zwraca True jeśli się przyjęło."""
    try:
        drv.execute_script(
            "const el=arguments[0], val=arguments[1];"
            "el.focus();"
            "el.value='';"
            "el.dispatchEvent(new Event('input', {bubbles:true}));"
            "el.value=val;"
            "el.dispatchEvent(new Event('input', {bubbles:true}));"
            "el.dispatchEvent(new Event('change', {bubbles:true}));",
            inp, value
        )
        time.sleep(0.05)
        return _value_of(inp) == (value or "").strip()
    except Exception:
        return False


def _clear_and_type_verified(drv, inp, value: str, retries: int, micro_pause: float = 0.05) -> bool:
    """Najpierw próba normalnego wpisania + weryfikacja; jeśli utnie (np. 'skie'→'kie'), to ustawienie JS-em."""
    want = (value or "").strip()
    for _ in range(max(1, retries)):
        try:
            inp.click()
            inp.send_keys(Keys.CONTROL, "a")
            inp.send_keys(Keys.DELETE)
            if want:
                for chunk in [want[i:i+16] for i in range(0, len(want), 16)]:
                    inp.send_keys(chunk)
                    time.sleep(micro_pause)
            time.sleep(micro_pause * 3)
        except Exception:
            pass
        if _value_of(inp) == want:
            return True
    return _set_value_js(drv, inp, want)


def _wait_for_results(drv, rows_xpath: str, spinner_xpath: Optional[str],
                      prev_signature: Optional[str], timeout: float) -> str:
    end = time.time() + max(1.0, timeout)
    last_sig = None
    while time.time() < end:
        if spinner_xpath:
            try:
                WebDriverWait(drv, 0.5).until(EC.invisibility_of_element_located((By.XPATH, spinner_xpath)))
            except Exception:
                pass
        rows = _by_xpath_all(drv, rows_xpath)
        n = len(rows)
        first = ""
        if n:
            try:
                first = rows[0].text.strip()[:64]
            except Exception:
                first = ""
        sig = f"{n}|{first}"
        if sig and sig != prev_signature:
            time.sleep(0.2)
            return sig
        time.sleep(0.2)
        last_sig = sig
    return last_sig or (prev_signature or "")


# --------- minimalny fallback dla pól (Miasto/Województwo) ----------
def _find_input_by_attrs_anywhere(drv, term_lower: str):
    """Znajduje widoczny <input> po atrybutach placeholder/aria-label/name/id zawierających frazę (także w iframe)."""
    def _query_in_doc():
        map_from = "AĄBCĆDEĘFGHIJKLŁMNŃOÓPRSŚTUWYZŹŻ"
        map_to   = "aąbcćdeęfghijklłmnńoóprsśtuwyzźż"
        xp = ("//input[("
              f"contains(translate(@placeholder,'{map_from}','{map_to}'), '{term_lower}') or "
              f"contains(translate(@aria-label,'{map_from}','{map_to}'), '{term_lower}') or "
              f"contains(translate(@name,'{map_from}','{map_to}'), '{term_lower}') or "
              f"contains(translate(@id,'{map_from}','{map_to}'), '{term_lower}'))]")
        els = _by_xpath_all(drv, xp)
        for el in els:
            try:
                if el.is_displayed() and el.is_enabled():
                    return el
            except Exception:
                continue
        return None

    el = _query_in_doc()
    if el:
        return el

    frames = _by_xpath_all(drv, "//iframe | //frame")
    for fr in frames:
        try:
            drv.switch_to.frame(fr)
            el = _query_in_doc()
            if el:
                return el
        except Exception:
            pass
        finally:
            try:
                drv.switch_to.default_content()
            except Exception:
                pass
    return None
# -------------------------------------------------------------------


def set_filters_and_wait(drv, args, woj: str, miasto: str, prev_sig: Optional[str]) -> str:
    inp_woj = _by_xpath_first(drv, args.xpath_woj_input)
    inp_miasto = _by_xpath_first(drv, args.xpath_miasto_input)

    if inp_woj is None and woj:
        inp_woj = _find_input_by_attrs_anywhere(drv, "wojewodztwo") or _find_input_by_attrs_anywhere(drv, "województwo")
    if inp_miasto is None:
        inp_miasto = _find_input_by_attrs_anywhere(drv, "miasto")

    if inp_woj is None and woj:
        print("[WARN] Nie znaleziono pola 'Województwo' — jadę tylko z 'Miasto'.")
    if inp_miasto is None:
        raise RuntimeError("Nie znaleziono pola 'Miasto' — podaj własny XPath przez --xpath-miasto-input.")

    if inp_woj and woj:
        ok = _clear_and_type_verified(drv, inp_woj, woj, retries=args.retype_retries)
        if not ok:
            print("[WARN] Pole 'Województwo' nie przyjęło pełnej wartości — kontynuuję.")

    ok = _clear_and_type_verified(drv, inp_miasto, miasto, retries=args.retype_retries)
    if not ok:
        print("[WARN] Pole 'Miasto' nie przyjęło pełnej wartości — kontynuuję mimo to.")
    try:
        inp_miasto.send_keys(Keys.ENTER)
    except Exception:
        pass

    time.sleep(args.page_wait if hasattr(args, "page_wait") else 0.8)
    return _wait_for_results(drv, args.rows_selector, args.spinner_xpath, prev_sig, args.wait_timeout)


# ========================= PAGINACJA (Prime paginator + SVG) ==================
def _scroll_into_view(drv, el):
    try:
        drv.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
        time.sleep(0.05)
    except Exception:
        pass


def _is_disabled_like(el) -> bool:
    try:
        cls = (el.get_attribute("class") or "").lower()
        aria = (el.get_attribute("aria-disabled") or "").lower()
        disabled = (el.get_attribute("disabled") is not None)
        return ("disabled" in cls) or ("inactive" in cls) or (aria == "true") or disabled
    except Exception:
        return False


def _find_next_button(drv, preferred_xpath: str):
    el = _by_xpath_first(drv, preferred_xpath)
    if el:
        return el
    el = _by_xpath_first(drv, "//a[@rel='next' or contains(., 'Następna') or contains(., '›') or contains(., '»')]")
    if el:
        return el
    el = _by_xpath_first(drv, "//button[contains(@class,'p-paginator-next')]")
    if el:
        return el
    el = _by_xpath_first(drv, "//*[name()='svg' and contains(@class,'p-paginator-icon')]/ancestor::button[1]")
    if el:
        return el
    el = _by_xpath_first(drv, "//*[name()='svg' and contains(@class,'p-icon')]/ancestor::*[self::a or self::button][1]")
    if el:
        return el
    return None


def click_next_and_wait(drv, next_xpath: str, rows_xpath: str,
                        spinner_xpath: Optional[str], current_sig: Optional[str],
                        timeout: float) -> Optional[str]:
    btn = _find_next_button(drv, next_xpath)
    if not btn or _is_disabled_like(btn):
        return None

    _scroll_into_view(drv, btn)
    clicked = False
    try:
        btn.click()
        clicked = True
    except Exception:
        try:
            drv.execute_script("arguments[0].click();", btn)
            clicked = True
        except Exception:
            pass
    if not clicked:
        return None

    new_sig = _wait_for_results(drv, rows_xpath, spinner_xpath, current_sig, timeout)
    if not new_sig or new_sig == current_sig:
        try:
            _scroll_into_view(drv, btn)
            drv.execute_script("arguments[0].click();", btn)
        except Exception:
            pass
        new_sig = _wait_for_results(drv, rows_xpath, spinner_xpath, current_sig, timeout)

    return new_sig if new_sig and new_sig != current_sig else None
# ==============================================================================


def guess_cell_by_header(tr, header_texts: Iterable[str]) -> Optional[str]:
    try:
        tds = tr.find_elements(By.CSS_SELECTOR, "td")
        for td in tds:
            label_bits = []
            for attr in ("aria-label", "title", "data-label"):
                try:
                    val = td.get_attribute(attr) or ""
                    if val:
                        label_bits.append(val.lower())
                except Exception:
                    pass
            joined = " ".join(label_bits)
            if any(h in joined for h in header_texts):
                return _safe_text(td)
        # fallback: 10 cyfr = potencjalny NIP
        for td in tds:
            text = _safe_text(td).replace(" ", "")
            if text.isdigit() and len(text) == 10:
                return text
    except Exception:
        pass
    return None


def extract_rows_from_page(drv, rows_xpath: str, name_css: str | None,
                           nip_css: str | None,
                           woj: str, miasto: str) -> List['ClientRow']:
    data: List[ClientRow] = []
    trs = _by_xpath_all(drv, rows_xpath)
    for tr in trs:
        nazwa = ""
        nip = ""
        el_name = _by_css_first(tr, name_css)
        if el_name:
            nazwa = _safe_text(el_name)
        el_nip = _by_css_first(tr, nip_css) if nip_css else None
        if el_nip:
            nip = _safe_text(el_nip)
        else:
            guessed = guess_cell_by_header(tr, {"nip"})
            if guessed:
                nip = guessed
        nip = "".join(ch for ch in nip if ch.isdigit())
        data.append(ClientRow(woj=woj, miasto=miasto, nip=nip, nazwa=nazwa))
    return data


def click_next_if_exists(*args, **kwargs) -> bool:
    """Alias dla zgodności (nieużywany)."""
    return False


# ======================== ZAPIS / ODCZYT EXCELA ================================
def write_full_excel(path: Path, rows: List[ClientRow], sheet_name: Optional[str]) -> Path:
    """Zapisuje pełny zestaw do Excela. Gdy plik zablokowany, zapisuje do *_autosave.xlsx."""
    df = pd.DataFrame([{
        "Województwo": r.woj, "Miasto": r.miasto, "NIP": r.nip,
        "Nazwa": r.nazwa, "Zebrano": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    } for r in rows])
    if df.empty:
        df = pd.DataFrame(columns=["Województwo","Miasto","NIP","Nazwa","Zebrano"])

    try:
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as xw:
            df.to_excel(xw, index=False, sheet_name=(sheet_name or "Arkusz1"))
        return path
    except Exception as e:
        alt = path.with_name(path.stem + "_autosave.xlsx")
        try:
            with pd.ExcelWriter(alt, engine="openpyxl", mode="w") as xw:
                df.to_excel(xw, index=False, sheet_name=(sheet_name or "Arkusz1"))
            print(f"[WARN] Nie udało się nadpisać {path.name} ({e}). Zapisano do {alt.name}.")
            return alt
        except Exception as e2:
            print(f"[ERROR] Zapis do Excela nie powiódł się: {e2}")
            return path


def read_existing_results(path: Path, sheet_name: Optional[str]) -> Tuple[List[ClientRow], Set[Tuple[str, str]]]:
    """Wczytuje istniejące wyniki i zwraca (lista ClientRow, zbiór znormalizowanych (woj, miasto))."""
    rows: List[ClientRow] = []
    done: Set[Tuple[str, str]] = set()
    if not path.exists():
        return rows, done
    try:
        df = pd.read_excel(path, sheet_name=(sheet_name or "Arkusz1"))
    except Exception as e:
        print(f"[WARN] Nie udało się odczytać istniejącego pliku wyników ({e}) — start od zera.")
        return rows, done

    if df is None or df.empty:
        return rows, done

    # mapowanie kolumn po normalizacji
    colmap = {}
    for c in df.columns:
        n = _normalize_pl(str(c))
        if "wojew" in n:
            colmap["woj"] = c
        elif "miasto" in n or "miejscow" in n:
            colmap["miasto"] = c
        elif n == "nip" or "nip" in n:
            colmap["nip"] = c
        elif "nazwa" in n:
            colmap["nazwa"] = c

    if "miasto" not in colmap:
        # jeśli nie ma kolumny Miasto, nie mamy jak wykryć duplikatów
        print("[WARN] W wynikach brak kolumny 'Miasto' — nie będę pomijał przetworzonych.")
        return rows, done

    for _, r in df.iterrows():
        woj = str(r.get(colmap.get("woj", ""), "")).strip()
        miasto = str(r.get(colmap["miasto"], "")).strip()
        nip = str(r.get(colmap.get("nip", ""), "")).strip()
        nazwa = str(r.get(colmap.get("nazwa", ""), "")).strip()
        rows.append(ClientRow(woj=woj, miasto=miasto, nip="".join(ch for ch in nip if ch.isdigit()), nazwa=nazwa))
        done.add((_normalize_pl(woj), _normalize_pl(miasto)))
    return rows, done
# ==============================================================================


def run() -> int:
    args = parse_args()
    base = _base_dir()

    towns_excel = Path(args.towns_excel).expanduser()
    if not towns_excel.is_absolute():
        towns_excel = (base / towns_excel).resolve()
    out_excel = Path(args.out_excel).expanduser()
    if not out_excel.is_absolute():
        out_excel = (base / out_excel).resolve()

    print(f"[INFO] Plik miejscowości: {towns_excel}")
    print(f"[INFO] Plik wyjściowy:   {out_excel}")

    # 1) Wczytaj listę miejscowości
    try:
        pairs = read_towns(towns_excel, args.sheet_towns, args.woj, args.miasto)
    except Exception as e:
        print(f"[ERROR] Nie udało się wczytać miejscowości: {e}")
        return 2

    print(f"[INFO] Wszystkich par (woj, miasto): {len(pairs)}")

    # 2) Wczytaj istniejące wyniki i zbiór przerobionych
    existing_rows, processed = read_existing_results(out_excel, args.sheet_out)
    if existing_rows:
        print(f"[INFO] Wczytano istniejące wyniki: {len(existing_rows)} wierszy, unikalnych miast: {len(processed)}")

    # 3) Ustal punkt startu – pierwszy nieprzerobiony
    start_idx = 0
    for i, (w, m) in enumerate(pairs):
        if (_normalize_pl(w), _normalize_pl(m)) not in processed:
            start_idx = i
            break
    else:
        print("[OK] Wszystkie miejscowości z listy już są w wynikach — nic do zrobienia.")
        return 0

    if start_idx > 0:
        print(f"[INFO] Pomijam {start_idx} przerobionych pozycji. Start od #{start_idx+1}: {pairs[start_idx][0] or '-'}, {pairs[start_idx][1]}")

    # 4) Start Selenium
    try:
        driver = attach_chrome(args.chrome_debug_port, args.implicit_wait)
    except Exception as e:
        print(f"[ERROR] Selenium/Chrome: {e}")
        return 3

    # 5) Zaczynamy z already existing rows, żeby nie nadpisać pliku przy flushu
    all_rows: List[ClientRow] = list(existing_rows)

    # 6) Główna pętla
    for idx, (woj, miasto) in enumerate(pairs[start_idx:], start=start_idx+1):
        # pomiń, jeśli jakimś cudem jest już przerobione
        if (_normalize_pl(woj), _normalize_pl(miasto)) in processed:
            print(f"[{idx}/{len(pairs)}] ({woj or '-'}, {miasto}) — pomijam (już w wynikach).")
            continue

        tag = f"[{idx}/{len(pairs)}] ({woj or '-'}, {miasto})"
        print(f"{tag} Ustawiam filtry...")

        try:
            prev_sig = None
            sig = set_filters_and_wait(driver, args, woj, miasto, prev_sig)
        except Exception as e:
            print(f"{tag} [ERROR] Filtry: {e}")
            # flush aktualnego stanu (dotychczasowe + to co zebraliśmy)
            written = write_full_excel(out_excel, all_rows, args.sheet_out)
            print(f"{tag} [FLUSH] Zapisano stan do: {written.name} (wierszy={len(all_rows)})")
            continue

        print(f"{tag} Zbieram wiersze...")
        seen_signatures: Set[str] = set([sig])
        page = 1
        city_rows_before = len(all_rows)
        while True:
            try:
                page_rows = extract_rows_from_page(
                    driver,
                    rows_xpath=args.rows_selector,
                    name_css=args.name_selector,
                    nip_css=args.nip_selector,
                    woj=woj, miasto=miasto
                )
                print(f"{tag} Strona {page}: wierszy={len(page_rows)}")
                all_rows.extend(page_rows)
            except Exception as e:
                print(f"{tag} [WARN] Ekstrakcja nie powiodła się na stronie {page}: {e}")

            if page >= args.max_pages:
                print(f"{tag} Limit stron ({args.max_pages}) osiągnięty.")
                break

            new_sig = click_next_and_wait(
                driver,
                next_xpath=args.next_selector_x,
                rows_xpath=args.rows_selector,
                spinner_xpath=args.spinner_xpath,
                current_sig=sig,
                timeout=args.wait_timeout
            )
            if not new_sig or new_sig in seen_signatures:
                break
            seen_signatures.add(new_sig)
            sig = new_sig
            page += 1

        # oznacz miasto jako przerobione (nawet jeśli brak wierszy – liczy się fakt przejścia)
        processed.add((_normalize_pl(woj), _normalize_pl(miasto)))

        # ---- FLUSH PO KAŻDYM MIEŚCIE ----
        written = write_full_excel(out_excel, all_rows, args.sheet_out)
        new_rows_cnt = len(all_rows) - city_rows_before
        print(f"{tag} [FLUSH] Zapisano stan do: {written.name} (przybyło {new_rows_cnt} wierszy; łącznie={len(all_rows)})")
        # ----------------------------------

    print(f"[INFO] Zapis końcowy: {out_excel}")
    final_written = write_full_excel(out_excel, all_rows, args.sheet_out)
    print(f"[OK] Gotowe. Wiersze zapisane łącznie: {len(all_rows)} → plik: {final_written.name}")
    try:
        driver.quit()
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(run())
