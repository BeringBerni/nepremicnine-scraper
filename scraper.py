"""
scraper.py – Nepremičnine.net scraper
======================================
Uporablja DrissionPage (pravi Chrome v stealth načinu) → zaobide Cloudflare Turnstile.

Namestitev:
    py -m pip install DrissionPage beautifulsoup4 lxml --only-binary=:all:

Zagon:
    python scraper.py                              # Gorenjska, hiše, prodaja, 10 strani
    python scraper.py --regija gorenjska --vrsta hisa
    python scraper.py --regija vse --vrsta vse      # vse regije, vse vrste
    python scraper.py --regije gorenjska,savinjska --vrste hisa,stanovanje
    python scraper.py --test --csv                 # 1 stran, samo CSV
    python scraper.py --strani 5 --delay 2
    python scraper.py --headless                   # brez okna brskalnika
"""

import sys
import io
import os
import re
import time
import random
import argparse
import csv
from datetime import datetime
from urllib.parse import urljoin

# Windows cp1252 ne podpira slovenskih znakov – vsili UTF-8 za stdout/stderr
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

try:
    from DrissionPage import ChromiumPage, ChromiumOptions
except ImportError:
    print("✗ Namesti: py -m pip install DrissionPage --only-binary=:all:")
    sys.exit(1)

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("✗ Namesti: py -m pip install beautifulsoup4 lxml")
    sys.exit(1)

# ── Konfiguracija ─────────────────────────────────────────────────────────────
BASE_URL = "https://www.nepremicnine.net"

# Regije: ime za prikaz → URL slug
REGIJE: dict[str, str] = {
    "LJ-mesto":          "ljubljana",
    "LJ-okolica":        "ljubljana-okolica",
    "Gorenjska":         "gorenjska",
    "J. Primorska":      "juzna-primorska",
    "S. Primorska":      "severna-primorska",
    "Notranjska":        "notranjska",
    "Savinjska":         "savinjska",
    "Podravska":         "podravska",
    "Koroška":           "koroska",
    "Dolenjska":         "dolenjska",
    "Posavska":          "posavska",
    "Zasavska":          "zasavska",
    "Pomurska":          "pomurska",
}

# Vrste nepremičnin: ime za prikaz → URL slug
VRSTE: dict[str, str] = {
    "Hiša":               "hisa",
    "Stanovanje":         "stanovanje",
    "Vikend":             "vikend",
    "Posest":             "posest",
    "Poslovni prostor":   "poslovni-prostor",
    "Garaža":             "garaza",
    "Počitniški objekt":  "pocitniski-objekt",
}

# Akcija: prodaja ali najem
AKCIJE: list[str] = ["prodaja", "najem"]

# Mapa za shranjevanje posameznih zagonov in akumulirane baze
SCRAPING_RUNS_DIR = "scraping_runs"

# Pomožne funkcije za slug ↔ ime pretvorbo
_REGIJE_SLUGS = {v: k for k, v in REGIJE.items()}   # slug → ime
_VRSTE_SLUGS  = {v: k for k, v in VRSTE.items()}    # slug → ime

def regija_slug(s: str) -> str:
    """Sprejme slug ali ime regije, vrne slug."""
    s = s.strip()
    if s in REGIJE:
        return REGIJE[s]
    if s in _REGIJE_SLUGS:
        return s
    raise ValueError(f"Neznana regija: {s!r}. Veljavne: {list(REGIJE.keys())}")

def vrsta_slug(s: str) -> str:
    """Sprejme slug ali ime vrste, vrne slug."""
    s = s.strip()
    if s in VRSTE:
        return VRSTE[s]
    if s in _VRSTE_SLUGS:
        return s
    raise ValueError(f"Neznana vrsta: {s!r}. Veljavne: {list(VRSTE.keys())}")

def list_path(akcija: str, regija: str, vrsta: str) -> str:
    return f"/oglasi-{akcija}/{regija}/{vrsta}/"

DB_CONFIG = {
    "host":     "localhost",
    "port":     3306,
    "user":     "root",
    "password": "",
    "database": "nepremicnine",
    "charset":  "utf8mb4",
}

# ── Brskalnik (DrissionPage – stealth Chrome) ──────────────────────────────────
_browser: ChromiumPage | None = None
_headless: bool = False  # privzeto vidno okno – CF lažje bypass; --headless za tiho


def get_browser() -> ChromiumPage:
    """Lazy-init stealth Chrome brskalnika. Enkrat ustvarjen, potem se reuse."""
    global _browser
    if _browser is None:
        co = ChromiumOptions()
        if _headless:
            co.headless()
        # Stealth nastavitve
        co.set_argument("--no-sandbox")
        co.set_argument("--disable-dev-shm-usage")
        co.set_argument("--disable-blink-features=AutomationControlled")
        co.set_argument("--lang=sl-SI")
        co.set_argument("--window-size=1400,900")
        co.set_pref("intl.accept_languages", "sl-SI,sl,en-US,en")
        _browser = ChromiumPage(addr_or_opts=co)
    return _browser


def _wait_for_cf(page: ChromiumPage, timeout: int = 30) -> bool:
    """Počaka, da Cloudflare challenge mine. Vrne True ko je prava vsebina dostopna.

    POZOR: ne preverjaj 'cloudflare' v celotnem HTML – ta beseda se pojavi
    v footerju/skriptih VSAKE CF-zaščitene strani, ne le pri challengu.
    Zanesljivi indikatorji CF challenge strani so:
      - naslov (title) == "Just a moment..."
      - prisotnost specifičnih CF challenge elementov v DOM-u
    """
    time.sleep(2)  # počakaj da se stran začne nalagati
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            title = (page.title or "").lower()
            html  = page.html.lower()
            is_challenge = (
                "just a moment" in title
                or "checking your browser" in title
                or 'id="cf-challenge-running"' in html
                or 'id="cf-challenge-body"'    in html
                or 'name="cf-turnstile-response"' in html
            )
            if not is_challenge:
                return True
        except Exception:
            pass
        time.sleep(1.5)
    return False


def get_html(url, referer=None):
    """Odpre URL v stealth Chrome, počaka na CF challenge, vrne HTML."""
    browser = get_browser()
    for attempt in range(3):
        try:
            browser.get(url)
            ok = _wait_for_cf(browser, timeout=30)
            if not ok:
                print(f"  ⚠ CF challenge ni minil ({url[:60]}...)")
            return browser.html
        except Exception as e:
            if attempt == 2:
                print(f"  ✗ {url}: {e}")
                return None
            time.sleep(2 ** attempt)
    return None


# ── Parsing helpers ───────────────────────────────────────────────────────────
def parse_cena(text):
    if not text:
        return None
    text = str(text).strip()
    # Direct float conversion handles meta content="495000.00"
    try:
        v = int(float(text))
        if 5_000 < v < 50_000_000:
            return v
    except ValueError:
        pass
    # Slovenian format "495.000,00 €": dot = thousands sep, comma = decimal sep
    # Strip trailing decimal part ",XX"
    t = re.sub(r",\d{1,2}\s*\S*\s*$", "", text)
    digits = re.sub(r"[^\d]", "", t)
    if digits:
        v = int(digits)
        return v if 5_000 < v < 50_000_000 else None
    return None


def parse_m2(text):
    if not text:
        return None
    m = re.search(r"([\d]+(?:[.,]\d+)?)", text)
    if m:
        v = float(m.group(1).replace(".", "").replace(",", "."))
        return v if 0 < v < 500_000 else None
    return None


def parse_leto(text):
    if not text:
        return None
    m = re.search(r"(19|20)\d{2}", text)
    if m:
        v = int(m.group())
        return v if v <= datetime.now().year + 2 else None
    return None


def set_attribute(rec, key, val):
    """Map HTML table key→value to record fields."""
    key = key.lower().replace(":", "").strip()
    val = val.strip()
    if not key or not val:
        return
    if "cena" in key and "m2" not in key and "m²" not in key:
        rec.setdefault("Cena", parse_cena(val))
    elif key in ("lokacija", "naslov") or "kraj" in key or "mesto" in key:
        rec.setdefault("Lokacija", val)
    elif "občin" in key or "obcin" in key:
        rec.setdefault("Obcina", val)
    elif ("vrst" in key or "tip obj" in key) and "energet" not in key:
        rec.setdefault("VrstaObjekta", val)
    elif ("površin" in key or "povrsina" in key or "velikost" in key or "neto" in key) \
            and "parcel" not in key and "zemlj" not in key:
        rec.setdefault("VelikostM2", parse_m2(val))
    elif "parcel" in key or "zemljišč" in key or "zemljisc" in key:
        rec.setdefault("ZemljisteM2", parse_m2(val))
    elif "sob" in key:
        rec.setdefault("StSob", val)
    elif "leto" in key or "gradnj" in key or "zgrajen" in key:
        rec.setdefault("LetoGradnje", parse_leto(val))
    elif "energetsk" in key or "razred" in key:
        rec.setdefault("EnergetskiRazred", val[:10] if len(val) <= 10 else val[:10])


# ── Parser za stran z oglasi (seznam) – vrača polne zapise ────────────────────
def _get_total_pages(soup: "BeautifulSoup") -> int:
    """Iz paginacije prebere skupno število strani. Vrne 1 če ni paginacije."""
    # Primarna metoda: <ul data-pages="N">
    pag_ul = soup.find("ul", attrs={"data-pages": True})
    if pag_ul:
        try:
            return max(1, int(pag_ul.get("data-pages", 1)))
        except (ValueError, TypeError):
            pass
    # Rezervna: zadnji <li class="paging_last"> → href vsebuje številko
    last_li = soup.find("li", class_="paging_last")
    if last_li:
        a = last_li.find("a")
        href = a.get("href", "") if a else ""
        m = re.search(r"/(\d+)/?$", href)
        if m:
            return max(1, int(m.group(1)))
    # Nič ni najdeno – samo 1 stran
    return 1


def parse_listing_page(page_num, delay, path):
    """Običe eno stran z oglasi in razčleni vse kartice.
    Vrne (records, total_pages) – total_pages je zanesljiv samo za page_num==1."""
    url = BASE_URL + path if page_num == 1 else f"{BASE_URL}{path}{page_num}/"
    html = get_html(url, referer=BASE_URL + "/")
    if not html:
        return [], 1

    soup       = BeautifulSoup(html, "lxml")
    total_pgs  = _get_total_pages(soup) if page_num == 1 else 1
    results    = []
    seen       = set()

    # Vsak oglas je znotraj <div itemprop="item"> (schema.org Offer)
    for box in soup.find_all("div", attrs={"itemprop": "item"}):
        rec = {"DatumScrapa": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

        # ── URL ──────────────────────────────────────────────────────────────
        main_page = box.find("meta", {"itemprop": "mainEntityOfPage"})
        if main_page:
            oglas_url = main_page.get("content", "").strip()
            if oglas_url:
                if not oglas_url.endswith("/"):
                    oglas_url += "/"
                rec["Url"] = oglas_url
        # Alternativno: data-href na .property-details
        if not rec.get("Url"):
            det = box.find(class_="property-details")
            if det:
                oglas_url = det.get("data-href", "").strip()
                if oglas_url:
                    if not oglas_url.endswith("/"):
                        oglas_url += "/"
                    rec["Url"] = oglas_url

        if not rec.get("Url") or rec["Url"] in seen:
            continue
        seen.add(rec["Url"])

        # ── Cena ─────────────────────────────────────────────────────────────
        price_meta = box.find("meta", {"itemprop": "price"})
        if price_meta:
            rec["Cena"] = parse_cena(price_meta.get("content", ""))
        if not rec.get("Cena"):
            h6 = box.find("h6")
            if h6:
                rec["Cena"] = parse_cena(h6.get_text())

        # ── Naslov ───────────────────────────────────────────────────────────
        # <a class="url-title-d" title="KRAJ, OPIS">
        title_a = (box.find("a", class_="url-title-d")
                   or box.find("a", class_="url-title-m"))
        if title_a:
            naslov = title_a.get("title", "").strip()
            if not naslov:
                h2 = title_a.find("h2")
                if h2:
                    naslov = h2.get_text(strip=True)
            if naslov:
                rec["Naslov"] = naslov[:500]

        # ── Lokacija: prvi del naslova pred vejico ────────────────────────────
        if rec.get("Naslov"):
            loc = rec["Naslov"].split(",")[0].strip().title()
            if loc:
                rec["Lokacija"] = loc

        # ── Vrsta objekta ─────────────────────────────────────────────────────
        tipi = box.find("span", class_="tipi")
        if tipi:
            rec["VrstaObjekta"] = tipi.get_text(strip=True)

        # ── Kratek opis ───────────────────────────────────────────────────────
        desc_p = box.find("p", {"itemprop": "description"})
        if desc_p:
            rec["Opis"] = desc_p.get_text(" ", strip=True)[:3000]

        # ── Podrobnosti: velikost, leto, zemljišče, sobe, EK ──────────────────
        # <ul itemprop="disambiguatingDescription">
        #   <li><img src=".../velikost.svg">114,00 m²</li>
        #   <li><img src=".../leto.svg">1986</li>
        #   <li><img src=".../nadstropje.svg">P+1+M</li>
        #   <li><img src=".../zemljisce.svg">367 m²</li>
        # </ul>
        detail_ul = box.find("ul", {"itemprop": "disambiguatingDescription"})
        if detail_ul:
            for li in detail_ul.find_all("li"):
                img  = li.find("img")
                text = li.get_text(strip=True)
                if not img:
                    continue
                src = (img.get("src", "") + img.get("data-src", "")).lower()
                if "velikost" in src:
                    rec.setdefault("VelikostM2", parse_m2(text))
                elif "leto" in src:
                    rec.setdefault("LetoGradnje", parse_leto(text))
                elif "zemljisce" in src or "zemlji" in src:
                    rec.setdefault("ZemljisteM2", parse_m2(text))
                elif "soba" in src:
                    rec.setdefault("StSob", text)
                elif "energet" in src:
                    rec.setdefault("EnergetskiRazred", text[:10])

        # ── Cena na m² ────────────────────────────────────────────────────────
        if rec.get("Cena") and rec.get("VelikostM2") and rec["VelikostM2"] > 0:
            rec["CenaNaM2"] = round(rec["Cena"] / rec["VelikostM2"], 2)

        results.append(rec)

    return results, total_pgs


# ── Baza (MariaDB) ────────────────────────────────────────────────────────────
def init_db():
    import pymysql
    conn = pymysql.connect(**DB_CONFIG)
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS Nepremicnine (
                Id               INT AUTO_INCREMENT PRIMARY KEY,
                Naslov           VARCHAR(500)    NULL,
                Url              VARCHAR(1000)   NOT NULL UNIQUE,
                Cena             DECIMAL(12,2)   NULL,
                Lokacija         VARCHAR(200)    NULL,
                Obcina           VARCHAR(200)    NULL,
                VrstaObjekta     VARCHAR(150)    NULL,
                VelikostM2       DECIMAL(10,2)   NULL,
                ZemljisteM2      DECIMAL(12,2)   NULL,
                StSob            VARCHAR(50)     NULL,
                LetoGradnje      SMALLINT        NULL,
                EnergetskiRazred VARCHAR(10)     NULL,
                Opis             TEXT            NULL,
                DatumScrapa      DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP,
                CenaNaM2         DECIMAL(10,2)   NULL,
                INDEX idx_cena      (Cena),
                INDEX idx_lokacija  (Lokacija),
                INDEX idx_leto      (LetoGradnje)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """)
    conn.commit()
    return conn


def insert_oglas(conn, rec):
    import pymysql
    sql = """
        INSERT IGNORE INTO Nepremicnine
            (Naslov, Url, Cena, Lokacija, Obcina, VrstaObjekta, VelikostM2,
             ZemljisteM2, StSob, LetoGradnje, EnergetskiRazred, Opis, DatumScrapa, CenaNaM2)
        VALUES
            (%(Naslov)s, %(Url)s, %(Cena)s, %(Lokacija)s, %(Obcina)s, %(VrstaObjekta)s,
             %(VelikostM2)s, %(ZemljisteM2)s, %(StSob)s, %(LetoGradnje)s,
             %(EnergetskiRazred)s, %(Opis)s, %(DatumScrapa)s, %(CenaNaM2)s)
    """
    defaults = {k: None for k in ["Naslov","Url","Cena","Lokacija","Obcina","VrstaObjekta",
                                    "VelikostM2","ZemljisteM2","StSob","LetoGradnje",
                                    "EnergetskiRazred","Opis","DatumScrapa","CenaNaM2"]}
    defaults.update(rec)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, defaults)
        conn.commit()
        return cur.rowcount > 0
    except pymysql.err.IntegrityError:
        return False  # duplicate URL


# ── Izvoz CSV ─────────────────────────────────────────────────────────────────
CSV_COLS = ["Naslov","Regija","VrstaNP","Akcija","Lokacija","Obcina","Cena",
            "VelikostM2","ZemljisteM2","StSob","LetoGradnje","EnergetskiRazred",
            "CenaNaM2","VrstaObjekta","DatumScrapa","Url"]

def export_csv(records, path="nepremicnine_export.csv"):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CSV_COLS, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(records)
    print(f"✓ CSV shranjen: {path}  ({len(records)} vrstic)")


# ── Mapa zagonov – pomožne funkcije ──────────────────────────────────────────
def get_runs_dir() -> str:
    """Vrne absolutno pot do mape scraping_runs (poleg tega skripta)."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, SCRAPING_RUNS_DIR)


def load_known_urls(master_path: str) -> set:
    """Naloži vse URL-je iz akumulirane baze (baza.csv).
    Vrne set URL-jev, ki jih pri naslednjem zagonu preskočimo."""
    known: set[str] = set()
    if not os.path.isfile(master_path):
        return known
    try:
        with open(master_path, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f, delimiter=";"):
                url = row.get("Url", "").strip()
                if url:
                    known.add(url)
    except Exception as e:
        print(f"  ⚠ Napaka pri branju baze: {e}")
    return known


def append_to_master(records: list, master_path: str):
    """Doda nove zapise v akumulirano bazo (baza.csv).
    Če datoteka ne obstaja, jo ustvari z glavo; sicer le doda vrstice."""
    if not records:
        return
    file_exists = os.path.isfile(master_path)
    mode = "a" if file_exists else "w"
    with open(master_path, mode, newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CSV_COLS, delimiter=";", extrasaction="ignore")
        if not file_exists:
            w.writeheader()
        w.writerows(records)
    verb = "Dodano v" if file_exists else "Ustvarjena"
    print(f"  ✓ {verb} bazno datoteko: {master_path}  (+{len(records)} vrstic)")


# ── Glavni tok ────────────────────────────────────────────────────────────────
def main():
    global _headless

    parser = argparse.ArgumentParser(description="Nepremičnine.net scraper")
    parser.add_argument("--strani",   type=int, default=0,
                        help="Maks. strani na kombinacijo (0 = vse, privzeto 0)")
    parser.add_argument("--test",     action="store_true",
                        help="Samo 1 stran (test)")
    parser.add_argument("--csv",      action="store_true",
                        help="Samo CSV, brez DB")
    parser.add_argument("--delay",    type=float, default=1.5,
                        help="Zakasnitev med zahtevki v sekundah (privzeto 1.5)")
    parser.add_argument("--headless", action="store_true",
                        help="Brez vidnega okna brskalnika")
    parser.add_argument("--visible",  action="store_true",
                        help="(zastarelo – privzeto je že vidno)")

    # Regija – ena vrednost ali "vse"
    parser.add_argument("--regija",  default=None,
                        help="Regija (slug ali ime, npr. gorenjska). Privzeto: gorenjska")
    # Vrsta – ena vrednost ali "vse"
    parser.add_argument("--vrsta",   default=None,
                        help="Vrsta nepremičnine (slug ali ime, npr. hisa). Privzeto: hisa")
    # Večkratna izbira (z vejicami)
    parser.add_argument("--regije",  default=None,
                        help="Več regij, ločenih z vejico, ali 'vse'")
    parser.add_argument("--vrste",   default=None,
                        help="Več vrst, ločenih z vejico, ali 'vse'")
    # Akcija
    parser.add_argument("--akcija",  default="prodaja",
                        choices=AKCIJE, help="prodaja ali najem (privzeto: prodaja)")
    # CSV ime
    parser.add_argument("--izhod",   default=None,
                        help="Ime izhodne CSV datoteke (privzeto samodejno)")

    a = parser.parse_args()
    _headless = a.headless
    max_pages = 1 if a.test else a.strani   # 0 = brez omejitve (vse strani)
    delay     = a.delay
    use_db    = not a.csv
    akcija    = a.akcija

    # ── Razreši seznam regij ──────────────────────────────────────────────────
    if a.regije:
        raw_r = a.regije.strip()
        if raw_r.lower() == "vse":
            reg_slugs = list(REGIJE.values())
        else:
            reg_slugs = [regija_slug(x.strip()) for x in raw_r.split(",") if x.strip()]
    elif a.regija:
        raw_r = a.regija.strip()
        if raw_r.lower() == "vse":
            reg_slugs = list(REGIJE.values())
        else:
            reg_slugs = [regija_slug(raw_r)]
    else:
        reg_slugs = ["gorenjska"]

    # ── Razreši seznam vrst ───────────────────────────────────────────────────
    if a.vrste:
        raw_v = a.vrste.strip()
        if raw_v.lower() == "vse":
            vrs_slugs = list(VRSTE.values())
        else:
            vrs_slugs = [vrsta_slug(x.strip()) for x in raw_v.split(",") if x.strip()]
    elif a.vrsta:
        raw_v = a.vrsta.strip()
        if raw_v.lower() == "vse":
            vrs_slugs = list(VRSTE.values())
        else:
            vrs_slugs = [vrsta_slug(raw_v)]
    else:
        vrs_slugs = ["hisa"]

    # ── Generiraj vse kombinacije ─────────────────────────────────────────────
    combinations = [(r, v) for r in reg_slugs for v in vrs_slugs]

    # ── Izhodna CSV datoteka ──────────────────────────────────────────────────
    if a.izhod:
        csv_path = a.izhod
    elif len(combinations) == 1:
        r0, v0 = combinations[0]
        csv_path = f"nepremicnine_{r0}_{v0}.csv"
    else:
        csv_path = f"nepremicnine_export_{akcija}.csv"

    print("=" * 64)
    print("  Nepremičnine.net Scraper")
    print(f"  {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    print("=" * 64)
    print(f"  Akcija:  {akcija}")
    print(f"  Regije:  {', '.join(_REGIJE_SLUGS.get(r, r) for r in reg_slugs)}")
    print(f"  Vrste:   {', '.join(_VRSTE_SLUGS.get(v, v) for v in vrs_slugs)}")
    print(f"  Kombinacij: {len(combinations)}  |  Maks strani: {'vse' if max_pages == 0 else max_pages}  |  Delay: {delay}s")
    print(f"  Izhod:   {csv_path}")
    print()

    # ── Nastavi mapo za zagone in preveri obstoječe zapise ────────────────────
    runs_dir     = get_runs_dir()
    os.makedirs(runs_dir, exist_ok=True)
    master_csv   = os.path.join(runs_dir, "baza.csv")
    run_ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_csv_path = os.path.join(runs_dir, f"scraping_{run_ts}.csv")

    print("  Preverjam obstoječe zapise v bazi…")
    known_urls = load_known_urls(master_csv)
    if known_urls:
        print(f"  ✓ Baza vsebuje {len(known_urls)} obstoječih oglasov – duplikati bodo preskočeni.")
    else:
        print("  ℹ  Baza je prazna – vsi najdeni oglasi bodo novi.")
    print()

    # Init DB
    conn = None
    if use_db:
        try:
            import pymysql
            conn = init_db()
            print("✓ Baza inicializirana (MariaDB nepremicnine)")
        except Exception as e:
            print(f"⚠ Baza ni dostopna: {e}")
            print("  Nadaljujem samo s CSV izvozom.")
            use_db = False

    # ── Scraping vseh kombinacij ──────────────────────────────────────────────
    all_records = []
    seen_urls   = set(known_urls)   # vključuje URL-je iz prejšnjih zagonov
    shranjenih  = 0
    napak       = 0
    preskocenih = 0   # že v bazi iz prejšnjih zagonov

    for combo_idx, (reg, vrs) in enumerate(combinations, 1):
        path = list_path(akcija, reg, vrs)
        reg_ime = _REGIJE_SLUGS.get(reg, reg)
        vrs_ime = _VRSTE_SLUGS.get(vrs, vrs)
        print(f"\n[{combo_idx}/{len(combinations)}] {reg_ime} / {vrs_ime}  ({akcija})")
        print(f"  URL: {BASE_URL}{path}")

        # ── Stran 1: preberi zapise IN skupno število strani ─────────────────
        print(f"  Stran   1: ", end="", flush=True)
        try:
            page_recs, total_pages = parse_listing_page(1, delay, path)
        except Exception as e:
            print(f"napaka: {e}")
            napak += 1
            continue

        if not page_recs:
            print("ni oglasov → preskočim")
            continue

        # Spoštuj omejitev max_pages (0 = brez omejitve)
        limit = total_pages if max_pages == 0 else min(total_pages, max_pages)
        print(f"{len(page_recs)} oglas(ov)  [skupaj strani: {total_pages}"
              f"{'' if max_pages == 0 else f', omejitev: {limit}'}]")

        novi = 0
        for rec in page_recs:
            rec["Regija"]  = reg_ime
            rec["VrstaNP"] = vrs_ime
            rec["Akcija"]  = akcija
            url_key = rec.get("Url", "")
            if url_key and url_key not in seen_urls:
                seen_urls.add(url_key)
                all_records.append(rec)
                novi += 1
                if conn:
                    try:
                        if insert_oglas(conn, rec):
                            shranjenih += 1
                    except Exception:
                        pass
            elif url_key and url_key in known_urls:
                preskocenih += 1   # oglas je že v bazi iz prejšnjega zagona
        print(f"           → {novi} novih  (skupaj: {len(all_records)})")
        time.sleep(delay + random.uniform(0, 0.8))

        # ── Strani 2..limit ───────────────────────────────────────────────────
        for page in range(2, limit + 1):
            print(f"  Stran {page:3d}: ", end="", flush=True)
            try:
                page_recs, _ = parse_listing_page(page, delay, path)
            except Exception as e:
                print(f"napaka: {e}")
                napak += 1
                continue

            if not page_recs:
                print("ni oglasov → konec")
                break

            novi = 0
            for rec in page_recs:
                rec["Regija"]  = reg_ime
                rec["VrstaNP"] = vrs_ime
                rec["Akcija"]  = akcija
                url_key = rec.get("Url", "")
                if url_key and url_key not in seen_urls:
                    seen_urls.add(url_key)
                    all_records.append(rec)
                    novi += 1
                    if conn:
                        try:
                            if insert_oglas(conn, rec):
                                shranjenih += 1
                        except Exception:
                            pass
                elif url_key and url_key in known_urls:
                    preskocenih += 1   # oglas je že v bazi iz prejšnjega zagona

            print(f"{len(page_recs)} oglas(ov), {novi} novih  (skupaj: {len(all_records)})")
            time.sleep(delay + random.uniform(0, 0.8))

    # ── Povzetek ──────────────────────────────────────────────────────────────
    print(f"\n{'=' * 64}")
    print(f"  Skupaj novih:          {len(all_records)}")
    print(f"  Preskočenih (v bazi):  {preskocenih}")
    if conn:
        print(f"  Shranjenih v DB:       {shranjenih}")
    if napak:
        print(f"  Napak:                 {napak}")

    if all_records:
        # 1. Shrani zapis zagona – samo novi oglasi tega zagona
        export_csv(all_records, run_csv_path)
        print(f"  📁 Zapis zagona:       {run_csv_path}")

        # 2. Dodaj nove oglase v akumulirano bazo (za trening modelov)
        append_to_master(all_records, master_csv)

        # 3. Shrani standardni izhod (novi oglasi tega zagona, za nazaj-kompatibilnost)
        export_csv(all_records, csv_path)

        # Kratek povzetek po kombinacijah
        from collections import Counter
        cnts = Counter((r.get("Regija","?"), r.get("VrstaNP","?")) for r in all_records)
        for (reg_i, vrs_i), cnt in sorted(cnts.items()):
            cene = [r["Cena"] for r in all_records
                    if r.get("Regija")==reg_i and r.get("VrstaNP")==vrs_i and r.get("Cena")]
            if cene:
                import statistics as _st
                print(f"  {reg_i:<20} / {vrs_i:<20}  N={cnt}  "
                      f"med={_st.median(cene):,.0f} €")
            else:
                print(f"  {reg_i:<20} / {vrs_i:<20}  N={cnt}")
    else:
        print("  Ni novih nepremičnin za shranjevanje.")
        if preskocenih:
            print(f"  (Vsi najdeni oglasi so že bili v bazi.)")

    print()
    print(f"  📂 Mapa zagonov:       {runs_dir}")
    print(f"  📊 Bazna datoteka:     {master_csv}")

    if conn:
        conn.close()

    if _browser is not None:
        try:
            _browser.quit()
        except Exception:
            pass

    print()
    print("  Za analizo zaženite:  python analyze.py")


if __name__ == "__main__":
    main()

