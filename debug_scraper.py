"""Debug skripta - preveri HTML posameznega oglasa"""
from DrissionPage import ChromiumPage, ChromiumOptions
import time, re

co = ChromiumOptions()
co.set_argument("--disable-blink-features=AutomationControlled")
co.set_argument("--window-size=1400,900")
page = ChromiumPage(addr_or_opts=co)

# Pojdi najprej na listing da dobiš cookies
listing_url = "https://www.nepremicnine.net/oglasi-prodaja/gorenjska/hisa/"
print(f"Listing: {listing_url}")
page.get(listing_url)
time.sleep(4)

# Sedaj pojdi na individualni oglas
oglas_url = "https://www.nepremicnine.net/oglasi-prodaja/ambroz-pod-krvavcem-mirna-lokacija-z-razgledom-hisa_7296865/"
print(f"Oglas:   {oglas_url}")
page.get(oglas_url)
time.sleep(4)

html = page.html or ""
print(f"NASLOV:  {page.title}")
print(f"HTML:    {len(html)} znakov")

# Shrani HTML oglas strani
with open("debug_oglas.html", "w", encoding="utf-8") as f:
    f.write(html)
print("Shranjen v debug_oglas.html")

# Poisci price elemente
import re
# meta price
metas = re.findall(r'<meta[^>]*itemprop=["\']price["\'][^>]*/>', html)
print(f"\nmeta price tags: {metas[:3]}")

# h1, h2, h6 with euro
eur_tags = re.findall(r'<(h[1-6])[^>]*>([^<]*€[^<]*)<', html)
print(f"\nEuro headings: {eur_tags[:5]}")

# lokacija
loc = re.findall(r'itemprop=["\']addressLocality["\'][^>]*>([^<]+)<', html)
print(f"\naddressLocality: {loc[:3]}")

# dt/dd atributi
dts = re.findall(r'<dt>([^<]+)</dt>\s*<dd>([^<]+)</dd>', html)
print(f"\ndt/dd pairs: {dts[:10]}")

# table rows
trs = re.findall(r'<t[dh]>([^<]+)</t[dh]>', html)
print(f"\ntable cells: {trs[:20]}")

page.quit()
print("\nKONEC")


url = "https://www.nepremicnine.net/oglasi-prodaja/gorenjska/hisa/"
print(f"Nalagam: {url}")
page.get(url)

print("Cakam 5 sekund da se stran nalozi...")
time.sleep(5)

title = page.title or ""
cur_url = page.url or ""
html = page.html or ""

print(f"\nNASLOV (title): {title}")
print(f"URL:            {cur_url}")
print(f"HTML dolzina:   {len(html)} znakov")

# Shrani cel HTML v datoteko ZA PREGLED (UTF-8, brez tiskanja)
with open("debug_html.html", "w", encoding="utf-8") as f:
    f.write(html)
print("\nCel HTML shranjen v: debug_html.html")

# Poisci razlicne vzorce linkov
patterns = {
    "/oglas/":          r'href=["\'](/oglas/[^"\'#]+)',
    "nepremicnine.net": r'href=["\'](https://www\.nepremicnine\.net/[^"\'#]+)',
    "vse href":         r'href=["\']([^"\'#]+)["\']',
}
for name, pat in patterns.items():
    found = re.findall(pat, html)
    print(f"\n{name}: {len(found)} najdenih")
    for l in found[:8]:
        print(f"  {l}")

page.quit()
print("\nKONEC")

