import pandas as pd
import requests
import re
import time
import urllib3
import io
import statistics
import pdfplumber
import pytesseract

from PIL import Image
from bs4 import BeautifulSoup
from ddgs import DDGS
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright

# =========================================
# CONFIG
# =========================================

INPUT_EXCEL = "tarifs_56_complet.xlsx"
OUTPUT_EXCEL = "tarifs_enrichi_ultimate.xlsx"

MAX_RESULTS = 5
DELAY = 20
TIMEOUT = 30000
SAVE_EVERY = 5

urllib3.disable_warnings()

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 "
        "(Windows NT 10.0; Win64; x64)"
    )
}

# =========================================
# TESSERACT
# =========================================

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\sylva\Desktop\Tesseract-OCR\tesseract.exe"
)

# =========================================
# DOMAINES A IGNORER
# =========================================

BAD_DOMAINS = [

    "facebook",
    "instagram",
    "youtube",
    "tiktok",
    "linkedin",

    "wikipedia",
    "larousse",

    "yandex",
    "etsy",
    "dealabs",
    "gralon",
    "restaurantguru",
    "linternaute",
    "societe.com",
    "mapado",
    "pagesjaunes",

    "amazon",
    "ebay",
]

# =========================================
# CATEGORIES
# =========================================

CATEGORY_RULES = {

    "restaurant": [
        "restaurant",
        "bar",
        "café",
        "cafe",
        "brasserie",
        "crêperie",
        "pizza",
    ],

    "cinema": [
        "cinéma",
        "cinema",
    ],

    "sport": [
        "golf",
        "squash",
        "tennis",
        "padel",
        "karting",
        "bowling",
        "base nautique",
    ],

    "hotel": [
        "hôtel",
        "hotel",
        "camping",
        "gîte",
        "gite",
        "village vacances",

        "moulin",
        "clos",
        "domaine",
        "villa",

        "maison d'hôtes",
        "chambre d'hôtes",
    ],

    "nature": [
        "plage",
        "anse",
        "lac",
        "rivière",
        "port",
    ],

    "patrimoine": [
        "église",
        "eglise",
        "chapelle",
        "abbaye",
    ]
}

# =========================================
# REGEX PRIX
# =========================================

PRICE_REGEX = re.compile(
    r'(\d{1,4}(?:[.,]\d{1,2})?)\s?€'
)

# =========================================
# DETECTION COLONNES
# =========================================

def find_column(df, keywords):

    for col in df.columns:

        c = str(col).lower()

        for k in keywords:

            if k in c:
                return col

    return None

# =========================================
# DETECTION CATEGORIE
# =========================================

def detect_category(name):

    n = name.lower()

    for category, words in CATEGORY_RULES.items():

        for w in words:

            if w in n:
                return category

    return "default"

# =========================================
# EXTRACTION PRIX
# =========================================

def extract_prices(text):

    prices = []

    matches = PRICE_REGEX.findall(text)

    for match in matches:

        try:

            value = float(
                match.replace(",", ".")
            )

            if 1 <= value <= 1000:
                prices.append(value)

        except:
            pass

    return prices

# =========================================
# FILTRE PRIX PAR CATEGORIE
# =========================================

def clean_prices(prices, category):

    filtered = []

    for p in prices:

        if category == "restaurant":

            if 8 <= p <= 80:
                filtered.append(p)

        elif category == "cinema":

            if 5 <= p <= 25:
                filtered.append(p)

        elif category == "sport":

            if 5 <= p <= 150:
                filtered.append(p)

        elif category == "hotel":

            if 30 <= p <= 700:
                filtered.append(p)

        elif category == "patrimoine":

            if 2 <= p <= 50:
                filtered.append(p)

        else:

            if 3 <= p <= 500:
                filtered.append(p)

    return filtered

# =========================================
# FILTRE URL PERTINENTE
# =========================================

def is_relevant_url(url, name):

    lower = url.lower()

    BAD_WORDS = [

        "recipe",
        "tarifleri",
        "yemek",
        "borek",

        "facebook",
        "instagram",
        "twitter",
        "linkedin",

        "youtube",
        "tiktok",

        "amazon",
        "ebay",

        "random",
    ]

    if any(
        bad in lower
        for bad in BAD_WORDS
    ):
        return False

    important_words = [

        w.lower()
        for w in name.split()
        if len(w) > 3
    ]

    match_count = 0

    for w in important_words:

        if w in lower:
            match_count += 1

    return match_count >= 1

# =========================================
# PLAYWRIGHT SCRAP
# =========================================

def scrape_page(url):

    try:

        with sync_playwright() as p:

            browser = p.chromium.launch(
                headless=True
            )

            page = browser.new_page()

            page.goto(
                url,
                timeout=TIMEOUT
            )

            page.wait_for_timeout(3000)

            page.mouse.wheel(0, 3000)

            page.wait_for_timeout(2000)

            html = page.content()

            browser.close()

            return html

    except Exception as e:

        print(f"Erreur Playwright {url}: {e}")

        return ""

# =========================================
# EXTRACTION PDF
# =========================================

def extract_pdf_text(pdf_url):

    try:

        response = requests.get(
            pdf_url,
            headers=HEADERS,
            timeout=20,
            verify=False
        )

        with open("temp.pdf", "wb") as f:
            f.write(response.content)

        text = ""

        with pdfplumber.open("temp.pdf") as pdf:

            for page in pdf.pages:

                page_text = page.extract_text()

                if page_text:
                    text += page_text + "\n"

        return text

    except Exception as e:

        print(f"Erreur PDF {pdf_url}: {e}")

        return ""

# =========================================
# OCR IMAGE
# =========================================

def extract_image_text(image_url):

    try:

        response = requests.get(
            image_url,
            headers=HEADERS,
            timeout=20,
            verify=False
        )

        image = Image.open(
            io.BytesIO(response.content)
        )

        text = pytesseract.image_to_string(
            image,
            lang="eng"
        )

        return text

    except Exception as e:

        print(f"Erreur OCR {image_url}: {e}")

        return ""

# =========================================
# DETECTER PDFS
# =========================================

def find_pdfs(soup, base_url):

    pdfs = []

    for link in soup.find_all("a", href=True):

        href = link["href"]

        if ".pdf" in href.lower():

            full = urljoin(base_url, href)

            pdfs.append(full)

    return pdfs

# =========================================
# DETECTER IMAGES UTILES
# =========================================

def find_images(soup, base_url):

    images = []

    for img in soup.find_all("img"):

        src = img.get("src")

        if not src:
            continue

        full = urljoin(base_url, src)

        lower = full.lower()

        if not any(
            ext in lower
            for ext in [
                ".jpg",
                ".jpeg",
                ".png",
                ".webp"
            ]
        ):
            continue

        BAD_IMAGE_WORDS = [

            "icon",
            "logo",
            "avatar",
            "banner",
            "cart",
            "sprite",

            "facebook",
            "instagram",
            "twitter",
            "linkedin",
            "youtube",

            "unknown",
            "placeholder",
            "default",
            "favicon",

            "social",
            "thumb",
            "mini",
            "small",
            "loading",
            "empty",

            "inactive",
            "nonactive",

            "tile",
            "map",
            "osm",
            "openstreetmap",

            "flag",
            "flags",

            "cloudly",
            "tourism-system",

            "blank",
            "ratio",

            "marker",

            "poster",
            "movie",
        ]

        if any(
            bad in lower
            for bad in BAD_IMAGE_WORDS
        ):
            continue

        images.append(full)

    images = list(set(images))

    return images[:5]

# =========================================
# RECHERCHE WEB
# =========================================

def search_tarif_page(name, city=None, category="default"):

    geo = ""

    if city and city != "nan":
        geo = f" {city}"

    queries = [
        f"{name}{geo} France tarif"
    ]

    if category == "hotel":

        queries.append(
            f"{name}{geo} booking"
        )

    urls = []

    try:

        with DDGS() as ddgs:

            for query in queries:

                time.sleep(10)

                try:

                    results = ddgs.text(
                        query,
                        max_results=MAX_RESULTS,
                        backend="lite"
                    )

                    for r in results:

                        url = r.get("href")

                        if not url:
                            continue

                        if any(
                            bad in url.lower()
                            for bad in BAD_DOMAINS
                        ):
                            continue

                        if not is_relevant_url(url, name):
                            continue

                        if url not in urls:
                            urls.append(url)

                except Exception as e:

                    print(f"Erreur recherche: {e}")

                    time.sleep(15)

    except Exception as e:

        print(f"Erreur globale: {e}")

    return urls[:10]

# =========================================
# NIVEAU PRIX
# =========================================

def get_price_level(avg):

    if avg == 0:
        return "gratuit"

    if avg < 10:
        return "€"

    elif avg < 30:
        return "€€"

    else:
        return "€€€"

# =========================================
# CONFIANCE
# =========================================

def get_confidence(url, prices):

    if not url:
        return 0

    score = 0.3

    u = url.lower()

    if "tarif" in u:
        score += 0.3

    if ".fr" in u:
        score += 0.2

    if len(prices) >= 2:
        score += 0.2

    return round(
        min(score, 1.0),
        2
    )

# =========================================
# LECTURE EXCEL
# =========================================

print("Lecture Excel...")

df = pd.read_excel(INPUT_EXCEL)

# =========================================
# DETECTION COLONNES
# =========================================

name_col = find_column(
    df,
    ["nom", "name", "titre"]
)

city_col = find_column(
    df,
    ["commune", "ville", "city"]
)

if name_col is None:
    name_col = df.columns[0]

# =========================================
# COLONNES IA
# =========================================

new_cols = [

    "ia_categorie",
    "ia_url_tarif",
    "ia_prix_min",
    "ia_prix_max",
    "ia_prix_moyen",
    "ia_niveau_prix",
    "ia_confiance",

]

for c in new_cols:

    if c not in df.columns:
        df[c] = None

# =========================================
# BOUCLE PRINCIPALE
# =========================================

for idx, row in df.iterrows():

    try:

        name = str(row[name_col])

        if not name or name == "nan":
            continue

        city = None

        if city_col:
            city = str(row[city_col])

        print(f"\n[{idx+1}/{len(df)}] {name}")

        category = detect_category(name)

        print(f" Catégorie : {category}")

        df.at[idx, "ia_categorie"] = category

        # nature gratuit
        if category == "nature":

            print(" Nature probablement gratuite")

            df.at[idx, "ia_prix_min"] = 0
            df.at[idx, "ia_prix_max"] = 0
            df.at[idx, "ia_prix_moyen"] = 0
            df.at[idx, "ia_niveau_prix"] = "gratuit"
            df.at[idx, "ia_confiance"] = 0.95

            continue

        urls = search_tarif_page(
            name,
            city,
            category
        )

        best_prices = []
        best_url = None

        for url in urls:

            print(f" -> {url}")

            html = scrape_page(url)

            if not html:
                continue

            soup = BeautifulSoup(
                html,
                "lxml"
            )

            text = soup.get_text(
                separator="\n"
            )

            # HTML
            prices = extract_prices(text)
            best_prices.extend(prices)

            # PDF
            pdfs = find_pdfs(
                soup,
                url
            )

            for pdf in pdfs[:3]:

                print(f" PDF -> {pdf}")

                pdf_text = extract_pdf_text(pdf)

                pdf_prices = extract_prices(
                    pdf_text
                )

                best_prices.extend(pdf_prices)

            # OCR
            images = find_images(
                soup,
                url
            )

            for img in images[:3]:

                print(f" OCR -> {img}")

                image_text = extract_image_text(
                    img
                )

                image_prices = extract_prices(
                    image_text
                )

                best_prices.extend(image_prices)

            if best_prices and best_url is None:
                best_url = url

        # CLEAN
        best_prices = clean_prices(
            best_prices,
            category
        )

        # RESULTATS
        if best_prices:

            pmin = min(best_prices)
            pmax = max(best_prices)

            # RESTAURANTS = moitié basse
            if category == "restaurant":

                sorted_prices = sorted(best_prices)

                keep = sorted_prices[
                    :max(3, len(sorted_prices)//2)
                ]

                pavg = round(
                    statistics.median(keep),
                    2
                )

            else:

                pavg = round(
                    statistics.median(best_prices),
                    2
                )

            confidence = get_confidence(
                best_url,
                best_prices
            )

            df.at[idx, "ia_url_tarif"] = best_url
            df.at[idx, "ia_prix_min"] = pmin
            df.at[idx, "ia_prix_max"] = pmax
            df.at[idx, "ia_prix_moyen"] = pavg
            df.at[idx, "ia_niveau_prix"] = get_price_level(pavg)
            df.at[idx, "ia_confiance"] = confidence

            print(
                f" OK -> {pavg}€ | confiance {confidence}"
            )

        else:

            print(" Aucun prix trouvé")

        # SAVE AUTO
        if idx % SAVE_EVERY == 0:

            df.to_excel(
                OUTPUT_EXCEL,
                index=False
            )

            print(" Sauvegarde automatique...")

        time.sleep(DELAY)

    except Exception as e:

        print(f"Erreur ligne {idx}: {e}")

print("\nSauvegarde finale...")

df.to_excel(
    OUTPUT_EXCEL,
    index=False
)

print(
    f"Fichier créé : {OUTPUT_EXCEL}"
)