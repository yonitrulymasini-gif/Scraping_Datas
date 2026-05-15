import pandas as pd
import requests
import re
import time
import urllib3

from bs4 import BeautifulSoup
from ddgs import DDGS
from urllib.parse import urlparse

# =========================================
# SSL
# =========================================

urllib3.disable_warnings()

# =========================================
# CONFIG
# =========================================

INPUT_EXCEL = "tarifs_56_complet.xlsx"
OUTPUT_EXCEL = "resultats/tarifs_enrichi_final.xlsx"

MAX_RESULTS = 5
DELAY = 5
TIMEOUT = 15
SAVE_EVERY = 10

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 "
        "(Windows NT 10.0; Win64; x64)"
    )
}

# =========================================
# DOMAINES A IGNORER
# =========================================

BAD_DOMAINS = [

    # social / bruit
    "facebook",
    "instagram",
    "youtube",
    "tiktok",
    "linkedin",

    # encyclopédies
    "wikipedia",
    "larousse",

    # annuaires / faux positifs
    "yandex",
    "etsy",
    "dealabs",
    "gralon",
    "restaurantguru",
    "linternaute",
    "societe.com",
    "mapado",
    "pagesjaunes",

    # marketplaces
    "amazon",
    "ebay",

    # bruit fréquent
    "leboncoin",
]

# =========================================
# BONUS DOMAINES
# =========================================

GOOD_DOMAINS = [
    ".fr",
    "tourisme",
    "office",
    "reservation",
    "billetterie",
    "booking",
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
        "creperie",
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
    ],

    "nature": [
        "plage",
        "anse",
        "lac",
        "rivière",
        "riviere",
        "sentier",
        "port",
    ],

    "patrimoine": [
        "église",
        "eglise",
        "chapelle",
        "abbaye",
        "château",
        "chateau",
        "moulin",
        "fort",
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
# SCRAP PAGE
# =========================================

def scrape_page(url):

    try:

        r = requests.get(
            url,
            headers=HEADERS,
            timeout=TIMEOUT,
            verify=False
        )

        if r.status_code != 200:
            return ""

        soup = BeautifulSoup(
            r.text,
            "lxml"
        )

        for tag in soup([
            "script",
            "style",
            "noscript"
        ]):
            tag.extract()

        text = soup.get_text(
            separator="\n"
        )

        return text

    except Exception as e:

        print(f"Erreur scraping {url}: {e}")

        return ""

# =========================================
# SCORE URL
# =========================================

def score_url(url):

    score = 0

    u = url.lower()

    # bonus domaines utiles
    for good in GOOD_DOMAINS:

        if good in u:
            score += 10

    # bonus tarif
    if "tarif" in u:
        score += 15

    # bonus prix
    if "prix" in u:
        score += 5

    return score

# =========================================
# RECHERCHE WEB
# =========================================

def search_tarif_page(name, city=None):

    geo = ""

    if city and city != "nan":
        geo = f" {city}"

    query = f"{name}{geo} tarif"

    urls = []

    try:

        with DDGS() as ddgs:

            results = ddgs.text(
                query,
                max_results=MAX_RESULTS,
                backend="lite"
            )

            for r in results:

                url = r.get("href")

                if not url:
                    continue

                # blacklist
                if any(
                    bad in url.lower()
                    for bad in BAD_DOMAINS
                ):
                    continue

                if url not in urls:
                    urls.append(url)

    except Exception as e:

        print(f"Erreur recherche: {e}")

        time.sleep(10)

    urls = sorted(
        urls,
        key=lambda x: score_url(x),
        reverse=True
    )

    return urls

# =========================================
# NIVEAU PRIX
# =========================================

def get_price_level(avg):

    if avg is None:
        return ""

    if avg == 0:
        return "gratuit"

    if avg < 10:
        return "€"

    elif avg < 30:
        return "€€"

    else:
        return "€€€"

# =========================================
# SCORE CONFIANCE IA
# =========================================

def get_confidence(url, prices):

    if not url:
        return 0.0

    score = 0.3

    u = url.lower()

    if "tarif" in u:
        score += 0.3

    if "tourisme" in u:
        score += 0.2

    if len(prices) >= 2:
        score += 0.1

    if ".fr" in u:
        score += 0.1

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
# CREATION COLONNES IA
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

        # lieux nature = gratuit
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
            city
        )

        best_prices = []
        best_url = None

        for url in urls:

            print(f" -> {url}")

            text = scrape_page(url)

            if not text:
                continue

            prices = extract_prices(text)

            # pages trop bruitées
            if len(prices) > 100:
                continue

            if prices:

                best_prices.extend(prices)

                if best_url is None:
                    best_url = url

        if best_prices:

            pmin = min(best_prices)
            pmax = max(best_prices)

            pavg = round(
                sum(best_prices) / len(best_prices),
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

        # sauvegarde auto
        if idx % SAVE_EVERY == 0:

            df.to_excel(
                OUTPUT_EXCEL,
                index=False
            )

            print(" Sauvegarde automatique...")

        time.sleep(DELAY)

    except Exception as e:

        print(f"Erreur ligne {idx}: {e}")

# =========================================
# EXPORT FINAL
# =========================================

print("\nSauvegarde finale...")

df.to_excel(
    OUTPUT_EXCEL,
    index=False
)

print(
    f"Fichier créé : {OUTPUT_EXCEL}"
)