# ============================================================
# TRAVELINE — Script unifié d'enrichissement v2
#
# pip install pandas openpyxl requests beautifulsoup4 lxml ddgs
#
# Étape 1 : Datatourisme  → parse tous les JSON du dossier
# Étape 2 : Google Places → note, horaires, tel, site, prix
# Étape 3 : Scraping web  → tarifs, email, réseaux sociaux
# ============================================================

import os, gzip, json, time, re, statistics, urllib3, requests, pandas as pd
from bs4 import BeautifulSoup
from ddgs import DDGS

urllib3.disable_warnings()

# ============================================================
# CONFIG
# ============================================================

DATA_FOLDER    = "" # CHEMIN DATATOURISME/FLUX
OUTPUT_EXCEL   = "traveline_departement_56.xlsx"
GOOGLE_API_KEY = "" # CLE API GOOGLE PLACES(NEW)

TEST_MODE  = True    # FALSE POUR TOUT TRAITER
TEST_LIMIT = 10

MAX_RESULTS = 8
TIMEOUT     = 15
DELAY       = 1
SAVE_EVERY  = 25

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# ============================================================
# COLONNES DU MODÈLE
# ============================================================

COLONNES = [
    "REFERENCE", "TITRE", "DESCRIPTION", "TYPE", "CATEGORIE", "CLASSIFICATION",
    "NOMBRE_MAX", "TAGS", "ADRESSE",
    "CONTACT_TEL", "CONTACT_EMAIL", "CONTACT_SITE",
    "CONTACT_INSTAGRAM", "CONTACT_FACEBOOK", "PHOTOS",
    "HORAIRES_LUNDI", "HORAIRES_MARDI", "HORAIRES_MERCREDI",
    "HORAIRES_JEUDI", "HORAIRES_VENDREDI", "HORAIRES_SAMEDI", "HORAIRES_DIMANCHE",
    "PUBLIC_ADULTE", "PUBLIC_ENFANT", "PUBLIC_PMR", "PUBLIC_ANIMAL",
    "EQUIPEMENT",
    "PRIX_MIN", "PRIX_MAX", "PRIX_MOYEN", "NIVEAUX_PRIX", "PRIX_CONFIANCE",
    "NOTE", "NOMBRE_AVIS",
]

JOURS_FR = ["LUNDI","MARDI","MERCREDI","JEUDI","VENDREDI","SAMEDI","DIMANCHE"]
JOURS_EN = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

# ============================================================
# REGEX
# ============================================================

PRICE_REGEX = re.compile(r'(\d{1,4}(?:[.,]\d{1,2})?)\s?€')
EMAIL_REGEX = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+')
PHONE_REGEX = re.compile(r'(?:\+33|0)[1-9](?:[\s\.-]?\d{2}){4}')
URL_REGEX   = re.compile(r'https?://[^\s"\'>]+')

# ============================================================
# MAPPING TYPE — FIX: "Hébergement" sans s
# ============================================================

TYPE_KEYWORDS = {
    "Hébergement": [
        "hotel","accommodation","camping","gite","gîte",
        "lodging","hostel","rental","residence","meublé",
        "bedandbreakfast","cottage","furnished",
    ],
    "Restaurants": [
        "restaurant","foodestablishment","cafeteria","catering",
    ],
}

# ============================================================
# MAPPING CATÉGORIE
# ============================================================

CATEGORY_RULES = [
    # Restaurants (ordre prioritaire)
    ("Restaurant étoilé",   ["michelin","étoilé","etoile"]),
    ("Brasserie",           ["brasserie"]),
    ("Burger",              ["burger","hamburger"]),
    ("Pizzeria",            ["pizza","pizzeria"]),
    ("Sushi",               ["sushi","maki"]),
    ("Asiatique",           ["asiatique","thai","chinois","japonais","vietnamien","coréen"]),
    ("Indien",              ["indien","india"]),
    ("Crêperie",            ["crêperie","creperie","galette","crêpe"]),
    ("Restaurant",          ["restaurant","bistrot","grill","snack","brasserie"]),
    # Hébergements
    ("Camping",             ["camping"]),
    ("Chambre d'hôte",      ["chambre d'hôte","chambre hote","bed and breakfast","b&b","bedandbreakfast"]),
    ("Auberge de jeunesse", ["auberge de jeunesse","hostel"]),
    ("Insolite",            ["cabane","yourte","bulle","tipi","treehouse","insolite"]),
    ("Résidence",           ["résidence","residence"]),
    ("Appartement",         ["appartement","studio","apartment"]),
    ("Gîte",               ["gîte","gite"]),
    ("Maison",              ["maison","villa","cottage"]),
    ("Hôtel",              ["hotel","hôtel"]),
    # Activités — ordre du plus spécifique au plus général
    ("Plage & Farniente",   ["plage","beach","anse","baignade","club de plage","bord de mer"]),
    ("Gastronomie",         ["dégustation","degustation","cidre","vin","cave","terroir","fromagerie","vignoble","biscuiterie","distillerie","fromage","chèvre","bière","produits du terroir"]),
    ("Thalasso & Bien-être",["thalasso","thalassothérapie","centre de bien-être","soin du corps","hammam","thermes","relaxation","massage bien-être","yoga","méditation","bain de vapeur"]),
    ("Monuments",           ["chapelle","église","eglise","abbaye","château","chateau","monument","basilique","fort","menhir","dolmen","mégalithe","alignement","calvaire","croix de mission","moulin à vent","moulin","tour médiévale","remparts","clocher","prieuré","manoir"]),
    ("Spectacles et Concerts",["concert","festival"," bal ","bal du","spectacle","cinéma","cinema","théâtre","theatre","opéra","cirque","danse","comédie","conte","marionnette","barde","one-man","one man"]),
    ("Musées & Culture",    ["musée","musee","museum","expo","exposition","galerie","collection","visite guidée","visite de","micro-folie","fresques","patrimoine culturel","bibliothèque","médiathèque","archives","centre culturel"]),
    ("Shopping",            ["boutique","marché artisanal","créateur","vide-grenier","brocante","puces"]),
    ("Activités Sportives", ["golf","tennis","karting","bowling","squash","surf","voile","kayak","vélo","velo","escalade","equitation","sport nautique","nautique","plongée","kite","paddle","char à voile","location de bateaux","bouée","jet ski","pêche en mer","pêche à pied","foulées","trail","run","escape game","escape room","laser game","accrobranche","paintball","casino","kasino","karaoké","laser","bowling","parc aventure","tyrolienne","saut"]),
    ("Randonnées & Nature", ["nature","forêt","foret","randonnée","randonnee","sentier","parc","jardin","rando","hiking","promenade","forêt domaniale","ferme pédagogique","animaux de la ferme","cueillette","potager","tourbière","marais","réserve naturelle","bain de forêt","balade nature","balade en kayak","balade en mer","balade équestre"]),
    ("Gastronomie",         ["marché"]),
]

def detect_type(type_str):
    t = type_str.lower()
    for main_type, keywords in TYPE_KEYWORDS.items():
        if any(k in t for k in keywords):
            return main_type
    return "Activités"

def detect_category(text, main_type):
    txt = str(text).lower()

    # Séparer les règles par type pour éviter les faux positifs
    if main_type == "Restaurants":
        resto_cats = [
            "Restaurant étoilé","Brasserie","Burger","Pizzeria",
            "Sushi","Asiatique","Indien","Crêperie","Restaurant"
        ]
        for cat, keywords in CATEGORY_RULES:
            if cat not in resto_cats: continue
            if any(k in txt for k in keywords): return cat
        return "Restaurant"

    elif main_type == "Hébergement":
        hebergement_cats = [
            "Camping","Chambre d'hôte","Auberge de jeunesse","Insolite",
            "Résidence","Appartement","Maison","Gîte","Hôtel"
        ]
        for cat, keywords in CATEGORY_RULES:
            if cat not in hebergement_cats: continue
            if any(k in txt for k in keywords): return cat
        return "Hébergement"

    else:
        # Activités — ne pas matcher les mots hébergement/resto
        activite_cats = [
            "Thalasso & Bien-être","Spectacles et Concerts","Musées & Culture",
            "Monuments","Gastronomie","Shopping","Plage & Farniente",
            "Activités Sportives","Randonnées & Nature"
        ]
        for cat, keywords in CATEGORY_RULES:
            if cat not in activite_cats: continue
            if any(k in txt for k in keywords): return cat
        return "Autre"  # Catégorie par défaut si rien ne matche

# ============================================================
# HELPERS PRIX
# ============================================================

def extract_prices(text, main_type):
    prices = []
    for m in PRICE_REGEX.findall(text):
        try:
            v = float(m.replace(",", "."))
            if main_type == "Hébergement":
                if 20 <= v <= 1500: prices.append(v)   # nuit/semaine
            elif main_type == "Restaurants":
                if 5 <= v <= 200: prices.append(v)      # repas/personne
            else:  # Activités
                if 1 <= v <= 200: prices.append(v)      # entrée/personne
        except: pass
    return prices

def niveau_prix(avg, main_type):
    if avg is None or avg == 0: return "Gratuit"
    if main_type == "Hébergement":
        if avg < 60:   return "€"
        if avg < 150:  return "€€"
        if avg < 300:  return "€€€"
        return "€€€€"
    elif main_type == "Restaurants":
        if avg < 15:   return "€"
        if avg < 35:   return "€€"
        if avg < 70:   return "€€€"
        return "€€€€"
    else:  # Activités
        if avg < 10:   return "€"
        if avg < 25:   return "€€"
        if avg < 60:   return "€€€"
        return "€€€€"

# ============================================================
# HELPERS SCRAPING
# ============================================================

BAD_DOMAINS = [
    "facebook","instagram","youtube","linkedin","tiktok",
    "amazon","ebay","dealabs","pagesjaunes","restaurantguru",
    "wikipedia","larousse","linternaute","francebleu","maville",
    "yandex","etsy","gralon","societe.com","mapado","leboncoin",
    "annuaire-mairie","eterritoire","jds.fr","alltrails","hika.app",
    "tripadvisor","booking.com","airbnb","expedia","easyflirt",
    "mappy","mitula","papvacances","ouest-france","breizh-info",
    "alentoor","sortir-en-bretagne","unidivers","infolocale",
    "zoofrance","parc-loisir","recreatiloups","abcsalles",
]

GOOD_DOMAINS = [
    ".fr","tourisme","office","reservation","billetterie","booking",
]

def is_bad_url(url):
    return any(bad in url.lower() for bad in BAD_DOMAINS)

def score_url(url):
    """Priorise les URLs les plus susceptibles d'avoir des tarifs"""
    score = 0
    u = url.lower()
    for good in GOOD_DOMAINS:
        if good in u: score += 10
    if "tarif" in u: score += 15
    if "prix"  in u: score += 5
    return score

def get_confidence(url, prices):
    """Score de confiance sur les prix trouvés"""
    if not url: return 0.0
    score = 0.3
    u = url.lower()
    if "tarif"    in u: score += 0.3
    if "tourisme" in u: score += 0.2
    if len(prices) >= 2: score += 0.1
    if ".fr"      in u: score += 0.1
    return round(min(score, 1.0), 2)

def extract_socials(text):
    insta, fb = "", ""
    for u in URL_REGEX.findall(text):
        low = u.lower()
        if "instagram.com" in low and not insta: insta = u
        if "facebook.com"  in low and not fb:    fb    = u
    return insta, fb

def search_urls(name, city=""):
    """Recherche DuckDuckGo + scoring des URLs"""
    geo   = f" {city}" if city and city != "nan" else ""
    query = f"{name}{geo} tarif"
    urls  = []
    # Mots du titre pour filtrer les résultats non pertinents
    name_words = [w.lower() for w in re.sub(r'[^\w\s]', '', name).split() if len(w) > 3]
    try:
        with DDGS() as ddgs:
            results = ddgs.text(query, max_results=MAX_RESULTS, backend="lite")
            for r in results:
                url = r.get("href","")
                if not url or is_bad_url(url): continue
                # Garder seulement si l'URL ou le snippet contient un mot du titre
                snippet = (r.get("body","") + " " + r.get("title","")).lower()
                url_low = url.lower()
                if any(w in url_low or w in snippet for w in name_words):
                    if url not in urls: urls.append(url)
    except Exception as e:
        print(f"  DDG erreur: {e}")
        time.sleep(10)
    return sorted(urls, key=score_url, reverse=True)

def scrape_page(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, verify=False)
        if r.status_code != 200: return ""
        soup = BeautifulSoup(r.text, "lxml")
        for tag in soup(["script","style","noscript"]): tag.extract()
        return soup.get_text(separator="\n")
    except: return ""

def search_social(titre, ville, platform):
    """Cherche le profil Instagram ou Facebook via DuckDuckGo"""
    query = f'"{titre}" {ville} {platform}'
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=5, backend="lite"))
            for r in results:
                url = r.get("href","")
                if platform == "instagram" and "instagram.com" in url:
                    # Ignorer les posts et pages génériques
                    if "/p/" not in url and "/reel/" not in url and "/explore/" not in url:
                        return url
                elif platform == "facebook" and "facebook.com" in url:
                    # Ignorer les posts, events et pages non officielles
                    if "/posts/" not in url and "/events/" not in url and "/groups/" not in url:
                        return url
    except: pass
    return ""

def search_email_ddg(titre, ville):
    """Cherche l'email de contact via DuckDuckGo"""
    query = f'"{titre}" {ville} contact email'
    BAD_EMAIL = ["noreply","no-reply","example","tripadvisor","booking","airbnb",
                 "pagesjaunes","morbihan","tourisme","bretagne","gmail.com"]
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=5, backend="lite"))
            for r in results:
                snippet = r.get("body","") + " " + r.get("title","")
                for mail in EMAIL_REGEX.findall(snippet):
                    if "@" in mail and "." in mail and len(mail) < 60:
                        if not any(b in mail.lower() for b in BAD_EMAIL):
                            return mail
    except: pass
    return ""

# ============================================================
# GOOGLE PLACES (nouvelle API)
# ============================================================

PLACES_URL = "https://places.googleapis.com/v1/places:searchText"
PLACES_HEADERS = {
    "Content-Type": "application/json",
    "X-Goog-Api-Key": GOOGLE_API_KEY,
    "X-Goog-FieldMask": ",".join([
        "places.displayName",
        "places.websiteUri",
        "places.internationalPhoneNumber",
        "places.regularOpeningHours",
        "places.priceLevel",
        "places.rating",
        "places.userRatingCount",
        "places.accessibilityOptions",
        "places.allowsDogs",
        "places.goodForChildren",
        "places.parkingOptions",
        "places.restroom",
        "places.outdoorSeating",
        "places.goodForGroups",
    ])
}

GOOGLE_DAY = {1:0, 2:1, 3:2, 4:3, 5:4, 6:5, 0:6}

PRICE_MAP_GOOGLE = {
    "PRICE_LEVEL_FREE":           "Gratuit",
    "PRICE_LEVEL_INEXPENSIVE":    "€",
    "PRICE_LEVEL_MODERATE":       "€€",
    "PRICE_LEVEL_EXPENSIVE":      "€€€",
    "PRICE_LEVEL_VERY_EXPENSIVE": "€€€€",
}

def format_horaire(period):
    try:
        o  = period.get("open",  {})
        c  = period.get("close", {})
        oh = str(o.get("hour",   0)).zfill(2)
        om = str(o.get("minute", 0)).zfill(2)
        ch = str(c.get("hour",   0)).zfill(2)
        cm = str(c.get("minute", 0)).zfill(2)
        return f"{oh}:{om}-{ch}:{cm}"
    except: return None

def google_places_search(titre, adresse):
    query = f"{titre} {adresse} France".strip()
    try:
        r = requests.post(
            PLACES_URL, headers=PLACES_HEADERS,
            json={"textQuery": query, "languageCode": "fr"}, timeout=10
        )
        if r.status_code != 200:
            print(f"  Google erreur HTTP {r.status_code}")
            return None
        places = r.json().get("places", [])
        return places[0] if places else None
    except Exception as e:
        print(f"  Google exception: {e}")
        return None

# ============================================================
# HELPERS PARSING JSON
# ============================================================

def get_fr_text(obj):
    """Extrait le texte FR depuis un objet multilangue"""
    if obj is None: return ""
    if isinstance(obj, (int, float)): return str(obj)
    if isinstance(obj, str): return obj.strip()
    if isinstance(obj, list):
        for item in obj:
            v = get_fr_text(item)
            if v: return v
        return ""
    if isinstance(obj, dict):
        # Format {'fr': ['texte']} ou {'fr': 'texte'}
        for lang in ["fr", "en"]:
            val = obj.get(lang)
            if val is not None:
                v = get_fr_text(val)
                if v: return v
        # Fallback toutes les valeurs
        for val in obj.values():
            v = get_fr_text(val)
            if v: return v
    return ""

def clean_str(val):
    """Convertit proprement n'importe quelle valeur en string"""
    if val is None: return ""
    if isinstance(val, list):
        val = val[0] if val else ""
    if isinstance(val, dict):
        # Prendre la valeur fr ou première valeur
        val = val.get("fr", val.get("en", next(iter(val.values()), "")))
        if isinstance(val, list): val = val[0] if val else ""
    return str(val).strip() if val else ""

# ============================================================
# ÉTAPE 1 — PARSE DATATOURISME
# ============================================================

print("\n" + "="*60)
print("ÉTAPE 1 — Datatourisme")
print("="*60)

rows = []
nb_files = 0
nb_errors = 0

for root, dirs, files in os.walk(DATA_FOLDER):
    for file in files:
        if not (file.endswith(".json") or file.endswith(".json.gz")):
            continue

        path = os.path.join(root, file)
        nb_files += 1

        try:
            if file.endswith(".json.gz"):
                with gzip.open(path, "rt", encoding="utf-8") as f:
                    raw = f.read().strip()
            else:
                with open(path, "r", encoding="utf-8") as f:
                    raw = f.read().strip()

            # Ignorer les fichiers vides
            if not raw:
                continue

            data = json.loads(raw)

            # Support @graph, liste ou objet direct
            if "@graph" in data:       items = data["@graph"]
            elif isinstance(data, list): items = data
            else:                        items = [data]

            for item in items:
                try:
                    row = {c: "" for c in COLONNES}  # Toutes les valeurs initialisées en string vide

                    # ── TITRE ──
                    title = get_fr_text(item.get("rdfs:label", ""))
                    if not title: continue
                    row["TITRE"] = title

                    # ── DESCRIPTION ──
                    # Priorité : hasDescription > dc:description > description
                    desc = ""
                    for has_d in item.get("hasDescription", []) if isinstance(item.get("hasDescription"), list) else [item.get("hasDescription", {})]:
                        if not has_d: continue
                        dc = has_d.get("dc:description", {})
                        desc = get_fr_text(dc)
                        if desc: break
                    if not desc:
                        for field in ["dc:description", "description", "schema:description"]:
                            desc = get_fr_text(item.get(field, ""))
                            if desc: break
                    row["DESCRIPTION"] = desc[:2000]

                    # ── TYPE ──
                    rdf_types = item.get("@type", [])
                    if isinstance(rdf_types, str): rdf_types = [rdf_types]
                    rdf_str = " ".join(rdf_types)
                    row["TYPE"] = detect_type(rdf_str)

                    # ── CATÉGORIE ──
                    full_text = f"{title} {row['DESCRIPTION']} {rdf_str}"
                    row["CATEGORIE"] = detect_category(full_text, row["TYPE"])

                    # ── CLASSIFICATION ──
                    for review in item.get("hasReview", []):
                        note = review.get("schema:ratingValue", "")
                        if note: row["CLASSIFICATION"] = f"{note} étoiles"; break
                    if not row["CLASSIFICATION"]:
                        for classif in item.get("hasClassification", []):
                            label = get_fr_text(classif.get("rdfs:label", ""))
                            if label: row["CLASSIFICATION"] = label; break
                    if not row["CLASSIFICATION"]:
                        for s in ["5 étoiles","4 étoiles","3 étoiles","2 étoiles","1 étoile"]:
                            if s in full_text.lower():
                                row["CLASSIFICATION"] = s; break

                    # ── NOMBRE MAX ──
                    for cf in ["schema:maximumAttendeeCapacity","capacity","peopleCapacity"]:
                        if item.get(cf):
                            row["NOMBRE_MAX"] = str(item[cf]); break

                    # ── LOCALISATION (adresse + horaires) ──
                    locs = item.get("isLocatedAt", [])
                    if isinstance(locs, dict): locs = [locs]
                    for loc in locs:
                        if not loc: continue

                        # Adresse
                        addrs = loc.get("schema:address", [])
                        if isinstance(addrs, dict): addrs = [addrs]
                        for addr in addrs:
                            street = clean_str(addr.get("schema:streetAddress", ""))
                            postal = clean_str(addr.get("schema:postalCode", ""))
                            city   = clean_str(addr.get("schema:addressLocality", ""))
                            parts  = [p for p in [street, f"{postal} {city}".strip(), "FRANCE"] if p.strip()]
                            row["ADRESSE"] = ", ".join(parts)
                            break

                        # Horaires
                        ohs = loc.get("schema:openingHoursSpecification", [])
                        if isinstance(ohs, dict): ohs = [ohs]
                        for oh in ohs:
                            days   = oh.get("schema:dayOfWeek", [])
                            opens  = clean_str(oh.get("schema:opens",  ""))
                            closes = clean_str(oh.get("schema:closes", ""))
                            if isinstance(days, str): days = [days]
                            for day in days:
                                day_clean = day.split("/")[-1]  # schema:Monday → Monday
                                if day_clean in JOURS_EN:
                                    idx_j = JOURS_EN.index(day_clean)
                                    col   = f"HORAIRES_{JOURS_FR[idx_j]}"
                                    if opens and closes and not row[col]:
                                        row[col] = f"{opens}-{closes}"

                    # ── CONTACT ──
                    contacts = item.get("hasContact", [])
                    if isinstance(contacts, dict): contacts = [contacts]
                    for contact in contacts:
                        def empty(v): return not v or str(v).strip() in ("", "nan")
                        if empty(row["CONTACT_TEL"]):
                            row["CONTACT_TEL"] = clean_str(contact.get("schema:telephone", ""))
                        if empty(row["CONTACT_EMAIL"]):
                            row["CONTACT_EMAIL"] = clean_str(contact.get("schema:email", ""))
                        if empty(row["CONTACT_SITE"]):
                            url_site = clean_str(contact.get("foaf:homepage", ""))
                            if url_site and not is_bad_url(url_site):
                                row["CONTACT_SITE"] = url_site
                        same_as = contact.get("schema:sameAs", [])
                        if isinstance(same_as, str): same_as = [same_as]
                        for url in same_as:
                            if "instagram" in url.lower() and empty(row["CONTACT_INSTAGRAM"]):
                                row["CONTACT_INSTAGRAM"] = url
                            elif "facebook" in url.lower() and empty(row["CONTACT_FACEBOOK"]):
                                row["CONTACT_FACEBOOK"] = url

                    # ── PUBLIC ──
                    row["PUBLIC_ADULTE"] = "oui"
                    desc_low = str(row.get("DESCRIPTION","")).lower()
                    title_low = title.lower()
                    combined_low = desc_low + " " + title_low

                    # ── PUBLIC déduit depuis description + titre ──
                    combined_low = (str(row.get("DESCRIPTION","")) + " " + title).lower()

                    # ENFANT — large spectre de détection
                    ENFANT_KW = [
                        "enfant","famille","familial","familiale","junior","kids","children",
                        "tout-petit","bébé","ado","adolescent","jeune public",
                        "mini-club","animation enfants","activité enfant","tarif enfant",
                        "poney","ferme pédagogique","parc de loisirs","aire de jeux",
                        "atelier enfant","spectacle jeunesse","conte","marionnette",
                    ]
                    if any(k in combined_low for k in ENFANT_KW):
                        row["PUBLIC_ENFANT"] = "oui"

                    # PMR — large spectre
                    PMR_KW = [
                        "pmr","handicap","mobilité réduite","fauteuil roulant",
                        "accessible","accessibilité","personne à mobilité",
                        "accès adapté","entrée de plain-pied","rampe d'accès",
                        "toilettes adaptées","stationnement adapté",
                    ]
                    if any(k in combined_low for k in PMR_KW):
                        row["PUBLIC_PMR"] = "oui"

                    # ANIMAL — distinguer animaux acceptés vs animaux présents sur place
                    ANIMAL_OUI_KW = [
                        "animaux acceptés","animaux bienvenus","chiens acceptés",
                        "chiens bienvenus","chien accepté","chien bienvenu",
                        "dog friendly","animaux de compagnie acceptés",
                        "animal accepté","animal bienvenu","nos compagnons",
                        "avec votre animal","avec votre chien","animaux admis",
                        "animaux autorisés","pets welcome","pet friendly",
                    ]
                    ANIMAL_NON_KW = [
                        "animaux non acceptés","animaux non admis","animaux interdits",
                        "chiens interdits","no pets","sans animal",
                    ]
                    if any(k in combined_low for k in ANIMAL_NON_KW):
                        row["PUBLIC_ANIMAL"] = "non"
                    elif any(k in combined_low for k in ANIMAL_OUI_KW):
                        row["PUBLIC_ANIMAL"] = "oui"

                    # ── ÉQUIPEMENTS depuis description ──
                    equip = []
                    desc_equip = (str(row.get("DESCRIPTION","")) + " " + title).lower()
                    equip_kw_map = {
                        "wifi":       ["wifi","wi-fi","connexion internet","internet gratuit"],
                        "parking":    ["parking","stationnement","place de parking","garage"],
                        "piscine":    ["piscine","bassin","baignade intérieure"],
                        "jacuzzi":    ["jacuzzi","balnéo","bain à remous","bain bouillonnant"],
                        "sauna":      ["sauna","hammam","bain de vapeur"],
                        "terrasse":   ["terrasse","balcon","véranda","jardin privatif"],
                        "barbecue":   ["barbecue","plancha","coin grill"],
                        "pétanque":   ["pétanque","boules"],
                        "vélo":       ["vélo","local à vélos","abri vélos","location de vélos"],
                        "spa":        ["espace spa","centre spa","soin spa","thalasso"],
                        "restauration":["restaurant sur place","restauration","repas inclus","petit-déjeuner inclus"],
                        "animaux":    ["animaux acceptés","animaux bienvenus","dog friendly"],
                        "baby":       ["lit bébé","chaise haute","équipement bébé"],
                        "lave-linge": ["lave-linge","machine à laver","buanderie"],
                        "climatisation":["climatisation","clim","air conditionné"],
                    }
                    for label, triggers in equip_kw_map.items():
                        if any(t in desc_equip for t in triggers):
                            equip.append(label)
                    row["EQUIPEMENT"] = ", ".join(equip)

                    # ── TARIFS OFFICIELS depuis hasOffer ──
                    prices_official = []
                    for offer in item.get("hasOffer", []):
                        for price_spec in offer.get("schema:priceSpecification", []):
                            price = price_spec.get("schema:price")
                            if price is not None:
                                try:
                                    v = float(str(price).replace(",", "."))
                                    if v >= 0: prices_official.append(v)
                                except: pass
                    if prices_official:
                        row["PRIX_MIN"]    = min(prices_official)
                        row["PRIX_MAX"]    = max(prices_official)
                        row["PRIX_MOYEN"]  = round(statistics.median(prices_official), 2)
                        row["NIVEAUX_PRIX"]= niveau_prix(row["PRIX_MOYEN"], row["TYPE"])
                        row["PRIX_CONFIANCE"] = 1.0  # Source officielle

                    rows.append(row)

                except Exception as e:
                    nb_errors += 1

        except Exception as e:
            print(f"  Erreur fichier {file}: {e}")
            nb_errors += 1

df = pd.DataFrame(rows, columns=COLONNES)
df = df.drop_duplicates(subset=["TITRE"])

if TEST_MODE:
    df = df.head(TEST_LIMIT)

# Typage strict dès la création
NUM_COLS  = ["PRIX_MIN","PRIX_MAX","PRIX_MOYEN","PRIX_CONFIANCE","NOTE","NOMBRE_AVIS","NOMBRE_MAX"]  # REFERENCE, TAGS, PHOTOS, TITRE... restent string
TEXT_COLS_NORM = [c for c in COLONNES if c not in NUM_COLS]
for col in TEXT_COLS_NORM:
    df[col] = df[col].fillna("").astype(str).replace("nan","").str.strip()
for col in NUM_COLS:
    df[col] = pd.to_numeric(df[col], errors="coerce")

print(f"  [OK] {len(df)} établissements extraits ({nb_files} fichiers, {nb_errors} erreurs)")
df.to_excel(OUTPUT_EXCEL, index=False)
print(f"  [SAVE] Sauvegarde initiale → {OUTPUT_EXCEL}")

# ============================================================
# MIGRATION — mettre à jour les anciens labels prix
# ============================================================

PRIX_MIGRATION = {
    "Petit budget": "€",
    "Confort":      "€€",
    "Premium":      "€€€",
    "Luxe":         "€€€€",
}
if "NIVEAUX_PRIX" in df.columns:
    df["NIVEAUX_PRIX"] = df["NIVEAUX_PRIX"].replace(PRIX_MIGRATION)

# ============================================================
# NORMALISATION — typage strict de toutes les colonnes
# ============================================================

NUM_COLS  = ["PRIX_MIN","PRIX_MAX","PRIX_MOYEN","PRIX_CONFIANCE","NOTE","NOMBRE_AVIS","NOMBRE_MAX"]  # REFERENCE, TAGS, PHOTOS, TITRE... restent string
TEXT_COLS_NORM = [c for c in COLONNES if c not in NUM_COLS]

for col in TEXT_COLS_NORM:
    if col in df.columns:
        df[col] = df[col].fillna("").astype(str).replace("nan", "").str.strip()

for col in NUM_COLS:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# ============================================================
# ÉTAPE 2 & 3 — ENRICHISSEMENT
# ============================================================

print("\n" + "="*60)
print("ÉTAPE 2 & 3 — Google Places + Scraping")
print("="*60)

for idx, row in df.iterrows():
    try:
        title     = str(row["TITRE"])
        address   = str(row["ADRESSE"])
        main_type = str(row["TYPE"])

        print(f"\n[{idx+1}/{len(df)}] {title}")

        # ── GOOGLE PLACES ──
        place = google_places_search(title, address)

        if place:
            # Note + avis
            cur_note = df.at[idx, "NOTE"]
            if pd.isna(cur_note) or str(cur_note).strip() in ("", "nan", "0", "0.0"):
                rating = place.get("rating")
                if rating is not None:
                    df.at[idx, "NOTE"]        = float(rating)
                    df.at[idx, "NOMBRE_AVIS"] = int(place.get("userRatingCount", 0))

            # Site web
            cur_site = df.at[idx, "CONTACT_SITE"]
            if (pd.isna(cur_site) or str(cur_site).strip() in ("", "nan")) and place.get("websiteUri"):
                df.at[idx, "CONTACT_SITE"] = place["websiteUri"]

            # Téléphone
            cur_tel = df.at[idx, "CONTACT_TEL"]
            if (pd.isna(cur_tel) or str(cur_tel).strip() in ("", "nan")) and place.get("internationalPhoneNumber"):
                df.at[idx, "CONTACT_TEL"] = place["internationalPhoneNumber"]

            # Horaires
            periods = place.get("regularOpeningHours", {}).get("periods", [])
            for period in periods:
                dg = period.get("open", {}).get("day")
                if dg is not None and dg in GOOGLE_DAY:
                    col = f"HORAIRES_{JOURS_FR[GOOGLE_DAY[dg]]}"
                    if not df.at[idx, col]:
                        h = format_horaire(period)
                        if h: df.at[idx, col] = h

            # Niveau prix Google (seulement si pas de tarifs officiels)
            cur_prix = df.at[idx, "NIVEAUX_PRIX"]
            if (pd.isna(cur_prix) or str(cur_prix).strip() in ("", "nan")) and place.get("priceLevel"):
                df.at[idx, "NIVEAUX_PRIX"] = PRICE_MAP_GOOGLE.get(place["priceLevel"], "")

            # PMR — Google Places
            access = place.get("accessibilityOptions", {})
            cur_pmr = df.at[idx, "PUBLIC_PMR"]
            if (pd.isna(cur_pmr) or str(cur_pmr).strip() in ("", "nan")):
                if access.get("wheelchairAccessibleEntrance") is True:
                    df.at[idx, "PUBLIC_PMR"] = "oui"
                elif access.get("wheelchairAccessibleEntrance") is False:
                    df.at[idx, "PUBLIC_PMR"] = "non"

            # Animaux — Google Places
            cur_animal = df.at[idx, "PUBLIC_ANIMAL"]
            if place.get("allowsDogs") is not None and (pd.isna(cur_animal) or str(cur_animal).strip() in ("", "nan")):
                df.at[idx, "PUBLIC_ANIMAL"] = "oui" if place["allowsDogs"] else "non"

            # Enfant — Google Places
            cur_enfant = df.at[idx, "PUBLIC_ENFANT"]
            if place.get("goodForChildren") is not None and (pd.isna(cur_enfant) or str(cur_enfant).strip() in ("", "nan")):
                df.at[idx, "PUBLIC_ENFANT"] = "oui" if place["goodForChildren"] else "non"

            # Équipements Google
            cur_equip  = df.at[idx, "EQUIPEMENT"]
            equip_str  = "" if pd.isna(cur_equip) or str(cur_equip).strip() == "nan" else str(cur_equip)
            equip_list = [e.strip() for e in equip_str.split(",") if e.strip()]
            parking = place.get("parkingOptions", {})
            if (parking.get("freeParkingLot") or parking.get("paidParkingLot")) and "parking" not in equip_list:
                equip_list.append("parking")
            if place.get("restroom")       and "toilettes" not in equip_list: equip_list.append("toilettes")
            if place.get("outdoorSeating") and "terrasse"  not in equip_list: equip_list.append("terrasse")
            if place.get("goodForGroups")  and "groupes"   not in equip_list: equip_list.append("groupes")
            df.at[idx, "EQUIPEMENT"] = ", ".join(equip_list)

            print(f"  Google OK note:{place.get('rating','')} prix:{place.get('priceLevel','')}")
        else:
            print("  Google X non trouvé")

        # ── NATURE/FORÊT/PLAGE → Gratuit seulement si Activités (pas Hébergement)
        NATURE_FREE = ["plage","beach","anse","forêt domaniale","foret domaniale","dune","sentier","rivière","lac","domaniale","marais","tourbière"]
        NATURE_TYPES = ["naturalheritage","naturalsite","beach"]
        is_nature = (
            main_type == "Activités" and (
                any(k in title.lower() for k in NATURE_FREE) or
                any(t in str(row.get("TYPE","")).lower() for t in NATURE_TYPES)
            )
        )
        if is_nature:
            df.at[idx, "NIVEAUX_PRIX"]   = "Gratuit"
            df.at[idx, "PRIX_MIN"]       = 0.0
            df.at[idx, "PRIX_MAX"]       = 0.0
            df.at[idx, "PRIX_MOYEN"]     = 0.0
            df.at[idx, "PRIX_CONFIANCE"] = 0.9
            print("  Nature → Gratuit (skipping scraping)")

        # ── RECHERCHE RÉSEAUX SOCIAUX + EMAIL via DuckDuckGo ──
        city = ""
        for part in str(row.get("ADRESSE","")).split(","):
            tokens = part.strip().split(" ", 1)
            if len(tokens) == 2 and tokens[0].isdigit() and len(tokens[0]) == 5:
                city = tokens[1].strip()
                break

        cur_insta = str(df.at[idx, "CONTACT_INSTAGRAM"]).strip()
        if cur_insta in ("", "nan"):
            insta = search_social(title, city, "instagram")
            if insta:
                df.at[idx, "CONTACT_INSTAGRAM"] = insta
                print(f"  Instagram OK {insta}")
            time.sleep(0.5)

        cur_fb = str(df.at[idx, "CONTACT_FACEBOOK"]).strip()
        if cur_fb in ("", "nan"):
            fb = search_social(title, city, "facebook")
            if fb:
                df.at[idx, "CONTACT_FACEBOOK"] = fb
                print(f"  Facebook OK {fb}")
            time.sleep(0.5)

        cur_email = str(df.at[idx, "CONTACT_EMAIL"]).strip()
        if cur_email in ("", "nan"):
            mail = search_email_ddg(title, city)
            if mail:
                df.at[idx, "CONTACT_EMAIL"] = mail
                print(f"  Email OK {mail}")
            time.sleep(0.5)

        # ── PRIX DEPUIS DESCRIPTION (avant scraping) ──
        desc_text = str(row.get("DESCRIPTION",""))
        if desc_text and (pd.isna(df.at[idx,"PRIX_MIN"]) or df.at[idx,"PRIX_MIN"] == 0):
            desc_prices = []
            for m in PRICE_REGEX.findall(desc_text):
                try:
                    v = float(m.replace(",","."))
                    if 1 <= v <= 500: desc_prices.append(v)
                except: pass
            if desc_prices:
                df.at[idx,"PRIX_MIN"]       = float(min(desc_prices))
                df.at[idx,"PRIX_MAX"]       = float(max(desc_prices))
                df.at[idx,"PRIX_MOYEN"]     = float(round(statistics.median(desc_prices),2))
                df.at[idx,"NIVEAUX_PRIX"]   = str(niveau_prix(df.at[idx,"PRIX_MOYEN"], main_type))
                df.at[idx,"PRIX_CONFIANCE"] = 0.95  # Point décimal
                print(f"  Prix description OK {min(desc_prices)}€-{max(desc_prices)}€")

        # ── SCRAPING (seulement si tarifs pas encore remplis) ──
        if is_nature:
            print("  Scraping skippé (lieu nature → gratuit)")
        elif df.at[idx, "PRIX_CONFIANCE"] >= 0.95:
            print("  Scraping skippé (tarifs déjà trouvés)")
        elif not is_nature:
            all_prices = []
            best_url   = None
            city       = ""
            if row["ADRESSE"]:
                # Extraire la ville depuis l'adresse
                for part in str(row["ADRESSE"]).split(","):
                    tokens = part.strip().split(" ", 1)
                    if len(tokens) == 2 and tokens[0].isdigit() and len(tokens[0]) == 5:
                        city = tokens[1].strip()
                        break

            urls = search_urls(title, city)

            for url in urls:
                print(f"  → {url[:80]}")
                text = scrape_page(url)
                if not text: continue

                # Email
                # Email — seulement depuis le site officiel ou une URL très proche
                cur_email = df.at[idx, "CONTACT_EMAIL"]
                cur_site_check = str(df.at[idx, "CONTACT_SITE"]).strip()
                # Extraire le domaine du site officiel
                site_domain = ""
                if cur_site_check and cur_site_check not in ("", "nan"):
                    try:
                        from urllib.parse import urlparse
                        site_domain = urlparse(cur_site_check).netloc.replace("www.","")
                    except: pass
                # Accepter email seulement si l'URL scrapée partage le même domaine
                url_domain = ""
                try:
                    from urllib.parse import urlparse
                    url_domain = urlparse(url).netloc.replace("www.","")
                except: pass
                if (pd.isna(cur_email) or str(cur_email).strip() in ("", "nan")):
                    # Email seulement si même domaine que le site officiel
                    if site_domain and url_domain == site_domain:
                        emails = EMAIL_REGEX.findall(text)
                        BAD_EMAIL_KW = ["noreply","no-reply","easyflirt","example","test@",
                                        "tripadvisor","donotreply","unsubscribe","mailer",
                                        "morbihan-affaires","cyclos","20minutes","gmail.com",
                                        "booking","airbnb","expedia","gites-de-france"]
                        for mail in emails:
                            if "@" in mail and "." in mail and len(mail) < 80:
                                if not any(bk in mail.lower() for bk in BAD_EMAIL_KW):
                                    df.at[idx, "CONTACT_EMAIL"] = mail
                                    break

                # Téléphone
                cur_tel = df.at[idx, "CONTACT_TEL"]
                if pd.isna(cur_tel) or str(cur_tel).strip() in ("", "nan"):
                    phones = PHONE_REGEX.findall(text)
                    if phones: df.at[idx, "CONTACT_TEL"] = phones[0]

                # Site officiel — score minimum 20 ET contient le nom ou est sur tarif/billetterie
                cur_site = df.at[idx, "CONTACT_SITE"]
                if (pd.isna(cur_site) or str(cur_site).strip() in ("", "nan")) and not is_bad_url(url):
                    url_score = score_url(url)
                    title_words = [w for w in title.lower().split() if len(w) > 3]
                    url_matches_title = any(w in url.lower() for w in title_words)
                    if url_score >= 20 or (url_score >= 10 and url_matches_title):
                        df.at[idx, "CONTACT_SITE"] = url

                # Réseaux sociaux — seulement depuis le site officiel du prestataire
                cur_site_now = str(df.at[idx, "CONTACT_SITE"]).strip()
                if cur_site_now and cur_site_now not in ("", "nan") and url == cur_site_now:
                    insta, fb = extract_socials(text)
                    cur_insta = df.at[idx, "CONTACT_INSTAGRAM"]
                    cur_fb    = df.at[idx, "CONTACT_FACEBOOK"]
                    if insta and (pd.isna(cur_insta) or str(cur_insta).strip() in ("", "nan")):
                        df.at[idx, "CONTACT_INSTAGRAM"] = insta
                    if fb and (pd.isna(cur_fb) or str(cur_fb).strip() in ("", "nan")):
                        df.at[idx, "CONTACT_FACEBOOK"] = fb

                # Pas d'équipements depuis le scraping — uniquement depuis Datatourisme
                # Prix — ignorer les pages trop bruitées (comme ton script original)
                prices = extract_prices(text, main_type)
                if len(prices) > 100:
                    print(f"    Page trop bruitée ({len(prices)} prix), ignorée")
                    continue
                if prices:
                    all_prices.extend(prices)
                    if best_url is None:
                        best_url = url  # Garder la 1ère URL avec des prix

            if all_prices:
                pmin = float(min(all_prices))
                pmax = float(max(all_prices))
                pavg = float(round(sum(all_prices) / len(all_prices), 2))
                conf = float(get_confidence(best_url, all_prices))
                df.at[idx, "PRIX_MIN"]       = pmin
                df.at[idx, "PRIX_MAX"]       = pmax
                df.at[idx, "PRIX_MOYEN"]     = pavg
                df.at[idx, "PRIX_CONFIANCE"] = conf
                df.at[idx, "NIVEAUX_PRIX"]   = str(niveau_prix(pavg, main_type))
                print(f"  Tarifs OK {pmin}€-{pmax}€ moy:{pavg}€ conf:{conf} → {niveau_prix(pavg, main_type)}")
            else:
                print("  Tarifs X aucun prix trouvé")

        # Sauvegarde auto
        if (idx + 1) % SAVE_EVERY == 0:
            for c in TEXT_COLS_NORM:
                if c in df.columns: df[c] = df[c].fillna("").astype(str).replace("nan","").str.strip()
            df.to_excel(OUTPUT_EXCEL, index=False)
            print(f"  [SAVE] Sauvegarde auto ({idx+1} traités)")

        time.sleep(DELAY)

    except Exception as e:
        print(f"  Erreur ligne {idx}: {e}")

# ============================================================
# EXPORT FINAL
# ============================================================

# Nettoyer les colonnes texte (enlever "nan")
TEXT_COLS = [c for c in COLONNES if c not in ["PRIX_MIN","PRIX_MAX","PRIX_MOYEN","PRIX_CONFIANCE","NOTE","NOMBRE_AVIS","NOMBRE_MAX"]]
for col in TEXT_COLS:
    if col in df.columns:
        df[col] = df[col].fillna("").astype(str).replace("nan","").str.strip()

# PUBLIC — "non" par défaut si vide après tout l'enrichissement
for col in ["PUBLIC_ADULTE","PUBLIC_ENFANT","PUBLIC_PMR","PUBLIC_ANIMAL"]:
    if col in df.columns:
        df[col] = df[col].fillna("non")
        df[col] = df[col].replace("","non").replace("nan","non")

# PRIX_CONFIANCE — forcer point décimal (pas virgule)
if "PRIX_CONFIANCE" in df.columns:
    df["PRIX_CONFIANCE"] = pd.to_numeric(
        df["PRIX_CONFIANCE"].astype(str).str.replace(",","."), errors="coerce"
    ).round(2)

# PUBLIC — mettre "non" si vide
for col in ["PUBLIC_ADULTE","PUBLIC_ENFANT","PUBLIC_PMR","PUBLIC_ANIMAL"]:
    if col in df.columns:
        df[col] = df[col].fillna("non").replace("","non").replace("nan","non")

print("\n" + "="*60)
print("BILAN FINAL")
print("="*60)
for col in COLONNES:
    if col in df.columns:
        filled = (df[col].astype(str).str.strip().replace("nan","") != "").sum()
        print(f"  {col:<25}: {filled}/{len(df)}")

df.to_excel(OUTPUT_EXCEL, index=False)
print(f"\n  Fichier final : {OUTPUT_EXCEL}")
print("="*60)