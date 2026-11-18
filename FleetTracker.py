"""MSC CRUISES & EXPLORA JOURNEYS FLEET TRACKER
Usa codici MMSI di ogni vessel per identificare posizione geografica, status e destinazione
Aggiornamento ogni 30 min, pubblica su GitHub Pages"""

import asyncio, json, websockets, openpyxl, os
from datetime import datetime, timezone
from pathlib import Path

API_KEY = os.environ["AISSTREAM_API_KEY"]
FOLDER = Path.cwd()
REFRESH = 30  # minuti
REPO_DIR = FOLDER
COLLECTION_MINUTES = 12

FLEET = {
    "357281000": ("MSC Armonia",       "Lirica",     "#C084FC", False),
    "356716000": ("MSC Sinfonia",      "Lirica",     "#7C3AED", False),
    "356042000": ("MSC Lirica",        "Lirica",     "#5B21B6", False),
    "357627000": ("MSC Opera",         "Lirica",     "#3B0764", False),
    "352003000": ("MSC Musica",        "Musica",     "#4ADE80", False),
    "372497000": ("MSC Orchestra",     "Musica",     "#16A34A", False),
    "355931000": ("MSC Poesia",        "Musica",     "#15803D", False),
    "352594000": ("MSC Magnifica",     "Musica",     "#052E16", False),
    "370648000": ("MSC Fantasia",      "Fantasia",   "#38BDF8", False),
    "357698000": ("MSC Splendida",     "Fantasia",   "#0EA5E9", False),
    "373178000": ("MSC Divina",        "Fantasia",   "#0284C7", False),
    "371861000": ("MSC Preziosa",      "Fantasia",   "#075985", False),
    "249973000": ("MSC Meraviglia",    "Meraviglia", "#E777E2", False),
    "248992000": ("MSC Bellissima",    "Meraviglia", "#C02FC0", False),
    "215325000": ("MSC Grandiosa",     "Meraviglia", "#861C78", False),
    "215610000": ("MSC Virtuosa",      "Meraviglia", "#C4056E", False),
    "256281000": ("MSC Euribia",       "Meraviglia", "#70034C", False),
    "248392000": ("MSC Seaside",       "Seaside",    "#FC802D", False),
    "248717000": ("MSC Seaview",       "Seaside",    "#EA580C", False),
    "215920000": ("MSC Seashore",      "Seaside",    "#9A3412", False),
    "256059000": ("MSC Seascape",      "Seaside",    "#431407", False),
    "256070000": ("MSC World Europa",  "World",      "#9CA3AF", False),
    "249193000": ("MSC World America", "World",      "#535E6E", False),
    "248715600": ("MSC World Asia",    "World",      "#374151", False),
    "256389000": ("Explora I",         "Explora",    "#FCD34D", True),
    "256798000": ("Explora II",        "Explora",    "#F59E0B", True),
    "247503500": ("Explora III",       "Explora",    "#D97706", True),
}

CLASS_ORDER = ["Explora", "World", "Seaside", "Meraviglia", "Fantasia", "Musica", "Lirica"]

# Porti ricavati da ship calendar
PORTS = {
    # ITALY
    "ITGOA": "Genoa",
    "ITCIV": "Civitavecchia",
    "ITPAL": "Palermo",
    "ITPMO": "Palermo",
    "ITNAP": "Naples",
    "ITBRI": "Bari",
    "ITTRS": "Trieste",
    "ITVCE": "Venice",
    "ITLIV": "Livorno",
    "ITANC": "Ancona",
    "ITMES": "Messina",
    "ITCAT": "Catania",
    "ITSPE": "La Spezia",
    "ITCAG": "Cagliari",
    "ITBDS": "Brindisi",
    "ITREG": "Reggio Calabria",
    "ITSIR": "Syracuse",
    "ITOLB": "Olbia",
    "ITMIL": "Milazzo",
    "ITSVN": "Savona",

    # SPAIN
    "ESBCN": "Barcelona",
    "ESVLC": "Valencia",
    "ESPAL": "Palma de Mallorca",
    "ESMAH": "Mahon",
    "ESLPA": "Las Palmas",
    "ESTCI": "Santa Cruz de Tenerife",
    "ESALI": "Alicante",
    "ESCAD": "Cadiz",
    "ESVGO": "Vigo",
    "ESGIJ": "Gijon",
    "ESLCG": "La Coruna",
    "ESBIO": "Bilbao",
    "ESIBZ": "Ibiza",
    "ESACE": "Arrecife",
    "ESFUE": "Puerto del Rosario",
    "ESSCT": "Santa Cruz de Tenerife",
    "ESSCP": "Santa Cruz de La Palma",
    "ESTGN": "Tarragona",

    # FRANCE
    "FRMAR": "Marseille",
    "FRMRS": "Marseille",
    "FRTLN": "Toulon",
    "FRAJC": "Ajaccio",
    "FRLEH": "Le Havre",
    "FRLRH": "La Rochelle",
    "FRBRE": "Brest",
    "FRCFR": "Cherbourg",
    "FRCNL": "Cannes",
    "FRSTN": "Saint-Nazaire",

    # GREECE
    "GRPIR": "Piraeus",
    "GRHER": "Heraklion",
    "GRRHO": "Rhodes",
    "GRSAN": "Santorini",
    "GRJTR": "Santorini",
    "GRJMK": "Mykonos",
    "GRCFU": "Corfu",
    "GRKAT": "Katakolon",
    "GRKAK": "Katakolon",
    "GRCEA": "Kefalonia",
    "GRSYR": "Syros",

    # TURKEY
    "TRIST": "Istanbul",
    "TRIZM": "Izmir",
    "TRBOD": "Bodrum",
    "TRKUS": "Kusadasi",
    "TRMRM": "Marmaris",

    # MOROCCO
    "MABAS": "Tanger Med",
    "MAAGA": "Agadir",
    "MACAS": "Casablanca",
    "MATNG": "Tangier",

    # EGYPT
    "EGPSD": "Port Said",
    "EGALY": "Alexandria",
    "EGSHK": "Sharm El-Sheikh",
    "EGSCN": "Suez Canal North",
    "EGSCS": "Suez Canal South",

    # UAE
    "AEDXB": "Dubai",
    "AEAUH": "Abu Dhabi",
    "AESRB": "Sir Bani Yas",

    # ASIA
    "SGSIN": "Singapore",
    "JPYOK": "Yokohama",
    "CNSHA": "Shanghai",
    "CNHKG": "Hong Kong",
    "JPOSA": "Osaka",
    "JPKOB": "Kobe",
    "JPFUK": "Fukuoka",
    "JPNGO": "Nagoya",
    "JPTOK": "Tokyo",
    "JPNAH": "Naha",
    "JPKAG": "Kagoshima",
    "JPNAK": "Nagasaki",
    "JPSAS": "Sasebo",
    "JPHAK": "Hakodate",
    "JPHIS": "Hiroshima",
    "JPISK": "Ishigaki",
    "JPMIY": "Miyako Island",
    "JPKAN": "Kanazawa",
    "JPSKA": "Sakata",
    "JPYAT": "Yatsushiro",
    "JPABR": "Aburatsu",
    "KRINC": "Incheon",
    "KRPUS": "Busan",
    "KRGAN": "Gangjeong",
    "TWKHH": "Kaohsiung",
    "TWKEL": "Keelung",
    "MYPGU": "Penang",
    "MYKLG": "Port Klang",
    "PHMNL": "Manila",
    "VNDAN": "Da Nang",

    # NORTH AMERICA
    "USNYC": "New York",
    "USMIA": "Miami",
    "USFLL": "Fort Lauderdale",
    "USGAL": "Galveston",
    "USSEA": "Seattle",
    "USLAX": "Los Angeles",
    "USSAN": "San Diego",
    "USBOS": "Boston",
    "USSJU": "San Juan",
    "USFPB": "Port Canaveral",
    "USHIL": "Hilo",
    "USHNL": "Honolulu",
    "USJNU": "Juneau",
    "USKET": "Ketchikan",
    "USSIT": "Sitka",
    "USICY": "Icy Strait Point",
    "USTRA": "Tracy Arm",

    # CARIBBEAN / CENTRAL AMERICA
    "MXCUN": "Cozumel",
    "MXCSB": "Cabo San Lucas",
    "MXPVR": "Puerto Vallarta",
    "MXCOS": "Costa Maya",
    "BMBDA": "Bermuda",
    "COBAQ": "Cartagena",
    "TTPOS": "Port of Spain",
    "BBBRG": "Bridgetown",
    "LCSLC": "Castries",
    "VCSVD": "Kingstown",
    "DOMSD": "Santo Domingo",
    "DOLRM": "La Romana",
    "DOLRN": "La Romana",
    "DOPOP": "Puerto Plata",
    "DOSAM": "Samana",
    "DOCBR": "Cabo Rojo",
    "DOAMB": "Amber Cove",
    "DOCAT": "Catalina Island",
    "HNRTN": "Roatan",
    "BZCZE": "Belize City",
    "CRCAL": "Puerto Limon",
    "CRPUN": "Puntarenas",
    "PAMIG": "Panama Canal (Balboa)",
    "PACTB": "Panama Canal (Cristobal)",
    "PALON": "Colon",
    "JMFAL": "Falmouth",
    "JMMBJ": "Montego Bay",
    "JMOCH": "Ocho Rios",
    "KYGEO": "George Town",
    "TCGDT": "Grand Turk",
    "BSFPO": "Freeport",
    "BSNAS": "Nassau",
    "AGSTJ": "St. John's",
    "DMROS": "Roseau",
    "GPPPT": "Pointe-à-Pitre",
    "MQFDF": "Fort-de-France",
    "VISTJ": "Charlotte Amalie",
    "VIRTW": "Road Town",
    "VIVGD": "Virgin Gorda",
    "SXPHI": "Philipsburg",
    "BQKRL": "Kralendijk",
    "AWORA": "Oranjestad",
    "CWWIL": "Willemstad",
    "GDSGT": "St. George's",
    "KNBST": "Basseterre",
    "MSOCN": "Ocean Cay",

    # SOUTH AMERICA
    "BRSSS": "Santos",
    "BRSSZ": "Santos",
    "ARBUE": "Buenos Aires",
    "BRRIJ": "Rio de Janeiro",
    "BRSAL": "Salvador",
    "BRMAO": "Maceio",
    "BRILS": "Ilheus",
    "BRITJ": "Itajai",
    "BRBZS": "Buzios",
    "BRABZ": "Buzios",
    "BRANG": "Angra dos Reis",
    "BRILB": "Ilhabela",
    "BRCPB": "Copacabana",
    "BRPOB": "Porto Belo",
    "BRBLN": "Balneario Camboriu",
    "UYMVD": "Montevideo",
    "UYPDE": "Punta del Este",

    # NORTHERN EUROPE
    "GBSOU": "Southampton",
    "GBDVR": "Dover",
    "GBLIV": "Liverpool",
    "GBBLF": "Belfast",
    "GBGRK": "Greenock",
    "GBSQF": "South Queensferry",
    "GBIVG": "Invergordon",
    "GBKWL": "Kirkwall",
    "GBLER": "Lerwick",
    "GBSTW": "Stornoway",
    "IECIE": "Cork",
    "IEDUN": "Dun Laoghaire",
    "DEHAM": "Hamburg",
    "DEWAR": "Warnemunde",
    "DEKIL": "Kiel",
    "NLRTM": "Rotterdam",
    "NLAMS": "Amsterdam",
    "BEZEE": "Zeebrugge",
    "DKCPH": "Copenhagen",
    "DKAAR": "Aarhus",
    "DKRNN": "Ronne",
    "DKSKA": "Skagen",
    "SESTO": "Stockholm",
    "SEVBY": "Visby",
    "FIHEL": "Helsinki",
    "PLGDY": "Gdynia",
    "EERTL": "Tallinn",
    "LVRIX": "Riga",
    "LTKLJ": "Klaipeda",

    # NORWAY
    "NOSVG": "Stavanger",
    "NOBGO": "Bergen",
    "NOOSL": "Oslo",
    "NOALE": "Alesund",
    "NOTRH": "Tromso",
    "NOTOS": "Tromso",
    "NOTRD": "Trondheim",
    "NOHGD": "Haugesund",
    "NOKRS": "Kristiansand",
    "NOLYN": "Lyngdal",
    "NOSND": "Sandnes",
    "NOEID": "Eidfjord",
    "NOFLM": "Flam",
    "NOLYR": "Longyearbyen",
    "NOSVJ": "Vik",
    "NOHVN": "Honningsvag",
    "NOGER": "Geiranger",
    "NOHEL": "Hellesylt",
    "NOOLD": "Olden",
    "NOMOL": "Molde",
    "NONFE": "Nordfjordeid",
    "NOULV": "Ulvik",
    "NOMLY": "Maloy",

    # ICELAND
    "ISREY": "Reykjavik",
    "ISISA": "Isafjordur",
    "ISAKR": "Akureyri",
    "ISSEY": "Seydisfjordur",

    # PORTUGAL
    "PTLIS": "Lisbon",
    "PTFNC": "Funchal",
    "PTPDL": "Ponta Delgada",
    "PTLEI": "Leixoes",

    # CROATIA / ADRIATIC
    "HRSPU": "Split",
    "HRDBV": "Dubrovnik",
    "HRZAD": "Zadar",
    "MEKOT": "Kotor",

    # MIDDLE EAST
    "JOAQJ": "Aqaba",
    "QADOH": "Doha",
    "OMMCT": "Muscat",
    "BHBAH": "Bahrain",

    # AFRICA
    "ZACPT": "Cape Town",
    "ZADUR": "Durban",
    "MZMPM": "Maputo",
    "MZPOM": "Pomene",
    "MZPTI": "Portuguese Island",
    "MGNOS": "Nosy Be",
    "MGDIE": "Antsiranana",
    "SCVTX": "Port Victoria",
    "MUPLU": "Port Louis",
    "RELAP": "La Possession",
    "CVMIN": "Mindelo",
    "NAWVB": "Walvis Bay",
    "TNLAG": "La Goulette",

    # MEDITERRANEAN ISLANDS
    "MTMLA": "Valletta",
    "MTVAL": "Valletta",
    "MCMCM": "Monte Carlo",
    "CYLEM": "Limassol",
    "GIBGI": "Gibraltar",
    "GIB": "Gibraltar",

    # CANADA
    "CAHAL": "Halifax",
    "CAVIC": "Victoria",

    # AUSTRALIA / NEW ZEALAND
    "AUSYD": "Sydney",
    "AUMEL": "Melbourne",
    "AUCNS": "Cairns",
    "AUTSV": "Townsville",
    "NZAKL": "Auckland",
    "NZWLG": "Wellington",

    # PACIFIC
    "WSAPW": "Apia",
    "ASPPG": "Pago Pago",
    "FJSUV": "Suva",
}

# Coordinate dei porti usate per riconoscere la posizione fisica della nave.
PORT_GEOFENCES = {
    # ITALY
    "Genoa": (44.4072, 8.9340, 18),
    "Naples": (40.8401, 14.2594, 18),
    "Civitavecchia": (42.0924, 11.7957, 18),
    "Palermo": (38.1250, 13.3700, 18),
    "Messina": (38.1938, 15.5540, 15),
    "Catania": (37.5025, 15.0950, 15),
    "Bari": (41.1310, 16.8700, 15),
    "Trieste": (45.6500, 13.7700, 18),
    "Venice": (45.4400, 12.3200, 18),
    "Livorno": (43.5510, 10.3010, 15),
    "La Spezia": (44.1050, 9.8250, 15),
    "Cagliari": (39.2130, 9.1090, 15),

    # FRANCE / SPAIN
    "Marseille": (43.3150, 5.3660, 20),
    "Barcelona": (41.3540, 2.1650, 18),
    "Valencia": (39.4400, -0.3150, 18),
    "Palma de Mallorca": (39.5630, 2.6300, 18),

    # NORTHERN EUROPE
    "Zeebrugge": (51.3400, 3.2040, 15),
    "Southampton": (50.8950, -1.4050, 15),
    "Rotterdam": (51.9500, 4.1500, 25),
    "Hamburg": (53.5350, 9.9500, 20),

    # GREECE / TURKEY
    "Piraeus": (37.9400, 23.6350, 18),
    "Santorini": (36.3869, 25.4286, 22),
    "Katakolon": (37.6450, 21.3200, 12),
    "Heraklion": (35.3430, 25.1450, 15),
    "Corfu": (39.6250, 19.9050, 15),
    "Istanbul": (41.0200, 28.9900, 20),
    "Izmir": (38.4167, 27.1500, 18),
    "Kusadasi": (37.8600, 27.2550, 15),

    # MEDITERRANEAN / PORTUGAL
    "Valletta": (35.8950, 14.5100, 15),
    "Lisbon": (38.7000, -9.1400, 18),
    "Funchal": (32.6450, -16.9100, 15),

    # NORTH AMERICA / CARIBBEAN
    "Miami": (25.7780, -80.1770, 18),
    "Port Canaveral": (28.4100, -80.6100, 18),
    "New York": (40.6900, -74.0300, 20),
    "Nassau": (25.0800, -77.3400, 15),
    "Ocean Cay": (25.4210, -79.2040, 12),

    # SOUTH AMERICA
    "Santos": (-23.9650, -46.3000, 20),
    "Buenos Aires": (-34.6000, -58.3700, 20),
    "Rio de Janeiro": (-22.8900, -43.1800, 20),
}

# salvataggio ultime pos in file json
live  = {}
known = {}

KNOWN_FILE = FOLDER / "fleet_known_positions.json"

def load_known():
    global known
    if KNOWN_FILE.exists():
        try:
            known = json.loads(KNOWN_FILE.read_text(encoding="utf-8"))
            print(f"  → caricate {len(known)} posizioni note da file")
        except Exception as e:
            print(f"  → impossibile leggere {KNOWN_FILE.name}: {e}")

def save_known():
    try:
        KNOWN_FILE.write_text(json.dumps(known, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"  → impossibile salvare known: {e}")

# codice del porto in nome
def port_name(raw):
    raw = (raw or "").strip().upper()
    if not raw:
        return ""

    # AIS può inviare rotte come ITGOA>ESLPA: la destinazione è l'ultimo codice.
    route_parts = [part.strip() for part in raw.replace("-", ">").split(">") if part.strip()]
    candidate = route_parts[-1] if route_parts else raw
    code = candidate[:5]

    if code in PORTS:
        return PORTS[code]

    print(f"  → PORTO NON ANCORA MAPPATO DA AGGIUNGERE: {raw}")
    return candidate.title()

# porto fisico vs navigazione
def distance_km(lat1, lon1, lat2, lon2):
    from math import radians, sin, cos, sqrt, atan2

    earth_radius_km = 6371.0

    d_lat = radians(lat2 - lat1)
    d_lon = radians(lon2 - lon1)

    a = (
        sin(d_lat / 2) ** 2
        + cos(radians(lat1))
        * cos(radians(lat2))
        * sin(d_lon / 2) ** 2
    )

    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    return earth_radius_km * c

def location(p):
    if not p:
        return ""

    lat = p.get("lat")
    lon = p.get("lon")
    sog = float(p.get("sog", 0) or 0)

    if lat is not None and lon is not None:
        nearest_port = None
        nearest_distance = None

        for port, (port_lat, port_lon, radius_km) in PORT_GEOFENCES.items():
            distance = distance_km(
                float(lat),
                float(lon),
                port_lat,
                port_lon
            )

            if distance <= radius_km:
                if nearest_distance is None or distance < nearest_distance:
                    nearest_port = port
                    nearest_distance = distance

        if nearest_port:
            return nearest_port

    # Non è vicino a un porto conosciuto e sta navigando
    if sog > 1:
        return "At Sea"
    return "In Port"

def darken(c, n=50):
    h = c.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "#{:02x}{:02x}{:02x}".format(max(0, r-n), max(0, g-n), max(0, b-n))

# icone nave MSC o EXPLORA (più piccole, luxury)
def ship_icon(color, luxury, offline=False):
    outline = "#132033"
    detail = "#ffffff"
    op = ' opacity="0.55"' if offline else ''

    w, h = (36, 21) if luxury else (48, 28)

    if luxury:
        body = (
            f'<path d="M4 36 L4 30 L8 30 L8 26 Q8 24 10 24 '
            f'L70 24 L76 24 Q80 24 84 27 L90 30 L92 30 '
            f'L92 34 Q70 40 48 40 Q24 40 4 36 Z" '
            f'fill="{color}" stroke="{outline}" stroke-width="2.8"/>'

            f'<rect x="10" y="16" width="55" height="8" rx="1" '
            f'fill="{detail}" stroke="{outline}" stroke-width="1"/>'

            f'<rect x="30" y="10" width="7" height="7" rx="1" '
            f'fill="{outline}"/>'

            f'<rect x="42" y="10" width="7" height="7" rx="1" '
            f'fill="{outline}"/>'

            f'<line x1="33" y1="5" x2="33" y2="10" '
            f'stroke="{outline}" stroke-width="1.6"/>'

            f'<line x1="45" y1="5" x2="45" y2="10" '
            f'stroke="{outline}" stroke-width="1.6"/>'

            f'<circle cx="20" cy="20" r="1.3" fill="{outline}"/>'
            f'<circle cx="27" cy="20" r="1.3" fill="{outline}"/>'
            f'<circle cx="55" cy="20" r="1.3" fill="{outline}"/>'

            f'<line x1="4" y1="36" x2="92" y2="34" '
            f'stroke="#dc2626" stroke-width="1.6"/>'
        )

    else:
        body = (
            f'<path d="M8 40 L8 34 Q8 32 10 32 L18 32 L18 24 '
            f'Q18 22 20 22 L30 22 L30 18 Q30 16 32 16 '
            f'L52 16 Q54 16 54 18 L54 22 L58 22 L58 16 '
            f'Q58 14 60 14 L66 14 Q68 14 68 16 L68 22 '
            f'L76 22 Q80 22 82 25 L88 32 L90 32 '
            f'Q92 32 92 34 L92 40 Q70 46 50 46 '
            f'Q28 46 8 40 Z" '
            f'fill="{color}" stroke="{outline}" stroke-width="3"/>'

            f'<rect x="19" y="23" width="36" height="8" rx="1.5" '
            f'fill="{detail}" stroke="{outline}" stroke-width="1"/>'

            f'<rect x="59" y="14.5" width="8" height="7" rx="1" '
            f'fill="{outline}"/>'

            f'<path d="M62 10 L64 14.5 L60 14.5 Z" '
            f'fill="{outline}"/>'

            f'<circle cx="25" cy="27" r="1.35" fill="{outline}"/>'
            f'<circle cx="33" cy="27" r="1.35" fill="{outline}"/>'
            f'<circle cx="41" cy="27" r="1.35" fill="{outline}"/>'
            f'<circle cx="49" cy="27" r="1.35" fill="{outline}"/>'

            f'<line x1="11" y1="37" x2="86" y2="37" '
            f'stroke="{detail}" stroke-width="1.4" opacity="0.85"/>'
        )

    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{w}" height="{h}" '
        f'viewBox="0 0 100 52"{op}>'
        f'{body}'
        f'</svg>'
    )

# sidebar/legenda
def sidebar():
    rows = []

    for cls in CLASS_ORDER:
        ships = [(m, v) for m, v in FLEET.items() if v[1] == cls]
        if not ships:
            continue

        rows.append(f'<div class="cls">{cls}</div>')

        for mmsi, (name, _, color, _) in ships:
            online = mmsi in live
            p = known.get(mmsi)
            loc = location(p) if p else ""
            spd = (
                f'{float(p.get("sog", 0) or 0):.0f} kn'
                if online and p
                else ""
            )
            dot_style = (
                f'background:{color}'
                + ('' if online else ';opacity:0.45')
            )
            ncls = 'n on' if online else ('n st' if p else 'n')
            fly = f'onclick="fly(\'{mmsi}\')"' if p else ''

            if online:
                badge = (
                    '<span class="b as">At Sea</span>'
                    if loc == "At Sea"
                    else (
                        f'<span class="b ip" title="{loc}">'
                        f'{loc[:16]}{"…" if len(loc) > 16 else ""}'
                        f'</span>'
                    )
                )
            elif p:
                ts = str(p.get("timestamp", ""))
                last_date = ""

                if len(ts) >= 10:
                    try:
                        _, month, day = ts[:10].split("-")
                        last_date = f"{day}/{month}"
                    except (ValueError, IndexError):
                        last_date = ""

                last_position = loc
                badge_text = f"Offline {last_date}" if last_date else "Offline"
                if last_position:
                    badge_text += f" · {last_position}"

                badge = (
                    f'<span class="b of" title="{badge_text}">'
                    f'{badge_text}'
                    f'</span>'
                )
            else:
                badge = ''

            rows.append(
                f'<div class="row" data-mmsi="{mmsi}" {fly}>'
                f'<div class="dot" style="{dot_style}"></div>'
                f'<span class="{ncls}">{name}</span>'
                f'{badge}'
                f'<span class="spd">{spd}</span>'
                f'</div>'
            )

    return "\n".join(rows)

#popup marker dettaglio nave su mappa
def markers():
    lines = []
    done  = set()
    for mmsi, p in live.items():
        v = FLEET.get(mmsi)
        if not v or p.get("lat") is None:
            continue
        name, cls, color, lux = v
        done.add(mmsi)
        svg  = ship_icon(color, lux).replace("`", "'").replace("\\", "\\\\")
        dest = port_name(p.get("destination", "")) or "—"
        loc  = location(p)
        spd  = float(p.get("sog", 0) or 0)
        hdg  = float(p.get("heading", 0) or 0)
        ts   = str(p.get("timestamp", ""))
        lat, lon = p["lat"], p["lon"]
        pop = (f'<div style=\\"min-width:200px\\">'
               f'<b style=\\"color:{color};font-size:13px\\">{name}</b>'
               f'<div style=\\"font-size:9px;color:#9ab4cc;letter-spacing:1px;margin:3px 0 10px\\">{cls} CLASS</div>'
               f'<table style=\\"width:100%;font-size:11px;border-collapse:collapse\\">'
               f'<tr><td style=\\"color:#9ab4cc\\">Current position</td><td style=\\"text-align:right;font-weight:600\\">{loc}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">AIS destination</td><td style=\\"text-align:right;font-weight:600\\">{dest}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Speed</td><td style=\\"text-align:right;font-weight:600\\">{spd:.1f} kn</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Heading</td><td style=\\"text-align:right;font-weight:600\\">{hdg:.0f}°</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Lat / Lon</td><td style=\\"text-align:right;font-weight:600\\">{lat:.4f}, {lon:.4f}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Updated</td><td style=\\"text-align:right;font-weight:600\\">{ts[:16]}</td></tr>'
               f'</table></div>')
        lines.append(f'mark({lat},{lon},`{svg}`,"{mmsi}","{pop}",false,{"true" if lux else "false"});')

    for mmsi, p in known.items():
        if mmsi in done or p.get("lat") is None:
            continue
        v = FLEET.get(mmsi)
        if not v:
            continue
        name, cls, color, lux = v
        svg  = ship_icon(color, lux, offline=True).replace("`", "'").replace("\\", "\\\\")
        dest = port_name(p.get("destination", "")) or "—"
        loc = location(p)
        spd  = float(p.get("sog", 0) or 0)
        hdg  = float(p.get("heading", 0) or 0)
        ts   = str(p.get("timestamp", ""))
        lat, lon = p["lat"], p["lon"]
        pop = (f'<div style=\\"min-width:200px\\">'
               f'<b style=\\"color:{color};opacity:0.6;font-size:13px\\">{name}</b>'
               f'<div style=\\"font-size:9px;color:#9ab4cc;letter-spacing:1px;margin:3px 0 10px\\">'
               f'{cls} CLASS &nbsp;·&nbsp; <span style=\\"color:#f59e0b\\">Offline</span></div>'
               f'<table style=\\"width:100%;font-size:11px;border-collapse:collapse\\">'
               f'<tr><td style=\\"color:#9ab4cc\\">Last position</td><td style=\\"text-align:right;font-weight:600\\">{loc}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">AIS destination</td><td style=\\"text-align:right;font-weight:600\\">{dest}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Speed</td><td style=\\"text-align:right;font-weight:600\\">{spd:.1f} kn</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Heading</td><td style=\\"text-align:right;font-weight:600\\">{hdg:.0f}°</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Lat / Lon</td><td style=\\"text-align:right;font-weight:600\\">{lat:.4f}, {lon:.4f}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Last Update</td><td style=\\"text-align:right;color:#f59e0b;font-weight:600\\">{ts[:16]}</td></tr>'
               f'</table></div>')
        lines.append(f'mark({lat},{lon},`{svg}`,"{mmsi}","{pop}",true,{"true" if lux else "false"});')

    return "\n".join(lines)

def write_github_page(html_content: str):
    output_file = REPO_DIR / "index.html"
    output_file.write_text(html_content, encoding="utf-8")
    print(f"  → generato {output_file.name}")

def write_html():
    now = datetime.now().strftime("%H:%M:%S")
    online = len(live)
    total = len(FLEET)
    n_off = sum(1 for m in FLEET if m not in live and m in known)
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="{REFRESH * 60}">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MSC & Explora Fleet Tracker</title>
<link rel="stylesheet" href="https://unpkg.com/maplibre-gl@5.6.1/dist/maplibre-gl.css">
<style>
* {{ box-sizing:border-box; margin:0; padding:0; font-family:Verdana,sans-serif; }}
body {{ background:#d8e8f0; height:100vh; display:flex; flex-direction:column; overflow:hidden; }}
header {{ background:linear-gradient(135deg,#0a1628,#122444,#0d1e3a); border-bottom:1px solid #1e3d70; padding:12px 24px 18px; display:flex; align-items:center; justify-content:space-between; flex-shrink:0; box-shadow:0 4px 16px rgba(0,0,0,.4); }}
.title {{ text-align:center; flex:1; }}
.title h1 {{ font-size:19px; font-weight:700; letter-spacing:2px; color:#fff; text-transform:uppercase; }}
.title p {{ font-size:13px; color:#7aaddb; letter-spacing:3px; text-transform:uppercase; margin-top:4px; }}
.info {{ display:flex; align-items:center; gap:20px; font-size:13px; color:#7aaddb; min-width:260px; justify-content:flex-end; }}
.dot {{ width:8px; height:8px; border-radius:50%; background:#22c55e; animation:pulse 2s infinite; display:inline-block; margin-right:5px; }}
@keyframes pulse {{ 0%,100% {{ opacity:1; box-shadow:0 0 0 0 rgba(34,197,94,.4); }} 50% {{ opacity:.5; box-shadow:0 0 0 5px rgba(34,197,94,0); }} }}
.on {{ color:#22c55e; font-weight:700; }}
.off {{ color:#f59e0b; font-weight:700; }}
#map {{ flex:1; min-height:0; }}
.panel {{ position:absolute; top:88px; right:14px; width:310px; background:rgba(255,255,255,.97); border:1px solid #b8cedd; border-radius:10px; box-shadow:0 4px 20px rgba(0,0,0,.13); z-index:10; display:flex; flex-direction:column; max-height:calc(100vh - 104px); overflow:hidden; }}
.panel-title {{ font-size:11px; font-weight:700; letter-spacing:2px; text-transform:uppercase; color:#5a7a9a; padding:9px 14px 7px; border-bottom:1px solid #d0e0ec; background:#eef4fa; }}
.list {{ flex:1; display:flex; flex-direction:column; overflow:hidden; }}
.cls {{ font-size:9px; font-weight:700; letter-spacing:2px; text-transform:uppercase; color:#9ab4cc; padding:4px 14px 2px; background:#f5f9fd; }}
.row {{ display:flex; align-items:center; gap:7px; padding:3px 11px; cursor:pointer; border-bottom:1px solid #eef3f8; transition:background .12s; }}
.row:hover {{ background:#e4eef8; }}
.row.selected-ship {{ background:#dbeafe; outline:2px solid #2563eb; outline-offset:-2px; }}
.row.selected-ship .n {{ color:#0a2450; font-weight:700; }}
.n {{ font-size:12px; color:#9ab4cc; flex:1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.n.on {{ color:#1a3050; font-weight:700; }}
.n.st {{ color:#7890a8; font-style:italic; }}
.b {{ font-size:10px; padding:1px 6px; border-radius:3px; flex-shrink:0; font-weight:700; max-width:150px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.as {{ background:#dbeafe; color:#1d4ed8; }}
.ip {{ background:#dcfce7; color:#166534; }}
.of {{ background:#fef3c7; color:#92400e; }}
.spd {{ font-size:10px; color:#9ab4cc; flex-shrink:0; }}
.ship-marker {{ cursor:pointer; transform-origin:center; }}
.maplibregl-popup-content {{ background:#fff; border:1px solid #b8cedd; border-radius:10px; font-family:Verdana,sans-serif; box-shadow:0 8px 24px rgba(0,0,0,.18); padding:12px; }}
.maplibregl-popup-tip {{ border-top-color:#fff !important; }}
.remote-help {{ position:absolute; left:55px; bottom:18px; z-index:10; background:rgba(10,22,40,.88); color:#dbeafe; border:1px solid #31547a; border-radius:6px; padding:6px 10px; font-size:10px; letter-spacing:.5px; box-shadow:0 2px 8px rgba(0,0,0,.25); pointer-events:none; }}
.remote-help strong {{ color:#fff; }}

@media (min-width:2560px) and (min-height:1400px) {{
  header {{ padding:20px 38px 28px; }}
  .title h1 {{ font-size:30px; letter-spacing:3px; }}
  .title p {{ font-size:20px; letter-spacing:5px; margin-top:7px; }}
  .info {{ gap:30px; font-size:19px; min-width:480px; }}
  .dot {{ width:12px; height:12px; margin-right:8px; }}
  .panel {{ top:142px; right:22px; width:500px; max-height:calc(100vh - 168px); border-radius:14px; }}
  .panel-title {{ font-size:17px; padding:14px 20px 11px; }}
  .cls {{ font-size:14px; padding:6px 20px 4px; }}
  .row {{ gap:10px; padding:5px 17px; }}
  .n {{ font-size:18px; }}
  .b {{ font-size:15px; padding:2px 9px; max-width:235px; }}
  .spd {{ font-size:15px; }}
  .remote-help {{ left:85px; bottom:28px; padding:10px 16px; font-size:15px; border-radius:9px; }}
  .maplibregl-ctrl-group button {{ width:44px; height:44px; }}
  .maplibregl-popup-content {{ font-size:16px; padding:18px; }}
}}
</style>
</head>
<body>
<header>
  <div style="min-width:140px"></div>
  <div class="title"><h1>MSC Cruises &amp; Explora Journeys</h1><p>Fleet Live Tracker</p></div>
  <div class="info"><span>Online: <span class="on">{online}</span>/{total} &nbsp;·&nbsp; Last known: <span class="off">{n_off}</span></span><span><span class="dot"></span>{now} &nbsp;·&nbsp; Refresh every {REFRESH} min</span></div>
</header>
<div id="map"></div>
<div class="remote-help">REMOTE · <strong>NAV</strong> · OK TO SELECT SHIP</div>
<div class="panel"><div class="panel-title">FLEET (Last known positions)</div><div class="list">{sidebar()}</div></div>
<script src="https://unpkg.com/maplibre-gl@5.6.1/dist/maplibre-gl.js"></script>
<script>
const map = new maplibregl.Map({{ container:'map', style:'https://tiles.openfreemap.org/styles/liberty', center:[10,30], zoom:3, attributionControl:true }});
map.addControl(new maplibregl.NavigationControl({{showCompass:false}}), 'top-left');
const mkrs = {{}};
const fleetBounds = new maplibregl.LngLatBounds();
function shipScale(lux) {{ const z=map.getZoom(); const zoomScale=z<=3?1.50:z<=4?1.30:z<=5?1.15:1; const screenScale=window.innerWidth>=2560?1.45:1; const scale=zoomScale*screenScale; return {{w:Math.round((lux?36:48)*scale),h:Math.round((lux?21:28)*scale)}}; }}
function sizeShipMarker(marker) {{ const d=shipScale(marker._shipLux); marker.getElement().style.width=d.w+'px'; marker.getElement().style.height=d.h+'px'; }}
function mark(lat,lon,svg,mmsi,popup,offline,lux) {{
  const el=document.createElement('div');
  el.className='ship-marker';
  el.innerHTML=svg;
  el.style.opacity=offline?'0.78':'1';
  el.style.display='flex';
  el.style.alignItems='center';
  el.style.justifyContent='center';
  el.querySelector('svg').style.width='100%';
  el.querySelector('svg').style.height='100%';
  const pop=new maplibregl.Popup({{offset:22,maxWidth:'290px',closeButton:true}}).setHTML(popup);
  const marker=new maplibregl.Marker({{element:el,anchor:'center'}}).setLngLat([lon,lat]).setPopup(pop).addTo(map);
  marker._shipLux=lux;
  mkrs[mmsi]=marker;
  fleetBounds.extend([lon,lat]);
  sizeShipMarker(marker);
}}
{markers()}
function resizeShipMarkers() {{ Object.values(mkrs).forEach(sizeShipMarker); }}
map.on('zoom', resizeShipMarkers);
map.on('load', function() {{
  if (!fleetBounds.isEmpty()) {{ const rightPadding=window.innerWidth>=2560?535:325; map.fitBounds(fleetBounds,{{padding:{{top:20,bottom:20,left:25,right:rightPadding}},maxZoom:5,duration:0}}); }}
}});
function fly(id) {{ const marker=mkrs[id]; if (!marker) return; map.flyTo({{center:marker.getLngLat(),zoom:Math.max(map.getZoom(),6),essential:true}}); marker.togglePopup(); }}
document.querySelectorAll('.row[data-mmsi]').forEach(row=>row.addEventListener('click',()=>fly(row.dataset.mmsi)));
let remoteMode='NAV';
let selectedShipIndex=-1;
const selectableShips=Object.keys(mkrs);
function updateRemoteHelp() {{ const help=document.querySelector('.remote-help'); help.innerHTML=remoteMode==='NAV'?'REMOTE · <strong>NAV</strong> · OK TO SELECT SHIP':'REMOTE · <strong>SHIPS</strong> · ↑↓ SELECT · ←→ ZOOM · OK DETAILS'; }}
function clearShipSelection() {{ document.querySelectorAll('.row').forEach(row=>row.classList.remove('selected-ship')); }}
function selectShip(index) {{
  if (!selectableShips.length) return;
  selectedShipIndex=(index+selectableShips.length)%selectableShips.length;
  const mmsi=selectableShips[selectedShipIndex];
  const marker=mkrs[mmsi];
  clearShipSelection();
  const row=document.querySelector('.row[data-mmsi="'+mmsi+'"]');
  if (row) {{ row.classList.add('selected-ship'); row.scrollIntoView({{behavior:'smooth',block:'nearest'}}); }}
  map.easeTo({{center:marker.getLngLat(),duration:450}});
}}
document.addEventListener('keydown',function(e) {{
  const key=e.key;
  if (remoteMode==='NAV') {{
    if (key==='ArrowUp') map.panBy([0,-180]);
    else if (key==='ArrowDown') map.panBy([0,180]);
    else if (key==='ArrowLeft') map.panBy([-180,0]);
    else if (key==='ArrowRight') map.panBy([180,0]);
    else if (key==='Enter') {{ remoteMode='SHIPS'; selectShip(0); updateRemoteHelp(); }}
    else return;
  }} else {{
    if (key==='ArrowUp') selectShip(selectedShipIndex-1);
    else if (key==='ArrowDown') selectShip(selectedShipIndex+1);
    else if (key==='ArrowLeft') map.zoomOut();
    else if (key==='ArrowRight') map.zoomIn();
    else if (key==='Enter') {{ const marker=mkrs[selectableShips[selectedShipIndex]]; if (marker) {{ map.flyTo({{center:marker.getLngLat(),zoom:Math.max(map.getZoom(),6),essential:true}}); marker.togglePopup(); }} }}
    else if (key==='Escape'||key==='Backspace') {{ remoteMode='NAV'; selectedShipIndex=-1; clearShipSelection(); updateRemoteHelp(); }}
    else return;
  }}
  e.preventDefault();
}});
</script>
</body>
</html>"""
    write_github_page(html)
    save_known()
    print(f"[{datetime.now():%H:%M:%S}] aggiornato - {online}/{total} online, {n_off} offline con posizione nota")

def write_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fleet"
    ws.append(["Name", "MMSI", "Class", "Status", "Location", "Lat", "Lon", "Speed_kn", "Heading", "Destination", "LastUpdate"])
    for mmsi, p in live.items():
        name, cls, *_ = FLEET.get(mmsi, ("?", "-", "#888", False))
        dest = port_name(p.get("destination", ""))
        ws.append([name, mmsi, cls, "Online", location(p), p.get("lat"), p.get("lon"),
                   p.get("sog"), p.get("heading"), dest, p.get("timestamp")])
    for mmsi, p in known.items():
        if mmsi in live:
            continue
        name, cls, *_ = FLEET.get(mmsi, ("?", "-", "#888", False))
        dest = port_name(p.get("destination", ""))
        ws.append([name, mmsi, cls, "Offline (last known)", location(p), p.get("lat"), p.get("lon"),
                   p.get("sog"), p.get("heading"), dest, p.get("timestamp")])
    wb.save(FOLDER / "fleet_positions.xlsx")

async def run():
    url = "wss://stream.aisstream.io/v0/stream"
    started_at = datetime.now(timezone.utc)

    try:
        async with websockets.connect(
            url,
            open_timeout=60,
            ping_interval=20,
            ping_timeout=20,
            close_timeout=10
        ) as ws:

            subscription = {
                "APIKey": API_KEY,
                "BoundingBoxes": [[[-90, -180], [90, 180]]],
                "FiltersShipMMSI": list(FLEET.keys()),
                "FilterMessageTypes": [
                    "PositionReport",
                    "ShipStaticData"
                ]
            }

            await ws.send(json.dumps(subscription))

            print(
                f"[{datetime.now():%H:%M:%S}] connesso - "
                f"{len(FLEET)} navi - raccolta per "
                f"{COLLECTION_MINUTES} minuti"
            )

            while True:
                elapsed = (
                    datetime.now(timezone.utc) - started_at
                ).total_seconds()

                remaining = COLLECTION_MINUTES * 60 - elapsed

                if remaining <= 0:
                    break

                try:
                    raw = await asyncio.wait_for(
                        ws.recv(),
                        timeout=min(30, remaining)
                    )
                except asyncio.TimeoutError:
                    continue

                try:
                    msg = json.loads(raw)
                except json.JSONDecodeError:
                    print(f"Messaggio AISStream non JSON: {raw}")
                    continue

                if msg.get("MessageType") == "SubscriptionConfirmation":
                    print("  → sottoscrizione AISStream confermata")
                    continue

                metadata = msg.get("MetaData")

                if not metadata:
                    print(
                        f"  → messaggio AISStream ignorato: "
                        f"{msg.get('MessageType', 'tipo sconosciuto')}"
                    )
                    continue

                mmsi = str(metadata.get("MMSI", ""))

                if not mmsi or mmsi not in FLEET:
                    continue

                message_type = msg.get("MessageType")
                message = msg.get("Message", {})

                if message_type == "PositionReport":
                    pr = message.get("PositionReport")

                    if not pr:
                        print(
                            "PositionReport senza dati: "
                            f"{msg}"
                        )
                        continue

                    lat = pr.get("Latitude")
                    lon = pr.get("Longitude")

                    if lat is None or lon is None:
                        continue

                    position = {
                        **known.get(mmsi, {}),
                        "lat": lat,
                        "lon": lon,
                        "sog": pr.get("Sog", 0),
                        "heading": (
                            pr.get("TrueHeading")
                            or pr.get("Cog", 0)
                        ),
                        "status": pr.get(
                            "NavigationalStatus",
                            0
                        ),
                        "timestamp": metadata.get(
                            "time_utc",
                            ""
                        ),
                    }

                    live[mmsi] = position
                    known[mmsi] = position

                    print(
                        f"  {FLEET[mmsi][0]}: "
                        f"{lat:.3f}, {lon:.3f}"
                    )

                elif message_type == "ShipStaticData":
                    static_data = message.get(
                        "ShipStaticData",
                        {}
                    )

                    destination = str(
                        static_data.get("Destination", "") or ""
                    ).strip()

                    # Non cancellare una destinazione valida già nota
                    # quando AIS invia ShipStaticData con Destination vuota.
                    if destination:
                        known.setdefault(mmsi, {})["destination"] = destination

                        if mmsi in live:
                            live[mmsi]["destination"] = destination

    except Exception as error:
        print(
            "Connessione AISStream terminata con errore: "
            f"{type(error).__name__}: {error}"
        )

    finally:
        for mmsi, position in live.items():
            known[mmsi] = position

        save_known()
        write_excel()
        write_html()

        print(
            f"[{datetime.now():%H:%M:%S}] "
            f"esecuzione completata - "
            f"{len(live)}/{len(FLEET)} navi ricevute "
            "in tempo reale"
        )

if __name__ == "__main__":
    print(f"MSC & Explora Fleet Tracker — {len(FLEET)} navi — refresh every {REFRESH} min")
    load_known()
    asyncio.run(run())
