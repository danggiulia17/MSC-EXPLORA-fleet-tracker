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
    "Savona": (44.3070, 8.4830, 15),
    "La Spezia": (44.1050, 9.8250, 15),
    "Livorno": (43.5510, 10.3010, 15),
    "Civitavecchia": (42.0924, 11.7957, 18),
    "Naples": (40.8401, 14.2594, 18),
    "Palermo": (38.1250, 13.3700, 18),
    "Messina": (38.1938, 15.5540, 15),
    "Catania": (37.5025, 15.0950, 15),
    "Syracuse": (37.0650, 15.2900, 15),
    "Reggio Calabria": (38.1150, 15.6500, 15),
    "Bari": (41.1310, 16.8700, 15),
    "Brindisi": (40.6450, 17.9450, 15),
    "Ancona": (43.6200, 13.5050, 15),
    "Trieste": (45.6500, 13.7700, 18),
    "Venice": (45.4400, 12.3200, 18),
    "Ravenna": (44.4900, 12.2950, 18),
    "Cagliari": (39.2130, 9.1090, 15),
    "Olbia": (40.9250, 9.5050, 15),
    "Milazzo": (38.2200, 15.2400, 12),

    # FRANCE / MONACO / CORSICA
    "Marseille": (43.3150, 5.3660, 20),
    "Toulon": (43.1200, 5.9300, 15),
    "Cannes": (43.5500, 7.0150, 15),
    "Nice": (43.6950, 7.2850, 15),
    "Monte Carlo": (43.7350, 7.4250, 12),
    "Ajaccio": (41.9200, 8.7400, 15),
    "Bastia": (42.7050, 9.4550, 15),
    "Le Havre": (49.4900, 0.1050, 20),
    "Cherbourg": (49.6450, -1.6200, 15),
    "Brest": (48.3850, -4.4950, 18),
    "La Rochelle": (46.1550, -1.1550, 15),
    "Saint-Nazaire": (47.2750, -2.2050, 15),

    # SPAIN - MEDITERRANEAN / BALEARIC ISLANDS
    "Barcelona": (41.3540, 2.1650, 18),
    "Tarragona": (41.0900, 1.2250, 15),
    "Valencia": (39.4400, -0.3150, 18),
    "Alicante": (38.3400, -0.4800, 15),
    "Cartagena Spain": (37.5950, -0.9850, 15),
    "Malaga": (36.7100, -4.4200, 18),
    "Cadiz": (36.5350, -6.2950, 18),
    "Palma de Mallorca": (39.5630, 2.6300, 18),
    "Mahon": (39.8900, 4.2700, 15),
    "Ibiza": (38.9100, 1.4400, 15),

    # PORTUGAL / ATLANTIC ISLANDS
    "Lisbon": (38.7000, -9.1400, 18),
    "Leixoes": (41.1850, -8.7050, 15),
    "Funchal": (32.6450, -16.9100, 15),
    "Ponta Delgada": (37.7350, -25.6650, 15),

    # MALTA / CYPRUS / GIBRALTAR
    "Valletta": (35.8950, 14.5100, 15),
    "Limassol": (34.6700, 33.0450, 18),
    "Gibraltar": (36.1400, -5.3550, 15),

    # GREECE
    "Piraeus": (37.9400, 23.6350, 18),
    "Santorini": (36.3869, 25.4286, 22),
    "Mykonos": (37.4500, 25.3300, 18),
    "Heraklion": (35.3430, 25.1450, 15),
    "Rhodes": (36.4450, 28.2300, 15),
    "Corfu": (39.6250, 19.9050, 15),
    "Katakolon": (37.6450, 21.3200, 12),
    "Kefalonia": (38.1800, 20.4900, 15),
    "Syros": (37.4400, 24.9450, 15),
    "Thessaloniki": (40.6300, 22.9350, 18),
    "Volos": (39.3550, 22.9450, 15),

    # TURKEY
    "Istanbul": (41.0200, 28.9900, 20),
    "Izmir": (38.4167, 27.1500, 18),
    "Kusadasi": (37.8600, 27.2550, 15),
    "Bodrum": (37.0350, 27.4300, 15),
    "Marmaris": (36.8500, 28.2750, 15),
    "Antalya": (36.8400, 30.6100, 20),
    "Mersin": (36.7900, 34.6400, 20),

    # CROATIA / MONTENEGRO / SLOVENIA / ADRIATIC
    "Split": (43.5050, 16.4400, 15),
    "Dubrovnik": (42.6600, 18.0850, 15),
    "Zadar": (44.1150, 15.2250, 15),
    "Kotor": (42.4250, 18.7700, 15),
    "Rijeka": (45.3250, 14.4450, 15),
    "Koper": (45.5550, 13.7400, 12),

    # NORTH AFRICA / EASTERN MEDITERRANEAN
    "Tangier": (35.7900, -5.8100, 18),
    "Tanger Med": (35.8900, -5.5000, 18),
    "Casablanca": (33.6050, -7.6200, 20),
    "Agadir": (30.4250, -9.6150, 18),
    "La Goulette": (36.8100, 10.3050, 15),
    "Alexandria": (31.1900, 29.8800, 20),
    "Port Said": (31.2550, 32.3000, 20),
    "Sharm El-Sheikh": (27.8550, 34.2850, 18),
    "Aqaba": (29.4900, 34.9950, 15),

    # UNITED KINGDOM / IRELAND
    "Southampton": (50.8950, -1.4050, 15),
    "Dover": (51.1250, 1.3300, 15),
    "Liverpool": (53.4050, -3.0000, 18),
    "Belfast": (54.6250, -5.9000, 18),
    "Greenock": (55.9500, -4.7600, 15),
    "South Queensferry": (55.9900, -3.4000, 15),
    "Invergordon": (57.6900, -4.1700, 15),
    "Kirkwall": (58.9850, -2.9600, 15),
    "Lerwick": (60.1550, -1.1450, 15),
    "Stornoway": (58.2100, -6.3900, 15),
    "Cork": (51.8500, -8.3000, 18),
    "Dun Laoghaire": (53.2950, -6.1350, 15),

    # NORTHERN EUROPE
    "Zeebrugge": (51.3400, 3.2040, 15),
    "Rotterdam": (51.9500, 4.1500, 25),
    "Amsterdam": (52.3950, 4.9000, 18),
    "Hamburg": (53.5350, 9.9500, 20),
    "Kiel": (54.3250, 10.1400, 15),
    "Warnemunde": (54.1800, 12.0900, 15),
    "Copenhagen": (55.7000, 12.6000, 18),
    "Aarhus": (56.1550, 10.2350, 15),
    "Ronne": (55.1000, 14.7000, 15),
    "Skagen": (57.7200, 10.5900, 15),
    "Stockholm": (59.3300, 18.1100, 22),
    "Visby": (57.6350, 18.2850, 15),
    "Helsinki": (60.1600, 24.9550, 18),
    "Tallinn": (59.4450, 24.7700, 15),
    "Riga": (57.0000, 24.1000, 20),
    "Klaipeda": (55.7100, 21.1300, 15),
    "Gdynia": (54.5300, 18.5500, 18),

    # NORWAY
    "Oslo": (59.9000, 10.7300, 18),
    "Bergen": (60.3950, 5.3200, 18),
    "Stavanger": (58.9700, 5.7300, 15),
    "Alesund": (62.4700, 6.1500, 15),
    "Tromso": (69.6500, 18.9650, 18),
    "Trondheim": (63.4400, 10.4000, 18),
    "Haugesund": (59.4100, 5.2700, 15),
    "Kristiansand": (58.1450, 8.0000, 15),
    "Eidfjord": (60.4650, 7.0700, 12),
    "Flam": (60.8600, 7.1150, 12),
    "Geiranger": (62.1050, 7.2050, 12),
    "Hellesylt": (62.0850, 6.8700, 12),
    "Olden": (61.8350, 6.8100, 12),
    "Nordfjordeid": (61.9050, 5.9900, 12),
    "Honningsvag": (70.9800, 25.9750, 15),
    "Longyearbyen": (78.2300, 15.6100, 18),
    "Molde": (62.7350, 7.1600, 15),
    "Maloy": (61.9350, 5.1150, 12),

    # ICELAND
    "Reykjavik": (64.1500, -21.9400, 18),
    "Isafjordur": (66.0700, -23.1250, 15),
    "Akureyri": (65.6800, -18.0900, 15),
    "Seydisfjordur": (65.2650, -14.0000, 15),

    # UNITED ARAB EMIRATES / PERSIAN GULF
    "Dubai": (25.2700, 55.2850, 20),
    "Abu Dhabi": (24.4950, 54.3700, 20),
    "Sir Bani Yas": (24.2800, 52.5900, 18),
    "Doha": (25.3000, 51.5500, 20),
    "Bahrain": (26.2000, 50.6100, 18),
    "Muscat": (23.6200, 58.5700, 20),
    "Khasab": (26.2050, 56.2500, 15),

    # SOUTH AFRICA / SOUTHERN AFRICA
    "Cape Town": (-33.9150, 18.4350, 22),
    "Durban": (-29.8700, 31.0450, 22),
    "Gqeberha": (-33.9650, 25.6300, 18),
    "Richards Bay": (-28.8000, 32.1000, 20),
    "Mossel Bay": (-34.1800, 22.1450, 18),
    "Walvis Bay": (-22.9500, 14.5000, 22),
    "Maputo": (-25.9700, 32.5800, 20),
    "Pomene": (-22.9100, 35.6000, 18),
    "Portuguese Island": (-26.0500, 32.9000, 15),
    "Nosy Be": (-13.4000, 48.2900, 18),
    "Antsiranana": (-12.2750, 49.2900, 18),
    "Port Victoria": (-4.6200, 55.4550, 15),
    "Port Louis": (-20.1600, 57.5000, 18),
    "La Possession": (-20.9300, 55.2900, 15),

    # MEXICO
    "Cozumel": (20.5110, -86.9500, 22),
    "Costa Maya": (18.7300, -87.6900, 18),
    "Cabo San Lucas": (22.8850, -109.9100, 20),
    "Puerto Vallarta": (20.6550, -105.2400, 18),
    "Progreso": (21.3000, -89.6700, 22),
    "Ensenada": (31.8550, -116.6250, 18),
    "Manzanillo": (19.0500, -104.3200, 20),
    "Acapulco": (16.8500, -99.9050, 18),

    # BAHAMAS / CARIBBEAN / GREATER ANTILLES
    "Nassau": (25.0800, -77.3400, 15),
    "Freeport": (26.5350, -78.7000, 18),
    "Ocean Cay": (25.4210, -79.2040, 12),
    "Miami": (25.7780, -80.1770, 18),
    "Port Canaveral": (28.4100, -80.6100, 18),
    "San Juan": (18.4600, -66.1000, 18),
    "Santo Domingo": (18.4600, -69.8900, 18),
    "La Romana": (18.4200, -68.9700, 18),
    "Puerto Plata": (19.8000, -70.7000, 18),
    "Amber Cove": (19.8350, -70.7250, 15),
    "Samana": (19.2000, -69.3350, 18),
    "Cabo Rojo": (17.9200, -71.6600, 15),
    "Falmouth": (18.4900, -77.6500, 15),
    "Montego Bay": (18.4550, -77.9400, 18),
    "Ocho Rios": (18.4100, -77.1050, 15),
    "George Town": (19.3000, -81.3830, 18),
    "Grand Turk": (21.4600, -71.1400, 15),
    "Roatan": (16.3150, -86.5300, 18),
    "Belize City": (17.5000, -88.1900, 18),

    # ANTILLES
    "Bridgetown": (13.1050, -59.6300, 15),
    "Castries": (14.0100, -61.0000, 15),
    "Kingstown": (13.1500, -61.2250, 15),
    "St. George's": (12.0500, -61.7500, 15),
    "Roseau": (15.3000, -61.3900, 15),
    "Pointe-a-Pitre": (16.2350, -61.5350, 15),
    "Fort-de-France": (14.6000, -61.0700, 15),
    "St. John's": (17.1200, -61.8500, 15),
    "Basseterre": (17.2950, -62.7250, 15),
    "Philipsburg": (18.0200, -63.0500, 15),
    "Charlotte Amalie": (18.3350, -64.9250, 15),
    "Road Town": (18.4250, -64.6150, 15),
    "Virgin Gorda": (18.4500, -64.4400, 15),
    "Kralendijk": (12.1500, -68.2750, 15),
    "Oranjestad": (12.5200, -70.0400, 15),
    "Willemstad": (12.1100, -68.9400, 15),
    "Port of Spain": (10.6500, -61.5150, 18),

    # CENTRAL AMERICA / COLOMBIA / PANAMA
    "Cartagena Colombia": (10.4000, -75.5300, 18),
    "Puerto Limon": (10.0000, -83.0350, 18),
    "Puntarenas": (9.9750, -84.8350, 18),
    "Colon": (9.3550, -79.9000, 18),
    "Panama Canal Atlantic": (9.3150, -79.9200, 18),
    "Panama Canal Pacific": (8.9500, -79.5650, 18),

    # BRAZIL
    "Santos": (-23.9650, -46.3000, 20),
    "Rio de Janeiro": (-22.8900, -43.1800, 20),
    "Salvador": (-12.9700, -38.5100, 20),
    "Maceio": (-9.6750, -35.7200, 18),
    "Ilheus": (-14.7800, -39.0300, 18),
    "Itajai": (-26.9050, -48.6600, 18),
    "Buzios": (-22.7550, -41.8850, 18),
    "Angra dos Reis": (-23.0100, -44.3150, 18),
    "Ilhabela": (-23.7800, -45.3550, 18),
    "Copacabana": (-22.9750, -43.1800, 15),
    "Porto Belo": (-27.1550, -48.5500, 15),
    "Balneario Camboriu": (-27.0000, -48.6100, 15),

    # ARGENTINA / URUGUAY
    "Buenos Aires": (-34.6000, -58.3700, 20),
    "Montevideo": (-34.9000, -56.2050, 18),
    "Punta del Este": (-34.9600, -54.9500, 18),

    # PERU
    "Callao": (-12.0550, -77.1500, 22),
    "Paracas": (-13.8350, -76.2500, 18),
    "Pisco": (-13.7100, -76.2200, 18),
    "Salaverry": (-8.2250, -78.9800, 18),
    "Matarani": (-17.0000, -72.1050, 18),
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
