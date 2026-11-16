"""MSC CRUISES & EXPLORA JOURNEYS FLEET TRACKER
Usa codici MMSI di ogni vessel per identificare posizione geografica, status e destinazione
Aggiornamento ogni 15 min, pubblica su GitHub Pages"""

import asyncio, json, websockets, openpyxl, os
from datetime import datetime, timezone
from pathlib import Path

API_KEY = os.environ["AISSTREAM_API_KEY"]
FOLDER = Path.cwd()
REFRESH = 15  # minuti
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
    # ITALIA
    "ITGOA": "Genova",
    "ITCIV": "Civitavecchia",
    "ITPAL": "Palermo",
    "ITNAP": "Napoli",
    "ITBRI": "Bari",
    "ITTRS": "Trieste",
    "ITVCE": "Venezia",
    "ITLIV": "Livorno",
    "ITANC": "Ancona",
    "ITMES": "Messina",
    "ITCAT": "Catania",
    "ITSPE": "La Spezia",
    "ITCAG": "Cagliari",
    "ITBDS": "Brindisi",
    "ITREG": "Reggio Calabria",
    "ITSIR": "Siracusa",
    "ITOLB": "Olbia",
    "ITMIL": "Milazzo",
    "ITSVN": "Savona",
    # SPAGNA
    "ESBCN": "Barcellona",
    "ESVLC": "Valencia",
    "ESPAL": "Palma di Maiorca",
    "ESMAH": "Mahon",
    "ESLPA": "Las Palmas",
    "ESTCI": "Santa Cruz de Tenerife",
    "ESALI": "Alicante",
    "ESCAD": "Cadice",
    "ESVGO": "Vigo",
    "ESGIJ": "Gijon",
    "ESLCG": "La Coruña",
    "ESBIO": "Bilbao",
    "ESIBZ": "Ibiza",
    "ESACE": "Arrecife di Lanzarote",
    "ESFUE": "Puerto del Rosario",
    "ESSCT": "Santa Cruz de Tenerife",
    "ESSCP": "Santa Cruz de La Palma",
    "ESTGN": "Tarragona",
    # FRANCIA
    "FRMAR": "Marsiglia",
    "FRTLN": "Tolone",
    "FRAJC": "Ajaccio",
    "FRLEH": "Le Havre",
    "FRLRH": "La Rochelle",
    "FRBRE": "Brest",
    "FRCFR": "Cherbourg",
    "FRCNL": "Cannes",
    "FRSTN": "St Nazaire",
    # GRECIA
    "GRPIR": "Pireo",
    "GRHER": "Heraklion",
    "GRRHO": "Rodi",
    "GRSAN": "Santorini",
    "GRJMK": "Mykonos",
    "GRCFU": "Corfù",
    "GRKAT": "Katakolon",
    "GRCEA": "Cefalonia",
    "GRSYR": "Syros",
    # TURCHIA
    "TRIST": "Istanbul",
    "TRIZM": "Smirne",
    "TRBOD": "Bodrum",
    "TRKUS": "Kusadasi",
    "TRMRM": "Marmaris",
    # MAROCCO
    "MABAS": "Tanger Med",
    "MAAGA": "Agadir",
    "MACAS": "Casablanca",
    "MATNG": "Tangeri",
    # EGITTO
    "EGPSD": "Port Said",
    "EGALY": "Alessandria",
    "EGSHK": "Sharm El-Sheikh",
    "EGSCN": "Canale di Suez Nord",
    "EGSCS": "Canale di Suez Sud",
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
    "JPMIY": "Miyako Is.",
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
    # AMERICA NORD
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
    "USICY": "Icy Strait / Hoonah",
    "USTRA": "Tracy Arm",
    # CARAIBI
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
    "DOSAM": "Samana Bay",
    "DOCBR": "Cabo Rojo",
    "DOAMB": "Amber Cove",
    "DOCAT": "Catalina Island",
    "HNRTN": "Isla de Roatan",
    "BZCZE": "Belize City",
    "CRCAL": "Puerto Limon",
    "CRPUN": "Puntarenas",
    "PAMIG": "Panama Canal (BLB)",
    "PACTB": "Panama Canal (CTB)",
    "PALON": "Colon",
    "JMFAL": "Falmouth",
    "JMMBJ": "Montego Bay",
    "JMOCH": "Ocho Rios",
    "KYGEO": "George Town",
    "TCGDT": "Grand Turk Island",
    "BSFPO": "Freeport",
    "BSNAS": "Nassau",
    "AGSTJ": "St John's",
    "DMROS": "Roseau",
    "GPPPT": "Pointe-à-Pitre",
    "MQFDF": "Fort de France",
    "VISTJ": "Charlotte Amalie",
    "VIRTW": "Road Town",
    "VIVGD": "Virgin Gorda",
    "SXPHI": "Philipsburg",
    "BQKRL": "Kralendijk",
    "AWORA": "Oranjestad",
    "CWWIL": "Willemstad",
    "GDSGT": "Saint George",
    "KNBST": "Basseterre",
    "MSOCN": "Ocean Cay",
    # AMERICA SUD
    "BRSSS": "Santos",
    "ARBUE": "Buenos Aires",
    "BRRIJ": "Rio de Janeiro",
    "BRSSZ": "Santos",
    "BRSAL": "Salvador",
    "BRMAO": "Maceió",
    "BRILS": "Ilheus",
    "BRITJ": "Itajai",
    "BRBZS": "Buzios",
    "BRANG": "Angra dos Reis",
    "BRILB": "Ilhabela",
    "BRCPB": "Copacabana",
    "BRPOB": "Porto Belo",
    "BRBLN": "Balneário Camboriú",
    "UYMVD": "Montevideo",
    "UYPDE": "Punta del Este",
    # EUROPA NORD
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
    "DEHAM": "Amburgo",
    "DEWAR": "Warnemünde",
    "DEKIL": "Kiel",
    "NLRTM": "Rotterdam",
    "NLAMS": "Amsterdam",
    "BEZEE": "Zeebrugge",
    "DKCPH": "Copenhagen",
    "DKAAR": "Aarhus",
    "DKRNN": "Ronne",
    "DKSKA": "Skagen",
    "SESTO": "Stoccolma",
    "SEVBY": "Visby",
    "FIHEL": "Helsinki",
    "PLGDY": "Gdynia",
    "EERTL": "Tallinn",
    "LVRIX": "Riga",
    "LTKLJ": "Klaipeda",
    # NORVEGIA
    "NOSVG": "Stavanger",
    "NOBGO": "Bergen",
    "NOOSL": "Oslo",
    "NOALE": "Ålesund",
    "NOTRH": "Tromsø",
    "NOTOS": "Tromsø",
    "NOTRD": "Trondheim",
    "NOHGD": "Haugesund",
    "NOKRS": "Kristiansand",
    "NOLYN": "Lyngdal",
    "NOSND": "Sandnes",
    "NOEID": "Eidfjord",
    "NOFLM": "Flåm",
    "NOLYR": "Longyearbyen",
    "NOSVJ": "Vik i Sogn",
    "NOHVN": "Honningsvåg",
    "NOGER": "Geiranger",
    "NOHEL": "Hellesylt",
    "NOOLD": "Olden",
    "NOMOL": "Molde Fjord",
    "NONFE": "Nordfjordeid",
    "NOULV": "Ulvik",
    "NOMLY": "Maløy",
    # ISLANDA
    "ISREY": "Reykjavik",
    "ISISA": "Ísafjörður",
    "ISAKR": "Akureyri",
    "ISSEY": "Seyðisfjörður",
    # PORTOGALLO
    "PTLIS": "Lisbona",
    "PTFNC": "Funchal",
    "PTPDL": "Ponta Delgada",
    "PTLEI": "Leixões",
    # CROAZIA / ADRIATICO
    "HRSPU": "Spalato",
    "HRDBV": "Dubrovnik",
    "HRZAD": "Zara",
    "MEKOT": "Kotor",
    # MEDIO ORIENTE
    "JOAQJ": "Aqaba",
    "QADOH": "Doha",
    "OMMCT": "Muscat",
    "BHBAH": "Bahrain",
    # AFRICA
    "ZACPT": "Città del Capo",
    "ZADUR": "Durban",
    "MZMPM": "Maputo",
    "MZPOM": "Pomene",
    "MZPTI": "Isola Portoghese",
    "MGNOS": "Nosy Be",
    "MGDIE": "Antsiranana",
    "SCVTX": "Port Victoria",
    "MUPLU": "Port Louis",
    "RELAP": "La Possession",
    "CVMIN": "Mindelo",
    "NAWVB": "Walvis Bay",
    "TNLAG": "La Goulette",
    # MEDITERRANEO ISOLE
    "MTMLA": "La Valletta",
    "MTVAL": "La Valletta",
    "MCMCM": "Monte Carlo",
    "CYLEM": "Limassol",
    "GIBGI": "Gibilterra",
    "GIB":   "Gibilterra",
    # CANADA
    "CAHAL": "Halifax",
    "CAVIC": "Victoria",
    # AUSTRALIA / NZ
    "AUSYD": "Sydney",
    "AUMEL": "Melbourne",
    "AUCNS": "Cairns",
    "AUTSV": "Townsville",
    "NZAKL": "Auckland",
    "NZWLG": "Wellington",
    # PACIFICO
    "WSAPW": "Apia",
    "ASPPG": "Pago Pago",
    "FJSUV": "Suva",
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
    return PORTS.get(raw[:5], raw.title()) if raw else ""

# porto vs at sea
def location(p):
    if float(p.get("sog", 0) or 0) > 1 and int(p.get("status", 0) or 0) in (0, 8):
        return "At Sea"
    dest = port_name(p.get("destination", ""))
    return dest or "In Port"


def darken(c, n=50):
    h = c.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "#{:02x}{:02x}{:02x}".format(max(0, r-n), max(0, g-n), max(0, b-n))

# icone nave MSC o EXPLORA (più piccole, luxury)
def ship_icon(color, luxury, offline=False):
    s  = darken(color)
    op = ' opacity="0.38"' if offline else ''
    w, h = (36, 21) if luxury else (48, 28)
    if luxury:
        body = (f'<path d="M4 36 L4 30 L8 30 L8 26 Q8 24 10 24 L70 24 L76 24 Q80 24 84 27 L90 30 L92 30 L92 34 Q70 40 48 40 Q24 40 4 36 Z" '
                f'fill="{color}" stroke="{s}" stroke-width="1.5"/>'
                f'<rect x="10" y="16" width="55" height="8" rx="1" fill="{s}" opacity="0.85"/>'
                f'<rect x="30" y="10" width="7" height="7" rx="1" fill="{s}"/>'
                f'<rect x="42" y="10" width="7" height="7" rx="1" fill="{s}"/>'
                f'<line x1="33" y1="5" x2="33" y2="10" stroke="{s}" stroke-width="1.2"/>'
                f'<line x1="45" y1="5" x2="45" y2="10" stroke="{s}" stroke-width="1.2"/>'
                f'<line x1="4" y1="36" x2="92" y2="34" stroke="#dc2626" stroke-width="1.2"/>')
    else:
        body = (f'<path d="M8 40 L8 34 Q8 32 10 32 L18 32 L18 24 Q18 22 20 22 L30 22 L30 18 '
                f'Q30 16 32 16 L52 16 Q54 16 54 18 L54 22 L58 22 L58 16 Q58 14 60 14 L66 14 '
                f'Q68 14 68 16 L68 22 L76 22 Q80 22 82 25 L88 32 L90 32 Q92 32 92 34 L92 40 '
                f'Q70 46 50 46 Q28 46 8 40 Z" fill="{color}" stroke="{s}" stroke-width="2"/>'
                f'<rect x="19" y="23" width="34" height="8" rx="1.5" fill="{s}" opacity="0.3"/>'
                f'<rect x="59" y="14.5" width="8" height="7" rx="1" fill="{s}"/>'
                f'<path d="M62 10 L64 14.5 L60 14.5 Z" fill="{s}"/>')
    return f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 100 52"{op}>{body}</svg>'

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
            p      = known.get(mmsi)
            loc    = location(p) if p else ""
            spd    = f'{float(p.get("sog", 0)):.0f} kn' if online and p else ""
            dot    = f'background:{color}' + ('' if online else ';opacity:0.28')
            ncls   = 'n on' if online else ('n st' if p else 'n')
            fly    = f'onclick="fly(\'{mmsi}\')"' if p else ''
            if online:
                badge = f'<span class="b as">At Sea</span>' if loc == "At Sea" else f'<span class="b ip" title="{loc}">{loc[:13]}{"…" if len(loc) > 13 else ""}</span>'
            elif p:
                badge = '<span class="b of">Offline</span>'
            else:
                badge = ''
            rows.append(f'<div class="row" {fly}><div class="dot" style="{dot}"></div>'
                        f'<span class="{ncls}">{name}</span>{badge}'
                        f'<span class="spd">{spd}</span></div>')
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
               f'<tr><td style=\\"color:#9ab4cc\\">Posizione</td><td style=\\"text-align:right;font-weight:600\\">{loc}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Destinazione</td><td style=\\"text-align:right;font-weight:600\\">{dest}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Velocità</td><td style=\\"text-align:right;font-weight:600\\">{spd:.1f} kn</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Rotta</td><td style=\\"text-align:right;font-weight:600\\">{hdg:.0f}°</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Lat / Lon</td><td style=\\"text-align:right;font-weight:600\\">{lat:.4f}, {lon:.4f}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Aggiornato</td><td style=\\"text-align:right;font-weight:600\\">{ts[:16]}</td></tr>'
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
        spd  = float(p.get("sog", 0) or 0)
        hdg  = float(p.get("heading", 0) or 0)
        ts   = str(p.get("timestamp", ""))
        lat, lon = p["lat"], p["lon"]
        pop = (f'<div style=\\"min-width:200px\\">'
               f'<b style=\\"color:{color};opacity:0.6;font-size:13px\\">{name}</b>'
               f'<div style=\\"font-size:9px;color:#9ab4cc;letter-spacing:1px;margin:3px 0 10px\\">'
               f'{cls} CLASS &nbsp;·&nbsp; <span style=\\"color:#f59e0b\\">Offline</span></div>'
               f'<table style=\\"width:100%;font-size:11px;border-collapse:collapse\\">'
               f'<tr><td style=\\"color:#9ab4cc\\">Destinazione</td><td style=\\"text-align:right;font-weight:600\\">{dest}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Velocità</td><td style=\\"text-align:right;font-weight:600\\">{spd:.1f} kn</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Rotta</td><td style=\\"text-align:right;font-weight:600\\">{hdg:.0f}°</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Lat / Lon</td><td style=\\"text-align:right;font-weight:600\\">{lat:.4f}, {lon:.4f}</td></tr>'
               f'<tr><td style=\\"color:#9ab4cc\\">Ultima nota</td><td style=\\"text-align:right;color:#f59e0b;font-weight:600\\">{ts[:16]}</td></tr>'
               f'</table></div>')
        lines.append(f'mark({lat},{lon},`{svg}`,"{mmsi}","{pop}",true,{"true" if lux else "false"});')

    return "\n".join(lines)

def write_github_page(html_content: str):
    output_file = REPO_DIR / "index.html"
    output_file.write_text(html_content, encoding="utf-8")
    print(f"  → generato {output_file.name}")


def write_html():
    now    = datetime.now().strftime("%H:%M:%S")
    online = len(live)
    total  = len(FLEET)
    n_off  = sum(1 for m in FLEET if m not in live and m in known)

    all_pos = {**known, **live}
    fly_js  = "\n".join(
        f'  if(id==="{m}") map.setView([{p["lat"]},{p["lon"]}],7);'
        for m, p in all_pos.items() if p.get("lat") is not None
    )

    html = f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="{REFRESH * 60}">
<title>MSC & Explora Fleet Tracker</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.min.css"/>
<style>
* {{ box-sizing:border-box; margin:0; padding:0; font-family:Verdana,sans-serif; }}
body {{ background:#d8e8f0; height:100vh; display:flex; flex-direction:column; overflow:hidden; }}
header {{
  background:linear-gradient(135deg,#0a1628,#122444,#0d1e3a);
  border-bottom:1px solid #1e3d70;
  padding:12px 24px 18px;
  display:flex; align-items:center; justify-content:space-between;
  flex-shrink:0; box-shadow:0 4px 16px rgba(0,0,0,.4);
}}
.title {{ text-align:center; flex:1; }}
.title h1 {{ font-size:19px; font-weight:700; letter-spacing:2px; color:#fff; text-transform:uppercase; }}
.title p  {{ font-size:10px; color:#7aaddb; letter-spacing:3px; text-transform:uppercase; margin-top:4px; }}
.info {{ display:flex; align-items:center; gap:20px; font-size:11px; color:#7aaddb; min-width:260px; justify-content:flex-end; }}
.dot  {{ width:8px; height:8px; border-radius:50%; background:#22c55e;
         animation:pulse 2s infinite; display:inline-block; margin-right:5px; }}
@keyframes pulse {{
  0%,100% {{ opacity:1; box-shadow:0 0 0 0 rgba(34,197,94,.4); }}
  50%      {{ opacity:.5; box-shadow:0 0 0 5px rgba(34,197,94,0); }}
}}
.on  {{ color:#22c55e; font-weight:700; }}
.off {{ color:#f59e0b; font-weight:700; }}
#map {{ flex:1; }}
.panel {{
  position:absolute; top:88px; right:14px; width:310px;
  background:rgba(255,255,255,.97); border:1px solid #b8cedd;
  border-radius:10px; box-shadow:0 4px 20px rgba(0,0,0,.13);
  z-index:900; display:flex; flex-direction:column;
  max-height:calc(100vh - 104px); overflow:hidden;
}}
.panel-title {{
  font-size:11px; font-weight:700; letter-spacing:2px;
  text-transform:uppercase; color:#5a7a9a;
  padding:9px 14px 7px; border-bottom:1px solid #d0e0ec; background:#eef4fa;
}}
.list {{ flex:1; display:flex; flex-direction:column; overflow:hidden; }}
.cls  {{
  font-size:9px; font-weight:700; letter-spacing:2px;
  text-transform:uppercase; color:#9ab4cc;
  padding:4px 14px 2px; background:#f5f9fd;
}}
.row {{
  display:flex; align-items:center; gap:7px;
  padding:3px 11px; cursor:pointer;
  border-bottom:1px solid #eef3f8; transition:background .12s;
}}
.row:hover {{ background:#e4eef8; }}
.n    {{ font-size:12px; color:#9ab4cc; flex:1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.n.on {{ color:#1a3050; font-weight:700; }}
.n.st {{ color:#7890a8; font-style:italic; }}
.b    {{ font-size:10px; padding:1px 6px; border-radius:3px; flex-shrink:0; font-weight:700;
         max-width:90px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.as   {{ background:#dbeafe; color:#1d4ed8; }}
.ip   {{ background:#dcfce7; color:#166534; }}
.of   {{ background:#fef3c7; color:#92400e; }}
.spd  {{ font-size:10px; color:#9ab4cc; flex-shrink:0; }}
.leaflet-popup-content-wrapper {{
  background:#fff !important; border:1px solid #b8cedd !important;
  border-radius:10px !important; font-family:Verdana,sans-serif !important;
  box-shadow:0 8px 24px rgba(0,0,0,.18) !important;
}}
.leaflet-popup-tip {{ background:#fff !important; }}
.leaflet-popup-tip {{ background:#fff !important; }}

.leaflet-tile-pane {{
  filter: brightness(0.78) contrast(1.16) saturate(0.72);
}}
</style>
</style>
</head>
<body>
<header>
  <div style="min-width:140px"></div>
  <div class="title">
    <h1>MSC Cruises &amp; Explora Journeys</h1>
    <p>Fleet Live Tracker</p>
  </div>
  <div class="info">
    <span>Online: <span class="on">{online}</span>/{total} &nbsp;·&nbsp; Last known: <span class="off">{n_off}</span></span>
    <span><span class="dot"></span>{now} &nbsp;·&nbsp; Refresh ogni {REFRESH} min</span>
  </div>
</header>
<div id="map"></div>
<div class="panel">
  <div class="panel-title">Fleet</div>
  <div class="list">
{sidebar()}
  </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.min.js"></script>
<script>
var map = L.map('map').setView([30,10],3);
L.tileLayer('https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
  attribution:'© OpenStreetMap contributors',
  maxZoom:19
}}).addTo(map);
var mkrs = {{}};
function mark(lat,lon,svg,mmsi,popup,offline,lux) {{
  var b64 = btoa(unescape(encodeURIComponent(svg)));
  var sz  = lux ? [36,21] : [48,28];
  var anc = lux ? [18,10] : [24,14];
  var ic  = L.icon({{iconUrl:'data:image/svg+xml;base64,'+b64, iconSize:sz, iconAnchor:anc, popupAnchor:[0,-16]}});
  mkrs[mmsi] = L.marker([lat,lon],{{icon:ic, opacity:offline?0.6:1}}).addTo(map).bindPopup(popup,{{maxWidth:270}});
}}
function fly(id) {{
{fly_js}
  if(mkrs[id]) mkrs[id].openPopup();
}}
{markers()}
</script>
</body>
</html>"""

    write_github_page(html)
    save_known()
    print(f"[{datetime.now():%H:%M:%S}] aggiornato — {online}/{total} online, {n_off} offline con posizione nota")


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

                    destination = static_data.get(
                        "Destination",
                        ""
                    )

                    known.setdefault(mmsi, {})[
                        "destination"
                    ] = destination

                    if mmsi in live:
                        live[mmsi][
                            "destination"
                        ] = destination

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
    print(f"MSC & Explora Fleet Tracker — {len(FLEET)} navi — refresh ogni {REFRESH} min")
    load_known()
    asyncio.run(run())