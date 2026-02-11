"""
Airport coordinate database and distance calculation utilities.

Contains ICAO airport coordinates and a haversine function for
calculating great-circle distances in nautical miles.

To add airports for your flying area, add entries to the AIRPORTS dict below:
    "ICAO": (latitude, longitude),  # Airport Name
"""

import math
import json
import os

# Airport coordinates database (lat, lon)
# Sources: FAA, ICAO databases
AIRPORTS = {
    # US - Florida
    "KFPR": (27.4951, -80.3683),  # Fort Pierce
    "KMLB": (28.1028, -80.6453),  # Melbourne
    "KVRB": (27.6556, -80.4178),  # Vero Beach
    "KTIX": (28.5148, -80.7992),  # Titusville
    "KFXE": (26.1973, -80.1707),  # Fort Lauderdale Exec
    "KPBI": (26.6832, -80.0956),  # Palm Beach Intl
    "KAPF": (26.1526, -81.7753),  # Naples
    "X14": (27.1814, -80.2222),   # La Belle
    "KFMY": (26.5866, -81.8633),  # Fort Myers
    "KOBE": (27.2626, -80.8490),  # Okeechobee
    "KSUA": (27.1817, -80.2211),  # Stuart
    "KSGJ": (29.9592, -81.3397),  # St Augustine
    "KGNV": (29.6900, -82.2718),  # Gainesville
    "KMKY": (25.9950, -80.4339),  # Marco Island
    "KTLH": (30.3965, -84.3503),  # Tallahassee
    "KJZI": (32.6986, -80.0028),  # Charleston Exec
    "KRZR": (35.1385, -83.8571),  # Franklin-Robbins
    "KJWN": (36.1824, -86.8867),  # Nashville West
    "KDTS": (30.4001, -86.4715),  # Destin
    "KDAB": (29.1799, -81.0581),  # Daytona Beach
    "KSPG": (27.7651, -82.6270),  # St Pete-Clearwater
    "KCTY": (30.5685, -85.9681),  # Cross City
    "KJES": (31.5540, -84.5955),  # Jesup
    "KLRO": (34.6097, -82.1734),  # Loerrach/Mt Pleasant
    "KIMM": (26.4332, -81.4010),  # Immokalee
    "X51": (25.4876, -80.5569),   # Homestead
    "KMTH": (24.7261, -81.0514),  # Marathon
    "KEYW": (24.5561, -81.7596),  # Key West
    "KFHB": (29.6863, -82.5916),  # Fernandina Beach
    "KSEF": (27.4564, -80.3718),  # Sebring
    "X59": (27.9831, -80.6816),   # Valkaria
    "KPGD": (26.9202, -81.9906),  # Punta Gorda
    "X60": (27.2450, -80.7267),   # Williston
    "KMIA": (25.7959, -80.2870),  # Miami Intl
    "KOPF": (25.9070, -80.2784),  # Opa-Locka
    "KAYS": (31.2491, -82.3955),  # Waycross
    "KNEW": (30.0424, -90.0283),  # Lakefront (New Orleans)
    "F95": (30.2183, -85.6828),   # Defuniak Springs
    "KZPH": (28.2282, -82.1559),  # Zephyrhills
    "KTPF": (27.9136, -82.4495),  # Tampa
    "X06": (27.2292, -80.9680),   # Arcadia
    "KMYR": (33.6797, -78.9283),  # Myrtle Beach
    "KSSI": (31.1518, -81.3913),  # St Simons Island
    "KBCT": (26.3785, -80.1077),  # Boca Raton
    "KZWN": (36.1824, -86.8867),  # Nashville (alias)
    "KGRD": (33.4689, -82.1607),  # Fort Gordon/Augusta
    "KMGR": (31.0849, -82.0387),  # Moultrie
    "1A3": (33.3556, -83.7564),   # Martin Campbell
    "24J": (29.6587, -82.5850),   # Suwannee County
    "28J": (29.9592, -81.6875),   # Palatka
    "42J": (30.0558, -81.5067),   # Keystone Heights
    "KLZU": (33.9781, -83.9624),  # Lawrenceville/Gwinnett

    # US - Northeast
    "KTEB": (40.8501, -74.0608),  # Teterboro
    "KHPN": (41.0670, -73.7076),  # White Plains
    "KJFK": (40.6413, -73.7781),  # JFK
    "KLGA": (40.7772, -73.8726),  # LaGuardia
    "KEWR": (40.6925, -74.1687),  # Newark
    "KBOS": (42.3656, -71.0096),  # Boston
    "KPHL": (39.8721, -75.2411),  # Philadelphia
    "KBWI": (39.1754, -76.6683),  # Baltimore
    "KIAD": (38.9445, -77.4558),  # Dulles
    "KRIC": (37.5052, -77.3197),  # Richmond
    "KORF": (36.8946, -76.2012),  # Norfolk
    "KROA": (37.3255, -79.9754),  # Roanoke
    "KBDL": (41.9389, -72.6832),  # Hartford
    "KPVD": (41.7268, -71.4282),  # Providence
    "KALB": (42.7483, -73.8017),  # Albany
    "KSWF": (41.5041, -74.1048),  # Stewart/Newburgh
    "KBXM": (44.8072, -68.8281),  # Bangor
    "KMMU": (40.7993, -74.4149),  # Morristown
    "KHVN": (41.2637, -72.8868),  # Tweed New Haven
    "KPSM": (43.0779, -70.8233),  # Portsmouth
    "KBED": (42.4700, -71.2890),  # Bedford/Hanscom
    "KORH": (42.2673, -71.8757),  # Worcester
    "KOXC": (41.4786, -73.1352),  # Oxford
    "KABE": (40.6521, -75.4408),  # Allentown
    "KILG": (39.6787, -75.6065),  # Wilmington DE
    "KACY": (39.4576, -74.5772),  # Atlantic City
    "KELM": (42.1600, -76.8916),  # Elmira
    "KBUF": (42.9405, -78.7322),  # Buffalo
    "KROC": (43.1189, -77.6724),  # Rochester
    "KERI": (42.0831, -80.1739),  # Erie
    "KHLG": (40.1750, -80.6463),  # Wheeling
    "KRMN": (39.7053, -77.6726),  # Martinsburg
    "KCHO": (38.1386, -78.4529),  # Charlottesville
    "KMRB": (39.4019, -77.9846),  # Martinsburg WV

    # US - Midwest
    "KMDW": (41.7868, -87.7522),  # Chicago Midway
    "KDAY": (39.9024, -84.2194),  # Dayton
    "KCMI": (40.0393, -88.2781),  # Champaign
    "KCWA": (44.7776, -89.6668),  # Wausau
    "KCAK": (40.9161, -81.4422),  # Akron-Canton
    "KMKE": (42.9472, -87.8966),  # Milwaukee
    "KIND": (39.7173, -86.2944),  # Indianapolis
    "KPIT": (40.4915, -80.2329),  # Pittsburgh
    "KGRR": (42.8808, -85.5228),  # Grand Rapids
    "KDET": (42.4092, -83.0099),  # Detroit Coleman
    "KDTW": (42.2124, -83.3534),  # Detroit Metro
    "KMCI": (39.2976, -94.7139),  # Kansas City
    "KMKC": (39.1227, -94.5928),  # Kansas City Downtown
    "KAPA": (39.5701, -104.8493), # Denver Centennial
    "KBJC": (39.9088, -105.1172), # Broomfield/Jeffco
    "KSUS": (38.6621, -90.6522),  # Spirit of St Louis
    "KFAR": (46.9207, -96.8158),  # Fargo
    "KDLH": (46.8421, -92.1936),  # Duluth
    "KFWA": (40.9785, -85.1951),  # Fort Wayne
    "KATW": (44.2581, -88.5191),  # Appleton
    "KPIA": (40.6642, -89.6933),  # Peoria
    "KLNK": (40.8510, -96.7592),  # Lincoln
    "KOMA": (41.3032, -95.8941),  # Omaha
    "KTOL": (41.5868, -83.8078),  # Toledo
    "KCMH": (39.9980, -82.8919),  # Columbus OH
    "KBKL": (41.5175, -81.6833),  # Cleveland Burke
    "KMQY": (36.0089, -86.5201),  # Smyrna TN
    "KYIP": (42.2379, -83.5304),  # Willow Run (Detroit)
    "KLBE": (40.2759, -79.4048),  # Latrobe

    # US - South
    "KHOU": (29.6454, -95.2789),  # Houston Hobby
    "KIAH": (29.9844, -95.3414),  # Houston Intl
    "KDWH": (30.0618, -95.5546),  # Houston Hooks
    "KDAL": (32.8471, -96.8518),  # Dallas Love
    "KDFW": (32.8998, -97.0403),  # Dallas-Fort Worth
    "KSAT": (29.5337, -98.4698),  # San Antonio
    "KPWA": (35.5342, -97.6471),  # Oklahoma City Wiley Post
    "KOKC": (35.3931, -97.6007),  # Oklahoma City
    "KRDU": (35.8776, -78.7875),  # Raleigh-Durham
    "KCLT": (35.2140, -80.9431),  # Charlotte
    "KAGS": (33.3700, -81.9645),  # Augusta
    "KRYY": (34.0132, -84.5971),  # McCollum (Atlanta area)
    "KPDK": (33.8756, -84.3020),  # Peachtree DeKalb
    "KTYS": (35.8110, -83.9940),  # Knoxville
    "KTRI": (36.4752, -82.4074),  # Tri-Cities
    "KBHM": (33.5629, -86.7535),  # Birmingham
    "KJAX": (30.4941, -81.6879),  # Jacksonville
    "KMEM": (35.0424, -89.9767),  # Memphis
    "KLBB": (33.6636, -101.8228), # Lubbock
    "KROW": (33.3016, -104.5307), # Roswell
    "KAIV": (33.1065, -88.1975),  # Aliceville
    "KLKR": (34.7230, -80.8549),  # Lancaster
    "KHDC": (34.4362, -82.6913),  # Dillingham

    # US - West
    "KPHX": (33.4373, -112.0078), # Phoenix Sky Harbor
    "KSDL": (33.6229, -111.9105), # Scottsdale
    "KIWA": (33.3078, -111.6553), # Phoenix-Mesa
    "KDVT": (33.6883, -112.0833), # Deer Valley (Phoenix)
    "KLAS": (36.0840, -115.1537), # Las Vegas
    "KSLC": (40.7884, -111.9778), # Salt Lake City
    "KTUS": (32.1161, -110.9410), # Tucson
    "KBFL": (35.4336, -119.0568), # Bakersfield
    "KSAC": (38.5125, -121.4935), # Sacramento
    "KSJC": (37.3626, -121.9290), # San Jose
    "KSAN": (32.7336, -117.1897), # San Diego
    "KSNA": (33.6757, -117.8683), # Orange County
    "KBUR": (34.2007, -118.3585), # Burbank
    "KBFI": (47.5300, -122.3020), # Boeing Field (Seattle)
    "KPAE": (47.9063, -122.2816), # Paine Field (Everett)
    "KRNO": (39.4991, -119.7681), # Reno
    "KMTJ": (38.5098, -107.8942), # Montrose
    "KGJT": (39.1224, -108.5267), # Grand Junction
    "KBZN": (45.7775, -111.1530), # Bozeman
    "KBTM": (45.9548, -112.4972), # Butte
    "KELP": (31.8072, -106.3778), # El Paso
    "KNYL": (32.6564, -114.6060), # Yuma
    "KCOS": (38.8058, -104.7007), # Colorado Springs
    "KCPR": (42.9080, -106.4644), # Casper
    "KFNL": (40.4518, -105.0113), # Fort Collins
    "KPVU": (40.2192, -111.7236), # Provo
    "KSGU": (37.0905, -113.5931), # St George UT
    "KIFP": (35.1575, -114.5594), # Bullhead City
    "KACV": (40.9781, -124.1086), # Arcata
    "KGEG": (47.6199, -117.5338), # Spokane
    "KABQ": (35.0402, -106.6090), # Albuquerque

    # Caribbean / Central America
    "TJSJ": (18.4394, -66.0018),  # San Juan PR
    "TNCM": (18.0441, -63.1089),  # St Maarten
    "MDPP": (18.5674, -68.3631),  # Punta Cana
    "MYNN": (25.0390, -77.4662),  # Nassau

    # Europe
    "LLBG": (32.0114, 34.8867),   # Tel Aviv Ben Gurion
    "LCLK": (34.8751, 33.6249),   # Larnaca
    "LCPH": (34.7180, 32.4857),   # Paphos
    "LIRA": (41.7994, 12.5949),   # Rome Ciampino
    "LIRQ": (43.8100, 11.2051),   # Florence
    "LSZS": (46.5344, 9.8844),    # Samedan (St Moritz)
    "EGLF": (51.2758, -0.7764),   # Farnborough
    "EGGW": (51.8747, -0.3684),   # Luton
    "EKCH": (55.6180, 12.6561),   # Copenhagen
    "EKRK": (51.8413, -8.4911),   # Cork Ireland
    "LEPA": (39.5517, 2.7388),    # Palma de Mallorca
    "LEBL": (41.2971, 2.0785),    # Barcelona
    "LEGE": (41.9010, 2.7606),    # Girona
    "LEVC": (39.4893, -0.4816),   # Valencia
    "LMML": (35.8575, 14.4775),   # Malta
    "GMMN": (33.3675, -7.5900),   # Casablanca
    "HECA": (30.1219, 31.4056),   # Cairo
    "HESH": (27.9778, 34.3950),   # Sharm El Sheikh
    "LBSF": (42.6952, 23.4114),   # Sofia
    "BIKF": (63.9850, -22.6056),  # Keflavik (Iceland)
    "CYYR": (53.3192, -60.4258),  # Goose Bay
    "CYYZ": (43.6772, -79.6306),  # Toronto Pearson
    "LIML": (45.4520, 9.2765),    # Milan Linate
    "LFMN": (43.6584, 7.2159),    # Nice
    "LFPB": (48.9694, 2.4414),    # Paris Le Bourget
    "LFSB": (47.5896, 7.5299),    # Basel
    "LTBS": (36.7131, 29.5847),   # Dalaman (Turkey)
    "LTAC": (40.1281, 32.9951),   # Ankara Esenboga
    "LTFE": (37.8554, 30.3282),   # Isparta
    "EHAM": (52.3086, 4.7639),    # Amsterdam
    "EDDL": (51.2895, 6.7668),    # Dusseldorf
    "EDDV": (52.4611, 9.6850),    # Hannover
    "LGIR": (35.3397, 25.1803),   # Heraklion (Crete)
    "LGMK": (37.4351, 25.3481),   # Mykonos
    "LGRP": (36.4054, 28.0862),   # Rhodes
    "LFDH": (48.5364, -3.3463),   # Lannion
    "LFLS": (45.3629, 5.3294),    # Grenoble

    # Middle East
    "OMDW": (24.8962, 55.1614),   # Al Maktoum / Dubai World Central
    "OBBI": (26.2708, 50.6336),   # Bahrain
    "OJAM": (31.7226, 35.9932),   # Amman Marka
    "VRDA": (7.1808, 79.8841),    # Placeholder (non-standard code)

    # Additional US airports
    "KCVC": (33.6322, -83.8466),  # Covington Municipal, GA
    "KEZM": (32.2164, -83.1287),  # Heart of Georgia Regional, Eastman GA
    "KMGW": (39.6428, -79.9164),  # Morgantown Municipal, WV
}

# Simulator/training device entries (no coordinates)
SIMULATOR_ENTRIES = {"FRASCA", "CAE HAWKER 800XP", "PA44 SIM"}


def haversine_nm(lat1, lon1, lat2, lon2):
    """Calculate great-circle distance between two points in nautical miles.

    Uses the Haversine formula.

    Args:
        lat1, lon1: Latitude and longitude of point 1 (in degrees)
        lat2, lon2: Latitude and longitude of point 2 (in degrees)

    Returns:
        Distance in nautical miles, rounded to 1 decimal place.
    """
    R_nm = 3440.065  # Earth radius in nautical miles
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    c = 2 * math.asin(math.sqrt(a))
    return round(R_nm * c, 1)


def load_custom_airports(filepath):
    """Load additional airports from a JSON file.

    JSON format: {"ICAO": [lat, lon], ...}

    Args:
        filepath: Path to JSON file with airport coordinates.

    Returns:
        Dict of ICAO code -> (lat, lon) tuples.
    """
    if not os.path.exists(filepath):
        return {}
    with open(filepath, 'r') as f:
        data = json.load(f)
    return {k: tuple(v) for k, v in data.items()}


def get_all_airports(custom_file=None):
    """Get combined airport database (built-in + custom).

    Args:
        custom_file: Optional path to JSON file with additional airports.

    Returns:
        Dict of ICAO code -> (lat, lon) tuples.
    """
    airports = dict(AIRPORTS)
    if custom_file:
        custom = load_custom_airports(custom_file)
        airports.update(custom)
    return airports
