"""
Tata Motors workshop location verifier.

Reads a Google My Maps KML export (or an .xlsx/.csv with Name/Latitude/Longitude),
checks each workshop against Google Places API (New), and writes an Excel report.

Usage:
    python verify.py Workshopdata.kml                 # full run
    python verify.py Workshopdata.kml --limit 30      # pilot on first 30 rows
    python verify.py Workshopdata.kml --offline       # only free pre-checks, no API calls

API key is read from the GOOGLE_MAPS_API_KEY environment variable
(set it as a GitHub Codespaces secret).
"""

import argparse
import csv
import json
import math
import os
import re
import sys
import time
import xml.etree.ElementTree as ET

import pandas as pd
import requests
from rapidfuzz import fuzz

# ----------------------------- settings -----------------------------
OK_DIST_M = 200          # accepted coordinate error
NEAR_DIST_M = 500        # 200-500 m => "check"
NAME_OK = 80             # fuzzy score (0-100) treated as a name match
NAME_WEAK = 60           # below this = no match
BIAS_RADIUS_M = 2000     # search is biased (not restricted) around the given coordinate
SLEEP_S = 0.15           # pause between API calls
CACHE_FILE = "places_cache.json"

TEXT_URL = "https://places.googleapis.com/v1/places:searchText"
NEARBY_URL = "https://places.googleapis.com/v1/places:searchNearby"
FIELDS = ("places.id,places.displayName,places.formattedAddress,"
          "places.location,places.types,places.businessStatus")

# words that say nothing about *which* workshop it is
GENERIC = {
    "pvt", "ltd", "limited", "private", "p", "i", "india", "co", "company", "the",
    "m", "s", "ms", "and", "tata", "commercial", "vehicle", "vehicles", "service",
    "services", "workshop", "centre", "center", "dealer", "dealership", "authorised",
    "authorized", "cv", "sales", "llp", "inc", "pltd", "plt",
}
# common in dealer names; kept for display but not enough on their own to call a match
WEAK_TOKENS = {
    "motors", "motor", "automobiles", "automobile", "auto", "autos", "automotive",
    "agencies", "agency", "enterprises", "enterprise", "garage", "trucks", "truck",
    "mega", "shree", "shri", "sri", "new",
}


# ----------------------------- input -----------------------------
def load_kml(path):
    ns = "{http://www.opengis.net/kml/2.2}"
    root = ET.parse(path).getroot()
    rows = []
    for pm in root.iter(ns + "Placemark"):
        r = {"Name": (pm.findtext(ns + "name") or "").strip()}
        for d in pm.iter(ns + "Data"):
            r[d.get("name")] = (d.findtext(ns + "value") or "").strip()
        coords = (pm.findtext(f".//{ns}coordinates") or "").strip()
        if coords:
            lon, lat = coords.split(",")[:2]
            r["lat"], r["lon"] = float(lat), float(lon)
        rows.append(r)
    return pd.DataFrame(rows)


def load_table(path):
    df = pd.read_excel(path) if path.lower().endswith((".xlsx", ".xls")) else pd.read_csv(path)
    cols = {c.lower().strip(): c for c in df.columns}
    df["lat"] = pd.to_numeric(df[cols.get("latitude", cols.get("lat"))], errors="coerce")
    df["lon"] = pd.to_numeric(df[cols.get("longitude", cols.get("lon", cols.get("lng")))], errors="coerce")
    if "Name" not in df.columns:
        df["Name"] = df[cols.get("name", cols.get("workshop name", df.columns[0]))]
    return df


# ----------------------------- helpers -----------------------------
def haversine_m(lat1, lon1, lat2, lon2):
    r = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def clean_name(name):
    s = str(name)
    s = re.sub(r"-\s*\d{5,}\s*$", "", s)          # trailing dealer code "-1009040"
    s = s.lower().replace("&", " and ")
    s = re.sub(r"\(\s*i\s*\)", " ", s)              # "(I)" in Popular Mega Motors (I)
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return " ".join(t for t in s.split() if t not in GENERIC)


def name_score(ours, theirs):
    a, b = clean_name(ours), clean_name(theirs)
    if not a or not b:
        return 0
    distinct_a = [t for t in a.split() if t not in WEAK_TOKENS]
    distinct_b = [t for t in b.split() if t not in WEAK_TOKENS]
    if distinct_a and distinct_b:
        # the distinctive part must match ("Popular", "Jai Baba Mungipa"); "Motors" alone is not enough
        return fuzz.token_set_ratio(" ".join(distinct_a), " ".join(distinct_b))
    if distinct_a or distinct_b:
        return 0                                   # e.g. "Hazi Motors" vs a generic "Tata Motors Service"
    return min(fuzz.token_set_ratio(a, b), 70)   # both sides only generic words: cap below NAME_OK


def search_query(row):
    q = re.sub(r"-\s*\d{5,}\s*$", "", str(row["Name"])).strip()
    city = str(row.get("City", "") or "").strip()
    if city and city.lower() not in q.lower() and city.lower() not in ("na", "nan"):
        q += f" {city}"
    return q


def maps_link(lat, lon, place_id=None):
    url = f"https://www.google.com/maps/search/?api=1&query={lat},{lon}"
    return url + (f"&query_place_id={place_id}" if place_id else "")


# ----------------------------- offline pre-checks -----------------------------
STATE_ALIASES = {
    "orissa": "odisha", "chattisgarh": "chhattisgarh", "telengana": "telangana",
    "up": "uttar pradesh", "uttart pradesh": "uttar pradesh", "br": "bihar",
    "delhi": "nct", "new delhi": "nct", "jammu & kashmir": "kashmir",
    "jammu and kashmir": "kashmir", "j&k": "kashmir", "leh-ladakh": "kashmir",
    "ladakh": "kashmir", "pune": "maharashtra", "puducherry": "pondicherry",
}


def load_india_places():
    try:
        import reverse_geocoder
    except ImportError:
        return None
    path = os.path.join(os.path.dirname(reverse_geocoder.__file__), "rg_cities1000.csv")
    out = {}
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r["cc"] == "IN":
                out.setdefault(r["admin1"].lower(), []).append((float(r["lat"]), float(r["lon"])))
    return out


def state_distance_km(places, state, lat, lon):
    """Distance from the coordinate to the nearest known town of the claimed state."""
    s = STATE_ALIASES.get(str(state).strip().lower(), str(state).strip().lower())
    pts = places.get(s)
    if not pts:
        return None
    return min(haversine_m(lat, lon, a, b) for a, b in pts) / 1000


def offline_checks(df):
    flags = [[] for _ in range(len(df))]
    dup = df.duplicated(["lat", "lon"], keep=False)
    places = load_india_places()
    for i in range(len(df)):
        lat, lon = df.iloc[i]["lat"], df.iloc[i]["lon"]
        if pd.isna(lat) or pd.isna(lon):
            flags[i].append("No coordinates")
            continue
        if not (6 <= lat <= 37.5 and 68 <= lon <= 97.5):
            flags[i].append("Coordinates outside India")
        if dup.iloc[i]:
            flags[i].append("Same coordinates as another workshop")
        state = df.iloc[i].get("State", "")
        if places and isinstance(state, str) and len(state.strip()) > 1 and state.strip().lower() != "na":
            d = state_distance_km(places, state, lat, lon)
            if d is not None and d > 75:
                flags[i].append(f"Coordinates ~{d:.0f} km from State '{state.strip()}' (check coords or State column)")
    df["Pre-check flags"] = ["; ".join(f) for f in flags]
    return df


# ----------------------------- Places API -----------------------------
class Places:
    def __init__(self, key):
        self.key = key
        self.calls = 0
        self.cache = json.load(open(CACHE_FILE)) if os.path.exists(CACHE_FILE) else {}

    def _post(self, url, body, cache_key):
        if cache_key in self.cache:
            return self.cache[cache_key]
        headers = {"X-Goog-Api-Key": self.key, "X-Goog-FieldMask": FIELDS,
                   "Content-Type": "application/json"}
        for attempt in range(5):
            resp = requests.post(url, json=body, headers=headers, timeout=30)
            if resp.status_code == 200:
                self.calls += 1
                places = resp.json().get("places", [])
                self.cache[cache_key] = places
                time.sleep(SLEEP_S)
                return places
            if resp.status_code in (429, 500, 502, 503, 504):
                time.sleep(2 ** attempt)
                continue
            sys.exit(f"\nAPI error {resp.status_code}: {resp.text[:500]}\n"
                     "Check the API key, that 'Places API (New)' is enabled, and billing is linked.")
        sys.exit("\nAPI kept failing (rate limit / server error). Re-run later; finished rows are cached.")

    def text_search(self, query, lat, lon):
        body = {"textQuery": query, "pageSize": 10, "languageCode": "en", "regionCode": "IN",
                "locationBias": {"circle": {"center": {"latitude": lat, "longitude": lon},
                                            "radius": BIAS_RADIUS_M}}}
        return self._post(TEXT_URL, body, f"T|{query}|{lat:.5f}|{lon:.5f}")

    def nearby_tata(self, lat, lon):
        body = {"textQuery": "Tata Motors service", "pageSize": 10, "languageCode": "en",
                "regionCode": "IN",
                "locationRestriction": {"rectangle": _box(lat, lon, NEAR_DIST_M)}}
        return self._post(TEXT_URL, body, f"N|{lat:.5f}|{lon:.5f}")

    def save(self):
        json.dump(self.cache, open(CACHE_FILE, "w"))


def _box(lat, lon, m):
    dlat = m / 111320
    dlon = m / (111320 * max(math.cos(math.radians(lat)), 0.01))
    return {"low": {"latitude": lat - dlat, "longitude": lon - dlon},
            "high": {"latitude": lat + dlat, "longitude": lon + dlon}}


def best_candidate(name, lat, lon, places):
    best = None
    for p in places:
        loc = p.get("location", {})
        if "latitude" not in loc:
            continue
        g_name = p.get("displayName", {}).get("text", "")
        score = name_score(name, g_name)
        dist = haversine_m(lat, lon, loc["latitude"], loc["longitude"])
        # prefer a good name match, then the closest
        key = (score >= NAME_OK, score >= NAME_WEAK, -dist if score >= NAME_WEAK else score)
        if best is None or key > best["key"]:
            best = {"key": key, "name": g_name, "score": score, "dist": dist,
                    "addr": p.get("formattedAddress", ""), "id": p.get("id", ""),
                    "glat": loc["latitude"], "glon": loc["longitude"],
                    "bstatus": p.get("businessStatus", "")}
    return best


def classify(b):
    if not b:
        return "NOT FOUND"
    s, d = b["score"], b["dist"]
    if s >= NAME_OK and d <= OK_DIST_M:
        return "VERIFIED"
    if s >= NAME_OK and d <= NEAR_DIST_M:
        return "CHECK - name OK, 200-500 m away"
    if s >= NAME_OK:
        return "WRONG COORDS? - name found far away"
    if s >= NAME_WEAK and d <= OK_DIST_M:
        return "CHECK - similar name nearby"
    return "NOT FOUND"


# ----------------------------- main -----------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output", default="verification_report.xlsx")
    ap.add_argument("--limit", type=int, help="only process the first N rows (pilot run)")
    ap.add_argument("--offline", action="store_true", help="only run free pre-checks")
    args = ap.parse_args()

    df = load_kml(args.input) if args.input.lower().endswith(".kml") else load_table(args.input)
    if args.limit:
        df = df.head(args.limit)
    df = df.reset_index(drop=True)
    print(f"Loaded {len(df)} workshops")
    df = offline_checks(df)

    out_cols = ["Status", "Name match %", "Distance (m)", "Google name", "Google address",
                "Google lat", "Google lon", "Google business status", "Tata place within 500 m",
                "Map: our point", "Map: Google place"]
    for c in out_cols:
        df[c] = pd.Series([None] * len(df), dtype=object)

    if not args.offline:
        key = os.getenv("GOOGLE_MAPS_API_KEY")
        if not key:
            sys.exit("GOOGLE_MAPS_API_KEY not set. Add it as a Codespaces secret and restart the codespace.")
        api = Places(key)
        try:
            for i, row in df.iterrows():
                lat, lon = row["lat"], row["lon"]
                df.at[i, "Map: our point"] = maps_link(lat, lon) if pd.notna(lat) else ""
                if pd.isna(lat) or pd.isna(lon):
                    df.at[i, "Status"] = "NO COORDINATES"
                    continue
                b = best_candidate(row["Name"], lat, lon, api.text_search(search_query(row), lat, lon))
                status = classify(b)
                if status in ("NOT FOUND", "CHECK - similar name nearby"):
                    # second chance: anything Tata-branded close to the point?
                    near = api.nearby_tata(lat, lon)
                    tata = [p.get("displayName", {}).get("text", "") for p in near
                            if "tata" in p.get("displayName", {}).get("text", "").lower()]
                    df.at[i, "Tata place within 500 m"] = " | ".join(tata[:3])
                    b2 = best_candidate(row["Name"], lat, lon, near)
                    if b2 and (not b or b2["key"] > b["key"]):
                        b, status = b2, classify(b2)
                    if status == "NOT FOUND" and tata:
                        status = "CHECK - Tata place nearby, name differs"
                df.at[i, "Status"] = status
                if b:
                    df.at[i, "Name match %"] = round(b["score"])
                    df.at[i, "Distance (m)"] = round(b["dist"])
                    df.at[i, "Google name"] = b["name"]
                    df.at[i, "Google address"] = b["addr"]
                    df.at[i, "Google lat"] = b["glat"]
                    df.at[i, "Google lon"] = b["glon"]
                    df.at[i, "Google business status"] = b["bstatus"]
                    df.at[i, "Map: Google place"] = maps_link(b["glat"], b["glon"], b["id"])
                if (i + 1) % 25 == 0:
                    api.save()
                    print(f"  {i + 1}/{len(df)} done  (API calls so far: {api.calls})", flush=True)
        finally:
            api.save()
        print(f"API calls made this run: {api.calls} (cached rows are not billed again)")
    else:
        df["Status"] = "OFFLINE ONLY"
        df["Map: our point"] = [maps_link(a, b) if pd.notna(a) else "" for a, b in zip(df.lat, df.lon)]

    write_report(df, args.output)


def write_report(df, path):
    front = ["Status", "Pre-check flags", "Name", "Workshop Code", "City", "State",
             "lat", "lon", "Name match %", "Distance (m)", "Google name", "Google address",
             "Google lat", "Google lon", "Google business status", "Tata place within 500 m",
             "Map: our point", "Map: Google place"]
    front = [c for c in front if c in df.columns]
    rest = [c for c in df.columns if c not in front and c not in ("Latitude", "Longitude", "Description")]
    df = df[front + rest].rename(columns={"lat": "Latitude", "lon": "Longitude"})

    summary = df["Status"].value_counts().rename_axis("Status").reset_index(name="Count")
    flags = df["Pre-check flags"].str.split("; ").explode()
    flags = flags[flags.astype(bool)].str.replace(r"Coordinates ~\d+ km from State.*", "Coordinates far from stated State",
                                                  regex=True)
    flag_summary = flags.value_counts().rename_axis("Pre-check flag").reset_index(name="Count")

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        summary.to_excel(xw, "Summary", index=False)
        flag_summary.to_excel(xw, "Summary", index=False, startrow=len(summary) + 3)
        df.to_excel(xw, "All workshops", index=False)
        flagged = df["Pre-check flags"].fillna("") != ""
        offline = (df["Status"] == "OFFLINE ONLY").all()
        todo = df[flagged] if offline else df[(df["Status"] != "VERIFIED") | flagged]
        todo.to_excel(xw, "To check manually", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions if ws.title != "Summary" else None
            for col in ws.columns:
                width = min(max(len(str(c.value or "")) for c in col[:200]) + 2, 50)
                ws.column_dimensions[col[0].column_letter].width = width
            if ws.title != "Summary":
                headers = [c.value for c in ws[1]]
                for name in ("Map: our point", "Map: Google place"):
                    if name in headers:
                        j = headers.index(name) + 1
                        for r in range(2, ws.max_row + 1):
                            cell = ws.cell(r, j)
                            if cell.value:
                                cell.hyperlink, cell.value, cell.style = cell.value, "Open", "Hyperlink"
    print(f"Report written: {path}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
