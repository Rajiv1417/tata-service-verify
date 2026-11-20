# Tata workshop location verifier

Checks each workshop in the My Maps export against Google Places and writes `verification_report.xlsx`.

## Run in Codespaces
```bash
pip install -r requirements.txt
# drag Workshopdata.kml into the file panel, then:
python verify.py Workshopdata.kml --offline        # free pre-checks only, no API calls
python verify.py Workshopdata.kml --limit 30       # pilot: first 30 rows
python verify.py Workshopdata.kml                  # full run
```
The API key comes from the `GOOGLE_MAPS_API_KEY` Codespaces secret.
Results are cached in `places_cache.json`. If the run stops, run the same command again: finished rows are not called (or billed) again.

## How it searches
1. Searches the workshop name **only within ~1 km of the given pin**.
2. If not found there, looks for any Tata-branded place within 500 m.
3. If still not found, searches the name (+ PIN code from the Address) over a wider area and
   compares Google's address with ours to tell "pin is wrong" from "that's another branch".

## Status meanings
| Status | Meaning | Action |
|---|---|---|
| VERIFIED | Name matches and place is within 200 m | none |
| CHECK - name OK, 200-500 m away | Small shift | glance at map |
| CHECK - name OK, 0.5-1.5 km away | Pin is off by up to ~1.5 km | fix pin |
| CHECK - similar name nearby | Partial name match within 200 m | confirm name |
| CHECK - Tata place near pin, name differs | A Tata listing is near the pin under another name | confirm |
| WRONG COORDS - workshop found at its address, away from pin | Google has it at the PIN/locality in our Address | replace with Google lat/lon after a look |
| NOT AT PIN - same name found elsewhere (other branch?) | Same dealer name, different town | check manually |
| NOT FOUND | Nothing on Google near the pin | check manually |

**Pre-check flags** (free, offline): duplicate coordinates, coordinates outside India, coordinates far from the stated State.

Thresholds are at the top of `verify.py` (`OK_DIST_M`, `NAME_OK`, etc.).
Never commit the workshop data or reports; `.gitignore` already excludes them.
