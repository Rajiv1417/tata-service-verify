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

## Status meanings
| Status | Meaning |
|---|---|
| VERIFIED | Name matches (≥80%) and the place is within 200 m |
| CHECK - name OK, 200-500 m away | Probably correct, small coordinate shift |
| WRONG COORDS? - name found far away | Google has this workshop somewhere else; use Google lat/lon after checking |
| CHECK - similar name nearby | Partial name match (60–80%) within 200 m |
| CHECK - Tata place nearby, name differs | A Tata-branded listing is within 500 m but under another name |
| NOT FOUND | Nothing matching on Google; check manually or the workshop may not be listed |

**Pre-check flags** (free, offline): duplicate coordinates, coordinates outside India, coordinates far from the stated State.

Thresholds are at the top of `verify.py` (`OK_DIST_M`, `NAME_OK`, etc.).
Never commit the workshop data or reports; `.gitignore` already excludes them.
