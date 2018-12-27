# JournalHoldingsCount

This script uses OCLC's WorldCat Search API to retrieve a list and count of academic libraries in Virginia holding certain journal titles.

Input spreadsheet has one journal title per row, with eISSN in column H and ISSN in column I.

List of holding libraries will be added as column U and count as column V.

---

Requires a config file (local_settings.ini) in the same directory. Example of local_settings.ini:
```
[WorldCat Search API]
wskey:key
```

To set up environment: ```conda env create --file environment.yaml```

To run: ```python JournalHoldingsCount.py spreadsheet.xlsx```