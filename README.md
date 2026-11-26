# SSB-Projects: NJ Rebate Lookup

Phone-friendly search over the NJ Clean Energy SSB rebate sheet. Type any part of an address,
business name, city or ZIP; results are grouped by building, with a stamp showing whether a
lighting rebate was ever paid there.

Live site: https://jermy-noienergygroup.github.io/SSB-Projects/

## Updating when a new sheet comes out
```
pip install openpyxl
python build_data.py "SSB-Projects-YYYY-MM-DD.xlsx"
```
This rewrites `data.json`. Upload the new `data.json` to the repo (replace the old one); the site updates in a minute or two.
The script looks for tabs whose names start with `Retro`, `New Construction`, `P4P EB` and `P4P NC`.
