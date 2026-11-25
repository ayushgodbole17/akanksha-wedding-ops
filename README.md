# Akanksha & Rishabh: wedding info page

Akanksha & Rishabh, 25–26 Nov 2026 (Kalyan) and 6 Dec 2026 (Bengaluru).

The family's plan lives in one Google Sheet. This page is a phone-friendly, information-only view of that sheet, with five screens: **Schedule** (what's on now, what's next, the full programme), **Rooms** (who's in which room at each hotel, searchable by name), **Drivers** (taxis and their numbers), **Contacts** (family and coordinators) and **Places** (venues and hotels with map links). Edits made in the sheet show up on the page within two minutes, or straight away with Refresh.

The page reads the **Guests**, **Rooms**, **Run sheet**, **Contacts** and **Vendors** tabs. Drivers are the rows in Contacts whose Role or Notes mention "taxi" or "driver".

This repo holds no guest data. The sheet's address travels inside the link you share.

## Setting it up (once)

1. Build the master workbook from the family's old Excel file:
   ```
   pip install openpyxl
   python tools/build_master.py "D:/Downloads/AK Invite List.xlsx" wedding-master.xlsx
   ```
   This prints a list of problems it found (conflicting spellings, people with no room, and so on). The same list is saved in the workbook's **Gaps** tab.
2. Open Google Sheets, then **File → Import → Upload**, choose `wedding-master.xlsx`, and pick **Replace spreadsheet**.
3. **Share → General access → Anyone with the link → Viewer.** Add the coordinators who will edit as **Editors**.
4. Open the page and paste the sheet's link. The page then gives you its own link (`…/?sheet=<id>`), which is the one to send to family. You can add `&tab=rooms` (or `schedule`, `drivers`, `contacts`, `places`) to open a particular screen.

## Rules for the sheet

1. Keep the tab names and column headings as they are, because the page looks for them.
2. Names in **In charge** (Run sheet) and **Point person** (Vendors) should match a name in **Contacts**. That's how the Call buttons find the number.
3. Room codes in **Guests → Room** must match **Rooms → Room**.

## Testing locally

```
python tools/export_csv.py wedding-master.xlsx test-data
python -m http.server 8000
```
Then open http://localhost:8000/?src=test-data. The `test-data/` folder and the `.xlsx` files are git-ignored so that guest data stays out of the repo.
