"""Turn the family's 'AK Invite List.xlsx' into one clean master workbook.

Usage:  python tools/build_master.py "D:/Downloads/AK Invite List.xlsx" wedding-master.xlsx

The source sheet holds the guest list twice (two slightly different copies),
plus three blocks of room groupings. This script merges the copies, links
people to rooms, adds the planning tabs (jobs, vendors, run sheet, contacts)
and lists every gap it finds, both on screen and in a 'Gaps' tab.
"""
import sys
from datetime import date, time

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SRC, OUT = sys.argv[1], sys.argv[2]

# ---------------------------------------------------------------- cleaning

def clean(s):
    return " ".join(str(s).split()) if s is not None else ""

# Spellings that differ between the copies or the room plans. Left side is
# what the sheet says, right side is the one spelling we keep. Every merge
# made here is listed in the Gaps tab so the family can check it.
CANON = {
    "anil deshmukh": "Sunil Deshmukh",        # copy 1 says Anil, copy 2 and room plan say Sunil
    "nevidata kulkarni": "Nivedita Kulkarni",
    "rraghavendra prasad": "Raghavendra Prasad",
    "raavi": "Raavi Burve",
    "neil": "Neil Burve",
    "raja shahpure mrs": "Mrs Raja Shahpure",
    "mrs shahpure": "Mrs Raja Shahpure",
    "mrudala khare": "Mrudula Khare",           # lists say Mrudala, room plan says Mrudula
    "shree prasad godbole": "Sree Prasad Godbole",
    "jothi": "Jothi Godbole",
    "girish kaka": "Girish Godbole",
    "aditi kaku": "Aditi Godbole",
}
canon_used = {}

def canon(name):
    n = clean(name)
    c = CANON.get(n.lower())
    if c and c != n:
        canon_used.setdefault(c, set()).add(n)
        return c
    return n

src = load_workbook(SRC, data_only=True)["Train Tickets"]

def v(ref):
    return clean(src[ref].value)

# ---------------------------------------------------------------- guests

# (list name in sheet, column, travel mode, coming from)
SECTIONS = [
    ("Bangalore to Kalyan", "C", "Train", "Bangalore"),
    ("Flight", "E", "Flight", ""),
    ("Pune", "G", "", "Pune"),
    ("Other places", "I", "", ""),
]
COPY1, COPY2 = range(6, 45), range(70, 109)

gaps = []  # (type, detail, what to do)
guests = []  # dicts
by_name = {}

for label, col, mode, city in SECTIONS:
    one = [canon(v(f"{col}{r}")) for r in COPY1 if v(f"{col}{r}")]
    two = [canon(v(f"{col}{r}")) for r in COPY2 if v(f"{col}{r}")]
    for n in sorted(set(one) - set(two)):
        gaps.append(("Copies disagree", f"{n} is in the first copy of the '{label}' list but not the second.",
                     "Kept them. Check they are still coming."))
    for n in sorted(set(two) - set(one)):
        gaps.append(("Copies disagree", f"{n} is in the second copy of the '{label}' list but not the first.",
                     "Kept them. Check they are still coming."))
    for n in two + [x for x in one if x not in two]:
        if n.lower() in by_name:
            gaps.append(("Listed twice", f"{n} appears in more than one travel list.", "Kept the first one."))
            continue
        g = dict(name=n, group=label, mode=mode, city=city, room="", notes=[])
        guests.append(g)
        by_name[n.lower()] = g

INVITE_COUNT = len(guests)

# ---------------------------------------------------------------- rooms

def read_rooms(set_name, rows, pairs):
    """Each (number col, name col) pair is a column of rooms: a number cell
    starts a new room, names under it belong to that room."""
    rooms, seen = [], {}
    for numcol, namecol in pairs:
        current = None
        for r in rows:
            num, nm = v(f"{numcol}{r}"), v(f"{namecol}{r}")
            if num:
                base = f"{set_name}-{num}"
                seen[base] = seen.get(base, 0) + 1
                code = base if seen[base] == 1 else f"{base}{'abcdefgh'[seen[base] - 1]}"
                if seen[base] == 2:
                    gaps.append(("Room number reused", f"Room number {num} is used more than once in room set {set_name}.",
                                 f"Renamed them {base}, {base}b, ... Give each its real room number."))
                current = dict(code=code, set=set_name, people=[])
                rooms.append(current)
            if nm and current is not None:
                current["people"].append(nm)
    return rooms

rooms = (read_rooms("A", range(51, 67), [("B", "C"), ("D", "E"), ("F", "G")])
         + read_rooms("B", range(114, 180), [("B", "C")])
         + read_rooms("C", range(114, 180), [("F", "G")]))

nums_b = sorted(int(r["code"].split("-")[1]) for r in rooms if r["set"] == "B")
missing_b = [n for n in range(1, max(nums_b) + 1) if n not in nums_b]
if missing_b:
    gaps.append(("Room number missing", f"Room set B skips number(s) {missing_b}.", "Check whether a room was left out."))

PLACEHOLDER = {"son", "daughter", "??"}
room_rows = []
for rm in rooms:
    occupants = []
    for raw in rm["people"]:
        if raw.lower() == "makeup artists":
            occupants.append("Makeup artist")
            continue
        n = canon(raw)
        if raw.lower() in PLACEHOLDER:
            first = occupants[0] if occupants else "?"
            n = "Unknown person (sheet says ??)" if raw == "??" else f"{raw.title()} of {first}"
            n = f"{n} [{rm['code']}]" if raw == "??" else n
        if n in occupants:  # same name written twice in one room = a second person
            n = f"{n} (2nd person)"
        g = by_name.get(n.lower())
        if g is None:
            g = dict(name=n, group="Not on invite list", mode="", city="", room="", notes=[])
            guests.append(g)
            by_name[n.lower()] = g
            gaps.append(("Not on invite list", f"{n} is in room {rm['code']} but not on any travel list.",
                         "Added to Guests. Fill in their real name and how they travel."))
        if g["room"]:
            gaps.append(("Two rooms", f"{n} is in room {g['room']} and room {rm['code']}.", "Pick one."))
        else:
            g["room"] = rm["code"]
        occupants.append(n)
    if not occupants:
        gaps.append(("Empty room", f"Room {rm['code']} has a number but nobody in it.", "Fill it or delete it."))
    room_rows.append(dict(code=rm["code"], set=rm["set"], people=occupants))

for c, variants in canon_used.items():
    gaps.insert(0, ("Spelling picked", f"Used '{c}' for: {', '.join(sorted(variants))}.",
                    "Correct it in the Guests tab if that's wrong."))

no_room = [g["name"] for g in guests if not g["room"]]
if no_room:
    gaps.append(("No room", f"{len(no_room)} people have no room: {', '.join(no_room)}.",
                 "Assign a room, or mark 'Day visit / own stay' in their Notes."))
gaps.append(("Hotel unknown", "None of the three room sets (A, B, C) says which hotel it belongs to.",
             "Fill the Hotel column in the Rooms tab."))
gaps.append(("Travel details", "No arrival dates, train/flight numbers, PNRs or phone numbers are in the sheet yet.",
             "Collect them into the Guests tab. Pickups are planned from this."))

# ---------------------------------------------------------------- plan tabs

D = date
JOBS = [
    # area, job, when, deadline, owner
    ("Coordination", "Overall coordination and link with the Singh family", "Until 6 Dec", None, "Ayush Godbole"),
    ("Coordination", "Final headcount for each meal, sent to caterers", "Before 25 Nov", D(2026, 11, 20), ""),
    ("Coordination", "Confirm which family arranges what at Saxena Lawns (reception, dinner, decor)", "26 Nov", D(2026, 11, 10), "Ayush Godbole"),
    ("Travel", "Bangalore to Kalyan train tickets (21 people): book and collect PNRs", "", D(2026, 11, 10), ""),
    ("Travel", "Collect flight details of flying guests (21 people)", "", D(2026, 11, 15), ""),
    ("Travel", "Collect travel plans of Pune and other-city guests", "", D(2026, 11, 15), ""),
    ("Travel", "Book cars for station and airport pickups", "24-26 Nov", D(2026, 11, 18), ""),
    ("Travel", "Pickup roster: who collects whom, and when", "24-26 Nov", D(2026, 11, 20), ""),
    ("Travel", "Return tickets and drops to station/airport", "27 Nov onwards", D(2026, 11, 20), ""),
    ("Stay", "Confirm hotel(s) and number of rooms", "", D(2026, 11, 10), ""),
    ("Stay", "Final room allocation (fill the Rooms tab)", "", D(2026, 11, 18), ""),
    ("Stay", "Check-in desk at hotel with room list and keys", "24-25 Nov", D(2026, 11, 24), ""),
    ("Stay", "Welcome kit in each room: water, snacks, printed schedule", "24 Nov", D(2026, 11, 23), ""),
    ("Venues", "Point person at Nakshatra Banquet Hall", "25-26 Nov", D(2026, 11, 15), ""),
    ("Venues", "Point person at Saxena Dream Lawns", "26 Nov", D(2026, 11, 15), ""),
    ("Venues", "Move guests from Nakshatra to Saxena Lawns", "26 Nov, 3-4 PM", D(2026, 11, 20), ""),
    ("Food", "Dinner on 25 Nov", "25 Nov", D(2026, 11, 15), ""),
    ("Food", "Breakfast and lunch on 26 Nov", "26 Nov", D(2026, 11, 15), ""),
    ("Food", "Vihinichi Pangat (feast for both sets of parents)", "26 Nov, 2:30 PM", D(2026, 11, 15), ""),
    ("Flowers & decor", "Decor and flowers at Nakshatra (Simant Pujan, mandap, Sangeet stage)", "25-26 Nov", D(2026, 11, 12), ""),
    ("Flowers & decor", "Garlands for the ceremony, bride's flowers, puja flowers", "25-26 Nov", D(2026, 11, 18), ""),
    ("Rituals", "Book the pandit (guruji) and agree dakshina", "25-26 Nov", D(2026, 11, 10), ""),
    ("Rituals", "Puja items for Simant Pujan and Vangnischay", "25 Nov", D(2026, 11, 20), ""),
    ("Rituals", "Puja items for Mangalashtak, Kanyadaan, Lajahom, Saptapadi, Sunmukh, Zhal", "26 Nov", D(2026, 11, 20), ""),
    ("Entertainment", "Book mehendi artist(s)", "25 Nov, 7:30 PM", D(2026, 11, 12), ""),
    ("Entertainment", "Sangeet: music, sound system, list of performances", "25 Nov, 7:30 PM", D(2026, 11, 18), ""),
    ("Bride", "Bride's outfits and jewellery, one packed bag per event", "25-26 Nov", D(2026, 11, 22), ""),
    ("Bride", "Makeup artists: booking, timings, their room (A-3)", "25-26 Nov", D(2026, 11, 12), ""),
    ("Photo & video", "Photographer and videographer: booking and list of must-have shots", "25-26 Nov", D(2026, 11, 12), ""),
    ("Gifts & money", "Gifts for the groom's family and return gifts for guests", "26 Nov", D(2026, 11, 20), ""),
    ("Gifts & money", "Cash, envelopes and tracking of vendor payments (Vendors tab)", "Ongoing", D(2026, 11, 22), ""),
    ("Guest care", "Elders' care: seating, medicines, wheelchair if needed", "25-26 Nov", D(2026, 11, 20), ""),
    ("Guest care", "First-aid kit and nearest hospital/doctor noted", "25-26 Nov", D(2026, 11, 22), ""),
    ("Guest care", "Bride's room at each venue, ladies' helper", "25-26 Nov", D(2026, 11, 20), ""),
    ("After", "Bidaai arrangements", "26 Nov, 10:30 PM", D(2026, 11, 22), ""),
    ("After", "Bengaluru reception: Vasavi Mahal, caterer, invites", "6 Dec, 11 AM", D(2026, 11, 20), ""),
    ("After", "Settle all vendor balances", "27 Nov - 7 Dec", D(2026, 12, 7), ""),
]

VENDORS = [
    ("Venue", "Nakshatra Banquet Hall", "25-26 Nov",
     "Kalyan Sports Club, near Adharwadi Chowk, Durgadi Birla College Road, Kalyan West 421301"),
    ("Venue", "Saxena Dream Lawns", "26 Nov",
     "Gandhari, Barave Road, opp. Mahada Colony, Godrej Hill, Khadakpada, Kalyan 421301"),
    ("Venue", "Vasavi Mahal", "6 Dec",
     "4 Vanivilas Road, 3rd Floor (above Vasavi Convention Hall), VV Puram, Basavanagudi, Bengaluru 560004"),
    ("Hotel", "", "24-27 Nov", ""), ("Caterer", "", "25-26 Nov", ""), ("Flowers & decor", "", "25-26 Nov", ""),
    ("Pandit", "", "25-26 Nov", ""), ("Mehendi artist", "", "25 Nov", ""), ("Sound / DJ", "", "25 Nov", ""),
    ("Makeup artists", "", "25-26 Nov", ""), ("Photographer", "", "25-26 Nov", ""),
    ("Videographer", "", "25-26 Nov", ""), ("Cars / taxis", "", "24-27 Nov", ""),
]

NAK, SAX, VAS, HOT = "Nakshatra Banquet Hall", "Saxena Dream Lawns", "Vasavi Mahal, Bengaluru", "Hotel"
RUN = [
    # date, start, end, what, venue, source
    (D(2026, 11, 24), None, None, "Early arrivals, pickups and hotel check-in", HOT, "Suggested"),
    (D(2026, 11, 25), None, time(14, 0), "Arrivals, pickups and hotel check-in", HOT, "Suggested"),
    (D(2026, 11, 25), time(14, 0), time(15, 30), "Decor check, puja items laid out", NAK, "Suggested"),
    (D(2026, 11, 25), time(15, 0), None, "Tilak (groom's side)", NAK, "Singh invite"),
    (D(2026, 11, 25), time(16, 0), time(18, 30), "Welcome the groom, Simant Pujan and Vangnischaya", NAK, "Programme"),
    (D(2026, 11, 25), time(19, 30), None, "Sangeet and Mehendi", NAK, "Programme"),
    (D(2026, 11, 25), time(20, 30), None, "Dinner", NAK, "Programme"),
    (D(2026, 11, 26), time(7, 0), time(9, 0), "Breakfast", NAK, "Programme"),
    (D(2026, 11, 26), time(8, 30), time(9, 15), "Mangalashtak and the wedding muhurat (9:15 AM)", NAK, "Programme"),
    (D(2026, 11, 26), time(9, 30), time(11, 0), "Blessings for the bride and groom from family", NAK, "Programme"),
    (D(2026, 11, 26), time(11, 30), time(14, 30), "Kanyadaan, Lajahom, Saptapadi, Sunmukh and Zhal", NAK, "Programme"),
    (D(2026, 11, 26), time(12, 0), None, "Lunch", NAK, "Programme"),
    (D(2026, 11, 26), time(14, 30), time(15, 0), "Vihinichi Pangat (feast for both sets of parents)", NAK, "Programme"),
    (D(2026, 11, 26), time(15, 0), time(16, 0), "Move guests to Saxena Lawns", "Nakshatra to Saxena", "Suggested"),
    (D(2026, 11, 26), time(16, 0), time(18, 0), "Wedding ceremonies as per Rajput traditions", SAX, "Programme"),
    (D(2026, 11, 26), time(19, 30), None, "Reception, followed by dinner", SAX, "Programme"),
    (D(2026, 11, 26), time(22, 30), None, "Bidaai", SAX, "Programme"),
    (D(2026, 11, 26), time(23, 0), None, "Drop guests back to hotel", "Saxena to hotel", "Suggested"),
    (D(2026, 12, 6), time(11, 0), None, "Reception in Bengaluru", VAS, "Godbole invite"),
]

CONTACTS = [
    ("Mahendra Godbole", "Bride's father", "+91 98450 46767"),
    ("Tejaswini Godbole", "Bride's mother", "+91 99455 67537"),
    ("Ayush Godbole", "Overall coordinator, link with Singh family", ""),
    ("Rajendra R Singh", "Groom's father", "+91 81085 15943"),
    ("Chhaya Rajendra Singh", "Groom's mother", "+91 90298 37129"),
]

# ---------------------------------------------------------------- writing

HEAD = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="6B2D5C")
WRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()

def sheet(title, headers, rows, widths):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for c in ws[1]:
        c.font, c.fill, c.alignment = HEAD, HEAD_FILL, WRAP
    for r in rows:
        ws.append(list(r))
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = WRAP
    return ws

def dropdown(ws, col, options, n=400):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col}2:{col}{n}")

how = wb.active
how.title = "How to use"
for line in [
    "Akanksha & Rishabh: wedding master sheet",
    "",
    "One row per thing. Edit cells directly; keep the column headings as they are (the phone page reads them).",
    "Guests: everyone coming, their travel, room and pickup. Room codes match the Rooms tab.",
    "Rooms: who sleeps where. Fill the Hotel column.",
    "Jobs: who is in charge of what. Put one name in Owner; set Status as it moves.",
    "Vendors: who is booked, what was agreed, what is paid.",
    "Run sheet: hour by hour for 25-26 Nov and 6 Dec, with who is in charge of each slot.",
    "Contacts: coordinators and key people. Names in Owner / Pickup by should match names here.",
    "Gaps: problems found when this sheet was built. Delete a row once it is sorted.",
    "",
    "Suggested deadlines and 'Suggested' run-sheet rows are proposals, change them freely.",
]:
    how.append([line])
how["A1"].font = Font(bold=True, size=14)
how.column_dimensions["A"].width = 110

GUEST_COLS = ["ID", "Name", "List in old sheet", "Travel mode", "Coming from", "Room", "Phone",
              "Attending 25 Nov", "Attending 26 Nov", "Attending 6 Dec",
              "Arrival date", "Arrival time", "Train / flight no", "PNR", "Arrives at",
              "Pickup needed", "Pickup by", "Departure date", "Departure time", "Departure train / flight", "Notes"]
gs = sheet("Guests", GUEST_COLS,
           [[f"G{i:03d}", g["name"], g["group"], g["mode"], g["city"], g["room"]] + [""] * 15
            for i, g in enumerate(guests, 1)],
           [7, 28, 18, 11, 13, 7, 16, 9, 9, 9, 12, 10, 14, 13, 18, 9, 18, 12, 10, 16, 30])
for col in "HIJP":
    dropdown(gs, col, ["Yes", "No", "Maybe"])
# Free-typed columns stay plain text so numbers (PNRs, phones) and words
# ("TBC") can share a column without Google's CSV export dropping either.
for col in "GKLMNQRSTU":
    for r in range(2, 401):
        gs[f"{col}{r}"].number_format = "@"
dropdown(gs, "D", ["Train", "Flight", "Car", "Bus", "Local"])
dropdown(gs, "O", ["Kalyan station", "Mumbai airport", "Thane station", "CSMT", "LTT", "Dadar", "Own transport"])

sheet("Rooms", ["Room", "Hotel", "Room number at hotel", "People", "Count", "Check-in", "Check-out", "Notes"],
      [[r["code"], "", "", ", ".join(r["people"]), len(r["people"]), "", "", ""] for r in room_rows],
      [8, 22, 12, 60, 7, 12, 12, 30])

js = sheet("Jobs", ["ID", "Area", "Job", "When", "Owner", "Helper", "Suggested deadline", "Status", "Notes"],
           [[f"J{i:02d}", a, j, w, o, "", d, "Not started", ""] for i, (a, j, w, d, o) in enumerate(JOBS, 1)],
           [6, 15, 55, 16, 18, 16, 12, 12, 30])
dropdown(js, "H", ["Not started", "In progress", "Done", "Stuck"])
for c in js["G"][1:]:
    c.number_format = "d mmm"

vs = sheet("Vendors", ["Category", "Vendor name", "For", "Address", "Contact person", "Phone", "Booked",
                       "Agreed price", "Advance paid", "Balance", "Point person", "Notes"],
           # venues are printed on the invitations, so they are booked
           [[c, n, f, a, "", "", "Yes" if c == "Venue" else "", None, None, None, "", ""] for c, n, f, a in VENDORS],
           [15, 24, 11, 40, 16, 16, 8, 11, 11, 11, 16, 30])
for r in range(2, len(VENDORS) + 2):
    vs[f"J{r}"] = f'=IF(H{r}="","",H{r}-N(I{r}))'
dropdown(vs, "G", ["Yes", "No", "Talking"])

rs = sheet("Run sheet", ["Date", "Start", "End", "What", "Where", "In charge", "Source", "Notes"],
           [[d, s, e, w, p, "", src_, ""] for d, s, e, w, p, src_ in RUN],
           [10, 8, 8, 50, 26, 18, 13, 30])
for row in rs.iter_rows(min_row=2):
    row[0].number_format = "d mmm"
    row[1].number_format = row[2].number_format = "h:mm AM/PM"

sheet("Contacts", ["Name", "Role", "Phone", "Side", "Notes"],
      [[n, r, p, "Singh" if "Singh" in n else "Godbole", ""] for n, r, p in CONTACTS],
      [24, 40, 18, 10, 30])

sheet("Gaps", ["Type", "What's wrong", "What to do"], gaps, [18, 80, 50])

wb.save(OUT)

# ---------------------------------------------------------------- report

print(f"Invite list: {INVITE_COUNT} people (sheet's own total: 101)")
print(f"Guests tab: {len(guests)} rows ({len(guests) - INVITE_COUNT} added from room plans)")
print(f"Rooms: {len(room_rows)} across sets A, B, C; people with no room: {len(no_room)}")
print(f"Gaps: {len(gaps)}\n")
for t, d, w in gaps:
    print(f"- [{t}] {d}")
