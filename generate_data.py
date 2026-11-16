import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(42)  # same "random" data every run
OUT = Path("data/raw")
OUT.mkdir(parents=True, exist_ok=True)
TODAY = date(2026, 10, 1)

FIRST = ["Ana", "Ben", "Chen", "Dana", "Eli", "Fatima", "Grace", "Hiro", "Ivan", "Jia",
         "Kofi", "Lena", "Marco", "Nia", "Omar", "Priya", "Quinn", "Rosa", "Sam", "Tara"]
LAST = ["Abbott", "Baker", "Cruz", "Diaz", "Evans", "Fischer", "Gupta", "Hughes", "Ito", "Jones",
        "Kim", "Lopez", "Mensah", "Novak", "Okafor", "Patel", "Reyes", "Silva", "Tanaka", "Wu"]
CHEMS = [("Acetone", "67-64-1", "Flammable"), ("Methanol", "67-56-1", "Flammable"),
         ("Ethanol", "64-17-5", "Flammable"), ("Isopropanol", "67-63-0", "Flammable"),
         ("Hexane", "110-54-3", "Flammable"), ("Toluene", "108-88-3", "Flammable"),
         ("Acetonitrile", "75-05-8", "Flammable"), ("Chloroform", "67-66-3", "Toxic"),
         ("Formaldehyde", "50-00-0", "Toxic"), ("Hydrochloric acid", "7647-01-0", "Corrosive"),
         ("Sulfuric acid", "7664-93-9", "Corrosive"), ("Sodium hydroxide", "1310-73-2", "Corrosive")]
TYPOS = {"Acetone": ["acetone", "Acetone ", "Acteone"], "Methanol": ["MeOH", "methanol"],
         "Isopropanol": ["Isopropyl alcohol", "IPA"], "Hydrochloric acid": ["HCl", "Hydrochloric Acid"]}
UNIT_VARIANTS = {"mL": ["ml", "ML", "milliliters"], "g": ["G", "grams"]}
HIGH_TURNOVER = {"L04", "L11"}  # seeded pattern your pipeline should detect


def write(name, rows):
    with open(OUT / name, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)


def rand_date(days_back):
    return TODAY - timedelta(days=random.randint(0, days_back))


# Labs and rooms
labs, rooms = [], []
for i in range(1, 21):
    lab_id = f"L{i:02d}"
    building = random.choice(["Building A", "Building B", "Building C", "Building D"])
    labs.append({"lab_id": lab_id, "pi_name": f"{random.choice(FIRST)} {random.choice(LAST)}",
                 "building": building})
    floor = random.randint(2, 9)
    for _ in range(2):
        rooms.append({"room_id": f"R{len(rooms) + 1:03d}", "lab_id": lab_id,
                      "building": building, "room_number": f"{floor}{random.randint(10, 49)}"})

# People (high-turnover labs have more stale contact records)
people = []
for lab in labs:
    stale_window = 700 if lab["lab_id"] in HIGH_TURNOVER else 300
    for _ in range(random.randint(4, 8)):
        fn, ln = random.choice(FIRST), random.choice(LAST)
        people.append({"person_id": f"P{len(people) + 1:03d}", "lab_id": lab["lab_id"],
                       "name": f"{fn} {ln}",
                       "email": f"{fn[0].lower()}{ln.lower()}{random.randint(1, 99)}@example.edu",
                       "role": random.choice(["Grad student", "Postdoc", "Lab manager", "Technician"]),
                       "last_confirmed": rand_date(stale_window).isoformat()})

# True inventory (what is really on the shelves)
inventory = []
for lab in labs:
    lab_rooms = [r["room_id"] for r in rooms if r["lab_id"] == lab["lab_id"]]
    for _ in range(random.randint(15, 40)):
        name, cas, hazard = random.choice(CHEMS)
        inventory.append({"container_id": f"C{len(inventory) + 1:05d}", "lab_id": lab["lab_id"],
                          "room_id": random.choice(lab_rooms), "chemical_name": name,
                          "cas_number": cas, "hazard_class": hazard,
                          "quantity": random.choice([100, 250, 500, 1000, 2500]),
                          "unit": random.choice(["mL", "g"]),
                          "last_verified": rand_date(400).isoformat()})

# Field verification: 14 visited labs, including both high-turnover labs
others = [l["lab_id"] for l in labs if l["lab_id"] not in HIGH_TURNOVER]
visited = sorted(set(random.sample(others, 12)) | HIGH_TURNOVER)
field = []
for c in inventory:
    if c["lab_id"] not in visited:
        continue
    err = 0.25 if c["lab_id"] in HIGH_TURNOVER else 0.08
    r = random.random()
    if r < err / 3:
        continue  # container is gone: missing on shelf
    row = {"visit_date": rand_date(30).isoformat(), "lab_id": c["lab_id"],
           "container_id": c["container_id"], "room_id": c["room_id"],
           "chemical_name": c["chemical_name"], "quantity": c["quantity"], "unit": c["unit"]}
    if r < 2 * err / 3:  # found in the lab's other room
        row["room_id"] = [x["room_id"] for x in rooms
                          if x["lab_id"] == c["lab_id"] and x["room_id"] != c["room_id"]][0]
    elif r < err:  # partly used up
        row["quantity"] = int(c["quantity"] * random.choice([0.2, 0.5]))
    field.append(row)
for lab_id in visited:  # containers on the shelf that were never entered
    extras = random.randint(3, 6) if lab_id in HIGH_TURNOVER else random.randint(0, 2)
    lab_rooms = [x["room_id"] for x in rooms if x["lab_id"] == lab_id]
    for k in range(extras):
        name, cas, hazard = random.choice(CHEMS)
        field.append({"visit_date": rand_date(30).isoformat(), "lab_id": lab_id,
                      "container_id": f"NEW-{lab_id}-{k + 1}", "room_id": random.choice(lab_rooms),
                      "chemical_name": name, "quantity": 500, "unit": "mL"})

# System export: the true inventory plus typical data-entry mess
system = []
for c in inventory:
    row = dict(c)
    if row["chemical_name"] in TYPOS and random.random() < 0.15:
        row["chemical_name"] = random.choice(TYPOS[row["chemical_name"]])
    if random.random() < 0.03:
        row["cas_number"] = ""
    if random.random() < 0.10:
        row["unit"] = random.choice(UNIT_VARIANTS[row["unit"]])
    if random.random() < 0.05:  # US-style date mixed in
        y, m, d = row["last_verified"].split("-")
        row["last_verified"] = f"{m}/{d}/{y}"
    system.append(row)
    if random.random() < 0.03:
        system.append(dict(row))  # duplicate record

# Inspection findings
inspections = []
cats = ["Chemical storage", "Labeling", "PPE", "Waste", "Eyewash/shower", "Training records"]
for lab in labs:
    n = random.randint(4, 8) if lab["lab_id"] in HIGH_TURNOVER else random.randint(1, 4)
    for _ in range(n):
        found = rand_date(365)
        closed = found + timedelta(days=random.randint(3, 90))
        is_open = closed > TODAY or random.random() < 0.2
        inspections.append({"finding_id": f"F{len(inspections) + 1:04d}", "lab_id": lab["lab_id"],
                            "date_found": found.isoformat(), "category": random.choice(cats),
                            "severity": random.choice(["Low", "Medium", "High"]),
                            "status": "Open" if is_open else "Closed",
                            "date_closed": "" if is_open else closed.isoformat()})

write("labs.csv", labs)
write("rooms.csv", rooms)
write("people.csv", people)
write("chemicals.csv", [{"cas_number": c, "chemical_name": n, "hazard_class": h} for n, c, h in CHEMS])
write("inventory_system_export.csv", system)
write("field_verification.csv", field)
write("inspection_findings.csv", inspections)
print(f"{len(system)} system rows, {len(field)} field rows, visited labs: {', '.join(visited)}")
