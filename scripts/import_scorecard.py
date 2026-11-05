"""Convert Bruce's intake spreadsheet into src/data/communities.json.

Usage:  python3 -I scripts/import_scorecard.py path/to/intake.xlsx

- Only rows marked Include = Y are kept.
- Bruce's Notes column is NOT written out: this repo is public, and notes are
  his private observations until he approves specific ones for publishing.
- Prints a review report of anything that needs a human look.
"""
import json, re, sys, datetime
from pathlib import Path
from openpyxl import load_workbook

TODAY = datetime.date.today()
OUT = Path(__file__).resolve().parent.parent / "src" / "data" / "communities.json"

# Obvious typos fixed on import (sheet value -> published value).
NAME_FIXES = {
    "Godlen Nest ALF": "Golden Nest ALF",
    "Renassiance at The Terraces": "Renaissance at The Terraces",
    "Discovery Village Naples - independent": "Discovery Village Naples – Independent Living",
    "Discovery Village Naples - ALF/MC": "Discovery Village Naples – AL/MC",
    "Aston Gardens at Pelican Marsh - IL": "Aston Gardens at Pelican Marsh – Independent Living",
    "Aston Gardens at Pelican Marsh - AL/MC": "Aston Gardens at Pelican Marsh – AL/MC",
    "Colliers at Naples": "The Collier of Naples",  # confirmed by Bruce 2026-11-05
}
ADDRESS_FIXES = {"Rue Ju Jardin": "Rue du Jardin"}
# Care-level fixes where the row's own name contradicts the entry.
CARE_FIXES = {"Aston Gardens at Pelican Marsh - AL/MC": ["AL", "MC"]}
# Date typos confirmed by Bruce 2026-11-05 (sheet had 2016).
DATE_FIXES = {
    "Marbella at Pelican Bay": datetime.date(2026, 1, 16),
    "The Cove at Marbella": datetime.date(2026, 1, 16),
}
APP_NAME_FIXES = {"Senior Simon i": "Senior Simon"}
# Bruce: "N/A" in the family-app column means no app is offered.


def s(v):
    return "" if v is None else str(v).strip()


def slugify(t):
    return re.sub(r"[^a-z0-9]+", "-", t.lower()).strip("-")


def care(raw):
    found = set(re.findall(r"IL|AL|MC", raw.upper()))
    return [c for c in ("IL", "AL", "MC") if c in found]


def emergency(raw):
    l = raw.lower()
    fee = re.search(r"\((\$[^)]*)\)", raw)
    detail = fee.group(1).replace("for pendant", "pendant fee").strip() if fee else ""
    if not l:
        return None, ""
    if l.startswith("no"):
        return "No", ""
    if "al only" in l:
        return "AL only", detail
    if "both" in l:
        return "Pendant + wall cords", detail
    if l.startswith("yes"):
        return "Yes", detail
    return raw, detail


def fall(raw):
    l = raw.lower()
    if not l:
        return None, ""
    if l.startswith("no"):
        return "No", ""
    if l.startswith("in mc"):
        return "Memory care only", raw.split(",", 1)[-1].strip()
    if l.startswith("yes"):
        d = raw.split("-", 1)[1].strip() if "-" in raw else ""
        return "Yes", d
    return raw, ""


def emar(raw):
    l = raw.lower()
    if not l:
        return None, ""
    if l.startswith("transitioning"):
        return "Transitioning", "Eldermark → Alis"
    if l.startswith("yes"):
        d = raw.split("-", 1)[1].strip() if "-" in raw else ""
        return "Yes", d
    if l.startswith("no"):
        return "No", ""
    return raw, ""


def generator(raw, levels):
    l = raw.lower().strip()
    if not l:
        return None
    if l == "none":
        return "None"
    if "except apts" in l:
        return "Common areas"
    care_units = ("all of al" in l or "all of mc" in l or "memory care" in l
                  or re.search(r"\bmc\b", l) or "all of al and mc" in l)
    whole = any(k in l for k in ("entire", "whole", "all of facility", "all of facilty"))
    if whole:
        return "Whole building"
    if "all of al and mc" in l and set(levels) <= {"AL", "MC"} and "common" not in l:
        return "Whole building"
    if "clubhouse" in l:
        return "Clubhouse only"
    if l.startswith("elevators only"):
        return "Elevators only"
    if care_units:
        return "Common areas + care units"
    if "common" in l or "dining" in l:
        return "Common areas"
    return raw


def main(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if "Community name" in [s(v) for v in r])
    hdr = [s(v) for v in rows[hi]]
    col = {h: hdr.index(h) for h in hdr if h}

    def g(r, name):
        return r[col[name]] if name in col else None

    out, flags, skipped, gen_map = [], [], [], {}
    for r in rows[hi + 1:]:
        name = s(g(r, "Community name"))
        if not name or name.startswith("EXAMPLE"):
            continue
        if s(g(r, "Include? (Y/N)")).lower() != "y":
            skipped.append((name, "Include = N"))
            continue
        filled = [s(g(r, c)) for c in ("Emergency call system", "Fall detection",
                                         "eMAR (electronic med records)", "Generator coverage")]
        if not any(filled):
            skipped.append((name, "no tech answers filled in"))
            continue

        levels = CARE_FIXES.get(name) or care(s(g(r, "Care levels (IL / AL / MC)")))
        ec, ec_d = emergency(s(g(r, "Emergency call system")))
        fd, fd_d = fall(s(g(r, "Fall detection")))
        em, em_d = emar(s(g(r, "eMAR (electronic med records)")))
        gen_raw = s(g(r, "Generator coverage"))
        gen = generator(gen_raw, levels)
        gen_map.setdefault(gen, set()).add(gen_raw)
        wander_raw = s(g(r, "Memory care wander management")).lower()
        wander = None if not wander_raw else ("Yes" if wander_raw.startswith("y") else "No")
        if "MC" not in levels:
            wander = "N/A"
        app_raw = s(g(r, "Family app or portal"))
        app = "Yes" if app_raw.lower().startswith("y") else ("No" if app_raw else None)
        app_name = s(g(r, "App name (if known)")) if app == "Yes" else ""
        app_name = APP_NAME_FIXES.get(app_name, app_name)

        d = DATE_FIXES.get(name) or g(r, "Date last toured / verified")
        verified, via = None, None
        if isinstance(d, (datetime.datetime, datetime.date)):
            dd = d.date() if isinstance(d, datetime.datetime) else d
            verified, via = dd.isoformat(), "tour"
            if dd > TODAY:
                flags.append(f"{name}: tour scheduled {dd} (shown as 'Tour scheduled' until then)")
            elif dd.year < 2024:
                flags.append(f"{name}: verified date {dd} looks like a typo")
        elif s(d).lower().startswith("email"):
            via = "email"

        addr = s(g(r, "Address"))
        for a, b in ADDRESS_FIXES.items():
            addr = addr.replace(a, b)
        pub = NAME_FIXES.get(name, name)
        out.append({
            "slug": slugify(pub),
            "name": pub,
            "city": s(g(r, "City")),
            "address": addr,
            "care": levels,
            "emergencyCall": ec, "emergencyCallDetail": ec_d,
            "fallDetection": fd, "fallDetectionDetail": fd_d,
            "familyApp": app, "familyAppName": app_name,
            "emar": em, "emarDetail": em_d,
            "wifi": "Yes" if s(g(r, "Resident Wi-Fi")).lower().startswith("y") else None,
            "generator": gen,
            "wanderManagement": wander,
            "verified": verified, "verifiedVia": via,
        })

    order = {"Naples": 0, "Marco Island": 1, "Bonita Springs": 2, "Estero": 3,
             "Fort Myers": 4, "North Fort Myers": 5, "Cape Coral": 6, "Lehigh Acres": 7}
    out.sort(key=lambda c: (order.get(c["city"], 9), c["name"]))
    payload = {"lastUpdated": TODAY.isoformat(), "communities": out}
    OUT.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")

    print(f"Wrote {len(out)} communities -> {OUT}")
    print("\nSkipped:")
    for n, why in skipped:
        print(f"  - {n}: {why}")
    print("\nGenerator mapping (category <- Bruce's wording):")
    for k, v in gen_map.items():
        print(f"  {k}: {sorted(v)}")
    print("\nFlags:")
    for f in flags:
        print(f"  - {f}")


if __name__ == "__main__":
    main(sys.argv[1])
