import pymupdf

SRC = "/workspace/26-27_Suites-and-Apt-Living-Agreement.pdf"
OUT = "/workspace/26-27_Suites-and-Apt-Living-Agreement.filled.tmp.pdf"

doc = pymupdf.open(SRC)

# Each entry: (page_index, y_order_index_among_placeholders_on_page, y1_limit, fontsize, fontname, text)
# y_order_index is the rank (0-based) of the placeholder on that page sorted top->bottom.

RESP_CLEAN = (
    "Kitchen: dishes are washed daily and never left overnight; the refrigerator, stove, "
    "and common counters are wiped down every Sunday. Bathroom: everyone wipes the sink "
    "after use and we deep-clean the toilet, shower, and floor weekly on a rotating basis. "
    "Chore plan: we rotate trash, bathroom, and common-area duties each week and track them "
    "on a shared whiteboard. Cleaning supplies are split evenly; whoever notices we're low "
    "restocks them and we settle up on a shared expense list. Allergies: Edwin has a mild cat "
    "allergy, so no cats visit the unit; Chris has no known allergies. We vacuum regularly and "
    "keep pet dander out of shared spaces."
)

RESP_GUESTS = (
    "Guests are welcome in common areas with a quick heads-up text, and at least a day's notice "
    "for overnight guests. We follow USC policy: no more than 3 consecutive nights, twice a month, "
    "and 2 guests per resident in a suite. Guests do not use a roommate's belongings without asking. "
    "If one of us wants a guest to leave, we'll say so directly and privately. Guests of any sex, "
    "family, and partners are all fine with advance notice and mutual respect."
)

RESP_SLEEP = (
    "Edwin is usually asleep by midnight and up around 8am; Chris tends to stay up later (around 1am) "
    "and wakes near 9am. If one of us comes in while the other is sleeping, we keep the lights off, "
    "use a phone flashlight, and stay quiet. Snooze is fine but limited to two alarms, then get up so "
    "it doesn't disturb the other person. During sleep hours we keep overhead lights off (desk lamp "
    "only), hold the thermostat around 70 F, and use headphones for music or video. We both honor USC "
    "Quiet Hours (11pm-8am Sun-Thu, midnight-8am Fri-Sat, and 24/7 during Study Days and finals)."
)

RESP_SHARING = (
    "Shareable with permission: snacks, coffee, and school supplies. Personal items (toiletries, "
    "clothing, chargers) are generally not shared. Fridge, closet, and storage are split evenly. "
    "We keep the space around 70 F at night and 72 F during the day, and open the windows in the "
    "morning when the weather is nice. While studying, calls, TV, and music go through headphones or "
    "move to common areas so the other person can focus."
)

RESP_SPACE = (
    "Video/phone calls in-room before 10pm; after that, take them to common areas. Study groups are "
    "fine with a heads-up. TV and noise quiet down by about 11pm on weeknights."
)

RESP_COMM = (
    "We communicate best by text for day-to-day things and in person for anything important or "
    "sensitive. If either of us wants privacy or alone time, we'll send a quick text or use a door "
    "signal, and the other will give space. If a roommate gets sick, let the other know by text so we "
    "can keep our distance, disinfect shared surfaces, and help pick up food or supplies if needed."
)

RESP_CONFLICT = (
    "We'll raise concerns directly, calmly, and early, in person when possible. If something can't be "
    "resolved between us, we'll bring in our RA together for support."
)

RESP_THINGS = (
    "Pet peeves: dishes left overnight and borrowing without asking. We're both focused on academics, "
    "so a quiet study environment matters. No smoking in the unit; alcohol only if of age and kept "
    "low-key. We'll respect each other's religious practices and health needs and check in regularly "
    "to keep things comfortable."
)

RESP_SIG_RES = "Edwin Huang  -  08/26/2026\nChris Gardner  -  08/26/2026"

RESP_SIG_RA = "Pending Resident Assistant / Professional Staff signature."

RESP_DATE = "08/26/2026"

PLACEHOLDERS = [
    (1, 0, 490.0, 9.0,  "helv", RESP_CLEAN),
    (2, 0, 393.0, 8.5,  "helv", RESP_GUESTS),
    (3, 0, 217.0, 9.0,  "helv", RESP_SLEEP),
    (3, 1, 521.0, 8.5,  "helv", RESP_SHARING),
    (3, 2, 697.0, 9.0,  "helv", RESP_SPACE),
    (4, 0, 372.0, 9.0,  "helv", RESP_COMM),
    (4, 1, 615.0, 9.0,  "helv", RESP_CONFLICT),
    (5, 0, 287.0, 8.5,  "helv", RESP_THINGS),
    (5, 1, 697.0, 11.0, "heit", RESP_SIG_RES),
    (6, 0, 211.0, 10.0, "helv", RESP_SIG_RA),
    (6, 1, 697.0, 11.0, "helv", RESP_DATE),
]

# Group placeholders on each page and find the rects for "Type response here."
page_rects = {}
for pno in range(doc.page_count):
    page = doc[pno]
    rects = page.search_for("Type response here.")
    rects.sort(key=lambda r: r.y0)
    page_rects[pno] = rects

# Build per-page response bands so we can also strip the decorative underlines
# that the template drew inside each answer area.
bands_by_page = {}
for (pno, idx, y1, fs, fname, text) in PLACEHOLDERS:
    r = page_rects[pno][idx]
    bands_by_page.setdefault(pno, []).append((r.y0 - 2.0, y1))

# First, redact placeholder texts AND the underlines within each answer band.
touched_pages = set(p[0] for p in PLACEHOLDERS)
for pno in touched_pages:
    page = doc[pno]
    for r in page_rects[pno]:
        page.add_redact_annot(r, fill=(1, 1, 1))
    # find horizontal rules inside any answer band and cover them
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) < 0.6:
                    ly = p1.y
                    for (by0, by1) in bands_by_page.get(pno, []):
                        if by0 <= ly <= by1:
                            page.add_redact_annot(
                                pymupdf.Rect(70.0, ly - 1.6, 542.0, ly + 1.6),
                                fill=(1, 1, 1),
                            )
                            break
    page.apply_redactions()

# Now insert answers.
for (pno, idx, y1, fs, fname, text) in PLACEHOLDERS:
    page = doc[pno]
    r = page_rects[pno][idx]
    x0 = 72.0
    x1 = 540.0
    y0 = r.y0 - 1.0
    rect = pymupdf.Rect(x0, y0, x1, y1)
    size = fs
    while size >= 6.0:
        rc = page.insert_textbox(
            rect, text, fontsize=size, fontname=fname,
            color=(0.10, 0.10, 0.35), align=0, lineheight=1.25,
        )
        if rc >= 0:
            break
        # didn't fit; clear anything drawn and shrink
        page.add_redact_annot(rect, fill=(1, 1, 1))
        page.apply_redactions()
        size -= 0.5
    if rc < 0:
        print(f"WARNING: page {pno} idx {idx} text did not fit even at size {size}")
    else:
        print(f"page {pno} idx {idx}: inserted at size {size:.1f} (leftover {rc:.1f}pt)")

import os
doc.save(OUT, garbage=4, deflate=True)
doc.close()
os.replace(OUT, SRC)
print("Saved", SRC)
