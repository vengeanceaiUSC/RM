"""Generate a fully completed Suite and Apartment Living Agreement.

Every question is answered inline, right next to the question itself, and every
section from the official template is included (Cleanliness, Guests, Sleeping,
Sharing, Space Use, Communication, Conflict, Things to Consider, and the
signature blocks). Official policy reminders are preserved verbatim.

The document is laid out with PyMuPDF's Story engine so the content reflows
cleanly across pages and nothing is left as an unanswered "Type response here."
placeholder.
"""

import html

import pymupdf

OUT = "/workspace/26-27_Suites-and-Apt-Living-Agreement.pdf"

# ---------------------------------------------------------------------------
# Header fields
# ---------------------------------------------------------------------------
HEADER = [
    ("Building", "La Sorbonne"),
    ("Apartment/Suite Number", "202"),
    ("Name of Roommate/Suitemates", "Edwin Huang, Chris Gardner"),
    ("Date", "08/25/2026"),
]

# ---------------------------------------------------------------------------
# Sections: each is (heading, [(question, answer), ...], [policy reminders])
# Answers are placed immediately next to each question.
# ---------------------------------------------------------------------------
SECTIONS = [
    (
        "Cleanliness",
        [
            ("How often will the trash be taken out?",
             "Once a week (Sundays), and sooner if it's full. We rotate who takes it out."),
            ("How often will the shared spaces and individual rooms be cleaned?",
             "Shared spaces are cleaned once a week; each of us keeps our own room tidy and does a deeper clean weekly."),
            ("How will we take turns completing these tasks?",
             "We rotate chores each week and track whose turn it is on a shared whiteboard / phone list."),
            ("If we share a kitchen, how often will we clean dishes, the refrigerator, stove, common spaces, etc.?",
             "Dishes are washed daily and never left overnight; the refrigerator, stove, and counters are wiped down every Sunday."),
            ("What expectations do we all have concerning bathroom cleanliness?",
             "Wipe the sink after use, keep personal items off the counter, and deep-clean the toilet, shower, and floor weekly on a rotating basis."),
            ("What is our chore plan?",
             "A weekly rotation of trash, bathroom, and common-area duties, tracked on our shared list so nothing gets missed."),
            ("How will we decide who purchases cleaning supplies?",
             "Costs are split evenly; whoever notices we're running low restocks them and we settle up on a shared expense list."),
            ("Do you have any allergies that are important for your roommates to know about (cats, dogs, peanuts, other)?",
             "Edwin has a mild cat allergy; Chris has no known allergies."),
            ("How can you optimize your cleaning schedule to keep those allergies in mind?",
             "No cats in the unit, vacuum regularly, and keep dust and pet dander out of shared spaces."),
        ],
        [],
    ),
    (
        "Guests",
        [
            ("When are guests allowed in the unit and individual rooms?",
             "In the common areas with a quick heads-up text; in an individual room only with that roommate's okay."),
            ("When are overnight guests allowed? What is the frequency you agree to guests being allowed in the space? "
             "Can guests use a roommate's belongings? (be specific of what they can and cannot use)",
             "Overnight guests need at least a day's notice and follow USC policy (max 3 consecutive nights, twice a month). "
             "Guests may use shared common-area items but not a roommate's personal belongings (bed, desk, food, toiletries, "
             "electronics) without that roommate's permission."),
            ("How will we communicate with each other if we want a guest to leave?",
             "We'll tell each other directly and privately, and the host will wrap things up promptly."),
            ("What does it look like to have guests of the opposite sex? Family members visiting? Partners?",
             "All are welcome with advance notice and mutual respect for shared space, sleep, and study time."),
        ],
        [
            "Policy Reminder: Guests must be registered and accompanied at all times. Guests can stay no more than "
            "three (3) consecutive nights and may do so twice within one month. In suites, each Resident is allowed a "
            "maximum of 2 guests at any single time. In apartments, each Resident is allowed a maximum of four (4) "
            "guests at any single time but shall not exceed two (2) overnight guests out of respect for their roommate.",
        ],
    ),
    (
        "Sleeping",
        [
            ("What are our sleep schedules?",
             "Edwin is usually asleep by midnight and up around 8am; Chris tends to stay up until about 1am and wakes near 9am."),
            ("What should we do if one person returns to the space while another is sleeping?",
             "Keep the overhead lights off, use a phone flashlight, and stay quiet."),
            ("Is it okay to use the snooze button? If yes, how many times is acceptable?",
             "Yes, but limited to two alarms, then get up so it doesn't disturb the other person."),
            ("What are the guidelines for lights, temperature, and noise during times when others are asleep?",
             "Overhead lights off (desk lamp only), thermostat around 70 F, and headphones for any music or video."),
        ],
        [
            "Policy Reminder: Quiet Hours are a minimum of 11pm to 8 am the following morning, Sunday through Thursday. "
            "On Friday and Saturday, Quiet Hours begin at midnight and end at 8 am the following morning. Quiet Hours "
            "extend to 24/7 on Study Days and during Final Exam periods. Courtesy Hours are in effect 24/7.",
        ],
    ),
    (
        "Sharing",
        [
            ("What items (food, toiletries, clothing, school supplies, etc.) can be shared with permission? "
             "Without permission? Never?",
             "With permission: snacks, coffee, and school supplies. Without permission: nothing personal. "
             "Never: toiletries and medication."),
            ("How will we divide shared space (refrigerator, bathroom, closets, storage)?",
             "Fridge, closet, bathroom, and storage are split evenly and labeled where it helps avoid mix-ups."),
            ("What temperature will we keep our space at night/during the day?",
             "Around 70 F at night and 72 F during the day."),
            ("When will we open the windows?",
             "In the mornings when the weather is nice and to air out the space after cooking."),
            ("While studying, is talking/video calling via phone, TV noise, or playing music allowed in individual "
             "rooms or common spaces?",
             "Calls, TV, and music go through headphones or move to the common area so the other person can focus."),
        ],
        [],
    ),
    (
        "Space Use",
        [
            ("What hours is it acceptable to make/receive video calls?",
             "In-room before 10pm; after that, take them to the common area."),
            ("Can we host study groups in our room?",
             "Yes, with a heads-up beforehand so the other roommate can plan around it."),
            ("If sharing a living space, around what time will tv/noise quiet down?",
             "By about 11pm on weeknights, earlier during exams."),
        ],
        [],
    ),
    (
        "Communication",
        [
            ("How do you do your best to communicate with others? Text, call, email, in person?",
             "Text for day-to-day things and in person for anything important or sensitive."),
            ("How will we communicate with one another if we want privacy/alone time?",
             "A quick text or an agreed door signal, and the other person will give space."),
            ("If a roommate becomes sick, how would you like to share/learn this news?",
             "Send a text so we can keep our distance, disinfect shared surfaces, and help with food or supplies."),
        ],
        [],
    ),
    (
        "Conflict",
        [
            ("How will we communicate issues and/or concerns when they arise with one another?",
             "Directly, calmly, and early, in person whenever possible."),
            ("Do you feel comfortable discussing concerns in person or some other way?",
             "Yes, in person; we can send a text first if it helps to start the conversation."),
            ("When will we connect with our RA to address concerns?",
             "If we can't resolve something ourselves, we'll bring it to our RA together for support."),
        ],
        [
            "Reminder: It is encouraged that you communicate conflicts with your roommate directly. If concerns are "
            "brought to an RA, they can assist in this process.",
        ],
    ),
    (
        "Things to Consider",
        [
            ("Anything not mentioned you would like your roommates or suitemates to be aware of? (pet peeves, stressors, "
             "communication preferences, religious accommodations, health needs/concerns, animal comfortabilities, "
             "views of alcohol, drugs/smoking, general living habits, comforts, etc.)",
             "Pet peeves: dishes left overnight and borrowing without asking. We're both focused on academics, so a "
             "quiet study environment matters. No smoking in the unit; alcohol only if of age and kept low-key. We'll "
             "respect each other's religious practices, health needs, and general living habits."),
            ("Are there any final thoughts you'd like to discuss?",
             "We'll check in with each other regularly and update this agreement if anything changes."),
        ],
        [],
    ),
]

POLICY_BLOCK = (
    "Living Agreements and Policies: As previously stated, it is best to communicate directly with your roommates to "
    "voice concerns or issues when they arise. Your RA is here to assist in that process should you need additional "
    "support or guidance voicing any concerns. If any resident reports concerns about any policy violations occurring "
    "in the community, those concerns may be subject to the Residential Review process. By signing the USC Housing "
    "Contract, you agree to comply with all policies as outlined."
)

ACK_BLOCK = (
    "Electronic Agreement Acknowledgement: By signing your electronic signature below, you agree to the terms and "
    "conditions outlined in this document. Failure to uphold the standards agreed upon in this document may result in "
    "disciplinary action. Please include your first and last name and the date of Agreement completion or revision below."
)

SIGNATURES = [
    ("Signature of Residents", "Edwin Huang &mdash; 08/26/2026<br/>Chris Gardner &mdash; 08/26/2026"),
    ("Signature of Resident Assistant or Professional Staff Member", "________________________________ (pending)"),
    ("Date of Document Completion or Revision", "08/26/2026"),
]


def esc(text: str) -> str:
    return html.escape(text)


# ---------------------------------------------------------------------------
# Build the HTML
# ---------------------------------------------------------------------------
parts = [
    "<h1>Suite and Apartment Living Agreement</h1>",
    "<p class='intro'>Sharing a space and living with others requires that all those living in a space discuss "
    "preferences, comfort levels, and abilities, and form agreements regarding shared spaces, guests visiting, "
    "cleanliness, and the use of each other's belongings, etc. Putting your shared understanding in writing is often "
    "the best way to avoid verbal agreements being forgotten or misinterpreted as time passes. This agreement is meant "
    "to help in discussing and compromising on topics that can cause friction within a living environment.</p>",
    "<p class='intro'>Please familiarize yourself with USC's Housing policies prior to completing this agreement.</p>",
]

parts.append("<div class='header'>")
for label, value in HEADER:
    parts.append(f"<p class='field'><b>{esc(label)}:</b> {esc(value)}</p>")
parts.append("</div>")

for heading, qas, policies in SECTIONS:
    parts.append(f"<h2>{esc(heading)}</h2>")
    for question, answer in qas:
        parts.append(
            f"<p class='qa'><span class='q'>{esc(question)}</span> "
            f"<span class='a'>{esc(answer)}</span></p>"
        )
    for policy in policies:
        parts.append(f"<p class='policy'>{esc(policy)}</p>")

parts.append(f"<p class='policy'>{POLICY_BLOCK}</p>")
parts.append(f"<p class='policy'>{ACK_BLOCK}</p>")

parts.append("<h2>Signatures</h2>")
for label, value in SIGNATURES:
    parts.append(f"<p class='qa'><span class='q'>{esc(label)}:</span> <span class='a'>{value}</span></p>")

HTML = "<html><body>" + "".join(parts) + "</body></html>"

CSS = """
* { font-family: sans-serif; }
h1 { font-size: 17px; color: #7a0019; text-align: center; margin-bottom: 8px; }
h2 { font-size: 13px; color: #7a0019; margin-top: 14px; margin-bottom: 4px;
     border-bottom: 1px solid #7a0019; padding-bottom: 2px; }
p { font-size: 10.5px; line-height: 1.35; margin: 3px 0; }
p.intro { color: #333333; }
p.field { font-size: 11px; margin: 2px 0; }
p.qa { margin: 5px 0; }
span.q { font-weight: bold; color: #111111; }
span.a { color: #1b3a7a; }
p.policy { font-size: 9.5px; font-style: italic; color: #444444; margin: 6px 0; }
"""

# USC Student Life / Residential Education logo, stamped at the top of every
# page to match the official template.
LOGO = "/workspace/usc_logo.png"
LOGO_RECT = pymupdf.Rect(67.5, 40.0, 197.25, 82.75)

story = pymupdf.Story(html=HTML, user_css=CSS)
writer = pymupdf.DocumentWriter(OUT)

MEDIABOX = pymupdf.paper_rect("letter")
# Leave room at the top for the logo.
WHERE = MEDIABOX + (54, 100, -54, -54)

more = 1
page_no = 0
while more:
    page_no += 1
    dev = writer.begin_page(MEDIABOX)
    more, _ = story.place(WHERE)
    story.draw(dev)
    writer.end_page()

writer.close()

# Stamp the logo onto every page.
doc = pymupdf.open(OUT)
for page in doc:
    page.insert_image(LOGO_RECT, filename=LOGO, keep_proportion=True)
doc.saveIncr()
doc.close()

print(f"Saved {OUT} ({page_no} pages)")
