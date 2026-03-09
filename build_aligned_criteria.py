"""
Restructure Faculty Performance Criteria into horizontally-aligned tables.
Each row = one criterion topic, columns = rating levels 1-4.

Improvements:
- Color-coded rating columns (red → yellow → light green → green)
- Alternating row zebra striping
- Larger readable fonts, multi-page with repeating headers
- Separator rows between core duties and stretch/promotion criteria
- Landscape orientation
"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.page import PageMargins

wb = Workbook()

# ── Styling ──
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
THICK_BOTTOM = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="medium"),
)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
WRAP_CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")

# ── UMN Brand Colors ──
# Maroon: #7A0019, Gold: #FFCC33
# Extended palette: Light gold #FFDE7A, Dark maroon #5B0013, Light maroon tint #F0D5DC

# Column header fills & fonts — UMN Maroon
HEADER_FILL = PatternFill(start_color="7A0019", end_color="7A0019", fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=9)

# Rating column background — neutral with warm UMN tint
ROW_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")         # white
ROW_FILL_ALT = PatternFill(start_color="FFF8EB", end_color="FFF8EB", fill_type="solid")    # very light gold

# Criterion column (col A) — light maroon tint
CRIT_FILL = PatternFill(start_color="F0D5DC", end_color="F0D5DC", fill_type="solid")        # light maroon
CRIT_FILL_ALT = PatternFill(start_color="E4BCC6", end_color="E4BCC6", fill_type="solid")    # slightly deeper maroon

# Promotion-highlight override — UMN Gold
PROMO_FILL = PatternFill(start_color="FFDE7A", end_color="FFDE7A", fill_type="solid")       # light gold
PROMO_FILL_ALT = PatternFill(start_color="FFCC33", end_color="FFCC33", fill_type="solid")   # UMN gold

# Separator row — UMN Maroon
SEP_FILL = PatternFill(start_color="7A0019", end_color="7A0019", fill_type="solid")
SEP_FONT = Font(name="Calibri", bold=True, size=8, color="FFFFFF")

# Body fonts
BODY_FONT = Font(name="Calibri", size=8.5)
BOLD_BODY_FONT = Font(name="Calibri", size=8.5, bold=True)
CRIT_FONT = Font(name="Calibri", bold=True, size=9)
CRIT_FONT_PROMO = Font(name="Calibri", bold=True, size=9, color="5B0013")  # dark maroon for promotion rows

# ── Data ──
# Each table has rows grouped into: ("core", [...rows...]) and ("promotion", [...rows...])
# A separator row is inserted between groups.
# Row format: (criterion_name, [col1, col2, col3, col4], is_promotion_relevant)

tables = {
    "Clinical": {
        "note": "ALL faculty with cFTE must meet ≥2. Clinician Track should meet mostly 3s. Promotion requires ≥2 of 4 bolded criteria. Academic Track, Clinical Focus must also have ≥1 scholarly peer-reviewed publication/year as first or senior author.",
        "groups": [
            ("Core Clinical Duties", [
                ("wRVU Productivity", [
                    "≤90% of benchmark",
                    "100% of benchmark",
                    "120% of benchmark",
                    ">120% of benchmark",
                ], True),
                ("Scheduling Template Utilization", [
                    "<80% utilization",
                    "Averages 80–90% for 46 weeks",
                    "Consistently ≥90% for 46 weeks",
                    "Consistently ≥90% for 46 weeks",
                ], True),
                ("Clinical Availability", [
                    "Unavailable for assigned duties; frequently cancels clinic or does not meet expected days/hours",
                    "Maintains reliable schedule; reschedules for cancellations; willing to provide inpatient coverage as required",
                    "Makes themselves available during times of clinical need (outpatient and inpatient)",
                    "Actively participates during times of clinical need (outpatient and inpatient)",
                ], False),
                ("Documentation", [
                    "Consistently late and/or incomplete; requires multiple prompts",
                    "Timely and complete with only minimal reminders",
                    "Timely and complete with no reminders",
                    "Timely and complete with no reminders",
                ], False),
                ("Knowledge & Skills", [
                    "Does not remain current; clinical work does not meet community standards",
                    "Knowledge is up-to-date; maintains skill set",
                    "Regularly attends Complex Case Conferences, Clinical Forums, and Clinical Council",
                    "Known for specialized skill set; sought after for consultation on complex cases; receives referrals regionally or nationally",
                ], True),
                ("Patient Interaction & Satisfaction", [
                    "Disrespectful to patients and/or clinic staff; surveys show consistently low rankings",
                    "Interacts professionally with patients and staff; surveys are adequate",
                    "Patient satisfaction surveys show consistently high ratings",
                    "Patient satisfaction surveys show consistently high ratings",
                ], False),
            ]),
            ("Scholarship, Leadership & Promotion Criteria", [
                ("Complex Case Conferences / Clinical Forums", [
                    "Does not attend Complex Case Conferences, Clinical Forums, or Clinical Council",
                    "Attends occasionally",
                    "Regularly attends",
                    "Active participant; presents cases and contributes to discussion",
                ], False),
                ("Clinical Publications", [
                    "No clinical publications",
                    "Contributes to a clinical publication as co-author",
                    "Publishes case report or case series",
                    "Publishes clinical review article",
                ], False),
                ("Clinical Innovation", [
                    "Resistant to changes in clinical care processes or technology",
                    "Open to and adopts innovations introduced by others",
                    "Supports innovations in clinical care, new program models, materials, products, and/or technology",
                    "Develops innovations in clinical care models, products, and/or technology",
                ], False),
                ("Speaking & Committee Service", [
                    "Does not participate in speaking or committee activities",
                    "Participates in departmental or local committees when asked",
                    "Invited to speak locally; serves on local/regional committees related to clinical expertise",
                    "Invited to speak regionally or nationally; serves on regional/national committees in professional societies",
                ], False),
                ("Quality Improvement", [
                    "Does not participate in quality improvement activities",
                    "Participates in QI activities when asked",
                    "Participates in quality improvement initiatives",
                    "Initiates/implements quality improvement initiatives and new models of care",
                ], False),
                ("Clinical Program Leadership", [
                    "Does not contribute to clinical program development",
                    "Supports clinical program operations as a team member",
                    "Takes on a leadership role within a clinical program",
                    "Successfully leads a clinical program",
                ], True),
                ("Awards & Recognition", [
                    "No recognition for clinical work",
                    "Receives informal positive feedback from colleagues or patients",
                    "Recognized within the department for clinical contributions",
                    "Nominated for and/or receives award for outstanding work as clinician or clinical team member",
                ], False),
            ]),
        ],
    },
    "Education": {
        "note": "Faculty with dedicated education aFTE must meet ≥3. Promotion requires ≥2 bolded criteria + 1 peer-reviewed pub/year as first or senior author.",
        "groups": [
            ("Core Education Duties", [
                ("Teaching Skills", [
                    "Limited teaching skills; difficulty interacting with learners",
                    "Average teaching skills and knowledge; interacts well with learners",
                    "Recognized by peers and learners as a skilled educator",
                    "Nominated for or receives award/recognition as a talented educator",
                ], True),
                ("Education Activity Participation", [
                    "Refuses to perform or does not show up for education activities; fails to perform supervision",
                    "Participates intermittently in relevant education activities",
                    "Seeks out and participates actively in educational activities",
                    "Highly active in education activities; holds a formal education leadership position",
                ], True),
                ("Student Evaluations", [
                    "Consistently low rankings",
                    "Average rankings",
                    "Consistently above-average rankings",
                    "Consistently high rankings",
                ], False),
                ("Education Materials & Preparation", [
                    "Unprepared; material is out-of-date or poor quality",
                    "Consistently prepared; material is up-to-date",
                    "Leads an education activity; develops novel education activities and materials",
                    "Has developed innovative education approaches or technologies that have significantly changed programs locally",
                ], True),
                ("Recruitment Interviews", [
                    "Does not participate in interviews for resident or other trainee recruitment",
                    "Participates in recruitment interviews when asked",
                    "Regularly participates in recruitment interviews; provides thoughtful evaluations",
                    "Takes a leadership role in recruitment; helps shape interview processes and candidate selection",
                ], False),
            ]),
            ("Scholarship, Leadership & Promotion Criteria", [
                ("Speaking Invitations", [
                    "Does not present on education topics",
                    "Presents on education topics within the department",
                    "Requested to speak on education topics locally and regionally",
                    "Requested to speak on education topics nationally or internationally",
                ], False),
                ("Education Publications & Dissemination", [
                    "No education-related publications or dissemination",
                    "Contributes as co-author to an education-related publication or presentation",
                    "Disseminates new educational insights, approaches, or materials through publications or online",
                    "First or last author in an education-focused publication in a peer-reviewed journal",
                ], True),
                ("Committee Service", [
                    "Does not serve on any education committees",
                    "Participates on a departmental education committee when asked",
                    "Sits on a departmental, medical school, or university education committee",
                    "Serves on national-level committees related to education",
                ], False),
                ("Journal Service", [
                    "No journal review activity",
                    "Reviews an education-related manuscript ad-hoc",
                    "Regularly reviews for an education journal",
                    "Serves on the editorial board of an education journal",
                ], False),
                ("Grant Funding", [
                    "No education grant activity",
                    "Contributes to an education grant application as co-investigator",
                    "Submits an education grant application as PI or co-PI",
                    "Demonstrates successful grant funding for education activities through internal or external peer-reviewed process",
                ], False),
                ("Mentorship", [
                    "Does not engage in education mentorship",
                    "Provides informal guidance to learners when approached",
                    "Actively mentors learners or junior faculty in education activities",
                    "Sought out by learners and junior faculty as an education mentor and advisor",
                ], False),
            ]),
        ],
    },
    "Research": {
        "note": "Faculty with research aFTE must meet mostly 2s and 3s. Promotion requires mostly 3s and 4s, ≥2 bolded criteria, ≥2 peer-reviewed pubs/year as first/senior author, and external funding.",
        "groups": [
            ("Core Research Activities", [
                ("Research Projects", [
                    "Does not have a current well-defined and active research project",
                    "Is conducting 1 well-defined research project",
                    "Is conducting 2+ well-defined research projects as PI (1 if early Asst. Prof)",
                    "Is PI of a well-funded and thriving lab group with 3+ projects",
                ], True),
                ("Research Funding", [
                    "Does not currently carry any research funding from any sources",
                    "Carries intramural research funding",
                    "Carries research funding from 2+ sources (at least 1 federal)",
                    "Holds 2 or more federal grants",
                ], True),
                ("Grant Submissions", [
                    "No national institute or foundation grant submissions in the prior year",
                    "Actively seeking federal or foundation funding through 2+ submissions/year",
                    "Actively seeking federal funding through 2+ submissions/year",
                    "Has a consistent record of multiple grant submissions",
                ], False),
                ("Grant Writing Group", [
                    "Does not participate in the Grant Writing Group",
                    "Regularly participates in the Grant Writing Group",
                    "Active and highly engaged participant in Grant Writing Group",
                    "Active and highly engaged participant in Grant Writing Group",
                ], False),
                ("First/Last Author Publications", [
                    "No first/last author publications in the prior year",
                    "First/last author on 1 published research paper",
                    "First/last author on at least 2 peer-reviewed papers",
                    "First or last author on >3 publications; 1 in high-impact journal (IF >10)",
                ], True),
                ("Co-Author Publications", [
                    "No co-authored publications",
                    "Co-author on 1+ peer-reviewed papers",
                    "Co-author on at least 2 papers in the prior year",
                    "Co-author on 3+ peer-reviewed papers",
                ], False),
                ("Collaborative Research", [
                    "No active collaborative work with a funded study",
                    "Actively collaborating on 1 funded research project",
                    "Actively collaborating on at least 2 funded research projects",
                    "Recognized for multiple significant and highly collaborative projects",
                ], False),
                ("Research Mentorship", [
                    "No research mentorship activities",
                    "At least 2 ongoing research mentorship activities",
                    "Actively engaged in at least 3 mentorship activities",
                    "Sought after as a research mentor; 4+ mentorship activities",
                ], True),
                ("Regulatory Compliance", [
                    "Has compliance / research conduct problems",
                    "Minor or past compliance issues; responsive to corrective actions",
                    "Maintains good standing with all regulatory requirements; completes training on time",
                    "No compliance or research conduct issues; proactively promotes research integrity and best practices",
                ], False),
            ]),
            ("Visibility, Service & Promotion Criteria", [
                ("Conference Presentations", [
                    "No presentations at national or international conferences",
                    "1 presentation or abstract (poster) at national/international conference",
                    "Invited to present nationally or internationally; 2 presentations or abstracts at conferences",
                    "Invited to present internationally (keynote speaker, symposium speaker)",
                ], False),
                ("Peer Review / Editorial Service", [
                    "No peer review or editorial activity",
                    "Asked to review 1 article ad-hoc",
                    "Serves on an editorial board",
                    "Serves as a journal editor or appointed to a research advisory board",
                ], False),
                ("Faculty Research Forum", [
                    "Did not attend any Faculty Research Forum events",
                    "Attended the Faculty Research Forum event",
                    "Attended all Faculty Research Forum events",
                    "Attended all events; actively presents or contributes to Faculty Research Forum",
                ], False),
                ("Research Council / IRB / Committees", [
                    "Does not participate in any research governance or committee activities",
                    "Attends research committee meetings when invited",
                    "Participates on an IRB panel",
                    "Serves on the Research Council; departmental or graduate research program committees",
                ], False),
            ]),
        ],
    },
    "Community Service": {
        "note": "ALL faculty must meet ≥2. Faculty with service aFTE must meet ≥3. Promotion (rare) requires mostly 4s including bolded items.",
        "groups": [
            ("Core Service Activities", [
                ("Participation in Service Activities", [
                    "Resistant or does not follow through when asked to participate",
                    "Willingly participates in advocacy or public engagement activities",
                    "Actively seeks out advocacy and/or community service and engagement activities",
                    "Frequently sought out to represent the dept or field in highly visible advocacy or communications activities",
                ], True),
                ("Communication with Community & Public", [
                    "Communicates in ways that undermine the medical school mission (inconsistent, negative, dismissive)",
                    "Makes efforts to be a good communicator with community members and the public; seeks guidance if needed",
                    "Seen as effective and creative in service and public communication activities",
                    "Has written or communicated at the national level (book, blog, YouTube) for community service, advocacy, or public education",
                ], False),
            ]),
            ("Advanced Service & Promotion Criteria", [
                ("Public/Community Engagement Reach", [
                    "No engagement with community partners or the public",
                    "Occasionally engages with community members or partners when opportunities arise",
                    "Sought out for communication with public and/or community partners (invited talks, trainings) in support of mental health",
                    "Recognized regionally or nationally as an advocate and resource in mental health",
                ], True),
                ("Community Partnerships & Relationships", [
                    "No relationships with external community partners or advocacy groups",
                    "Maintains basic professional relationships with community partners",
                    "Forms active and positive relationships with legislature, advocacy groups, media, and/or community partners",
                    "Has played a key role in driving change in legislative, public policy, or public awareness issues related to mental health",
                ], False),
                ("Innovation in Advocacy & Engagement", [
                    "No involvement in advocacy or engagement innovation",
                    "Supports existing advocacy or public engagement initiatives",
                    "Develops new methods, approaches, and partnerships for advocacy, public engagement, and/or community partnerships",
                    "Has developed new methods, studies, or advocacy strategies and published on their effects in peer-reviewed journals",
                ], False),
            ]),
        ],
    },
    "Departmental Culture": {
        "note": "ALL faculty must meet 3 (High Satisfactory). 0.05 aFTE is automatically funded for these activities.",
        "groups": [
            ("Core Departmental Expectations", [
                ("Administrative Responsiveness", [
                    "Does not perform needed admin tasks in a timely manner; needs multiple prompts with consequences",
                    "Performs needed admin tasks but can be late or require multiple prompting",
                    "Responds to administrative responsibilities in a timely, independent, positive manner",
                    "Successfully serves in an administrative leadership role",
                ], False),
                ("Email / Communication Responsiveness", [
                    "Consistently does not answer emails from admin staff, department administrators, or the Head",
                    "Eventually responds to admin emails or messages from the Head, but is often late",
                    "Responds to admin emails or messages from the Head in a timely manner",
                    "Consistently highly effective at solving problems, charting new directions, finding innovative solutions",
                ], False),
                ("Faculty Meeting Attendance", [
                    "Does not attend Faculty Meetings",
                    "Attends Faculty Meetings infrequently",
                    "Attends Faculty Meetings regularly with camera on",
                    "Consistently an active participant in Faculty Meetings and other dept meetings (camera on)",
                ], False),
                ("Grand Rounds Attendance", [
                    "Does not attend Grand Rounds",
                    "Attends Grand Rounds infrequently",
                    "Attends Grand Rounds regularly",
                    "Consistently an active participant at Grand Rounds",
                ], False),
                ("Ethics & Integrity", [
                    "Unethical in work with patients, families, or research",
                    "Engages in ethical behavior",
                    "Manifests personal integrity and high ethical standards",
                    "Embodies & fosters integrity and the highest ethical standards; sought out for input during complex or conflictual issues",
                ], False),
                ("Departmental Action Plans", [
                    "Does not adhere to action plans that are in the best interest of the department",
                    "Generally follows departmental action plans with occasional reminders",
                    "Consistently adheres to and supports departmental action plans",
                    "Champions departmental action plans; helps others understand and follow through",
                ], False),
            ]),
            ("Culture, Growth & Leadership", [
                ("Workplace Culture & Collegiality", [
                    "Negative attitude — fosters complaints or problems instead of solutions",
                    "Maintains cordial relationships with colleagues, trainees, patients, and families",
                    "Helps support methods to increase well-being, productivity, and efficient processes",
                    "Widely recognized as a positive culture leader; actively mentors others and models collaborative problem-solving",
                ], False),
                ("Inclusive Excellence & Well-Being (IEWB)", [
                    "Undermines a culture of Inclusive Excellence",
                    "Participates in one IEWB activity per year",
                    "Participates in at least two activities per year supported by IEWB",
                    "Is an active participant in IEWB and/or similar activities",
                ], False),
                ("Feedback & Growth", [
                    "Shows difficulty accepting feedback",
                    "Willing to accept feedback",
                    "Seeks and accepts feedback; offers respectful feedback when appropriate",
                    "Widely recognized for strong listening skills; frequently seeks to receive and offer feedback with respect, openness, and clarity",
                ], False),
                ("Crucial Conversations & Problem-Solving", [
                    "Unwilling to engage in productive Crucial Conversations",
                    "Willing to engage in department problem-solving activities if asked",
                    "Contributes actively to a positive problem-solving work environment; engages in productive Crucial Conversations when needed",
                    "Proactively engages in growth-promoting Crucial Conversations that support learning and healthy change",
                ], False),
            ]),
        ],
    },
}

COL_HEADERS = ["Criterion", "1 – Unsatisfactory", "2 – Low Satisfactory", "3 – High Satisfactory", "4 – Outstanding"]

for idx, (sheet_name, data) in enumerate(tables.items()):
    ws = wb.active if idx == 0 else wb.create_sheet()
    ws.title = sheet_name

    # ── Page setup: landscape, 8.5x11, fit to 1 page WIDE but multi-page tall ──
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # allow multiple pages vertically
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.4, header=0.2, footer=0.2)
    ws.print_title_rows = "1:4"  # repeat title + headers on every page

    # ── Title row ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    title_cell = ws.cell(row=1, column=1, value=f"{sheet_name.upper()} PERFORMANCE CRITERIA")
    title_cell.font = Font(name="Calibri", bold=True, size=13, color="7A0019")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── Note row ──
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)
    note_cell = ws.cell(row=2, column=1, value=data["note"])
    note_cell.font = Font(name="Calibri", italic=True, size=8.5)
    note_cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Row 3 = spacer
    ws.row_dimensions[3].height = 6

    # ── Column headers (row 4) ──
    for col_idx, header in enumerate(COL_HEADERS, start=1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = WRAP_CENTER
        cell.border = THIN_BORDER
    ws.row_dimensions[4].height = 28

    # ── Data rows ──
    current_row = 5
    data_row_count = 0  # for zebra striping (resets per group)

    for group_idx, (group_label, rows) in enumerate(data["groups"]):
        # Separator row
        if group_idx > 0:
            # Add a thin spacer row before separator
            ws.row_dimensions[current_row].height = 4
            current_row += 1

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        sep_cell = ws.cell(row=current_row, column=1, value=f"  {group_label}")
        sep_cell.font = SEP_FONT
        sep_cell.fill = SEP_FILL
        sep_cell.alignment = Alignment(vertical="center")
        sep_cell.border = THIN_BORDER
        # Fill all merged cells with border
        for c in range(2, 6):
            ws.cell(row=current_row, column=c).fill = SEP_FILL
            ws.cell(row=current_row, column=c).border = THIN_BORDER
        ws.row_dimensions[current_row].height = 20
        current_row += 1
        data_row_count = 0

        for criterion, values, is_promo in rows:
            is_alt = (data_row_count % 2 == 1)

            # Column A: criterion name
            c = ws.cell(row=current_row, column=1, value=criterion)
            if is_promo:
                c.font = CRIT_FONT_PROMO
                c.fill = PROMO_FILL_ALT if is_alt else PROMO_FILL
            else:
                c.font = CRIT_FONT
                c.fill = CRIT_FILL_ALT if is_alt else CRIT_FILL
            c.alignment = WRAP_TOP
            c.border = THIN_BORDER

            # Columns B-E: rating values with color coding
            for col_idx, val in enumerate(values, start=2):
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = BOLD_BODY_FONT if is_promo else BODY_FONT
                cell.alignment = WRAP_TOP
                cell.border = THIN_BORDER

                if is_promo:
                    cell.fill = PROMO_FILL_ALT if is_alt else PROMO_FILL
                else:
                    cell.fill = ROW_FILL_ALT if is_alt else ROW_FILL

            data_row_count += 1
            current_row += 1

    # ── Column widths for landscape (~10in printable) ──
    ws.column_dimensions["A"].width = 24
    for col_letter in ["B", "C", "D", "E"]:
        ws.column_dimensions[col_letter].width = 34

    # Freeze panes below header
    ws.freeze_panes = "A5"

wb.save("/home/kolim/FacultyEvaluation/Faculty_Performance_Criteria_Aligned.xlsx")
print("Done! Saved to Faculty_Performance_Criteria_Aligned.xlsx")
