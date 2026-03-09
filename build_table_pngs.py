"""
Generate a PNG table image for each of the 5 mission areas.
Uses Playwright (headless Chromium) to render styled HTML tables.

Usage:
    python build_table_pngs.py          # outputs to png/ directory
"""
import os
import html as html_mod
from playwright.sync_api import sync_playwright

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "png")

# ── Reuse the same data dict from build_aligned_criteria.py ──
# (kept here to avoid import side-effects from the openpyxl script)

tables = {
    "Clinical": {
        "note": "ALL faculty with cFTE must meet \u22652. Clinician Track should meet mostly 3s. Promotion requires \u22652 of 4 bolded criteria. Academic Track, Clinical Focus must also have \u22651 scholarly peer-reviewed publication/year as first or senior author.",
        "groups": [
            ("Core Clinical Duties", [
                ("wRVU Productivity", ["\u226490% of benchmark", "100% of benchmark", "120% of benchmark", ">120% of benchmark"], True),
                ("Scheduling Template Utilization", ["<80% utilization", "Averages 80\u201390% for 46 weeks", "Consistently \u226590% for 46 weeks", "Consistently \u226590% for 46 weeks"], True),
                ("Clinical Availability", ["Unavailable for assigned duties; frequently cancels clinic or does not meet expected days/hours", "Maintains reliable schedule; reschedules for cancellations; willing to provide inpatient coverage as required", "Makes themselves available during times of clinical need (outpatient and inpatient)", "Actively participates during times of clinical need (outpatient and inpatient)"], False),
                ("Documentation", ["Consistently late and/or incomplete; requires multiple prompts", "Timely and complete with only minimal reminders", "Timely and complete with no reminders", "Timely and complete with no reminders"], False),
                ("Knowledge & Skills", ["Does not remain current; clinical work does not meet community standards", "Knowledge is up-to-date; maintains skill set", "Regularly attends Complex Case Conferences, Clinical Forums, and Clinical Council", "Known for specialized skill set; sought after for consultation on complex cases; receives referrals regionally or nationally"], True),
                ("Patient Interaction & Satisfaction", ["Disrespectful to patients and/or clinic staff; surveys show consistently low rankings", "Interacts professionally with patients and staff; surveys are adequate", "Patient satisfaction surveys show consistently high ratings", "Patient satisfaction surveys show consistently high ratings"], False),
            ]),
            ("Scholarship, Leadership & Promotion Criteria", [
                ("Complex Case Conferences / Clinical Forums", ["Does not attend Complex Case Conferences, Clinical Forums, or Clinical Council", "Attends occasionally", "Regularly attends", "Active participant; presents cases and contributes to discussion"], False),
                ("Clinical Publications", ["No clinical publications", "Contributes to a clinical publication as co-author", "Publishes case report or case series", "Publishes clinical review article"], False),
                ("Clinical Innovation", ["Resistant to changes in clinical care processes or technology", "Open to and adopts innovations introduced by others", "Supports innovations in clinical care, new program models, materials, products, and/or technology", "Develops innovations in clinical care models, products, and/or technology"], False),
                ("Speaking & Committee Service", ["Does not participate in speaking or committee activities", "Participates in departmental or local committees when asked", "Invited to speak locally; serves on local/regional committees related to clinical expertise", "Invited to speak regionally or nationally; serves on regional/national committees in professional societies"], False),
                ("Quality Improvement", ["Does not participate in quality improvement activities", "Participates in QI activities when asked", "Participates in quality improvement initiatives", "Initiates/implements quality improvement initiatives and new models of care"], False),
                ("Clinical Program Leadership", ["Does not contribute to clinical program development", "Supports clinical program operations as a team member", "Takes on a leadership role within a clinical program", "Successfully leads a clinical program"], True),
                ("Awards & Recognition", ["No recognition for clinical work", "Receives informal positive feedback from colleagues or patients", "Recognized within the department for clinical contributions", "Nominated for and/or receives award for outstanding work as clinician or clinical team member"], False),
            ]),
        ],
    },
    "Education": {
        "note": "Faculty with dedicated education aFTE must meet \u22653. Promotion requires \u22652 bolded criteria + 1 peer-reviewed pub/year as first or senior author.",
        "groups": [
            ("Core Education Duties", [
                ("Teaching Skills", ["Limited teaching skills; difficulty interacting with learners", "Average teaching skills and knowledge; interacts well with learners", "Recognized by peers and learners as a skilled educator", "Nominated for or receives award/recognition as a talented educator"], True),
                ("Education Activity Participation", ["Refuses to perform or does not show up for education activities; fails to perform supervision", "Participates intermittently in relevant education activities", "Seeks out and participates actively in educational activities", "Highly active in education activities; holds a formal education leadership position"], True),
                ("Student Evaluations", ["Consistently low rankings", "Average rankings", "Consistently above-average rankings", "Consistently high rankings"], False),
                ("Education Materials & Preparation", ["Unprepared; material is out-of-date or poor quality", "Consistently prepared; material is up-to-date", "Leads an education activity; develops novel education activities and materials", "Has developed innovative education approaches or technologies that have significantly changed programs locally"], True),
                ("Recruitment Interviews", ["Does not participate in interviews for resident or other trainee recruitment", "Participates in recruitment interviews when asked", "Regularly participates in recruitment interviews; provides thoughtful evaluations", "Takes a leadership role in recruitment; helps shape interview processes and candidate selection"], False),
            ]),
            ("Scholarship, Leadership & Promotion Criteria", [
                ("Speaking Invitations", ["Does not present on education topics", "Presents on education topics within the department", "Requested to speak on education topics locally and regionally", "Requested to speak on education topics nationally or internationally"], False),
                ("Education Publications & Dissemination", ["No education-related publications or dissemination", "Contributes as co-author to an education-related publication or presentation", "Disseminates new educational insights, approaches, or materials through publications or online", "First or last author in an education-focused publication in a peer-reviewed journal"], True),
                ("Committee Service", ["Does not serve on any education committees", "Participates on a departmental education committee when asked", "Sits on a departmental, medical school, or university education committee", "Serves on national-level committees related to education"], False),
                ("Journal Service", ["No journal review activity", "Reviews an education-related manuscript ad-hoc", "Regularly reviews for an education journal", "Serves on the editorial board of an education journal"], False),
                ("Grant Funding", ["No education grant activity", "Contributes to an education grant application as co-investigator", "Submits an education grant application as PI or co-PI", "Demonstrates successful grant funding for education activities through internal or external peer-reviewed process"], False),
                ("Mentorship", ["Does not engage in education mentorship", "Provides informal guidance to learners when approached", "Actively mentors learners or junior faculty in education activities", "Sought out by learners and junior faculty as an education mentor and advisor"], False),
            ]),
        ],
    },
    "Research": {
        "note": "Faculty with research aFTE must meet mostly 2s and 3s. Promotion requires mostly 3s and 4s, \u22652 bolded criteria, \u22652 peer-reviewed pubs/year as first/senior author, and external funding.",
        "groups": [
            ("Core Research Activities", [
                ("Research Projects", ["Does not have a current well-defined and active research project", "Is conducting 1 well-defined research project", "Is conducting 2+ well-defined research projects as PI (1 if early Asst. Prof)", "Is PI of a well-funded and thriving lab group with 3+ projects"], True),
                ("Research Funding", ["Does not currently carry any research funding from any sources", "Carries intramural research funding", "Carries research funding from 2+ sources (at least 1 federal)", "Holds 2 or more federal grants"], True),
                ("Grant Submissions", ["No national institute or foundation grant submissions in the prior year", "Actively seeking federal or foundation funding through 2+ submissions/year", "Actively seeking federal funding through 2+ submissions/year", "Has a consistent record of multiple grant submissions"], False),
                ("Grant Writing Group", ["Does not participate in the Grant Writing Group", "Regularly participates in the Grant Writing Group", "Active and highly engaged participant in Grant Writing Group", "Active and highly engaged participant in Grant Writing Group"], False),
                ("First/Last Author Publications", ["No first/last author publications in the prior year", "First/last author on 1 published research paper", "First/last author on at least 2 peer-reviewed papers", "First or last author on >3 publications; 1 in high-impact journal (IF >10)"], True),
                ("Co-Author Publications", ["No co-authored publications", "Co-author on 1+ peer-reviewed papers", "Co-author on at least 2 papers in the prior year", "Co-author on 3+ peer-reviewed papers"], False),
                ("Collaborative Research", ["No active collaborative work with a funded study", "Actively collaborating on 1 funded research project", "Actively collaborating on at least 2 funded research projects", "Recognized for multiple significant and highly collaborative projects"], False),
                ("Research Mentorship", ["No research mentorship activities", "At least 2 ongoing research mentorship activities", "Actively engaged in at least 3 mentorship activities", "Sought after as a research mentor; 4+ mentorship activities"], True),
                ("Regulatory Compliance", ["Has compliance / research conduct problems", "Minor or past compliance issues; responsive to corrective actions", "Maintains good standing with all regulatory requirements; completes training on time", "No compliance or research conduct issues; proactively promotes research integrity and best practices"], False),
            ]),
            ("Visibility, Service & Promotion Criteria", [
                ("Conference Presentations", ["No presentations at national or international conferences", "1 presentation or abstract (poster) at national/international conference", "Invited to present nationally or internationally; 2 presentations or abstracts at conferences", "Invited to present internationally (keynote speaker, symposium speaker)"], False),
                ("Peer Review / Editorial Service", ["No peer review or editorial activity", "Asked to review 1 article ad-hoc", "Serves on an editorial board", "Serves as a journal editor or appointed to a research advisory board"], False),
                ("Faculty Research Forum", ["Did not attend any Faculty Research Forum events", "Attended the Faculty Research Forum event", "Attended all Faculty Research Forum events", "Attended all events; actively presents or contributes to Faculty Research Forum"], False),
                ("Research Council / IRB / Committees", ["Does not participate in any research governance or committee activities", "Attends research committee meetings when invited", "Participates on an IRB panel", "Serves on the Research Council; departmental or graduate research program committees"], False),
            ]),
        ],
    },
    "Community Service": {
        "note": "ALL faculty must meet \u22652. Faculty with service aFTE must meet \u22653. Promotion (rare) requires mostly 4s including bolded items.",
        "groups": [
            ("Core Service Activities", [
                ("Participation in Service Activities", ["Resistant or does not follow through when asked to participate", "Willingly participates in advocacy or public engagement activities", "Actively seeks out advocacy and/or community service and engagement activities", "Frequently sought out to represent the dept or field in highly visible advocacy or communications activities"], True),
                ("Communication with Community & Public", ["Communicates in ways that undermine the medical school mission (inconsistent, negative, dismissive)", "Makes efforts to be a good communicator with community members and the public; seeks guidance if needed", "Seen as effective and creative in service and public communication activities", "Has written or communicated at the national level (book, blog, YouTube) for community service, advocacy, or public education"], False),
            ]),
            ("Advanced Service & Promotion Criteria", [
                ("Public/Community Engagement Reach", ["No engagement with community partners or the public", "Occasionally engages with community members or partners when opportunities arise", "Sought out for communication with public and/or community partners (invited talks, trainings) in support of mental health", "Recognized regionally or nationally as an advocate and resource in mental health"], True),
                ("Community Partnerships & Relationships", ["No relationships with external community partners or advocacy groups", "Maintains basic professional relationships with community partners", "Forms active and positive relationships with legislature, advocacy groups, media, and/or community partners", "Has played a key role in driving change in legislative, public policy, or public awareness issues related to mental health"], False),
                ("Innovation in Advocacy & Engagement", ["No involvement in advocacy or engagement innovation", "Supports existing advocacy or public engagement initiatives", "Develops new methods, approaches, and partnerships for advocacy, public engagement, and/or community partnerships", "Has developed new methods, studies, or advocacy strategies and published on their effects in peer-reviewed journals"], False),
            ]),
        ],
    },
    "Departmental Culture": {
        "note": "ALL faculty must meet 3 (High Satisfactory). 0.05 aFTE is automatically funded for these activities.",
        "groups": [
            ("Core Departmental Expectations", [
                ("Administrative Responsiveness", ["Does not perform needed admin tasks in a timely manner; needs multiple prompts with consequences", "Performs needed admin tasks but can be late or require multiple prompting", "Responds to administrative responsibilities in a timely, independent, positive manner", "Successfully serves in an administrative leadership role"], False),
                ("Email / Communication Responsiveness", ["Consistently does not answer emails from admin staff, department administrators, or the Head", "Eventually responds to admin emails or messages from the Head, but is often late", "Responds to admin emails or messages from the Head in a timely manner", "Consistently highly effective at solving problems, charting new directions, finding innovative solutions"], False),
                ("Faculty Meeting Attendance", ["Does not attend Faculty Meetings", "Attends Faculty Meetings infrequently", "Attends Faculty Meetings regularly with camera on", "Consistently an active participant in Faculty Meetings and other dept meetings (camera on)"], False),
                ("Grand Rounds Attendance", ["Does not attend Grand Rounds", "Attends Grand Rounds infrequently", "Attends Grand Rounds regularly", "Consistently an active participant at Grand Rounds"], False),
                ("Ethics & Integrity", ["Unethical in work with patients, families, or research", "Engages in ethical behavior", "Manifests personal integrity and high ethical standards", "Embodies & fosters integrity and the highest ethical standards; sought out for input during complex or conflictual issues"], False),
                ("Departmental Action Plans", ["Does not adhere to action plans that are in the best interest of the department", "Generally follows departmental action plans with occasional reminders", "Consistently adheres to and supports departmental action plans", "Champions departmental action plans; helps others understand and follow through"], False),
            ]),
            ("Culture, Growth & Leadership", [
                ("Workplace Culture & Collegiality", ["Negative attitude \u2014 fosters complaints or problems instead of solutions", "Maintains cordial relationships with colleagues, trainees, patients, and families", "Helps support methods to increase well-being, productivity, and efficient processes", "Widely recognized as a positive culture leader; actively mentors others and models collaborative problem-solving"], False),
                ("Inclusive Excellence & Well-Being (IEWB)", ["Undermines a culture of Inclusive Excellence", "Participates in one IEWB activity per year", "Participates in at least two activities per year supported by IEWB", "Is an active participant in IEWB and/or similar activities"], False),
                ("Feedback & Growth", ["Shows difficulty accepting feedback", "Willing to accept feedback", "Seeks and accepts feedback; offers respectful feedback when appropriate", "Widely recognized for strong listening skills; frequently seeks to receive and offer feedback with respect, openness, and clarity"], False),
                ("Crucial Conversations & Problem-Solving", ["Unwilling to engage in productive Crucial Conversations", "Willing to engage in department problem-solving activities if asked", "Contributes actively to a positive problem-solving work environment; engages in productive Crucial Conversations when needed", "Proactively engages in growth-promoting Crucial Conversations that support learning and healthy change"], False),
            ]),
        ],
    },
}

COL_HEADERS = ["Criterion", "1 \u2013 Unsatisfactory", "2 \u2013 Low Satisfactory", "3 \u2013 High Satisfactory", "4 \u2013 Outstanding"]

# ── HTML / CSS ──

CSS = """\
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: 'Segoe UI', Calibri, Arial, sans-serif;
    background: white;
    padding: 24px 28px;
    -webkit-font-smoothing: antialiased;
}
.title {
    color: #7A0019;
    font-size: 20px;
    font-weight: 700;
    text-align: center;
    margin-bottom: 6px;
    letter-spacing: 0.5px;
}
.note {
    font-size: 11px;
    color: #444;
    font-style: italic;
    text-align: center;
    margin-bottom: 14px;
    line-height: 1.4;
    max-width: 960px;
    margin-left: auto;
    margin-right: auto;
}
table {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
    font-size: 12px;
}
col.criterion { width: 18%; }
col.rating    { width: 20.5%; }
th {
    background: #7A0019;
    color: white;
    font-weight: 600;
    font-size: 11.5px;
    padding: 8px 10px;
    text-align: center;
    border: 1px solid #5B0013;
}
/* Separator rows */
tr.separator td {
    background: #7A0019;
    color: white;
    font-weight: 600;
    font-size: 11.5px;
    padding: 6px 10px;
    border: 1px solid #5B0013;
    letter-spacing: 0.3px;
}
/* Data rows */
td {
    padding: 7px 10px;
    vertical-align: top;
    border: 1px solid #D0D0D0;
    line-height: 1.45;
    font-size: 11.5px;
}
td.criterion {
    font-weight: 600;
    font-size: 12px;
    background: #F0D5DC;
}
/* Zebra striping */
tr.alt td { background: #FFF8EB; }
tr.alt td.criterion { background: #E4BCC6; }
/* Promotion rows */
tr.promo td {
    background: #FFDE7A;
    font-weight: 600;
}
tr.promo td.criterion {
    background: #FFDE7A;
    color: #5B0013;
}
tr.promo.alt td {
    background: #FFCC33;
}
tr.promo.alt td.criterion {
    background: #FFCC33;
    color: #5B0013;
}
"""


def esc(text: str) -> str:
    return html_mod.escape(text)


def build_html(sheet_name: str, data: dict) -> str:
    rows_html = []
    data_row_count = 0

    for group_idx, (group_label, rows) in enumerate(data["groups"]):
        # Separator row
        rows_html.append(
            f'<tr class="separator"><td colspan="5">{esc(group_label)}</td></tr>'
        )
        data_row_count = 0

        for criterion, values, is_promo in rows:
            cls = []
            if is_promo:
                cls.append("promo")
            if data_row_count % 2 == 1:
                cls.append("alt")
            cls_str = f' class="{" ".join(cls)}"' if cls else ""

            cells = f'<td class="criterion">{esc(criterion)}</td>'
            for val in values:
                cells += f"<td>{esc(val)}</td>"
            rows_html.append(f"<tr{cls_str}>{cells}</tr>")
            data_row_count += 1

    headers = "".join(f"<th>{esc(h)}</th>" for h in COL_HEADERS)

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{CSS}</style></head>
<body>
<div class="title">{esc(sheet_name.upper())} PERFORMANCE CRITERIA</div>
<div class="note">{esc(data['note'])}</div>
<table>
<colgroup>
  <col class="criterion">
  <col class="rating"><col class="rating"><col class="rating"><col class="rating">
</colgroup>
<thead><tr>{headers}</tr></thead>
<tbody>
{"".join(rows_html)}
</tbody>
</table>
</body></html>"""


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch()

        for sheet_name, data in tables.items():
            html_content = build_html(sheet_name, data)
            page = browser.new_page(
                viewport={"width": 1100, "height": 800},
                device_scale_factor=2,
            )
            page.set_content(html_content, wait_until="networkidle")

            # Resize viewport to full content height so screenshot captures everything
            content_height = page.evaluate("document.body.scrollHeight")
            page.set_viewport_size({"width": 1100, "height": content_height + 48})

            filename = sheet_name.replace(" ", "_") + ".png"
            filepath = os.path.join(OUTPUT_DIR, filename)
            page.screenshot(path=filepath, full_page=True)
            page.close()
            print(f"  Saved {filepath}")

        browser.close()
    print(f"\nDone! {len(tables)} PNGs saved to {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
