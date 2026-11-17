import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from docx import Document

DATA = Path("data")
REPORTS = Path("reports")
DOCS = Path("docs")

DATA.mkdir(exist_ok=True)
REPORTS.mkdir(exist_ok=True)
DOCS.mkdir(exist_ok=True)

sites = ["SITE-101", "SITE-102", "SITE-103", "SITE-104", "SITE-105"]

milestones = [
    ("Site Survey", 0, 2, "Field Engineer"),
    ("Permitting", 3, 7, "PM Coordinator"),
    ("Equipment Delivery", 8, 10, "Logistics"),
    ("Installation", 11, 14, "Installation Crew"),
    ("Acceptance", 15, 16, "Project Manager"),
]

base_date = pd.Timestamp("2027-05-03")

rows = []

for site_index, site in enumerate(sites):
    offset = site_index * 3

    for milestone_index, (milestone, start_offset, end_offset, owner) in enumerate(milestones):
        start = base_date + pd.Timedelta(days=start_offset + offset)
        end = base_date + pd.Timedelta(days=end_offset + offset)

        dependency = (
            "None"
            if milestone_index == 0
            else milestones[milestone_index - 1][0]
        )

        rows.append({
            "Site ID": site,
            "Milestone": milestone,
            "Owner": owner,
            "Planned Start": start.date(),
            "Planned End": end.date(),
            "Dependency": dependency,
            "Status": "On Track",
            "Percent Complete": 0
        })

schedule = pd.DataFrame(rows)

# Simulate progress
schedule.loc[
    (schedule["Site ID"] == "SITE-101") &
    (schedule["Milestone"].isin(["Site Survey", "Permitting"])),
    ["Status", "Percent Complete"]
] = ["Complete", 100]

schedule.loc[
    (schedule["Site ID"] == "SITE-102") &
    (schedule["Milestone"] == "Site Survey"),
    ["Status", "Percent Complete"]
] = ["Complete", 100]

schedule.to_csv(DATA / "deployment_schedule.csv", index=False)

scenarios = pd.DataFrame([
    {
        "Scenario ID": "SCN-01",
        "Site ID": "SITE-101",
        "Issue": "Permit approval delayed",
        "Affected Milestone": "Permitting",
        "Delay Days": 4,
        "Resource Conflict": "No",
        "Proposed Action": "Move equipment staging forward while permit remains pending"
    },
    {
        "Scenario ID": "SCN-02",
        "Site ID": "SITE-102",
        "Issue": "Installation crew unavailable",
        "Affected Milestone": "Installation",
        "Delay Days": 3,
        "Resource Conflict": "Yes",
        "Proposed Action": "Reassign Crew B after SITE-101 installation"
    },
    {
        "Scenario ID": "SCN-03",
        "Site ID": "SITE-103",
        "Issue": "Equipment delivery delayed",
        "Affected Milestone": "Equipment Delivery",
        "Delay Days": 5,
        "Resource Conflict": "No",
        "Proposed Action": "Resequence site preparation activities"
    },
    {
        "Scenario ID": "SCN-04",
        "Site ID": "SITE-104",
        "Issue": "Survey engineer double-booked",
        "Affected Milestone": "Site Survey",
        "Delay Days": 2,
        "Resource Conflict": "Yes",
        "Proposed Action": "Move survey to alternate engineer"
    },
    {
        "Scenario ID": "SCN-05",
        "Site ID": "SITE-105",
        "Issue": "Acceptance engineer unavailable",
        "Affected Milestone": "Acceptance",
        "Delay Days": 2,
        "Resource Conflict": "Yes",
        "Proposed Action": "Schedule remote acceptance review"
    },
    {
        "Scenario ID": "SCN-06",
        "Site ID": "SITE-103",
        "Issue": "Permit and crew timing overlap",
        "Affected Milestone": "Installation",
        "Delay Days": 3,
        "Resource Conflict": "Yes",
        "Proposed Action": "Shift crew to SITE-105 until permit clears"
    }
])

scenarios.to_csv(DATA / "delay_scenarios.csv", index=False)

actions = pd.DataFrame([
    ["ACT-001", "SITE-101", "Confirm revised permit date", "PM Coordinator", "2027-05-12", "Open"],
    ["ACT-002", "SITE-102", "Confirm alternate installation crew", "Project Manager", "2027-05-15", "In Progress"],
    ["ACT-003", "SITE-103", "Confirm equipment shipment ETA", "Logistics", "2027-05-18", "Open"],
    ["ACT-004", "SITE-104", "Assign alternate survey engineer", "Resource Manager", "2027-05-10", "Complete"],
    ["ACT-005", "SITE-105", "Schedule acceptance reviewer", "Project Manager", "2027-05-25", "Open"],
])

actions.columns = [
    "Action ID",
    "Site ID",
    "Action",
    "Owner",
    "Due Date",
    "Status"
]

actions.to_csv(DATA / "action_log.csv", index=False)

change_request = pd.DataFrame([
    {
        "Change ID": "CR-001",
        "Site ID": "SITE-103",
        "Requested Change": "Move installation milestone by 3 days",
        "Reason": "Equipment delivery delay",
        "Schedule Impact": "Installation and acceptance move by 3 days",
        "Cost Impact": "No estimated cost increase",
        "Decision": "Pending Review"
    }
])

change_request.to_csv(DATA / "change_request.csv", index=False)

excel_file = REPORTS / "network_site_deployment_tracker.xlsx"

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    schedule.to_excel(writer, sheet_name="Deployment Schedule", index=False)
    scenarios.to_excel(writer, sheet_name="Delay Scenarios", index=False)
    actions.to_excel(writer, sheet_name="Action Log", index=False)
    change_request.to_excel(writer, sheet_name="Change Request", index=False)

wb = load_workbook(excel_file)

for ws in wb.worksheets:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            try:
                max_length = max(max_length, len(str(cell.value or "")))
            except Exception:
                pass

        ws.column_dimensions[column_letter].width = min(max_length + 2, 45)

schedule_ws = wb["Deployment Schedule"]

status_column = None

for cell in schedule_ws[1]:
    if cell.value == "Percent Complete":
        status_column = cell.column_letter
        break

if status_column:
    schedule_ws.conditional_formatting.add(
        f"{status_column}2:{status_column}26",
        CellIsRule(
            operator="lessThan",
            formula=["50"],
            fill=PatternFill(
                start_color="FFC7CE",
                end_color="FFC7CE",
                fill_type="solid"
            )
        )
    )

wb.save(excel_file)

doc = Document()

doc.add_heading(
    "Network Site Deployment Schedule and Resource Tracker",
    0
)

doc.add_paragraph(
    "This project demonstrates PMO-style planning and coordination "
    "for five sample cellular network deployment sites."
)

doc.add_heading("Project Scope", level=1)

doc.add_paragraph(
    "The tracker covers site survey, permitting, equipment delivery, "
    "installation, and acceptance activities."
)

doc.add_heading("Schedule Summary", level=1)
doc.add_paragraph("Sites tracked: 5")
doc.add_paragraph("Milestones tracked: 25")
doc.add_paragraph("Delay/resource scenarios tested: 6")

doc.add_heading("PMO Process", level=1)

for text in [
    "Establish planned milestones and owners.",
    "Define milestone dependencies.",
    "Monitor schedule status and completion.",
    "Identify delays and resource conflicts.",
    "Document corrective actions.",
    "Evaluate schedule changes through a change request.",
    "Maintain an action log until issues are resolved."
]:
    doc.add_paragraph(text, style="List Bullet")

doc.add_heading("Key Learning", level=1)

doc.add_paragraph(
    "A deployment schedule is not only a list of dates. Effective project "
    "coordination requires understanding dependencies, resource availability, "
    "schedule impact, ownership, and follow-up actions."
)

doc.save(DOCS / "project_summary.docx")

print("PROJECT COMPLETE")
print("25 milestones created")
print("6 delay/resource scenarios created")
print("Excel tracker:", excel_file)
print("Word summary:", DOCS / "project_summary.docx")
