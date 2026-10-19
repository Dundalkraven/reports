import sys
import re
import requests
from docxtpl import DocxTemplate
import os
import datetime

# Check if an Issue ID was passed to the script via Tampermonkey / Command Line
if len(sys.argv) > 1:
    raw_issue_id = sys.argv[1]
else:
    raise ValueError("Error: No Issue ID provided. Usage: python script.py <ISSUE_ID>")

# Clean up input to extract raw issue ID (handles raw ID or full URL if passed)
match = re.search(r"([a-zA-Z0-9]+-\d+)", raw_issue_id)
ISSUE_ID = match.group(1) if match else raw_issue_id.strip()

print(f"=== PROCESSING ISSUE: {ISSUE_ID} ===")

# --- CONFIG ---
TOKEN = "perm-ci5oYW1pbGk=.NzctMTM3.2DmnYBuMYzQf04AlOrYtisyZago3Rm"

YOUTRACK_URL = f"https://youtrack.meteocontrol.de/api/issues/{ISSUE_ID}"
WORKITEMS_URL = f"https://youtrack.meteocontrol.de/api/issues/{ISSUE_ID}/timeTracking/workItems"

TEMPLATE_FILE = r"C:\Users\r.hamili\youtrack-report\spenttimetest.docx"
OUTPUT_FILE = f"Rapport_{ISSUE_ID}.docx"

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
}

params = {
    "fields": "summary,spentTime(presentation),customFields(name,value(name,fullName,presentation,text)),workItems(date,duration(presentation),text,type(name))"
}


def get_issue_data():
    print("Calling Issue API...")
    response = requests.get(YOUTRACK_URL, headers=headers, params=params)
    print("Issue Status:", response.status_code)

    if response.status_code != 200:
        print("Error:", response.text)
        return None

    issue = response.json()

    nom_assignee = "Non assigné"
    project_num = "N/A"
    client_name = "N/A"
    subcontractor = "N/A"

    for field in issue.get("customFields", []):
        name = field.get("name")
        val = field.get("value")

        if not val:
            continue

        # Safe extraction for both dict objects and raw strings
        extracted_val = val.get("name", str(val)) if isinstance(val, dict) else str(val)

        if name == "Assignee":
            nom_assignee = val.get("fullName", "Non assigné") if isinstance(val, dict) else str(val)
        elif name in ["Project Number", "Project Numero", "Project Numéro"]:
            project_num = extracted_val
        elif name == "Client":
            client_name = extracted_val
        elif name in ["Subcontractor", "Sous-traitant"]:
            subcontractor = extracted_val

    raw_work_items = issue.get("workItems", [])
    
    if not raw_work_items:
        print("Fetching from dedicated timeTracking endpoint...")
        work_response = requests.get(WORKITEMS_URL, headers=headers, params={"fields": "date,duration(presentation),text,type(name)"})
        if work_response.status_code == 200:
            raw_work_items = work_response.json()

    # Sort from Oldest to Newest
    raw_work_items.sort(key=lambda x: x.get("date", 0))

    items_list = []
    index = 1

    for item in raw_work_items:
        # Extract the Work Type Category securely
        work_type_obj = item.get("type")
        work_type_name = ""
        if isinstance(work_type_obj, dict):
            work_type_name = work_type_obj.get("name") or ""

        # Safe extraction of text (handles None, empty string, or missing key)
        raw_text = item.get("text")
        description = (raw_text or "").strip()

        # STRICT FILTERING:
        is_commissioning_type = "commissioning" in work_type_name.lower() or "mise en service" in work_type_name.lower()
        has_commissioning_text = "commissioning" in description.lower() or "mise en service" in description.lower()

        if work_type_name != "" and not is_commissioning_type:
            # Excludes explicit categories like Support, Travel Time, Meeting, etc.
            continue
        elif work_type_name == "" and not has_commissioning_text:
            # If no category is assigned, skip unless text indicates commissioning work
            continue

        timestamp = item.get("date")
        date_str = "N/A"
        if timestamp:
            date_str = datetime.datetime.fromtimestamp(timestamp / 1000).strftime("%d %b %Y")
        
        duration = (item.get("duration") or {}).get("presentation", "0m")
        if not description:
            description = "MISE EN SERVICE"

        items_list.append({
            "HEADER_LINE": f"JOUR {index}: {date_str}",
            "DURATION": f" – {duration.upper()} – ",
            "DESCRIPTION": description.upper()
        })
        index += 1

    if not items_list:
        items_list.append({
            "HEADER_LINE": "Aucun historique",
            "DURATION": " - ",
            "DESCRIPTION": "Aucune donnée de mise en service trouvée."
        })

    return {
        "Project": issue.get("summary", "N/A"),
        "ProjectNumero": project_num,
        "Client": client_name,
        "subcontractor": subcontractor,
        "Assignee": nom_assignee,
        "items": items_list
    }


data = get_issue_data()

if not data:
    raise Exception("No data retrieved")

if not os.path.exists(TEMPLATE_FILE):
    raise Exception(f"Template not found: {TEMPLATE_FILE}")

doc = DocxTemplate(TEMPLATE_FILE)
doc.render(data)
doc.save(OUTPUT_FILE)

print(" File created successfully:", OUTPUT_FILE)