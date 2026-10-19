import requests
from docxtpl import DocxTemplate
import os
import datetime

# --- CONFIG ---
ISSUE_ID = "sysdes-33612"
TOKEN = "perm-ci5oYW1pbGk=.NzctMTM3.2DmnYBuMYzQf04AlOrYtisyZago3Rm"

YOUTRACK_URL = f"https://youtrack.meteocontrol.de/api/issues/{ISSUE_ID}"

TEMPLATE_FILE = r"C:\Users\r.hamili\youtrack-report\modelefinale.docx"
OUTPUT_FILE = f"Rapport_{ISSUE_ID}.docx"

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
}

params = {
    "fields": "summary,spentTime(presentation),customFields(name,value(name,fullName,presentation,text)),workItems(date,duration(presentation),text,type(name))"
}

print("=== START ===")


def get_issue_data():
    print("Calling API...")
    response = requests.get(YOUTRACK_URL, headers=headers, params=params)
    print("Status:", response.status_code)

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

        if name == "Assignee":
            nom_assignee = val.get("fullName", "Non assigné")
        elif name in ["Project Number", "Project Numero", "Project Numéro"]:
            project_num = val.get("name", str(val))
        elif name == "Client":
            client_name = val.get("name", str(val))
        elif name in ["Subcontractor", "Sous-traitant"]:
            subcontractor = val.get("name", str(val))

    raw_work_items = issue.get("workItems", [])
    
    # Sort from Oldest to Newest so JOUR 1 is the earliest date
    raw_work_items.sort(key=lambda x: x.get("date", 0))

    items_list = []
    index = 1

    for item in raw_work_items:
        work_type_obj = item.get("type")
        work_type_name = work_type_obj.get("name", "") if work_type_obj else ""

        # Include it if it's Commissioning or if no explicit type is set
        if "commissioning" not in work_type_name.lower() and work_type_name != "":
            continue

        timestamp = item.get("date")
        date_str = "N/A"
        if timestamp:
            date_str = datetime.datetime.fromtimestamp(timestamp / 1000).strftime("%d %b %Y")
        
        duration = item.get("duration", {}).get("presentation", "0m")
        description = item.get("text", "").strip()
        if not description:
            description = "Mise en service"

        items_list.append({
            "HEADER_LINE": f"JOUR {index}: {date_str}",
            "DURATION": f" – {duration.upper()} – ",
            "DESCRIPTION": description.upper()
        })
        index += 1

    # If no items were found, add a placeholder row so it's not blank
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
        "items": items_list  # Matches the lowercase {% for ITEM in items %}
    }


data = get_issue_data()

if not data:
    raise Exception("No data retrieved")

if not os.path.exists(TEMPLATE_FILE):
    raise Exception(f"Template not found: {TEMPLATE_FILE}")

doc = DocxTemplate(TEMPLATE_FILE)
doc.render(data)
doc.save(OUTPUT_FILE)

print("✅ File created successfully:", OUTPUT_FILE)