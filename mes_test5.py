import requests
from docxtpl import DocxTemplate
import os

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
    "fields": "summary,spentTime(presentation),customFields(name,value(name,fullName,presentation,text))"
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

    spent_time = issue.get("spentTime")
    temps_passe = spent_time.get("presentation") if spent_time else "0h"

    return {
        "Project_Name": issue.get("summary", "N/A"),
        "Project_Number": project_num,
        "Client": client_name,
        "Assignee": nom_assignee,
        "spenttime": temps_passe
    }


data = get_issue_data()

if not data:
    raise Exception("No data retrieved")

print("Data:", data)

if not os.path.exists(TEMPLATE_FILE):
    raise Exception(f"Template not found: {TEMPLATE_FILE}")

doc = DocxTemplate(TEMPLATE_FILE)
doc.render(data)
doc.save(OUTPUT_FILE)

print("File created:", OUTPUT_FILE)
