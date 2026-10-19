import os
import time
import requests
from docxtpl import DocxTemplate
import datetime

# --- CONFIGURATION ---
DOWNLOADS_FOLDER = r"C:\Users\r.hamili\Downloads"
TRIGGER_FILE = os.path.join(DOWNLOADS_FOLDER, "generate.txt")

TOKEN = "perm-ci5oYW1pbGk=.NzctMTM3.2DmnYBuMYzQf04AlOrYtisyZago3Rm"
TEMPLATE_FILE = r"C:\Users\r.hamili\youtrack-report\modelefinale.docx"
OUTPUT_DIR = r"C:\Users\r.hamili\youtrack-report"  # Where reports are saved

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
}

PARAMS = {
    "fields": "summary,spentTime(presentation),customFields(name,value(name,fullName,presentation,text)),workItems(date,duration(presentation),text,type(name))"
}


def process_issue(issue_id):
    youtrack_url = f"https://youtrack.meteocontrol.de/api/issues/{issue_id}"
    workitems_url = f"https://youtrack.meteocontrol.de/api/issues/{issue_id}/timeTracking/workItems"

    print(f"\nProcessing Issue ID: {issue_id}...")
    response = requests.get(youtrack_url, headers=HEADERS, params=PARAMS)
    
    if response.status_code != 200:
        print(f" Failed to fetch issue: {response.status_code} - {response.text}")
        return

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
        work_resp = requests.get(workitems_url, headers=HEADERS, params={"fields": "date,duration(presentation),text,type(name)"})
        if work_resp.status_code == 200:
            raw_work_items = work_resp.json()

    raw_work_items.sort(key=lambda x: x.get("date", 0))

    items_list = []
    index = 1

    for item in raw_work_items:
        work_type_obj = item.get("type")
        work_type_name = work_type_obj.get("name", "") if isinstance(work_type_obj, dict) else ""

        raw_text = item.get("text")
        description = (raw_text or "").strip()

        is_commissioning_type = "commissioning" in work_type_name.lower() or "mise en service" in work_type_name.lower()
        has_commissioning_text = "commissioning" in description.lower() or "mise en service" in description.lower()

        if work_type_name != "" and not is_commissioning_type:
            continue
        elif work_type_name == "" and not has_commissioning_text:
            continue

        timestamp = item.get("date")
        date_str = datetime.datetime.fromtimestamp(timestamp / 1000).strftime("%d %b %Y") if timestamp else "N/A"
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

    data = {
        "Project": issue.get("summary", "N/A"),
        "ProjectNumero": project_num,
        "Client": client_name,
        "subcontractor": subcontractor,
        "Assignee": nom_assignee,
        "items": items_list
    }

    output_path = os.path.join(OUTPUT_DIR, f"Rapport_{issue_id}.docx")
    doc = DocxTemplate(TEMPLATE_FILE)
    doc.render(data)
    doc.save(output_path)

    print(f" Report generated successfully: {output_path}")


# --- WATCHER LOOP ---
print(" Listening for generate.txt in Downloads folder...")
while True:
    try:
        if os.path.exists(TRIGGER_FILE):
            time.sleep(0.5)  # Brief pause to ensure browser finishes writing the file
            with open(TRIGGER_FILE, "r", encoding="utf-8") as f:
                issue_id = f.read().strip()

            if issue_id:
                process_issue(issue_id)

            # Cleanup trigger file so it can fire again next time
            os.remove(TRIGGER_FILE)
            print(" Cleaned up generate.txt. Waiting for next trigger...")

    except Exception as e:
        print(f"⚠️ Error: {e}")

    time.sleep(2)  # Check every 2 seconds
