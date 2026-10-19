import sys
import re
import requests
from docxtpl import DocxTemplate
import os
import datetime
import time

# --- CONFIG ---
ISSUE_ID = "sysdes-26150"
TOKEN = "perm-ci5oYW1pbGk=.NzctMTM3.2DmnYBuMYzQf04AlOrYtisyZago3Rm"

YOUTRACK_URL = f"https://youtrack.meteocontrol.de/api/issues/{ISSUE_ID}"
WORKITEMS_URL = f"https://youtrack.meteocontrol.de/api/issues/{ISSUE_ID}/timeTracking/workItems"

TEMPLATE_FILE = r"C:\Users\r.hamili\youtrack-report\modelefinale.docx"
OUTPUT_FILE = f"Rapport_{ISSUE_ID}.docx"

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
}

params = {
    "fields": "summary,spentTime(presentation),customFields(name,value(name,fullName,presentation,text,minutes))"
}

def parse_duration_to_minutes(duration_str):
    """Converts YouTrack duration strings (e.g. '1w 1d 7h 30m', '2h', '45m') into total minutes."""
    if not duration_str:
        return 0
    total_minutes = 0
    weeks = re.search(r"(\d+)\s*w", duration_str, re.I)
    days = re.search(r"(\d+)\s*d", duration_str, re.I)
    hours = re.search(r"(\d+)\s*h", duration_str, re.I)
    mins = re.search(r"(\d+)\s*m", duration_str, re.I)
    
    if weeks: total_minutes += int(weeks.group(1)) * 5 * 8 * 60  # 5d/wk, 8h/day
    if days: total_minutes += int(days.group(1)) * 8 * 60         # 8h/day
    if hours: total_minutes += int(hours.group(1)) * 60
    if mins: total_minutes += int(mins.group(1))
    
    return total_minutes

def format_minutes_to_display(total_minutes):
    """Formats total minutes into human-readable format e.g. '31h 30m' or '31h'."""
    if total_minutes <= 0:
        return "0m"
    hours = total_minutes // 60
    mins = total_minutes % 60
    if hours > 0 and mins > 0:
        return f"{hours}h {mins}m"
    elif hours > 0:
        return f"{hours}h"
    else:
        return f"{mins}m"

def get_issue_data():
    print("Calling Issue API...")
    response = requests.get(YOUTRACK_URL, headers=headers, params=params)
    if response.status_code != 200:
        print("Error:", response.text)
        return None

    issue = response.json()

    nom_assignee = "Non assigné"
    project_num = "N/A"
    client_name = "N/A"
    subcontractor = "N/A"
    total_spent_time = ""

    # 1. Target Spent time in Custom Fields directly
    for field in issue.get("customFields", []):
        name = field.get("name", "")
        val = field.get("value")

        if not val:
            continue

        if name.lower() in ["spent time", "temps passé"]:
            if isinstance(val, dict):
                total_spent_time = val.get("presentation") or val.get("name") or ""
            elif isinstance(val, (int, float)):
                total_spent_time = format_minutes_to_display(val)
            else:
                total_spent_time = str(val)

        extracted_val = val.get("name", str(val)) if isinstance(val, dict) else str(val)

        if name == "Assignee":
            nom_assignee = val.get("fullName", "Non assigné") if isinstance(val, dict) else str(val)
        elif name in ["Project Number", "Project Numero", "Project Numéro"]:
            project_num = extracted_val
        elif name == "Client":
            client_name = extracted_val
        elif name in ["Subcontractor", "Sous-traitant"]:
            subcontractor = extracted_val

    # 2. Fallback to root spentTime if custom field was empty
    if not total_spent_time:
        spent_time_obj = issue.get("spentTime") or {}
        if isinstance(spent_time_obj, dict):
            total_spent_time = spent_time_obj.get("presentation", "")

    # 3. Fetch Work Items
    work_response = requests.get(WORKITEMS_URL, headers=headers, params={"fields": "date,duration(presentation),text,type(name)"})
    raw_work_items = work_response.json() if work_response.status_code == 200 else []

    raw_work_items.sort(key=lambda x: x.get("date", 0))

    items_list = []
    index = 1
    calculated_total_minutes = 0

    for item in raw_work_items:
        duration_str = (item.get("duration") or {}).get("presentation", "0m")
        calculated_total_minutes += parse_duration_to_minutes(duration_str)

        raw_text = item.get("text") or ""
        description = raw_text.strip() or "MISE EN SERVICE"

        timestamp = item.get("date")
        date_str = datetime.datetime.fromtimestamp(timestamp / 1000).strftime("%d %b %Y") if timestamp else "N/A"

        items_list.append({
            "HEADER_LINE": f"JOUR {index}: {date_str}",
            "DURATION": f" – {duration_str.upper()} – ",
            "DESCRIPTION": description.upper()
        })
        index += 1

    if not items_list:
        items_list.append({
            "HEADER_LINE": "Aucun historique",
            "DURATION": " - ",
            "DESCRIPTION": "Aucune donnée de mise en service trouvée."
        })

    # 4. Fallback calculation if YouTrack field returned nothing
    if not total_spent_time or total_spent_time.strip() == "":
        total_spent_time = format_minutes_to_display(calculated_total_minutes)

    formatted_spent_time = total_spent_time.upper()

    return {
        "Project": issue.get("summary", "N/A"),
        "ProjectNumero": project_num,
        "Client": client_name,
        "subcontractor": subcontractor,
        "Assignee": nom_assignee,
        "TotalSpentTime": formatted_spent_time,
        "TOTALSPENTTIME": formatted_spent_time,
        "items": items_list
    }

def render_report():
    print(f"\n[{datetime.datetime.now().strftime('%H:%M:%S')}] Rendering report...")
    data = get_issue_data()

    if not data:
        print("❌ Failed to retrieve data.")
        return

    if not os.path.exists(TEMPLATE_FILE):
        print(f"❌ Template not found: {TEMPLATE_FILE}")
        return

    try:
        doc = DocxTemplate(TEMPLATE_FILE)
        doc.render(data)
        doc.save(OUTPUT_FILE)
        print(f"✅ File updated successfully! Extracted Spent Time: {data.get('TotalSpentTime')}")
    except PermissionError:
        print(f"⚠️ Permission Error: Please close '{OUTPUT_FILE}' in Word so the script can write to it.")
    except Exception as e:
        print(f"❌ Error rendering document: {e}")

# --- WATCHER LOOP ---
if __name__ == "__main__":
    print(f"👀 Watching template file for changes: {TEMPLATE_FILE}")
    print("Press Ctrl+C to stop watching.")

    # Render once on startup
    render_report()

    last_mtime = os.path.getmtime(TEMPLATE_FILE) if os.path.exists(TEMPLATE_FILE) else 0

    try:
        while True:
            time.sleep(1)  # Check every second
            if os.path.exists(TEMPLATE_FILE):
                current_mtime = os.path.getmtime(TEMPLATE_FILE)
                if current_mtime != last_mtime:
                    last_mtime = current_mtime
                    print(f"\n🔄 Change detected in template file!")
                    render_report()
    except KeyboardInterrupt:
        print("\n👋 Stopped watching file.")