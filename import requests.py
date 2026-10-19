import requests
from docxtpl import DocxTemplate
import os
import sys

# --- CONFIGURATION ---
ISSUE_ID = "FRANCE-8"
TOKEN = "perm-ci5oYW1pbGk=.NzctMTM3.2DmnYBuMYzQf04AlOrYtisyZago3Rm"
YOUTRACK_URL = f"https://youtrack.meteocontrol.de/api/issues/{ISSUE_ID}"

TEMPLATE_FILE = "template.docx"
OUTPUT_FILE = f"Rapport_{ISSUE_ID}.docx"

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
}

# Explicitly requesting customFields and their presentations to get "1w 1d 7h"
params = {
    "fields": "summary,customFields(name,value(name,fullName,presentation,text))"
}

def get_issue_data():
    print(f"--- Fetching Data for {ISSUE_ID} ---")
    try:
        response = requests.get(YOUTRACK_URL, headers=headers, params=params, timeout=10)
        
        if response.status_code == 401:
            print("❌ Error: Unauthorized. Check your API Token.")
            return None
        elif response.status_code == 404:
            print(f"❌ Error: Issue {ISSUE_ID} not found.")
            return None
        
        response.raise_for_status()
        issue = response.json()

        # Initialize context with keys matching placeholders in image_839498.png
        context = {
            "Project": issue.get("summary", "N/A"),
            "ProjectNumero": "N/A",
            "Client": "N/A",
            "subcontractor": "N/A",
            "Assignee": "Non assigné",
            "Spent_time": "0h"
        }

        # Iterating through custom fields to find values shown in image_79a3e7.png
        for field in issue.get("customFields", []):
            name = field.get("name")
            val = field.get("value")

            if not val:
                continue

            # 1. Project Number
            if name in ["Project Number", "Project Numero", "Project Numéro"]:
                context["ProjectNumero"] = val.get("name", str(val))

            # 2. Client
            elif name == "Client":
                context["Client"] = val.get("name", str(val))

            # 3. Intervenant (subcontractor)
            elif name in ["Intervenant", "Subcontractor", "INTERVENANT(S) COMPLEMENTAIRE"]:
                if isinstance(val, list):
                    context["subcontractor"] = ", ".join([v.get("name", "") for v in val])
                else:
                    context["subcontractor"] = val.get("name", str(val))

            # 4. Assignee (Matches 'Patrick Nonki' in image_79aad4.png)
            elif name == "Assignee":
                if isinstance(val, list):
                    context["Assignee"] = ", ".join([v.get("fullName", "") for v in val])
                else:
                    context["Assignee"] = val.get("fullName", "")

            # 5. Spent Time (Matches '1w 1d 7h' in image_79a3e7.png)
            elif name == "Spent time":
                if isinstance(val, dict):
                    context["Spent_time"] = val.get("presentation", "0h")
                else:
                    context["Spent_time"] = str(val)

        return context

    except Exception as e:
        print(f"❌ API Request Failed: {e}")
        return None

def main():
    # 1. Get Data
    data = get_issue_data()
    if not data:
        print("🚨 Script stopped: Could not retrieve data.")
        return

    print("✅ Data Extracted Successfully:")
    for k, v in data.items():
        print(f"   {k}: {v}")

    # 2. Check Template
    if not os.path.exists(TEMPLATE_FILE):
        print(f"❌ Error: Template file '{TEMPLATE_FILE}' not found.")
        return

    # 3. Render and Save
    try:
        doc = DocxTemplate(TEMPLATE_FILE)
        doc.render(data)
        doc.save(OUTPUT_FILE)
        print(f"\n🎉 SUCCESS: File created -> {OUTPUT_FILE}")
        
    except PermissionError:
        # Prevents the crash shown in image_79aa57.png
        print(f"❌ PERMISSION ERROR: Close '{OUTPUT_FILE}' in Word and run the script again.")
    except Exception as e:
        print(f"❌ Rendering Error: {e}")

if __name__ == "__main__":
    main()