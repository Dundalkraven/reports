import os
import requests
import re
import datetime
from docxtpl import DocxTemplate

# --- CONFIGURATION & URL INPUT ---
YOUTRACK_URL = "https://youtrack.meteocontrol.de/issue/sysdes-31219/Photovolt-SAS-Autoconso-UP8-Commissioning"

TOKEN = "perm-ci5oYW1pbGk=.NzctMTM3.2DmnYBuMYzQf04AlOrYtisyZago3Rm"
BASE_URL = "https://youtrack.meteocontrol.de/api"
TEMPLATE_PATH = "template.docx"

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/json"
}

def extract_id_from_url(url):
    match = re.search(r'/issue/([A-Za-z]+-\d+)', url)
    if match:
        issue_id = match.group(1).lower()
        print(f"Dynamic ID Extracted from URL: {issue_id}")
        return issue_id
    return None

def get_french_label(work_type_obj, description_text):
    """
    Categorizes the work item precisely to align with the template categories.
    """
    t_name = ""
    if work_type_obj and isinstance(work_type_obj, dict):
        t_name = str(work_type_obj.get('name', '')).lower().strip()
    
    desc = str(description_text).lower().strip()
    
    # 1. Check for missed appointments / delays / unsuccessful attempts
    if any(kw in desc for kw in ["delay", "non honoré", "unsuccessful", "annulé", "reporté"]):
        return "RDV NON HONORE : MOBILISATION"
        
    # 2. Check for technical documentation tasks
    if "documentation" in desc or "rapport" in desc:
        return "DOCUMENTATION - ECHANGES"

    # 3. Check for commissioning (remote execution)
    if "commissioning" in t_name or "mise en service" in t_name or "commissioning" in desc:
        return "MISE EN SERVICE A DISTANCE"
        
    # 4. Check for standard support and communication tracks
    if any(kw in t_name or kw in desc for kw in ["support", "email", "follow-up", "call", "téléphone", "échange"]):
        return "ASSISTANCE / ECHANGES"
        
    # Safe fallback if undetermined
    return "ASSISTANCE / ECHANGES"

def format_duration_clean(minutes):
    hrs = minutes // 60
    mins = minutes % 60
    return f"{hrs}H{mins:02d}" if hrs > 0 else f"{mins}M"

def get_issue_data(issue_id):
    print(f"--- Fetching Field Attributes for {issue_id} ---")
    
    params = {
        "fields": "id,summary,customFields(name,value(name,fullName,presentation,text))"
    }
    
    try:
        response = requests.get(f"{BASE_URL}/issues/{issue_id}", headers=headers, params=params, timeout=10)
        if response.status_code != 200:
            print(f"API Error: {response.status_code}")
            return None
        
        issue = response.json()
        internal_id = issue.get("id")
        summary = issue.get("summary", "")
        parts = [p.strip() for p in summary.split("-")]
        
        installation_tasks = [
            "Analyse des risques", "Etat des lieux de l’installation",
            "Contrôle des composants après transport/ installation",
            "Vérification polarité 230V", "Mise en place Fusibles Batterie",
            "Vérification polarité 24VDC", "Alimentation des composants"
        ]
        
        context = {
            "Client": parts[0] if len(parts) > 0 else "N/A",
            "Project": parts[1] if len(parts) > 1 else "N/A",
            "Subsystem": parts[2] if len(parts) > 2 else "N/A",
            "ProjectNumero": "N/A",
            "OrderNo": "N/A",
            "CustomerPA": "N/A",
            "subcontractor": "N/A",
            "Assignee": "Non assigné",
            "Spent_time": "0h",
            "energy_tasks": installation_tasks,
            
            # This handles your primary Chronology table tracking
            "COMM_SITE": []
        }

        # Sidebar properties processing
        for field in issue.get("customFields", []):
            name = field.get("name")
            val = field.get("value")
            if not val: continue
            if name == "Project Number":
                context["ProjectNumero"] = val.get("name", str(val)) if isinstance(val, dict) else str(val)
            elif name == "Order No":
                context["OrderNo"] = str(val)
            elif name == "Customer pA":
                context["CustomerPA"] = str(val)
            elif name == "Assignee":
                context["Assignee"] = val.get("fullName", "")
            elif name == "Spent time":
                context["Spent_time"] = val.get("presentation", "0h") if isinstance(val, dict) else str(val)

        # Log extraction pipeline
        work_url = f"{BASE_URL}/issues/{internal_id}/timeTracking/workItems"
        work_params = {"fields": "id,duration(minutes),date,text,workType(name)"}
        work_res = requests.get(work_url, headers=headers, params=work_params, timeout=10)
        
        if work_res.status_code == 200:
            items = work_res.json()
            
            # Sort items chronologically so Row 1 starts from the earliest action logged
            items.sort(key=lambda x: (x.get('date', 0), x.get('id', '')))
            
            for index, target_item in enumerate(items, start=1):
                description = target_item.get('text') or ""
                w_type_obj = target_item.get('workType')
                
                # Dynamic translation logic mapping
                display_category = get_french_label(w_type_obj, description)
                
                date_str = datetime.datetime.fromtimestamp(target_item.get('date') / 1000).strftime('%d/%m/%Y')
                duration_str = format_duration_clean(target_item.get('duration', {}).get('minutes', 0))
                
                # Append matching object items into your loop tracker
                context["COMM_SITE"].append({
                    "HEADER_LINE": f"JOUR {index}: {date_str}",
                    "DURATION": f"{duration_str} – {display_category}",
                    "DESCRIPTION": description.strip()
                })
                
            print(f"✔ Successfully compiled and categorized {len(context['COMM_SITE'])} items for the report.")
        else:
            print(f"Error fetching time tracking items from API: {work_res.status_code}")

        return context

    except Exception as e:
        print(f"Processing Error: {e}")
        return None

def main():
    issue_id = extract_id_from_url(YOUTRACK_URL)
    if not issue_id: return

    data = get_issue_data(issue_id)
    if not data: return

    output_path = f"Rapport_{issue_id}.docx"
    if not os.path.exists(TEMPLATE_PATH):
        print(f"Error: Template file '{TEMPLATE_PATH}' was not found.")
        return

    try:
        doc = DocxTemplate(TEMPLATE_PATH)
        doc.render(data)
        doc.save(output_path)
        print(f" SUCCESS: Finished populating activities! Output saved to -> [{output_path}].")
    except PermissionError:
        print(f"ERROR: Close '{output_path}' in Word before running this script.")
    except Exception as e:
        print(f"Render Exception: {e}")

if __name__ == "__main__":
    main()
