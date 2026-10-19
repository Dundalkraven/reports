import datetime
import os
from docxtpl import DocxTemplate
from flask import Flask, request, jsonify
from flask_cors import CORS
import requests

app = Flask(__name__)
CORS(app)  # Enables cross-origin requests from YouTrack

# --- CONFIG ---
TOKEN = "perm-ci5oYW1pbGk=.NzctMTM3.2DmnYBuMYzQf04AlOrYtisyZago3Rm"
TEMPLATE_FILE = r"C:\Users\r.hamili\youtrack-report\modelefinale.docx"

headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}

params = {
    "fields": "summary,customFields(name,value(name,fullName,presentation,text,minutes)),workItems(date,duration(presentation),text,type(name))"
}


def generate_commissioning_report(issue_id):
  youtrack_url = f"https://youtrack.meteocontrol.de/api/issues/{issue_id}"
  workitems_url = (
      f"https://youtrack.meteocontrol.de/api/issues/{issue_id}/timeTracking/workItems"
  )
  output_file = f"Rapport_{issue_id}.docx"

  print(f"=== START FOR ISSUE: {issue_id} ===")

  print("Calling Issue API...")
  response = requests.get(youtrack_url, headers=headers, params=params)
  print("Issue Status:", response.status_code)

  if response.status_code != 200:
    raise Exception(f"YouTrack API Error: {response.text}")

  issue = response.json()

  nom_assignee = "Non assigné"
  project_num = "N/A"
  client_name = "N/A"
  subcontractor = "N/A"
  total_spent_time = "N/A"

  for field in issue.get("customFields", []):
    name = field.get("name")
    val = field.get("value")

    if not val:
      continue

    extracted_val = (
        val.get("name", str(val)) if isinstance(val, dict) else str(val)
    )

    if name == "Assignee":
      nom_assignee = (
          val.get("fullName", "Non assigné")
          if isinstance(val, dict)
          else str(val)
      )
    elif name in ["Project Number", "Project Numero", "Project Numéro"]:
      project_num = extracted_val
    elif name == "Client":
      client_name = extracted_val
    elif name in ["Subcontractor", "Sous-traitant"]:
      subcontractor = extracted_val
    elif name in ["Spent time", "Spent Time", "Temps passé"]:
      if isinstance(val, dict):
        total_spent_time = val.get("presentation", str(val))
      else:
        total_spent_time = str(val)

  raw_work_items = issue.get("workItems", [])

  if not raw_work_items:
    print("Fetching from dedicated timeTracking endpoint...")
    work_response = requests.get(
        workitems_url,
        headers=headers,
        params={"fields": "date,duration(presentation),text,type(name)"},
    )
    if work_response.status_code == 200:
      raw_work_items = work_response.json()

  raw_work_items.sort(key=lambda x: x.get("date", 0))

  items_list = []
  index = 1

  for item in raw_work_items:
    work_type_obj = item.get("type")
    work_type_name = ""
    if isinstance(work_type_obj, dict):
      work_type_name = work_type_obj.get("name", "")

    description = item.get("text", "").strip()

    is_commissioning_type = (
        "commissioning" in work_type_name.lower()
        or "mise en service" in work_type_name.lower()
    )
    has_commissioning_text = (
        "commissioning" in description.lower()
        or "mise en service" in description.lower()
    )

    if work_type_name != "" and not is_commissioning_type:
      continue
    elif work_type_name == "" and not has_commissioning_text:
      continue

    timestamp = item.get("date")
    date_str = "N/A"
    if timestamp:
      date_str = datetime.datetime.fromtimestamp(timestamp / 1000).strftime(
          "%d %b %Y"
      )

    duration = item.get("duration", {}).get("presentation", "0m")
    if not description:
      description = "MISE EN SERVICE"

    items_list.append({
        "HEADER_LINE": f"JOUR {index}: {date_str}",
        "DURATION": f" – {duration.upper()} – ",
        "DESCRIPTION": description.upper(),
    })
    index += 1

  if not items_list:
    items_list.append({
        "HEADER_LINE": "Aucun historique",
        "DURATION": " - ",
        "DESCRIPTION": "Aucune donnée de mise en service trouvée.",
    })

  data = {
      "Project": issue.get("summary", "N/A"),
      "ProjectNumero": project_num,
      "Client": client_name,
      "subcontractor": subcontractor,
      "Assignee": nom_assignee,
      "TotalSpentTime": total_spent_time,
      "items": items_list,
  }

  if not os.path.exists(TEMPLATE_FILE):
    raise Exception(f"Template not found: {TEMPLATE_FILE}")

  doc = DocxTemplate(TEMPLATE_FILE)
  doc.render(data)
  doc.save(output_file)

  print(" File created successfully:", output_file)
  return output_file


@app.route("/", methods=["GET"])
def home():
  return (
      jsonify(
          {
              "status": "active",
              "message": "YouTrack Report Server is running.",
          }
      ),
      200,
  )


@app.route("/generate-report", methods=["POST"])
def trigger_report():
  req_data = request.get_json()

  if not req_data or "issueId" not in req_data:
    return (
        jsonify(
            {"status": "error", "message": "Missing 'issueId' in request body"}
        ),
        400,
    )

  issue_id = req_data["issueId"]

  try:
    output_file = generate_commissioning_report(issue_id)
    return (
        jsonify(
            {
                "status": "success",
                "message": f"Report generated successfully: {output_file}",
            }
        ),
        200,
    )
  except Exception as e:
    print(" Error generating report:", str(e))
    return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == "__main__":
  app.run(host="0.0.0.0", port=5000, ssl_context="adhoc")
