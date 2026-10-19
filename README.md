# 🚀 Commissionning Report Automation solution

An automated utility and Flask server integration designed to extract technical issue data, custom fields, and time tracking work logs from YouTrack and render professional Word commissioning reports (`.docx`) 
using Python.


Note:  This solutio is a temporaty solution to test the report generation since no dedicted server is available. Other options to consider:
    - use of cloud server( Azure, or AWS) >> paid on the go
    - replace the local Flask server with a Power Automate Cloud Flow. Tampermonkey sends the YouTrack issueId to an HTTP Webhook inside Power Automate. Power Automate then fetches the data from YouTrack,
      populates a Word template in OneDrive or SharePoint, converts it to a PDF or .docx, and saves or emails the final report.
---

## 🛠️ Tech Stack & Tools

* **Language:** Python
* **Libraries:** `requests`, `docxtpl`, `Flask`, `flask-cors`, `re`, `datetime`
* **Data Source:** YouTrack REST API (Issues & Time Tracking endpoints)
* **Output Format:** Microsoft Word Document Templates (`.docx`)
* **Integration:** Local Flask API server triggered via browser extensions (Tampermonkey)

---

## ✨ Key Features

* **Automated Data Extraction:** Pulls summary metadata, project numbers, clients, subcontractors, and assigned engineers directly from YouTrack custom fields.
* **Intelligent Time Tracking & Fallbacks:** Automatically parses YouTrack's `spentTime` or calculates total duration dynamically from individual work logs.
* **Strict Work Item Filtering:** Filters work items specifically by commissioning categories or descriptive keywords (`commissioning` / `mise en service`).
* **Template Rendering:** Populates structured Word document tables and chronological logs using `docxtpl`.
* **Local Web Server (Flask):** Exposes an endpoint (`/generate-report`) to trigger report generation instantly from browser workflows.

---

## ⚙️ Quick Start

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/your-username/your-repo-name.git](https://github.com/your-username/your-repo-name.git)
   cd your-repo-name

   OR

   Download the scripts from this repository:   https://github.com/Dundalkraven/reports/

   Important:  Since i m using a private assigned IP, only users within LAN range can access the flask server (running on my local PC). Users outside the local network should use office VPN when connecting remotely.
   If this does not work, we would have to set up peer to peer netwoking to be able to access user PCs
  
