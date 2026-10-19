import requests
import urllib3

# Suppress self-signed certificate warnings in test script
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://192.168.160.79:5000/generate-report"
payload = {"issueId": "sysdes-31219"}

try:
    response = requests.post(url, json=payload, verify=False)
    print("Status Code:", response.status_code)
    print("Response:", response.text)
except Exception as e:
    print("Connection failed:", e)