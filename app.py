import os
import io
from datetime import datetime

import requests
import dropbox
from flask import Flask, request, jsonify
from docx import Document

app = Flask(__name__)

MONDAY_API_KEY = os.environ.get("MONDAY_API_KEY")
BOARD_ID = os.environ.get("BOARD_ID")
DROPBOX_TOKEN = os.environ.get("DROPBOX_TOKEN")

dbx = dropbox.Dropbox(DROPBOX_TOKEN)

COLUMN_MAP = {
    "{{Engineer}}": "text_mm12tcvb",
    "{{Street}}": "text_mm12vcy9",
    "{{Number}}": "text_mm12jf0w",
    "{{Apartment}}": "text_mm127a33",
    "{{Hit date}}": "date_mm1rnmvg",
    "{{Date}}": "dup__of_90__timeline",
    "{{months}}": "text_mm129ef8"
}

def monday_query(query):
    response = requests.post(
        "https://api.monday.com/v2",
        json={"query": query},
        headers={"Authorization": MONDAY_API_KEY}
    )
    response.raise_for_status()
    return response.json()

def get_item_data(item_id):
    query = f"""
    query {{
      items(ids: [{item_id}]) {{
        name
        column_values {{
          id
          text
        }}
      }}
    }}
    """
    data = monday_query(query)
    item = data["data"]["items"][0]
    cols = {c["id"]: c.get("text", "") for c in item["column_values"]}
    cols["name"] = item["name"]
    return cols

def create_report(data):
    template_path = "/Templates/Report_Template.docx"
    _, res = dbx.files_download(template_path)
    doc = Document(io.BytesIO(res.content))

    replacements = {}
    for placeholder, col_id in COLUMN_MAP.items():
        replacements[placeholder] = data.get(col_id, "")

    replacements["{{date_today}}"] = datetime.now().strftime("%d/%m/%Y")

    for p in doc.paragraphs:
        for old, new in replacements.items():
            if old in p.text:
                p.text = p.text.replace(old, new)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def process_item(item_id):
    data = get_item_data(item_id)

    project_name = data.get("name", f"project_{item_id}")
    folder_path = f"/Projects/{project_name}"
    file_path = f"{folder_path}/Report.docx"

    try:
        dbx.files_create_folder_v2(folder_path)
    except Exception:
        pass

    report = create_report(data)
    dbx.files_upload(report.read(), file_path, mode=dropbox.files.WriteMode.overwrite)

    link = dbx.sharing_create_shared_link_with_settings(file_path).url
    return link

@app.route("/", methods=["GET"])
def home():
    return "OK", 200

@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    if request.method == "GET":
        return "Webhook endpoint is live", 200

    data = request.get_json(silent=True) or {}

    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})

    try:
        item_id = (
            data.get("event", {}).get("pulseId")
            or data.get("event", {}).get("itemId")
            or data.get("pulseId")
            or data.get("itemId")
        )

        if not item_id:
            return jsonify({"error": "No item id found", "payload": data}), 400

        link = process_item(int(item_id))
        return jsonify({"status": "success", "link": link}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
