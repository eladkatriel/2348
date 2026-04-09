# ===== FINAL STABLE VERSION - FIXED DROPBOX ROOT PATH =====

import os
import io
import re
import json
from datetime import datetime

import requests
import dropbox
from flask import Flask, request, jsonify
from docx import Document

app = Flask(__name__)

# =========================
# CONFIG
# =========================
MONDAY_API_KEY = os.environ.get("MONDAY_API_KEY")
BOARD_ID = os.environ.get("BOARD_ID")

DROPBOX_APP_KEY = os.environ.get("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET = os.environ.get("DROPBOX_APP_SECRET")
DROPBOX_REFRESH_TOKEN = os.environ.get("DROPBOX_REFRESH_TOKEN")

LINK_COLUMN_ID = os.environ.get("LINK_COLUMN_ID", "link_mm27m1ce")

BASE_REPORTS_PATH = "/20260228 - שאגת הארי"
TEMPLATE_DIR = "/Template/23-48"

# =========================
# DROPBOX INIT
# =========================
dbx = dropbox.Dropbox(
    oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
    app_key=DROPBOX_APP_KEY,
    app_secret=DROPBOX_APP_SECRET,
)

def path_exists(path):
    try:
        dbx.files_get_metadata(path)
        return True
    except:
        return False

if not path_exists(BASE_REPORTS_PATH):
    raise Exception(f"BASE_REPORTS_PATH NOT FOUND: {BASE_REPORTS_PATH}")

if not path_exists(TEMPLATE_DIR):
    raise Exception(f"TEMPLATE_DIR NOT FOUND: {TEMPLATE_DIR}")

print("Dropbox paths validated successfully")

# =========================
# MONDAY
# =========================
def monday_query(query, variables=None):
    response = requests.post(
        "https://api.monday.com/v2",
        json={"query": query, "variables": variables or {}},
        headers={"Authorization": MONDAY_API_KEY},
    )
    data = response.json()
    if "errors" in data:
        raise Exception(data["errors"])
    return data

def get_item_data(item_id):
    query = """
    query ($item_ids: [ID!]) {
      items(ids: $item_ids) {
        name
        column_values {
          id
          text
        }
      }
    }
    """
    data = monday_query(query, {"item_ids": [item_id]})
    item = data["data"]["items"][0]
    cols = {c["id"]: c["text"] for c in item["column_values"]}
    cols["name"] = item["name"]
    return cols

# =========================
# DATE
# =========================
def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d.%m.%Y")
    except:
        return date_str or ""

def extract_date(data):
    timeline = data.get("dup__of_90__timeline", "")
    match = re.search(r"\d{4}-\d{2}-\d{2}", timeline or "")
    if match:
        return format_date(match.group(0))
    return format_date(data.get("date_mm1rnmvg"))

# =========================
# TEMPLATE
# =========================
def select_template(value):
    mapping = {
        "קבלן": "Contractor_template.docx",
        "מהנדס": "Engineer_template.docx",
        "להריסה": "Engineer_template.docx",
        "נזק ישן, אין פינוי ואין פיצוי": "Contractor_template.docx",
    }
    return mapping.get(value, "Contractor_template.docx")

# =========================
# WORD REPLACEMENT
# =========================
def replace_doc(doc, rep):
    for p in doc.paragraphs:
        for k, v in rep.items():
            if k in p.text:
                p.text = p.text.replace(k, v)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for k, v in rep.items():
                    if k in c.text:
                        c.text = c.text.replace(k, v)

# =========================
# MAIN
# =========================
def process_item(item_id):
    data = get_item_data(item_id)

    city = data.get("text_mm264acy", "")
    name = data.get("name", "")

    base = f"{BASE_REPORTS_PATH}/{city}/{name}"
    photos = f"{base}/תמונות"
    findings = f"{base}/ממצאים ראשוניים + כ.כמויות"

    for p in [base, photos, findings]:
        try:
            dbx.files_create_folder_v2(p)
        except:
            pass

    date_val = extract_date(data)

    filename = f"מימצאים מבניים והנחיות ראשוניות רחוב {data.get('text_mm12vcy9')} {data.get('text_mm12jf0w')} - דירה {data.get('text_mm127a33')}- {city}, {date_val}.docx"
    path = f"{findings}/{filename}"

    template_name = select_template(data.get("color_mm12nfzt"))
    template_path = f"{TEMPLATE_DIR}/{template_name}"

    _, res = dbx.files_download(template_path)
    doc = Document(io.BytesIO(res.content))

    rep = {
        "{{Engineer}}": data.get("text_mm12tcvb", ""),
        "{{Street}}": data.get("text_mm12vcy9", ""),
        "{{Number}}": data.get("text_mm12jf0w", ""),
        "{{Apartment}}": data.get("text_mm127a33", ""),
        "{{City}}": city,
        "{{Case}}": data.get("text_mm12qp1q", ""),
        "{{ID}}": data.get("text_mm12vayb", ""),
        "{{Date}}": date_val,
        "{{date_today}}": datetime.now().strftime("%d.%m.%Y"),
    }

    replace_doc(doc, rep)

    buffer = io.BytesIO()
    doc.save(buffer)

    dbx.files_upload(buffer.getvalue(), path, mode=dropbox.files.WriteMode.overwrite)

    link = dbx.sharing_create_shared_link_with_settings(path).url

    monday_query("""
    mutation ($board_id: ID!, $item_id: ID!, $column_values: JSON!) {
      change_multiple_column_values(
        board_id: $board_id,
        item_id: $item_id,
        column_values: $column_values
      ) { id }
    }
    """, {
        "board_id": BOARD_ID,
        "item_id": item_id,
        "column_values": json.dumps({
            LINK_COLUMN_ID: {"url": link, "text": "דו\"ח"}
        })
    })

    return link

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})

    item_id = data["event"]["pulseId"]
    link = process_item(item_id)

    return jsonify({"status": "ok", "link": link})

@app.route("/")
def home():
    return "OK"

if __name__ == "__main__":
    app.run(port=10000)
