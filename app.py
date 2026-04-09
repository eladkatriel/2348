# ===== FINAL VERSION - LOCKED TO SHARED FOLDER =====

import os
import io
import re
import json
from datetime import datetime
from pathlib import PurePosixPath

import requests
import dropbox
from flask import Flask, request, jsonify
from docx import Document

app = Flask(__name__)

MONDAY_API_KEY = os.environ.get("MONDAY_API_KEY")
BOARD_ID = os.environ.get("BOARD_ID")

DROPBOX_APP_KEY = os.environ.get("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET = os.environ.get("DROPBOX_APP_SECRET")
DROPBOX_REFRESH_TOKEN = os.environ.get("DROPBOX_REFRESH_TOKEN")

LINK_COLUMN_ID = os.environ.get("LINK_COLUMN_ID", "link_mm27m1ce")

DROPBOX_SHARED_LINK = "https://www.dropbox.com/scl/fo/5vychlhm7el3kjn5t8ah9/h?rlkey=fzledyaec01ixjsnd1zgbqoax"

TARGET_FOLDER_NAME = "20260228 - שאגת הארי"

dbx = dropbox.Dropbox(
    oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
    app_key=DROPBOX_APP_KEY,
    app_secret=DROPBOX_APP_SECRET,
)

# =========================
# RESOLVE REAL PATH FROM LINK
# =========================
def resolve_shared_root():
    meta = dbx.sharing_get_shared_link_metadata(DROPBOX_SHARED_LINK)
    shared_id = meta.shared_folder_id

    folders = dbx.sharing_list_folders().entries
    for f in folders:
        if f.shared_folder_id == shared_id:
            print("LOCKED ROOT:", f.path_display)
            return f.path_display

    raise Exception("Shared folder not found")

SHARED_ROOT = resolve_shared_root()

# =========================
# PATHS
# =========================
BASE_REPORTS_PATH = f"{SHARED_ROOT}/{TARGET_FOLDER_NAME}"
TEMPLATE_DIR = "/Template/23-48"

print("BASE_REPORTS_PATH:", BASE_REPORTS_PATH)

# =========================
# HELPERS
# =========================
def create_folder(path):
    try:
        dbx.files_create_folder_v2(path)
    except:
        pass

def format_date(d):
    try:
        return datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y")
    except:
        return ""

def extract_date(data):
    t = data.get("dup__of_90__timeline", "")
    m = re.search(r"\d{4}-\d{2}-\d{2}", t or "")
    if m:
        return format_date(m.group(0))
    return format_date(data.get("date_mm1rnmvg"))

# =========================
# MONDAY
# =========================
def monday(query, vars=None):
    r = requests.post(
        "https://api.monday.com/v2",
        json={"query": query, "variables": vars or {}},
        headers={"Authorization": MONDAY_API_KEY},
    )
    return r.json()

def get_item(item_id):
    q = """
    query ($id: [ID!]) {
      items(ids: $id) {
        name
        column_values { id text }
      }
    }
    """
    d = monday(q, {"id": [item_id]})
    item = d["data"]["items"][0]
    cols = {c["id"]: c["text"] for c in item["column_values"]}
    cols["name"] = item["name"]
    return cols

# =========================
# MAIN
# =========================
def process(item_id):
    data = get_item(item_id)

    city = data.get("text_mm264acy", "")
    name = data.get("name", "")

    base = f"{BASE_REPORTS_PATH}/{city}/{name}"
    photos = f"{base}/תמונות"
    findings = f"{base}/ממצאים ראשוניים + כ.כמויות"

    for p in [base, photos, findings]:
        create_folder(p)

    date_val = extract_date(data)

    filename = f"דוח {city} {date_val}.docx"
    path = f"{findings}/{filename}"

    _, res = dbx.files_download(f"{TEMPLATE_DIR}/Engineer_template.docx")
    doc = Document(io.BytesIO(res.content))

    for p in doc.paragraphs:
        p.text = p.text.replace("{{City}}", city)

    buffer = io.BytesIO()
    doc.save(buffer)

    dbx.files_upload(buffer.getvalue(), path, mode=dropbox.files.WriteMode.overwrite)

    link = dbx.sharing_create_shared_link_with_settings(path).url

    monday("""
    mutation ($b: ID!, $i: ID!, $v: JSON!) {
      change_multiple_column_values(board_id: $b, item_id: $i, column_values: $v){id}
    }
    """, {
        "b": BOARD_ID,
        "i": item_id,
        "v": json.dumps({LINK_COLUMN_ID: {"url": link, "text": "דו\"ח"}})
    })

    return link

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})

    item_id = data["event"]["pulseId"]
    link = process(item_id)

    return jsonify({"ok": True, "link": link})

@app.route("/")
def home():
    return "OK"

if __name__ == "__main__":
    app.run(port=10000)
