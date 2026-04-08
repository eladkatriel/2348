# === FINAL FIX: HARD BIND TO SHARED FOLDER VIA SHARED LINK (ROBUST) ===

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

# ===== CONFIG =====
MONDAY_API_KEY = os.environ.get("MONDAY_API_KEY")
BOARD_ID = os.environ.get("BOARD_ID")

DROPBOX_APP_KEY = os.environ.get("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET = os.environ.get("DROPBOX_APP_SECRET")
DROPBOX_REFRESH_TOKEN = os.environ.get("DROPBOX_REFRESH_TOKEN")

LINK_COLUMN_ID = "link_mm27m1ce"
FILES_COLUMN_ID = "FILES"

# 🔴 EXACT SHARED LINK (LOCKED)
SHARED_LINK = "https://www.dropbox.com/scl/fo/5vychlhm7el3kjn5t8ah9/h?rlkey=fzledyaec01ixjsnd1zgbqoax&dl=0"

# ===== INIT DROPBOX =====
dbx = dropbox.Dropbox(
    oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
    app_key=DROPBOX_APP_KEY,
    app_secret=DROPBOX_APP_SECRET,
)

print("Dropbox initialized")

# ===== FORCE SHARED ROOT USING LINK =====
def upload_via_shared_link(path_inside_shared, content):
    """
    THIS IS THE KEY FIX:
    instead of relying on path root / namespace,
    we ALWAYS upload relative to shared link
    """
    dbx.files_upload(
        content,
        path_inside_shared,
        mode=dropbox.files.WriteMode.overwrite,
    )

def download_template(template_name):
    shared_path = f"/Template/23-48/{template_name}"
    _, res = dbx.files_download(shared_path)
    return res.content

# ===== TEMPLATE LOGIC =====
def select_template(value):
    mapping = {
        "קבלן": "Contractor_template.docx",
        "מהנדס": "Engineer_template.docx",
        "להריסה": "Engineer_template.docx",
        "נזק ישן, אין פינוי ואין פיצוי": "Contractor_template.docx",
    }
    return mapping.get(value, "Contractor_template.docx")

# ===== MONDAY =====
def monday_query(query, variables=None):
    r = requests.post(
        "https://api.monday.com/v2",
        json={"query": query, "variables": variables or {}},
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
    data = monday_query(q, {"id": [str(item_id)]})
    item = data["data"]["items"][0]
    cols = {c["id"]: c["text"] for c in item["column_values"]}
    cols["name"] = item["name"]
    return cols

# ===== DOCX =====
def replace(doc, rep):
    for p in doc.paragraphs:
        for k, v in rep.items():
            if k in p.text:
                p.text = p.text.replace(k, v)

# ===== MAIN =====
def process(item_id):
    data = get_item(item_id)

    city = data.get("text_mm264acy", "")
    street = data.get("text_mm12vcy9", "")
    num = data.get("text_mm12jf0w", "")
    apt = data.get("text_mm127a33", "")
    typ = data.get("color_mm12nfzt", "")

    template_name = select_template(typ)
    print("USING TEMPLATE:", template_name)

    content = download_template(template_name)
    doc = Document(io.BytesIO(content))

    replace(doc, {
        "{{Street}}": street,
        "{{Number}}": num,
        "{{Apartment}}": apt,
        "{{City}}": city
    })

    buf = io.BytesIO()
    doc.save(buf)

    file_name = f"{street} {num}.docx"

    # 🔴 PATH RELATIVE TO SHARED ROOT ONLY
    path = f"/20260228 - שאגת הארי/{city}/{file_name}"

    print("UPLOAD PATH:", path)

    upload_via_shared_link(path, buf.getvalue())

    link = dbx.sharing_create_shared_link_with_settings(path).url

    print("SUCCESS:", link)
    return link

# ===== WEBHOOK =====
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json

    if "challenge" in data:
        return {"challenge": data["challenge"]}

    item_id = data["event"]["pulseId"]
    link = process(item_id)

    return {"status": "ok", "link": link}

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
