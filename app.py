import os
import io
from datetime import datetime

import requests
import dropbox
from flask import Flask, request, jsonify
from docx import Document

app = Flask(__name__)

# ===== CONFIG =====
MONDAY_API_KEY = os.environ.get("MONDAY_API_KEY")
BOARD_ID = os.environ.get("BOARD_ID")
DROPBOX_TOKEN = os.environ.get("DROPBOX_TOKEN")

if not MONDAY_API_KEY:
    raise ValueError("Missing MONDAY_API_KEY environment variable")

if not BOARD_ID:
    raise ValueError("Missing BOARD_ID environment variable")

if not DROPBOX_TOKEN:
    raise ValueError("Missing DROPBOX_TOKEN environment variable")

dbx = dropbox.Dropbox(DROPBOX_TOKEN)

# ===== PATHS =====
BASE_REPORTS_PATH = "/YOE/חרבות ברזל 2023/20260228 - שאגת הארי"
TEMPLATE_PATH = "/Template/23-48/Contractor_template.docx"

# ===== COLUMN IDS =====
CITY_COLUMN_ID = "text_mm264acy"

# ===== PLACEHOLDER -> MONDAY COLUMN MAP =====
COLUMN_MAP = {
    "{{Engineer}}": "text_mm12tcvb",
    "{{Street}}": "text_mm12vcy9",
    "{{Number}}": "text_mm12jf0w",
    "{{Apartment}}": "text_mm127a33",
    "{{Hit date}}": "date_mm1rnmvg",
    "{{Date}}": "dup__of_90__timeline",
    "{{months}}": "text_mm129ef8",
}

# ===== MONDAY API =====
def monday_query(query: str):
    response = requests.post(
        "https://api.monday.com/v2",
        json={"query": query},
        headers={
            "Authorization": MONDAY_API_KEY,
            "Content-Type": "application/json",
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def get_item_data(item_id: int):
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
    items = data.get("data", {}).get("items", [])

    if not items:
        raise Exception(f"No item found for item_id={item_id}")

    item = items[0]
    cols = {c["id"]: c.get("text", "") for c in item["column_values"]}
    cols["name"] = item["name"]

    return cols


# ===== HELPERS =====
def sanitize_filename(name: str) -> str:
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid_chars:
        name = name.replace(ch, "-")
    return " ".join(name.split()).strip()


def create_folder_if_needed(folder_path: str):
    try:
        dbx.files_create_folder_v2(folder_path)
        print("FOLDER CREATED:", folder_path)
    except Exception as e:
        print("FOLDER CREATE SKIPPED OR FAILED:", folder_path, str(e))


def get_or_create_shared_link(file_path: str) -> str:
    try:
        existing_links = dbx.sharing_list_shared_links(path=file_path, direct_only=True).links
        if existing_links:
            print("EXISTING SHARE LINK FOUND")
            return existing_links[0].url
    except Exception as e:
        print("SHARED LINK LOOKUP FAILED:", str(e))

    link = dbx.sharing_create_shared_link_with_settings(file_path).url
    print("NEW SHARE LINK CREATED:", link)
    return link


# ===== WORD REPLACEMENT =====
def replace_in_paragraph(paragraph, replacements: dict):
    if not paragraph.text:
        return

    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return

    new_text = full_text
    changed = False

    for old, new in replacements.items():
        if old in new_text:
            new_text = new_text.replace(old, str(new))
            changed = True

    if not changed:
        return

    # שומרים את כל הטקסט ב-run הראשון ומרוקנים את היתר
    paragraph.runs[0].text = new_text
    for run in paragraph.runs[1:]:
        run.text = ""


def replace_in_table(table, replacements: dict):
    for row in table.rows:
        for cell in row.cells:
            replace_in_document_parts(cell, replacements)


def replace_in_document_parts(container, replacements: dict):
    for paragraph in container.paragraphs:
        replace_in_paragraph(paragraph, replacements)

    for table in container.tables:
        replace_in_table(table, replacements)


def replace_everywhere(doc: Document, replacements: dict):
    # גוף המסמך
    replace_in_document_parts(doc, replacements)

    # Header / Footer
    for section in doc.sections:
        replace_in_document_parts(section.header, replacements)
        replace_in_document_parts(section.footer, replacements)


# ===== DOCUMENT CREATION =====
def create_report(data: dict) -> io.BytesIO:
    print("DOWNLOADING TEMPLATE FROM:", TEMPLATE_PATH)

    _, res = dbx.files_download(TEMPLATE_PATH)
    doc = Document(io.BytesIO(res.content))

    replacements = {}
    for placeholder, col_id in COLUMN_MAP.items():
        replacements[placeholder] = data.get(col_id, "") or ""

    replacements["{{date_today}}"] = datetime.now().strftime("%d/%m/%Y")
    replacements["{{City}}"] = data.get(CITY_COLUMN_ID, "") or ""
    replacements["{{ProjectName}}"] = data.get("name", "") or ""

    print("REPLACEMENTS:", replacements)

    replace_everywhere(doc, replacements)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    print("REPORT CREATED IN MEMORY")
    return buffer


# ===== MAIN PROCESS =====
def process_item(item_id: int):
    print("START process_item with item_id:", item_id)

    data = get_item_data(item_id)
    print("ITEM DATA:", data)

    project_name = (data.get("name", "") or "").strip()
    if not project_name:
        project_name = f"project_{item_id}"

    city_name = (data.get(CITY_COLUMN_ID, "") or "").strip()
    if not city_name:
        raise Exception(f"City field is empty or missing (column id: {CITY_COLUMN_ID})")

    report_folder = f"{BASE_REPORTS_PATH}/{city_name}/{project_name}"
    photos_folder = f"{report_folder}/תמונות"
    findings_folder = f"{report_folder}/ממצאים ראשוניים + כ.כמויות"

    street = (data.get("text_mm12vcy9", "") or "").strip()
    number = (data.get("text_mm12jf0w", "") or "").strip()
    apartment = (data.get("text_mm127a33", "") or "").strip()
    report_date = (data.get("dup__of_90__timeline", "") or "").strip()

    file_name = sanitize_filename(
        f"מימצאים מבניים והנחיות ראשוניות רחוב {street} {number} - "
        f"דירה {apartment}- {city_name}, {report_date}.docx"
    )

    file_path = f"{findings_folder}/{file_name}"

    print("REPORT FOLDER:", report_folder)
    print("PHOTOS FOLDER:", photos_folder)
    print("FINDINGS FOLDER:", findings_folder)
    print("FILE PATH:", file_path)

    create_folder_if_needed(f"{BASE_REPORTS_PATH}/{city_name}")
    create_folder_if_needed(report_folder)
    create_folder_if_needed(photos_folder)
    create_folder_if_needed(findings_folder)

    report = create_report(data)

    dbx.files_upload(
        report.read(),
        file_path,
        mode=dropbox.files.WriteMode.overwrite
    )
    print("FILE UPLOADED TO DROPBOX")

    link = get_or_create_shared_link(file_path)
    print("PROCESS SUCCESS. LINK:", link)

    return link


# ===== ROUTES =====
@app.route("/", methods=["GET"])
def home():
    return "OK", 200


@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    if request.method == "GET":
        return "Webhook endpoint is live", 200

    data = request.get_json(silent=True) or {}
    print("INCOMING DATA:", data)

    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]}), 200

    try:
        item_id = (
            data.get("event", {}).get("pulseId")
            or data.get("event", {}).get("itemId")
            or data.get("pulseId")
            or data.get("itemId")
        )

        print("ITEM ID:", item_id)

        if not item_id:
            return jsonify({"error": "No item id found", "payload": data}), 400

        link = process_item(int(item_id))
        return jsonify({"status": "success", "link": link}), 200

    except Exception as e:
        print("PROCESS ERROR:", str(e))
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
