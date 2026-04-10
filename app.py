# FINAL CLEAN VERSION - NO HARD PATHS / NO VALIDATION CRASH
import os
import io
from urllib.parse import quote_plus

import dropbox
from flask import Flask
from docx import Document
from docx.shared import Inches

from playwright.sync_api import sync_playwright

app = Flask(__name__)

MAP_WIDTH_INCHES = 6.5
MAP_HEIGHT_INCHES = 4.4

dbx = dropbox.Dropbox(os.environ.get("DROPBOX_TOKEN"))
print("Dropbox initialized - FINAL CLEAN")

def create_govmap_image(address_text: str) -> bytes:
    url = f"https://www.govmap.gov.il/?b=2&q={quote_plus(address_text)}&z=10"
    print("GOVMAP URL:", url)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 1000})
        page.goto(url)
        page.wait_for_timeout(5000)
        img = page.screenshot()
        browser.close()
        return img

def insert_map(doc, img_bytes):
    for p in doc.paragraphs:
        if "{{Map}}" in p.text:
            p.text = ""
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes),
                            width=Inches(MAP_WIDTH_INCHES),
                            height=Inches(MAP_HEIGHT_INCHES))

@app.route("/")
def home():
    return "OK FINAL CLEAN", 200
