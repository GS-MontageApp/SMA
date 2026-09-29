"""
=============================================================================
MICROSERVICE: backend/main.py (Schlauchmanagement-App Backend)
=============================================================================
Zweck: Echter Server-Export via Python und openpyxl zur 100% verlustfreien 
Erhaltung aller Logos, Grafiken, Rahmenlinien, Formeln und Dropdown-Menüs.
Inklusive Live-Einbindung des Logos von GitHub und Schutz für A1:K1.
"""

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image
import os
import tempfile
import re
import requests
import io
import time
import uuid

app = FastAPI(title="SMA Export Microservice", version="1.0.21")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_TEMPLATE_PATH = os.path.join(BASE_DIR, "..", "templates", "XLSX-Muster.xlsx")
if not os.path.exists(LOCAL_TEMPLATE_PATH):
    LOCAL_TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "XLSX-Muster.xlsx")

GITHUB_RAW_TEMPLATE_URL = "https://raw.githubusercontent.com/GS-MontageApp/SMA/main/templates/XLSX-Muster.xlsx"
GITHUB_RAW_LOGO_URL = "https://raw.githubusercontent.com/GS-MontageApp/SMA/main/templates/logo.png"

class CellUpdate(BaseModel):
    sheet_name: str
    row: int  # 1-based index
    col: int  # 1-based index
    value: str | int | float | None

class ExportRequest(BaseModel):
    filename: str
    updates: list[CellUpdate]

@app.get("/")
def read_root():
    github_reachable = False
    github_logo_reachable = False
    try:
        test_url = f"{GITHUB_RAW_TEMPLATE_URL}?cb={uuid.uuid4()}"
        r = requests.head(test_url, headers={"Cache-Control": "no-cache"}, timeout=3)
        github_reachable = (r.status_code == 200)

        test_logo_url = f"{GITHUB_RAW_LOGO_URL}?cb={uuid.uuid4()}"
        r_logo = requests.head(test_logo_url, headers={"Cache-Control": "no-cache"}, timeout=3)
        github_logo_reachable = (r_logo.status_code == 200)
    except:
        pass

    return {
        "status": "online", 
        "service": "SMA Excel Export Microservice",
        "version": "1.0.21",
        "github_template_reachable": github_reachable,
        "github_logo_reachable": github_logo_reachable,
        "local_fallback_found": os.path.exists(LOCAL_TEMPLATE_PATH)
    }

def get_live_workbook():
    try:
        cache_buster = f"?cb={uuid.uuid4()}&t={int(time.time())}"
        target_url = GITHUB_RAW_TEMPLATE_URL + cache_buster
        
        headers = {
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0"
        }

        response = requests.get(target_url, headers=headers, timeout=10)
        if response.status_code == 200:
            return openpyxl.load_workbook(io.BytesIO(response.content))
    except Exception as e:
        print(f"Warnung: Live-Download von GitHub fehlgeschlagen: {e}. Nutze lokale Fallback-Vorlage.")

    if os.path.exists(LOCAL_TEMPLATE_PATH):
        return openpyxl.load_workbook(LOCAL_TEMPLATE_PATH)
    
    raise HTTPException(
        status_code=404, 
        detail="Weder von GitHub noch lokal konnte eine gültige Mustervorlage (XLSX-Muster.xlsx) geladen werden."
    )

@app.post("/api/export")
def export_excel(payload: ExportRequest):
    temp_dir = tempfile.mkdtemp()
    output_filename = f"Aktualisiert_{payload.filename}"
    output_path = os.path.join(temp_dir, output_filename)

    try:
        wb = get_live_workbook()

        # Logo von GitHub herunterladen und in das erste Tabellenblatt einfügen
        try:
            logo_cache_buster = f"?cb={uuid.uuid4()}&t={int(time.time())}"
            logo_response = requests.get(GITHUB_RAW_LOGO_URL + logo_cache_buster, timeout=5)
            if logo_response.status_code == 200:
                img_io = io.BytesIO(logo_response.content)
                img = Image(img_io)
                primary_ws = wb.active
                primary_ws.add_image(img, "A1")
        except Exception as logo_err:
            print(f"Hinweis: Logo konnte nicht von GitHub geladen werden (unkritisch): {logo_err}")

        max_row_written = 5
        for update in payload.updates:
            sheet_name = update.sheet_name
            
            # SCHUTZREGEL: Zeile 1, Spalten A bis K (col 1 bis 11) absolut vor Überschreiben schützen (Logo & Header)
            if update.row == 1 and 1 <= update.col <= 11:
                continue

            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                target_row = update.row
                target_col = update.col
                
                # Robustes Merged-Cell-Handling: Auf die obere linke Master-Zelle umleiten
                for cr in ws.merged_cells.ranges:
                    if target_row >= cr.min_row and target_row <= cr.max_row and target_col >= cr.min_col and target_col <= cr.max_col:
                        target_row = cr.min_row
                        target_col = cr.min_col
                        break

                target_cell = ws.cell(row=target_row, column=target_col)
                
                # Falls es trotz Umleitung ein schreibgeschütztes MergedCell-Objekt bleibt, sicher überspringen
                if type(target_cell).__name__ == 'MergedCell' or isinstance(target_cell, MergedCell):
                    continue

                target_cell.value = update.value
                if target_row > max_row_written:
                    max_row_written = target_row
            else:
                ws = wb.active
                ws.cell(row=update.row, column=update.col, value=update.value)
                if update.row > max_row_written:
                    max_row_written = update.row

        try:
            for sheetname in wb.sheetnames:
                ws = wb[sheetname]
                if hasattr(ws, 'data_validations') and ws.data_validations.dataValidation:
                    for dv in ws.data_validations.dataValidation:
                        current_sqref = str(dv.sqref)
                        first_range = current_sqref.split()[0] if current_sqref else ""
                        if ":" in first_range:
                            start_col_row = first_range.split(":")[0]
                            col_match = re.match(r"([A-Z]+)", start_col_row)
                            if col_match:
                                col_letters = col_match.group(1)
                                new_end_row = max(max_row_written, 150)
                                dv.sqref = f"{col_letters}5:{col_letters}{new_end_row}"
        except Exception as val_err:
            print(f"Hinweis bei Datenvalidierung (unkritisch): {str(val_err)}")

        wb.save(output_path)
        wb.close()

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Fehler bei der Excel-Verarbeitung: {str(e)}")
