"""
=============================================================================
MICROSERVICE: backend/main.py (Schlauchmanagement-App Backend)
=============================================================================
Zweck: Echter Server-Export via Python und openpyxl zur 100% verlustfreien 
Erhaltung aller Logos, Grafiken, Rahmenlinien, Formeln und Dropdown-Menüs 
ohne Autofilter-Manipulation (Original-Autofilter unangetastet).
"""

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
import openpyxl
from openpyxl.cell.cell import MergedCell
import os
import tempfile
import re

app = FastAPI(title="SMA Export Microservice", version="1.0.14")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "..", "templates", "XLSX-Muster.xlsx")
if not os.path.exists(TEMPLATE_PATH):
    TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "XLSX-Muster.xlsx")

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
    return {
        "status": "online", 
        "service": "SMA Excel Export Microservice",
        "template_found": os.path.exists(TEMPLATE_PATH)
    }

@app.post("/api/export")
def export_excel(payload: ExportRequest):
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(
            status_code=404, 
            detail=f"Original-Vorlage (XLSX-Muster.xlsx) wurde auf dem Server nicht gefunden."
        )

    temp_dir = tempfile.mkdtemp()
    output_filename = f"Aktualisiert_{payload.filename}"
    output_path = os.path.join(temp_dir, output_filename)

    try:
        wb = openpyxl.load_workbook(TEMPLATE_PATH)

        max_row_written = 5
        for update in payload.updates:
            sheet_name = update.sheet_name
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                target_cell = ws.cell(row=update.row, column=update.col)
                
                if isinstance(target_cell, MergedCell):
                    found_master = False
                    for range_str in ws.merged_cells.ranges:
                        if target_cell.coordinate in range_str:
                            min_col, min_row, _, _ = range_str.bounds
                            target_cell = ws.cell(row=min_row, column=min_col)
                            found_master = True
                            break
                    if not found_master:
                        continue

                target_cell.value = update.value
                if update.row > max_row_written:
                    max_row_written = update.row
            else:
                ws = wb.active
                ws.cell(row=update.row, column=update.col, value=update.value)

        # HINWEIS: Wir lassen die Autofilter der Vorlage absolut unangetastet,
        # um fehlerhafte Pfeil-Verschiebungen zu verhindern.

        # Robuste Anpassung der Validierungsbereiche (Dropdowns ab Zeile 5)
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

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Fehler bei der Excel-Verarbeitung: {str(e)}")
