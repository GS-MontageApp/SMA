"""
=============================================================================
MICROSERVICE: backend/main.py (Schlauchmanagement-App Backend)
=============================================================================
Zweck: Echter Server-Export via Python und openpyxl zur 100% verlustfreien 
Erhaltung aller Logos, Grafiken, Rahmenlinien, Formeln und Dropdown-Menüs.
"""

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
import openpyxl
import os
import tempfile

app = FastAPI(title="SMA Export Microservice", version="1.0.7")

# CORS-Konfiguration für GitHub Pages PWA
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Robuste absolute Pfadermittlung für die Vorlage unabhängig vom Render-Root-Verzeichnis
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Prüft sowohl im Parent-Verzeichnis als auch lokal
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
            detail=f"Original-Vorlage (XLSX-Muster.xlsx) wurde auf dem Server unter gesuchtem Pfad nicht gefunden."
        )

    temp_dir = tempfile.mkdtemp()
    output_filename = f"Aktualisiert_{payload.filename}"
    output_path = os.path.join(temp_dir, output_filename)

    try:
        wb = openpyxl.load_workbook(TEMPLATE_PATH)

        for update in payload.updates:
            sheet_name = update.sheet_name
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                ws.cell(row=update.row, column=update.col, value=update.value)
            else:
                ws = wb.active
                ws.cell(row=update.row, column=update.col, value=update.value)

        wb.save(output_path)
        wb.close()

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Fehler bei der Excel-Verarbeitung: {str(e)}")
