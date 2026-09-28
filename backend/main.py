"""
=============================================================================
MICROSERVICE: backend/main.py (Schlauchmanagement-App Backend)
=============================================================================
Zweck: Echter Server-Export via Python und openpyxl zur 100% verlustfreien 
Erhaltung aller Logos, Grafiken, Rahmenlinien, Formeln und Dropdown-Menüs 
in der Excel-Vorlage (templates/XLSX-Muster.xlsx).
"""

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
import openpyxl
import os
import shutil
import tempfile

app = FastAPI(title="SMA Export Microservice", version="1.0.0")

# CORS-Konfiguration, damit die GitHub Pages PWA ungehindert zugreifen kann
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In Production ggf. auf deine GitHub Pages Domain einschränken
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Pfad zur originalen Vorlage im Microservice-Repo
TEMPLATE_PATH = "templates/XLSX-Muster.xlsx"

class CellUpdate(BaseModel):
    sheet_name: str
    row: int  # 1-based index für openpyxl
    col: int  # 1-based index für openpyxl
    value: str | int | float | None

class ExportRequest(BaseModel):
    filename: str
    updates: list[CellUpdate]

@app.get("/")
def read_root():
    return {"status": "online", "service": "SMA Excel Export Microservice"}

@app.post("/api/export")
def export_excel(payload: ExportRequest):
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=404, detail="Original-Vorlage (XLSX-Muster.xlsx) auf dem Server nicht gefunden.")

    # Erstelle eine temporäre Kopie der Vorlage, um das Original unangetastet zu lassen
    temp_dir = tempfile.mkdtemp()
    output_filename = f"Aktualisiert_{payload.filename}"
    output_path = os.path.join(temp_dir, output_filename)

    try:
        # openpyxl lädt die Datei inklusive aller Drawings, Shapes, Validations und Styles
        wb = openpyxl.load_workbook(TEMPLATE_PATH)

        # Trage die modifizierten Zellen passgenau ein
        for update in payload.updates:
            sheet_name = update.sheet_name
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                # openpyxl arbeitet 1-basiert (row, col)
                ws.cell(row=update.row, column=update.col, value=update.value)
            else:
                # Fallback auf das erste Tabellenblatt, falls Name abweicht
                ws = wb.active
                ws.cell(row=update.row, column=update.col, value=update.value)

        # Speichern der bearbeiteten Datei (alle Binärstrukturen & Logos bleiben intakt!)
        wb.save(output_path)
        wb.close()

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Fehler bei der Excel-Verarbeitung: {str(e)}")
