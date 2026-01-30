from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from typing import List, Optional
import os

app = FastAPI()

# ===============================
# RUTAS BASE
# ===============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")

TEMPLATES_WORD = {
    "mantenimiento": os.path.join(BASE_DIR, "templates_word", "mantenimiento.docx"),
    "sustitucion": os.path.join(BASE_DIR, "templates_word", "sustitucion_equipo.docx"),
}

os.makedirs(UPLOAD_DIR, exist_ok=True)

# ===============================
# STATIC Y HTML
# ===============================
app.mount(
    "/static",
    StaticFiles(directory=os.path.join(BASE_DIR, "static")),
    name="static"
)
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

# ===============================
# HOME
# ===============================
@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# ===============================
# GENERAR DOCUMENTO
# ===============================
@app.post("/generar")
async def generar(
    # -------- tipo de documento --------
    tipo_documento: str = Form(...),

    # -------- datos generales --------
    municipio: str = Form(""),
    afiliacion: str = Form(""),
    fecha_solicitud: str = Form(""),
    fecha_atencion: str = Form(""),
    fecha_cierre: str = Form(""),
    tecnologia: str = Form(""),
    ticket: str = Form(""),

    # -------- mantenimiento --------
    ajuste: str = Form(None),
    reparacion: str = Form(None),
    reubicacion: str = Form(None),
    cambio: str = Form(None),
    siniestro: str = Form(None),
    otros: str = Form(None),

    falla: str = Form(""),
    solucion: str = Form(""),
    observaciones: str = Form(""),

    # -------- sustitución --------
    er_descripcion: str = Form(""),
    er_marca: str = Form(""),
    er_modelo: str = Form(""),
    er_serie: str = Form(""),

    ei_descripcion: str = Form(""),
    ei_marca: str = Form(""),
    ei_modelo: str = Form(""),
    ei_serie: str = Form(""),

    justificacion: str = Form(""),
    tipo_sustitucion: str = Form(""),

    # -------- firmas --------
    nombre_coordinador: str = Form(""),
    nombre_responsable: str = Form(""),

    # -------- fotos (REALMENTE OPCIONALES) --------
    fotos: Optional[List[UploadFile]] = File(None)
):
    # ===============================
    # CARGAR PLANTILLA
    # ===============================
    template_path = TEMPLATES_WORD.get(tipo_documento)

    if not template_path:
        return HTMLResponse("Tipo de documento no válido", status_code=400)

    doc = DocxTemplate(template_path)

    # ===============================
    # PROCESAR FOTOS SOLO SI EXISTEN
    # ===============================
    fotos_word = ["", "", "", ""]

    if fotos is not None:
        for i, foto in enumerate(fotos[:4]):
            if foto.filename:
                ruta = os.path.join(UPLOAD_DIR, foto.filename)
                with open(ruta, "wb") as f:
                    f.write(await foto.read())
                fotos_word[i] = InlineImage(doc, ruta, Cm(5))

    # ===============================
    # CONTEXTO
    # ===============================
    contexto = {
        "municipio": municipio,
        "afiliacion": afiliacion,
        "fecha_solicitud": fecha_solicitud,
        "fecha_atencion": fecha_atencion,
        "fecha_cierre": fecha_cierre,
        "tecnologia": tecnologia,
        "ticket": ticket,

        "ajuste": "✔" if ajuste else "",
        "reparacion": "✔" if reparacion else "",
        "reubicacion": "✔" if reubicacion else "",
        "cambio": "✔" if cambio else "",
        "siniestro": "✔" if siniestro else "",
        "otros": "✔" if otros else "",

        "falla": falla,
        "solucion": solucion,
        "observaciones": observaciones,

        "er_descripcion": er_descripcion,
        "er_marca": er_marca,
        "er_modelo": er_modelo,
        "er_serie": er_serie,

        "ei_descripcion": ei_descripcion,
        "ei_marca": ei_marca,
        "ei_modelo": ei_modelo,
        "ei_serie": ei_serie,

        "justificacion": justificacion,
        "tipo_sustitucion": tipo_sustitucion,

        "nombre_coordinador": nombre_coordinador,
        "nombre_responsable": nombre_responsable,

        "foto1": fotos_word[0],
        "foto2": fotos_word[1],
        "foto3": fotos_word[2],
        "foto4": fotos_word[3],
    }

    # ===============================
    # GENERAR ARCHIVO
    # ===============================
    nombre_base = ticket if ticket else "documento"
    output_file = f"{tipo_documento}_{nombre_base.replace('/', '_')}.docx"

    doc.render(contexto)
    doc.save(output_file)

    return FileResponse(
        output_file,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=output_file
    )
