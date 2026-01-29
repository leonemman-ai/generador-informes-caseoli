from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
import os

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates_word", "mantenimiento.docx")

os.makedirs(UPLOAD_DIR, exist_ok=True)

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generar")
async def generar(
    municipio: str = Form(...),
    afiliacion: str = Form(...),
    fecha_solicitud: str = Form(...),
    fecha_atencion: str = Form(...),
    fecha_cierre: str = Form(...),
    tecnologia: str = Form(...),

    ajuste: str = Form(None),
    reparacion: str = Form(None),
    reubicacion: str = Form(None),
    cambio: str = Form(None),
    siniestro: str = Form(None),
    otros: str = Form(None),

    ticket: str = Form(...),
    falla: str = Form(...),
    solucion: str = Form(...),
    observaciones: str = Form(""),

    nombre_coordinador: str = Form(""),
    nombre_responsable: str = Form(""),

    fotos: list[UploadFile] = File(...)
):
    doc = DocxTemplate(TEMPLATE_PATH)

    fotos_word = ["", "", "", ""]
    for i, foto in enumerate(fotos[:4]):
        ruta = os.path.join(UPLOAD_DIR, foto.filename)
        with open(ruta, "wb") as f:
            f.write(await foto.read())
        fotos_word[i] = InlineImage(doc, ruta, Cm(5))

    contexto = {
        "municipio": municipio,
        "afiliacion": afiliacion,
        "fecha_solicitud": fecha_solicitud,
        "fecha_atencion": fecha_atencion,
        "fecha_cierre": fecha_cierre,
        "tecnologia": tecnologia,

        "ajuste": "✔" if ajuste else "",
        "reparacion": "✔" if reparacion else "",
        "reubicacion": "✔" if reubicacion else "",
        "cambio": "✔" if cambio else "",
        "siniestro": "✔" if siniestro else "",
        "otros": "✔" if otros else "",

        "ticket": ticket,
        "falla": falla,
        "solucion": solucion,
        "observaciones": observaciones,

        "nombre_coordinador": nombre_coordinador,
        "nombre_responsable": nombre_responsable,

        "foto1": fotos_word[0],
        "foto2": fotos_word[1],
        "foto3": fotos_word[2],
        "foto4": fotos_word[3],
    }

    output_file = f"Reporte_{ticket.replace('/', '_')}.docx"
    doc.render(contexto)
    doc.save(output_file)

    return FileResponse(
        output_file,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=output_file
    )
