from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import sqlite3
import os
import html
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from typing import List, Optional

DB_PATH = "controlpremium.db"
PDF_DIR = "presupuestos"
LOGO_PATH = "logo_empresa.png"

app = FastAPI(
    title="Control Premium Web API",
    description="Servicio web para gestionar clientes, stock y presupuestos desde Render.",
    version="1.0.0"
)

app.mount("/static", StaticFiles(directory="static"), name="static")

os.makedirs(PDF_DIR, exist_ok=True)


def get_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.executescript(r"""
    CREATE TABLE IF NOT EXISTS clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT UNIQUE,
        telefono TEXT,
        direccion TEXT,
        email TEXT,
        dni_cuit TEXT
    );

    CREATE TABLE IF NOT EXISTS proveedores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT,
        telefono TEXT,
        contacto TEXT
    );

    CREATE TABLE IF NOT EXISTS stock (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        producto TEXT UNIQUE,
        cantidad INTEGER,
        proveedor_id INTEGER,
        FOREIGN KEY(proveedor_id) REFERENCES proveedores(id)
    );

    CREATE TABLE IF NOT EXISTS presupuestos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER,
        fecha TEXT,
        obra_descripcion TEXT,
        fecha_inicio TEXT,
        condiciones TEXT,
        subtotal REAL,
        iva REAL,
        monto REAL,
        estado TEXT DEFAULT 'Pendiente',
        FOREIGN KEY(cliente_id) REFERENCES clientes(id)
    );

    CREATE TABLE IF NOT EXISTS presupuesto_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        presupuesto_id INTEGER,
        concepto TEXT,
        unidad TEXT,
        cantidad REAL,
        precio_unitario REAL,
        subtotal REAL,
        FOREIGN KEY(presupuesto_id) REFERENCES presupuestos(id)
    );

    CREATE TABLE IF NOT EXISTS movimientos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT,
        descripcion TEXT,
        monto REAL,
        fecha TEXT
    );

    CREATE TABLE IF NOT EXISTS configuracion (
        clave TEXT PRIMARY KEY,
        valor TEXT
    );
    """)
    conn.commit()

    cursor.execute("SELECT COUNT(*) FROM proveedores")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO proveedores(nombre, telefono, contacto) VALUES (?, ?, ?)", ("Proveedor A", "", ""))
        cursor.execute("INSERT INTO proveedores(nombre, telefono, contacto) VALUES (?, ?, ?)", ("Proveedor B", "", ""))
        conn.commit()

    cursor.execute("SELECT COUNT(*) FROM stock")
    if cursor.fetchone()[0] == 0:
        cursor.execute("SELECT id FROM proveedores LIMIT 1")
        proveedor_id = cursor.fetchone()[0]
        sample_stock = [
            ("Placa Standard", 50, proveedor_id),
            ("Placa Anti-humedad", 30, proveedor_id),
            ("Placa Decorativa", 20, proveedor_id),
            ("Pintura Blanca", 100, proveedor_id),
            ("Pintura Color", 80, proveedor_id),
            ("Barniz", 50, proveedor_id),
        ]
        cursor.executemany("INSERT INTO stock(producto, cantidad, proveedor_id) VALUES (?, ?, ?)", sample_stock)
        conn.commit()

    default_config = {
        "EMPRESA_NOMBRE": "Stoplac",
        "EMPRESA_CUIT": "30-xxxxxxxx-x",
        "EMPRESA_DIRECCION": "Calle Falsa 123, Ciudad",
        "EMPRESA_TELEFONO": "11-1234-5678",
        "EMPRESA_CONTACTO": "info@stoplac.com",
        "EMPRESA_LOGO": LOGO_PATH,
    }
    for clave, valor in default_config.items():
        cursor.execute("INSERT OR IGNORE INTO configuracion(clave, valor) VALUES (?, ?)", (clave, valor))
    conn.commit()
    conn.close()


class Cliente(BaseModel):
    nombre: str
    telefono: Optional[str] = ""
    direccion: Optional[str] = ""
    email: Optional[str] = ""
    dni_cuit: Optional[str] = ""


class PresupuestoItem(BaseModel):
    concepto: str
    unidad: str
    cantidad: float
    precio_unitario: float
    subtotal: Optional[float] = None


class PresupuestoCreate(BaseModel):
    cliente_id: int
    obra_descripcion: str
    fecha_inicio: Optional[str] = ""
    condiciones: Optional[str] = ""
    subtotal: float
    iva: float
    monto: float
    items: List[PresupuestoItem]


@app.on_event("startup")
def startup_event():
    init_db()


@app.get("/")
def read_root():
    return RedirectResponse(url="/static/index.html")


@app.get("/source", response_class=HTMLResponse)
def view_source():
    source_path = "ControlPremium.py"
    if not os.path.exists(source_path):
        raise HTTPException(status_code=404, detail="Source file not found")
    with open(source_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
    escaped = html.escape(content)
    html_body = f"""
    <html>
        <head>
            <title>ControlPremium.py</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; background: #f4f4f7; color: #222; }}
                pre {{ white-space: pre-wrap; word-wrap: break-word; font-family: monospace; background: #1e1e1e; color: #f8f8f2; padding: 16px; border-radius: 8px; overflow-x: auto; }}
                a {{ color: #0078D7; text-decoration: none; }}
                a:hover {{ text-decoration: underline; }}
            </style>
        </head>
        <body>
            <h1>ControlPremium.py</h1>
            <p><a href="/">Volver a la app</a></p>
            <pre>{escaped}</pre>
        </body>
    </html>
    """
    return HTMLResponse(content=html_body)


@app.get("/health")
def health_check():
    return {"status": "ok", "timestamp": datetime.utcnow().isoformat()}


@app.get("/clientes")
def list_clientes():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, telefono, direccion, email, dni_cuit FROM clientes ORDER BY nombre")
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows


@app.post("/clientes")
def create_cliente(cliente: Cliente):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            "INSERT INTO clientes(nombre, telefono, direccion, email, dni_cuit) VALUES (?, ?, ?, ?, ?)",
            (cliente.nombre, cliente.telefono, cliente.direccion, cliente.email, cliente.dni_cuit)
        )
        conn.commit()
        cliente_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        conn.close()
        raise HTTPException(status_code=400, detail="El cliente ya existe")
    conn.close()
    return {"id": cliente_id, **cliente.dict()}


@app.get("/stock")
def list_stock():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT s.id, s.producto, s.cantidad, s.proveedor_id, p.nombre AS proveedor FROM stock s LEFT JOIN proveedores p ON s.proveedor_id = p.id ORDER BY s.producto")
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows


@app.get("/presupuestos")
def list_presupuestos():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT p.id, c.nombre AS cliente, p.fecha, p.obra_descripcion, p.monto, p.estado FROM presupuestos p LEFT JOIN clientes c ON p.cliente_id = c.id ORDER BY p.id DESC"
    )
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows


@app.get("/presupuestos/{presupuesto_id}")
def get_presupuesto(presupuesto_id: int):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT p.id, c.nombre AS cliente, p.fecha, p.obra_descripcion, p.fecha_inicio, p.condiciones, p.subtotal, p.iva, p.monto, p.estado FROM presupuestos p LEFT JOIN clientes c ON p.cliente_id = c.id WHERE p.id = ?",
        (presupuesto_id,)
    )
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Presupuesto no encontrado")

    cursor.execute(
        "SELECT concepto, unidad, cantidad, precio_unitario, subtotal FROM presupuesto_items WHERE presupuesto_id = ?",
        (presupuesto_id,)
    )
    items = [dict(item) for item in cursor.fetchall()]
    conn.close()
    resultado = dict(row)
    resultado["items"] = items
    return resultado


@app.post("/presupuestos")
def create_presupuesto(data: PresupuestoCreate):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM clientes WHERE id = ?", (data.cliente_id,))
    if cursor.fetchone() is None:
        conn.close()
        raise HTTPException(status_code=404, detail="Cliente no encontrado")

    cursor.execute(
        "INSERT INTO presupuestos(cliente_id, fecha, obra_descripcion, fecha_inicio, condiciones, subtotal, iva, monto) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (data.cliente_id, datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'), data.obra_descripcion, data.fecha_inicio, data.condiciones, data.subtotal, data.iva, data.monto)
    )
    presupuesto_id = cursor.lastrowid

    for item in data.items:
        subtotal_item = item.subtotal if item.subtotal is not None else item.cantidad * item.precio_unitario
        cursor.execute(
            "INSERT INTO presupuesto_items(presupuesto_id, concepto, unidad, cantidad, precio_unitario, subtotal) VALUES (?, ?, ?, ?, ?, ?)",
            (presupuesto_id, item.concepto, item.unidad, item.cantidad, item.precio_unitario, subtotal_item)
        )

    conn.commit()
    conn.close()
    return {
        "id": presupuesto_id,
        "message": "Presupuesto creado",
        "subtotal": data.subtotal,
        "iva": data.iva,
        "monto": data.monto,
    }


def draw_logo_on_canvas(c):
    if os.path.exists(LOGO_PATH):
        try:
            logo = ImageReader(LOGO_PATH)
            c.drawImage(logo, 40, 760, width=120, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass


def generar_pdf_presupuesto(presupuesto_id: int, path: str):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT p.id, c.nombre AS cliente, p.fecha, p.obra_descripcion, p.fecha_inicio, p.condiciones, p.subtotal, p.iva, p.monto FROM presupuestos p LEFT JOIN clientes c ON p.cliente_id = c.id WHERE p.id = ?",
        (presupuesto_id,)
    )
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise ValueError("Presupuesto no encontrado")

    cursor.execute(
        "SELECT concepto, unidad, cantidad, precio_unitario, subtotal FROM presupuesto_items WHERE presupuesto_id = ?",
        (presupuesto_id,)
    )
    items = [dict(item) for item in cursor.fetchall()]
    conn.close()

    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4
    draw_logo_on_canvas(c)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(200, height - 80, "Presupuesto Control Premium")
    c.setFont("Helvetica", 11)
    c.drawString(40, height - 120, f"Presupuesto ID: {row['id']}")
    c.drawString(40, height - 135, f"Cliente: {row['cliente']}")
    c.drawString(40, height - 150, f"Fecha: {row['fecha']}")
    c.drawString(40, height - 165, f"Obra: {row['obra_descripcion']}")
    c.drawString(40, height - 180, f"Fecha inicio: {row['fecha_inicio']}")

    data = [["Concepto", "Unidad", "Cantidad", "P. Unitario", "Subtotal"]]
    for item in items:
        data.append([
            item["concepto"],
            item["unidad"],
            f"{item['cantidad']}",
            f"${item['precio_unitario']:.2f}",
            f"${item['subtotal']:.2f}"
        ])

    table = Table(data, colWidths=[180, 70, 70, 90, 90])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    table.wrapOn(c, width, height)
    table.drawOn(c, 40, height - 360)

    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(width - 40, 120, f"Subtotal: ${row['subtotal']:.2f}")
    c.drawRightString(width - 40, 100, f"IVA: ${row['iva']:.2f}")
    c.drawRightString(width - 40, 80, f"Total: ${row['monto']:.2f}")

    c.showPage()
    c.save()


@app.get("/presupuestos/{presupuesto_id}/pdf")
def download_presupuesto_pdf(presupuesto_id: int):
    filename = f"presupuesto_{presupuesto_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    path = os.path.join(PDF_DIR, filename)
    try:
        generar_pdf_presupuesto(presupuesto_id, path)
    except ValueError as error:
        raise HTTPException(status_code=404, detail=str(error))
    return FileResponse(path, filename=filename, media_type="application/pdf")
