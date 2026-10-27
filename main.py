"""INJUL - Consolidado Maestro de CCTV (Versión Completa y Definitiva)."""

import json
import logging
import os
import re
from contextlib import contextmanager
from datetime import date, datetime
from io import BytesIO
from xml.sax.saxutils import escape

import openpyxl
import psycopg2
import redis
from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)

try:
    from zoneinfo import ZoneInfo
    ZONA_HORARIA = ZoneInfo("America/Santo_Domingo")
except Exception:
    from datetime import timedelta, timezone
    ZONA_HORARIA = timezone(timedelta(hours=-4), "AST")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("injul-cctv")

app = Flask(__name__)

DATABASE_URL = os.environ.get("DATABASE_URL")
REDIS_URL = os.environ.get("REDIS_URL")

CATALOGO_CAMARAS = {
    "Caribbean Agregados": [
        "Frente Oficina", "Puerta desde CR", "Lado Atras", "IP Camera 04",
        "Frente Lado desde Caribbean", "Lado Atras esquina puerta", "Puerta lado norte",
        "Frente hacia lado norte", "Oficina Gabinete data", "IP Camera 10", "IP Camera 11",
        "IP Camera 12", "IP Camera 13", "IP Camera 14", "IP Camera 15", "IP Camera 16",
    ],
    "Marmotech Blue Mall": [
        "Show 2da planta", "Area Clientes 1ra planta", "Area 1ra planta puerta trasera",
        "2da Planta banos", "Area 1ra planta puerta trasera (2)", "Gabinete TI",
        "IP Camera 07", "2da Planta banos (2)", "IP Camera 09", "IP Camera 10",
        "IP Camera 11", "IP Camera 12", "IP Camera 13", "IP Camera 14", "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech DCM": [
        "Puerta Principal", "Cocina", "IP Camera 03", "S1 Entrada Vehiculos",
        "S2 Pasillo lateral", "IP Camera 06", "1ra planta asensor visita", "IP Camera 09",
        "IP Camera 11", "segundo nivel banera", "seg piso griferia", "IP Camera 14",
        "IP Camera 15", "Entrada edificio", "S2 Salida", "IP Camera 18", "IP Camera 19",
        "5to piso onix miel", "S2 Salida V", "S2 Salida Planta E", "Entrada Calle Parqueos",
        "Frente Edificio", "Salida S1 Vehiculos", "S1 Entrada al S2",
        "S1 Entrada vista de frente", "S2 Area Piezas y Vehiculos", "IP Camera 30",
        "almacen griferia piiso 1", "piso 1 entrada plancha", "IP Camera 33",
        "IP Camera 34", "IP Camera 35", "IP Camera 36", "IP Camera 37", "IP Camera 38",
        "IP Camera 39", "IP Camera 40", "IP Camera 41", "IP Camera 42", "IP Camera 43",
        "IP Camera 45", "Oficina Almacen", "Entrada a cobro lobby", "IP Camera 49",
        "Salida Asensor c", "IP Camera 51", "IP Camera 52", "IP Camera 53", "IP Camera 54",
        "IP Camera 55", "IP Camera 56",
    ],
    "Marmotech Outlet RP": [
        "IP Camera 01", "IP Camera 02", "IP Camera 03", "IP Camera 04", "IP Camera 05",
        "IP Camera 06", "IP Camera 07", "IP Camera 08", "IP Camera 09", "IP Camera 10",
        "IP Camera 11", "IP Camera 12", "IP Camera 13", "IP Camera 14", "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech Punta Cana": [
        "Puerta Entrada Principal", "Area trasera 1", "Area Trasera 2", "Area Trasera 5",
        "Area trasera 3", "Area Trasera 4", "Oficina 1", "Area Gabinete datos", "Almacen",
        "Area Lateral Derecha", "Zona Carga / Descarga", "Almacen 1", "Frente Edificacion",
        "Instalacion", "Entrada Oficinas", "Almacen 2 Escalera Subir Techo Oficina",
        "Almacen 3",
    ],
    "Sucursal San Cristóbal": [
        "Almacen Repuesto", "Parqueo 1 Tienda", "Entrada Tienda S", "Parqueo 2 Tienda",
        "Almacen 1 Tienda", "C/Entrada Tienda", "Almacen 2 Tienda", "Almacen 3 Tienda",
        "Calle Entrada 2", "Oficina Tienda S", "Parqueo 3 Tiend", "Almacen 4 Tienda",
        "Almacen 5 Tienda", "Area Caja Oficin", "MODULARES 1",
        "Entrada Mantenimiento norte", "MODULARES Area F", "MODULARES 4", "Alm. Gene 04",
        "Alm Gene 06", "Cuarto de Compresores", "Alm Gene 01", "Alm Gene 02", "MOTORES 2",
        "ENTRADA P1", "Puerta Entrada", "ARTECH 02", "PRODUCCION 1", "ARTECH 01",
        "Alm Gene 03", "PULIDORAS", "ARTECH03", "PULIDORA 02", "DATA CENTER 1",
        "Carga Area 1 Alm", "Puerta Entrada A", "Area Motores Alm", "Carga 2 Alm Gene",
        "Gasoil entrada /", "Area Parqueo 1 F", "Salida Fabrica A", "Gasoil Dispensad",
        "Gasoil", "Parqueo Fabrica", "Puerta Entrada G", "Garita Guardian",
        "Escritorio Garit", "Frente Puerta no", "Almacen Repuesto", "Almacen Repuesto (2)",
        "Ing y Calidad", "Alm Gene 07", "Area de Planta", "ARTECH 1", "ARTECH 2",
        "Alm Gene 05", "IP Camera 57", "IP Camera 58", "IP Camera 59",
    ],
    "Marmotech Santiago 1": [
        "Almacen Repuesto", "Parqueo 1 Tienda", "Entrada Tienda S", "Parqueo 2 Tienda",
        "Almacen 1 Tienda", "C/Entrada Tienda", "Almacen 2 Tienda", "Almacen 3 Tienda",
        "Calle Entrada 2", "Oficina Tienda S", "Parqueo 3 Tiend", "Almacen 4 Tienda",
        "Almacen 5 Tienda", "Area Caja Oficin", "MODULARES 1",
        "Entrada Mantenimiento norte",
    ],
    "Marmotech Santiago 2": [
        "IP Camera 01", "IP Camera 02", "IP Camera 03", "IP Camera 04", "IP Camera 05",
        "IP Camera 06", "IP Camera 07", "IP Camera 08", "IP Camera 09", "IP Camera 10",
        "IP Camera 11", "IP Camera 12", "IP Camera 13", "IP Camera 14", "IP Camera 15",
        "IP Camera 16",
    ],
    "Santiago 2 - Sección Muebles": [
        "MUEBLES", "ENTRADA", "COCINA", "GRIFERIA", "PARQUEO O-E", "BANOS", "PISOS",
        "SALA REUNION", "PARQUEO E-O", "COUNTER", "MASETEROS", "IP Camera 12",
        "IP Camera 13", "IP Camera 14", "IP Camera 15", "IP Camera 16",
    ],
}

CATALOGO_CICLICO = {
    "Sucursal San Cristóbal": [
        "Entrada puerta 1 - entrada p1", "Área de diseño artech - artech02",
        "Área de diseño artech 2 - artech01", "Pulidoras", "Pulidora 02",
        "Carga almacén general - carga área 1 alm", "Carga almacén general 2",
        "Entrada a fábrica puerta 3 - salida fábrica A",
        "Garita punchado - puerta entrada G", "Calle principal puerta 2 - frente puerta",
    ],
    "Marmotech DCM": [
        "IP Camera 30", "Frente de la puerta, utensilios - IP Camera 32",
        "Parqueo de motores - IP Camera 34", "Terraza lado céntrico - IP Camera 36",
        "IP Camera 38", "Parqueo sótano 1 - IP Camera 40",
        "Pasillos de exhibidores parte céntrica - IP Camera 42", "IP Camera 44",
        "IP Camera 46", "IP Camera 48",
    ],
    "Marmotech Santiago 1": [
        "Calle frente a la empresa - IP Camera 02", "Parte trasera - IP Camera 04",
        "Calle principal - IP Camera 06", "Mostrador - IP Camera 08",
        "Almacén interior 2 - IP Camera 10",
    ],
    "Marmotech Santiago 2": [
        "Mostrador - Muebles", "Cocina", "Parqueo principal - Parqueo 0-E", "Pisos",
        "Parqueo E-O",
    ],
    "Marmotech Blue Mall": [
        "Show 2da Planta", "Área 1ra Planta Puerta Trasera",
        "Área 1ra Planta Puerta Trasera Norte", "Área Clientes 1ra Planta", "Gabinete TI",
    ],
    "Marmotech Outlet RP": [
        "Parqueo lado derecho - IP Camera 10", "Parte atrás, compresores A/C - IP Camera 12",
        "Almacén General - IP Camera 14", "Exhibidores de venta - IP Camera 16",
        "Cámara escalera lado izquierdo - IP Camera 02",
    ],
    "Marmotech Punta Cana": [
        "Área Trasera 1", "Área Trasera 5", "Área Trasera 4", "Área Gabinete Datos",
        "Puerta Entrada Principal",
    ],
    "Marmotech Caribbean": [
        "CB 6-7-8 Planta Eléctrica", "Taller Soldar", "Oficina Almacén Repuesto",
        "Almacén Repuesto", "Taller Fabricación",
    ],
    "Caribbean Agregados": [
        "Frente Lado desde Caribbean", "Lado Atrás esquina puerta", "Puerta lado norte",
        "Frente hacia lado norte", "Oficina Gabinete data",
    ],
}

ANCLA_ROTACION = date(2026, 10, 24)
INVENTARIO_DE = {k: k for k in CATALOGO_CAMARAS}

def _construir_pool(clave):
    base = list(CATALOGO_CICLICO[clave])
    vistos = {n.lower() for n in base}
    nums = {m for n in base for m in re.findall(r"ip camera\s*0*(\d+)", n.lower())}
    extra = []
    for n in CATALOGO_CAMARAS.get(INVENTARIO_DE.get(clave, ""), []):
        ip = re.findall(r"ip camera\s*0*(\d+)$", n.lower())
        if n.lower() in vistos or (ip and ip[0] in nums):
            continue
        vistos.add(n.lower())
        extra.append(n)
    return base + extra

POOL_CICLICO = {clave: _construir_pool(clave) for clave in CATALOGO_CICLICO}

ESTADOS = [
    ("Operativa", "Operativa (Sí)"),
    ("No visible", "No visible (Rojo)"),
    ("Por verificar", "Por verificar"),
]
ESTADOS_VALIDOS = {valor for valor, _ in ESTADOS}
CLASES_ESTADO = {"Operativa": "estado-op", "No visible": "estado-nv", "Por verificar": "estado-pv"}
MENSAJES = {
    "ticket_ok": ("ok", "Ticket registrado correctamente."),
    "sin_db": ("error", "Base de datos no configurada."),
    "error": ("error", "No se pudo guardar."),
}

def ahora():
    return datetime.now(ZONA_HORARIA) if ZONA_HORARIA else datetime.now()

def local_txt(momento):
    if momento is None:
        return ""
    if ZONA_HORARIA and momento.tzinfo:
        momento = momento.astimezone(ZONA_HORARIA)
    return momento.strftime("%d/%m/%Y %I:%M:%S %p")

def ip_cliente():
    reenviado = request.headers.get("X-Forwarded-For", "")
    return (reenviado.split(",")[0].strip() or request.remote_addr or "")[:60]

@contextmanager
def db_cursor():
    conn = psycopg2.connect(DATABASE_URL, connect_timeout=5)
    try:
        with conn:
            with conn.cursor() as cur:
                yield cur
    finally:
        conn.close()

def inicializar_base_de_datos():
    if not DATABASE_URL:
