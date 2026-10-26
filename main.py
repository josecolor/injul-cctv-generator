import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import psycopg2
from flask import Flask, redirect, render_template_string, request, send_file, url_for
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

app = Flask(__name__)

# --- CATÁLOGO OFICIAL DE CÁMARAS POR SUCURSAL ---
CATALOGO_CAMARAS = {
    "Caribbean Agregados": [
        "Frente Oficina",
        "Puerta desde CR",
        "Lado Atras",
        "IP Camera 04",
        "Frente Lado desde Caribbean",
        "Lado Atras esquina puerta",
        "Puerta lado norte",
        "Frente hacia lado norte",
        "Oficina Gabinete data",
        "IP Camera 10",
        "IP Camera 11",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech Blue Mall": [
        "Show 2da planta",
        "Area Clientes 1ra planta",
        "Area 1ra planta puerta trasera",
        "2da Planta banos",
        "Area 1ra planta puerta trasera (2)",
        "Gabinete TI",
        "IP Camera 07",
        "2da Planta banos (2)",
        "IP Camera 09",
        "IP Camera 10",
        "IP Camera 11",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech DCM": [
        "Puerta Principal",
        "Cocina",
        "IP Camera 03",
        "S1 Entrada Vehiculos",
        "S2 Pasillo lateral",
        "IP Camera 06",
        "1ra planta asensor visita",
        "IP Camera 09",
        "IP Camera 11",
        "segundo nivel banera",
        "seg piso griferia",
        "IP Camera 14",
        "IP Camera 15",
        "Entrada edificio",
        "S2 Salida",
        "IP Camera 18",
        "IP Camera 19",
        "5to piso onix miel",
        "S2 Salida V",
        "S2 Salida Planta E",
        "Entrada Calle Parqueos",
        "Frente Edificio",
        "Salida S1 Vehiculos",
        "S1 Entrada al S2",
        "S1 Entrada vista de frente",
        "S2 Area Piezas y Vehiculos",
        "IP Camera 30",
        "almacen griferia piiso 1",
        "piso 1 entrada plancha",
        "IP Camera 33",
        "IP Camera 34",
        "IP Camera 35",
        "IP Camera 36",
        "IP Camera 37",
        "IP Camera 38",
        "IP Camera 39",
        "IP Camera 40",
        "IP Camera 41",
        "IP Camera 42",
        "IP Camera 43",
        "IP Camera 45",
        "Oficina Almacen",
        "Entrada a cobro lobby",
        "IP Camera 49",
        "Salida Asensor c",
        "IP Camera 51",
        "IP Camera 52",
        "IP Camera 53",
        "IP Camera 54",
        "IP Camera 55",
        "IP Camera 56",
    ],
    "Marmotech Outlet RP": [
        "IP Camera 01",
        "IP Camera 02",
        "IP Camera 03",
        "IP Camera 04",
        "IP Camera 05",
        "IP Camera 06",
        "IP Camera 07",
        "IP Camera 08",
        "IP Camera 09",
        "IP Camera 10",
        "IP Camera 11",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech Punta Cana": [
        "Puerta Entrada Principal",
        "Area trasera 1",
        "Area Trasera 2",
        "Area Trasera 5",
        "Area trasera 3",
        "Area Trasera 4",
        "Oficina 1",
        "Area Gabinete datos",
        "Almacen",
        "Area Lateral Derecha",
        "Zona Carga / Descarga",
        "Almacen 1",
        "Frente Edificacion",
        "Instalacion",
        "Entrada Oficinas",
        "Almacen 2 Escalera Subir Techo Oficina",
        "Almacen 3",
    ],
    "Sucursal San Cristóbal": [
        "Almacen Repuesto",
        "Parqueo 1 Tienda",
        "Entrada Tienda S",
        "Parqueo 2 Tienda",
        "Almacen 1 Tienda",
        "C/Entrada Tienda",
        "Almacen 2 Tienda",
        "Almacen 3 Tienda",
        "Calle Entrada 2",
        "Oficina Tienda S",
        "Parqueo 3 Tiend",
        "Almacen 4 Tienda",
        "Almacen 5 Tienda",
        "Area Caja Oficin",
        "MODULARES 1",
        "Entrada Mantenimiento norte",
        "MODULARES Area F",
        "MODULARES 4",
        "Alm. Gene 04",
        "Alm Gene 06",
        "Cuarto de Compresores",
        "Alm Gene 01",
        "Alm Gene 02",
        "MOTORES 2",
        "ENTRADA P1",
        "Puerta Entrada",
        "ARTECH 02",
        "PRODUCCION 1",
        "ARTECH 01",
        "Alm Gene 03",
        "PULIDORAS",
        "ARTECH03",
        "PULIDORA 02",
        "DATA CENTER 1",
        "Carga Area 1 Alm",
        "Puerta Entrada A",
        "Area Motores Alm",
        "Carga 2 Alm Gene",
        "Gasoil entrada /",
        "Area Parqueo 1 F",
        "Salida Fabrica A",
        "Gasoil Dispensad",
        "Gasoil",
        "Parqueo Fabrica",
        "Puerta Entrada G",
        "Garita Guardian",
        "Escritorio Garit",
        "Frente Puerta no",
        "Almacen Repuesto (2)",
        "Ing y Calidad",
        "Alm Gene 07",
        "Area de Planta",
        "ARTECH 1",
        "ARTECH 2",
        "Alm Gene 05",
        "IP Camera 57",
        "IP Camera 58",
        "IP Camera 59",
    ],
    "Marmotech Santiago 1": [
        "Almacen Repuesto",
        "Parqueo 1 Tienda",
        "Entrada Tienda 5",
        "Parqueo 2 Tienda",
        "Almacen 1 Tienda",
        "C/Entrada Tienda",
        "Almacen 2 Tienda",
        "Almacen 3 Tienda",
        "Calle Entrada 2",
        "Oficina Tienda S",
        "Parqueo 3 Tiend",
        "Almacen 4 Tienda",
        "Almacen 5 Tienda",
        "Area Caja Oficin",
        "MODULARES 1",
        "Entrada Mantenimiento norte",
    ],
    "Marmotech Santiago 2": [
        "IP Camera 01",
        "IP Camera 02",
        "IP Camera 03",
        "IP Camera 04",
        "IP Camera 05",
        "IP Camera 06",
        "IP Camera 07",
        "IP Camera 08",
        "IP Camera 09",
        "IP Camera 10",
        "IP Camera 11",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
    "Santiago 2 - Sección Muebles": [
        "MUEBLES",
        "ENTRADA",
        "COCINA",
        "GRIFERIA",
        "PARQUEO O-E",
        "BANOS",
        "PISOS",
        "SALA REUNION",
        "PARQUEO E-O",
        "COUNTER",
        "MASETEROS",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>INJUL - Consolidado CCTV</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 15px; }
        .container { max-width: 950px; margin: 0 auto; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        h1 { color: #1F497D; font-size: 20px; border-bottom: 2px solid #1F497D; padding-bottom: 8px; }
        h3 { color: #1F497D; margin-top: 5px; }
        .card { background: #f8f9fa; padding: 15px; margin-bottom: 15px; border-radius: 6px; border-left: 4px solid #1F497D; }
        .btn { display: inline-block; background: #1F497D; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; font-weight: bold; margin-top: 5px; margin-right: 5px; border: none; cursor: pointer; font-size: 14px; }
        .btn-excel { background: #217346; }
        .btn-ticket { background: #D9534F; }
        .btn:hover { opacity: 0.9; }
        .form-group { margin-bottom: 10px; }
        label { display: block; font-weight: bold; font-size: 13px; margin-bottom: 3px; color: #444; }
        input, select, textarea { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; font-size: 12px; }
        th { background-color: #1F497D; color: white; }
        .camera-row { background: #fff; }
    </style>
    <script>
        function actualizarCamaras() {
            var select = document.getElementById('sucursal-select');
            var sucursalSeleccionada = select.value;
            var divs = document.getElementsByClassName('lista-sucursal');
            for (var i = 0; i < divs.length; i++) {
                divs[i].style.display = 'none';
            }
            var activo = document.getElementById('sucursal-' + sucursalSeleccionada);
            if (activo) {
                activo.style.display = 'block';
            }
            var selectedOption = select.options[select.selectedIndex];
            document.getElementById('nombre-sucursal').value = selectedOption.text;
        }
        window.onload = function() {
            actualizarCamaras();
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>INJUL - Sistema de Monitoreo CCTV</h1>
        
        <div class="card">
            <h3>Seleccionar Sucursal y Revisar Cámaras</h3>
            <div class="form-group">
                <label>Sucursal:</label>
                <select id="sucursal-select" onchange="actualizarCamaras()">
                    {% for suc in catalogo.keys() %}
                    <option value="{{ loop.index0 }}">{{ suc }}</option>
                    {% endfor %}
                </select>
            </div>

            <form action="/guardar-sucursal" method="POST">
                <input type="hidden" id="nombre-sucursal" name="sucursal" value="">
                
                {% for suc, camaras in catalogo.items() %}
                <div id="sucursal-{{ loop.index0 }}" class="lista-sucursal" style="display:none; margin-top: 15px;">
                    <h4 style="color: #1F497D; margin-bottom: 5px;">Cámaras de {{ suc }}</h4>
                    <p style="font-size: 12px; color: #666;">Selecciona el estado de cada cámara para el reporte:</p>
                    <table>
                        <tr>
                            <th>#</th>
                            <th>Nombre de la Cámara (EZStation)</th>
                            <th>Estado</th>
                            <th>Observación</th>
                        </tr>
                        {% for cam in camaras %}
                        <tr class="camera-row">
                            <td>{{ loop.index }}</td>
                            <td style="text-align: left; font-weight: bold;">{{ cam }}</td>
                            <td>
                                <select name="estado_{{ loop.index0 }}" style="padding: 4px;">
                                    <option value="Operativa">Operativa (Sí)</option>
                                    <option value="No visible">No visible (Rojo)</option>
                                    <option value="Por verificar">Por verificar</option>
                                </select>
                            </td>
                            <td>
                                <input type="text" name="obs_{{ loop.index0 }}" placeholder="Detalle si falla..." style="padding: 4px;">
                            </td>
                        </tr>
                        {% endfor %}
                    </table>
                </div>
                {% endfor %}

                <div style="margin-top: 15px;">
                    <button type="submit" class="btn">💾 Guardar Estado de esta Sucursal</button>
                </div>
            </form>
        </div>

        <div class="card">
            <h3>Registrar Ticket o Novedad con Proveedor</h3>
            <form action="/guardar-ticket" method="POST">
                <div class="form-group">
                    <label>Número de Ticket:</label>
                    <input type="text" name="ticket" required placeholder="Ej: 10437 o SIN TICKET">
                </div>
                <div class="form-group">
                    <label>Sucursal o Área Afectada:</label>
                    <input type="text" name="sucursal_area" required placeholder="Ej: Marmotech DCM">
                </div>
