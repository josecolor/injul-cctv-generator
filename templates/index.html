<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>INJUL - Consolidado Maestro de CCTV</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body { background-color: #f8f9fa; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .navbar { background-color: #1b263b; }
        .card-resumen { border-left: 4px solid #415a77; }
        .estado-op { background-color: #d1e7dd !important; color: #0f5132; }
        .estado-nv { background-color: #f8d7da !important; color: #842029; }
        .estado-pv { background-color: #fff3cd !important; color: #664d03; }
    </style>
</head>
<body>

    <!-- Menú Superior -->
    <nav class="navbar navbar-expand-lg navbar-dark shadow-sm">
        <div class="container-fluid">
            <a class="navbar-brand" href="/"><i class="fa-solid fa-video me-2"></i>INJUL CCTV</a>
            <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
                <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNav">
                <ul class="navbar-nav ms-auto">
                    <li class="nav-item"><a class="nav-link active" href="/"><i class="fa-solid fa-house me-1"></i> Resumen</a></li>
                    <li class="nav-item"><a class="nav-link" href="/ciclico"><init class="fa-solid fa-rotate me-1"></init> Monitoreo Cíclico</a></li>
                    <li class="nav-item"><a class="nav-link" href="/historial"><i class="fa-solid fa-clock-rotate-left me-1"></i> Historial</a></li>
                </ul>
            </div>
        </div>
    </nav>

    <!-- Contenido Principal -->
    <div class="container my-4">
        
        <!-- Tarjetas de Resumen Global -->
        <div class="row text-center mb-4">
            <div class="col-md-3 mb-2">
                <div class="card shadow-sm p-3 card-resumen">
                    <h5>Total Cámaras</h5>
                    <h3 class="text-primary">{{ resumen.total }}</h3>
                </div>
            </div>
            <div class="col-md-3 mb-2">
                <div class="card shadow-sm p-3 card-resumen">
                    <h5>Operativas</h5>
                    <h3 class="text-success">{{ resumen.operativas }}</h3>
                </div>
            </div>
            <div class="col-md-3 mb-2">
                <div class="card shadow-sm p-3 card-resumen">
                    <h5>No Visibles</h5>
                    <h3 class="text-danger">{{ resumen.no_visibles }}</h3>
                </div>
            </div>
            <div class="col-md-3 mb-2">
                <div class="card shadow-sm p-3 card-resumen">
                    <h5>Disponibilidad</h5>
                    <h3 class="text-info">{{ "%.1f"|format(resumen.disponibilidad) }}%</h3>
                </div>
            </div>
        </div>

        <!-- Selector de Sucursales y Listado de Cámaras -->
        <div class="card shadow-sm p-4">
            <h4 class="mb-3"><i class="fa-solid fa-building-shield me-2"></i>Seleccionar Sucursal y Revisar Cámaras</h4>
            
            <div class="mb-3">
                <label for="selectSucursal" class="form-label fw-bold">Sucursal:</label>
                <select id="selectSucursal" class="form-select" onchange="cambiarSucursal(this.value)">
                    {% for suc in catalogo.keys() %}
                        <option value="{{ loop.index0 }}" {% if seleccion == loop.index0 %}selected{% endif %}>{{ suc }}</option>
                    {% endfor %}
                </select>
            </div>

            {% set sucursal_actual = list(catalogo.keys())[seleccion] if seleccion < catalogo.keys()|list|length else catalogo.keys()|list|index(0) %}
            
            <h5 class="mt-4 text-secondary">Cámaras de <span class="text-dark">{{ sucursal_actual }}</span></h5>
            
            <div class="table-responsive">
                <table class="table table-bordered table-striped align-middle">
                    <thead class="table-dark">
                        <tr>
                            <th style="width: 5%;">#</th>
                            <th style="width: 30%;">Cámara (EZStation)</th>
                            <th style="width: 15%;">En vivo</th>
                            <th style="width: 15%;">Grabación</th>
                            <th style="width: 20%;">Estado</th>
                            <th style="width: 15%;">Observación</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for cam in catalogo[sucursal_actual] %}
                        {% set info = detalles.get(sucursal_actual, {}).get(cam, {}) %}
                        {% set est = info.get('estado', 'Por verificar') %}
                        <tr>
                            <td>{{ loop.index }}</td>
                            <td><strong>{{ cam }}</strong></td>
                            <td>
                                <input type="text" class="form-control form-control-sm" value="{{ info.get('en_vivo', '') }}" onchange="guardarCamara('{{ sucursal_actual }}', {{ loop.index0 }}, this)">
                            </td>
                            <td>
                                <select class="form-select form-select-sm" onchange="guardarCamara('{{ sucursal_actual }}', {{ loop.index0 }}, this)">
                                    <option value="Normal" {% if info.get('grabacion') == 'Normal' %}selected{% endif %}>Normal</option>
                                    <option value="Por movimiento" {% if info.get('grabacion') == 'Por movimiento' %}selected{% endif %}>Por movimiento</option>
                                    <option value="No graba" {% if info.get('grabacion') == 'No graba' %}selected{% endif %}>No graba</option>
                                </select>
                            </td>
                            <td>
                                <select class="form-select form-select-sm {% if est == 'Operativa' %}estado-op{% elif est == 'No visible' %}estado-nv{% else %}estado-pv{% endif %}" onchange="guardarCamara('{{ sucursal_actual }}', {{ loop.index0 }}, this)">
                                    {% for valor, texto in estados %}
                                    <option value="{{ valor }}" {% if est == valor %}selected{% endif %}>{{ texto }}</option>
                                    {% endfor %}
                                </select>
                            </td>
                            <td>
                                <input type="text" class="form-control form-control-sm" value="{{ info.get('observacion', '') }}" placeholder="Detalle..." onchange="guardarCamara('{{ sucursal_actual }}', {{ loop.index0 }}, this)">
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>

    </div>

    <script>
        function cambiarSucursal(idx) {
            window.location.href = "/?s=" + idx;
        }

        function guardarCamara(sucursal, idx, elemento) {
            // Lógica asíncrona para guardar cambios sin recargar toda la página
            let fila = elemento.closest('tr');
            let inputs = fila.querySelectorAll('input, select');
            
            let datos = {
                sucursal: sucursal,
                idx: idx,
                en_vivo: inputs[0].value,
                grabacion: inputs[1].value,
                estado: inputs[2].value,
                observacion: inputs[3].value
            };

            fetch('/api/camara', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(datos)
            }).then(res => res.json()).then(res => {
                if(res.ok) {
                    // Actualiza clase visual del selector de estado
                    let selEstado = inputs[2];
                    selEstado.className = "form-select form-select-sm " + (selEstado.value === 'Operativa' ? 'estado-op' : (selEstado.value === 'No visible' ? 'estado-nv' : 'estado-pv'));
                }
            });
        }
    </script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
