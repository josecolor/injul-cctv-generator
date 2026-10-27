"""Módulo de Monitoreo Cíclico para INJUL CCTV."""

from flask import render_template, request, redirect, url_for

def registrar(app, catalogo_camaras, pool_ciclico, ancla_rotacion, camaras_no_visibles, inventario_de, db_cursor, ahora, local_txt, ip_cliente, database_url):
    
    @app.route("/ciclico")
    def ciclico_index():
        try:
            hoy = ahora().date()
            dias_dif = (hoy - ancla_rotacion).days
        except Exception:
            dias_dif = 0
            
        sucursales = list(pool_ciclico.keys())
        idx_sucursal = abs(dias_dif) % len(sucursales) if sucursales else 0
        sucursal_hoy = sucursales[idx_sucursal] if sucursales else "Sucursal San Cristóbal"
        
        pool = pool_ciclico.get(sucursal_hoy, [])
        camaras_dia = []
        for i in range(5):
            if pool:
                camaras_dia.append(pool[(dias_dif + i) % len(pool)])
                
        return render_template(
            "ciclico.html",
            sucursal_hoy=sucursal_hoy,
            camaras_dia=camaras_dia,
            sucursales=sucursales,
            dias_dif=dias_dif
        )

    @app.route("/ciclico/guardar", methods=["POST"])
    def ciclico_guardar():
        if not database_url:
            return redirect(url_for("ciclico_index"))
        
        sucursal = request.form.get("sucursal")
        camara = request.form.get("camara")
        estado = request.form.get("estado")
        observacion = request.form.get("observacion", "")
        
        try:
            with db_cursor() as cur:
                cur.execute("""
                    INSERT INTO ciclico_camaras (sucursal, camara, estado, observacion, actualizado_en)
                    VALUES (%s, %s, %s, %s, NOW())
                    ON CONFLICT (sucursal, camara) DO UPDATE
                    SET estado = EXCLUDED.estado, observacion = EXCLUDED.observacion, actualizado_en = NOW()
                """, (sucursal, camara, estado, observacion))
        except Exception:
            pass
            
        return redirect(url_for("ciclico_index"))
