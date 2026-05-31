import streamlit as st
import pandas as pd
import os
import io
from datetime import date, datetime
#streamlit run Application.py

st.set_page_config(
    page_title="Captura de Costos - Plancha de Acero",
    page_icon="-",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Archivo CSV donde se guardan todos los registros
CSV_PATH = "registros_produccion.csv"

# Costos estándar por producto (S/ por unidad) — basados en datos reales
# Fuente: hoja CostoEstandar del Excel original
COSTOS_ESTANDAR = {
    "Barra de Acero 6m": {"mp": 40, "mod": 10, "cif": 8, "total": 58},
    "Varilla Corrugada": {"mp": 35, "mod": 9, "cif": 7, "total": 51},
    "Alambrón": {"mp": 30, "mod": 8, "cif": 6, "total": 44},
    "Clavo 2\"": {"mp": 5, "mod": 2, "cif": 1, "total": 8},
    "Plancha de Acero": {"mp": 60, "mod": 12, "cif": 10, "total": 82},
}

# Precio de venta por unidad (S/)
PRECIO_VENTA = {
    "Barra de Acero 6m": 85,
    "Varilla Corrugada": 78,
    "Alambrón": 65,
    "Clavo 2\"": 12,
    "Plancha de Acero": 120,
}

# Materiales disponibles para consumo de MP
MATERIALES_MP = ["Chatarra", "Mineral", "Palanquilla", "Scrap de acero", "Otro"]

# Tipos de CIF disponibles
TIPOS_CIF = ["Energía", "Mantenimiento", "Depreciación", "Agua", "Otro"]

# Columnas del CSV (orden exacto en que se guardan)
COLUMNAS_CSV = [
    "FechaRegistro",  # Timestamp de cuando se guardó
    "FechaProduccion",  # Fecha de la orden de producción
    "OrdenID",  # Número de orden (ej: OP008)
    "Planta",  # Planta donde se produjo
    "Producto",  # Nombre del producto
    "CantidadProducida",  # Unidades producidas
    "Material_MP",  # Tipo de materia prima usada
    "CantidadKg_MP",  # Kilogramos de MP consumidos
    "CostoUnitario_MP",  # Costo por kg de MP (S/)
    "CostoTotal_MP",  # Costo total de MP = kg × costo/kg
    "HorasTrabajadas",  # Horas de mano de obra directa
    "CostoHora_MOD",  # Costo por hora de MOD (S/)
    "CostoTotal_MOD",  # Costo total MOD = horas × costo/hora
    "TipoCIF",  # Tipo de costo indirecto
    "CostoTotal_CIF",  # Monto total de CIF
    "CostoTotal_Real",  # Suma MP + MOD + CIF
    "CostoRealPorUnidad",  # CostoTotal_Real / CantidadProducida
    "CostoStd_PorUnidad",  # Costo estándar según tabla de referencia
    "Varianza_PorUnidad",  # Real - Estándar (negativo = eficiente)
    "Estado_Varianza",  # Eficiente / Alerta / Crítico
    "Ingresos",  # PrecioVenta × CantidadProducida
    "UtilidadBruta",  # Ingresos - CostoTotal_Real
    "MargenBruto_Pct",  # (UtilidadBruta / Ingresos) × 100
    "Observaciones",  # Notas libres del operario
]


# ─────────────────────────────────────────────────────────────
# FUNCIONES AUXILIARES
# ─────────────────────────────────────────────────────────────

def cargar_datos() -> pd.DataFrame:
    """
    Carga el CSV existente si ya existe.
    Si no existe, devuelve un DataFrame vacío con las columnas correctas.
    Nunca sobreescribe datos previos.
    """
    if os.path.exists(CSV_PATH):
        try:
            df = pd.read_csv(CSV_PATH, encoding="utf-8-sig")
            # Asegura que todas las columnas existan aunque el archivo sea antiguo
            for col in COLUMNAS_CSV:
                if col not in df.columns:
                    df[col] = ""
            return df[COLUMNAS_CSV]  # Devuelve solo las columnas esperadas en orden
        except Exception as e:
            st.error(f"Error al leer el archivo CSV: {e}")
            return pd.DataFrame(columns=COLUMNAS_CSV)
    else:
        return pd.DataFrame(columns=COLUMNAS_CSV)


def guardar_registro(nuevo_registro: dict) -> bool:
    """
    Agrega una nueva fila al CSV sin sobreescribir los datos existentes.
    Retorna True si se guardó correctamente, False si hubo un error.
    """
    try:
        df_nuevo = pd.DataFrame([nuevo_registro])
        # mode='a' = append (agregar), header solo si el archivo no existe
        df_nuevo.to_csv(
            CSV_PATH,
            mode="a",
            header=not os.path.exists(CSV_PATH),
            index=False,
            encoding="utf-8-sig",  # utf-8-sig evita problemas con tildes en Excel
        )
        return True
    except Exception as e:
        st.error(f"Error al guardar: {e}")
        return False


def clasificar_varianza(varianza: float) -> str:
    """
    Clasifica la varianza de costo respecto al estándar.
    Negativo = gastó MENOS que el estándar = eficiente (bueno).
    Positivo = gastó MÁS que el estándar = problema.
    """
    if varianza < 0:
        return "Eficiente"
    elif varianza < 3:
        return "Alerta"
    else:
        return "Crítico"


def validar_formulario(
        fecha_prod, orden_id, cantidad, kg_mp, costo_unit_mp,
        horas_mod, costo_hora_mod, costo_cif
) -> list[str]:
    """
    Ejecuta todas las validaciones de negocio y devuelve
    una lista de mensajes de error. Lista vacía = sin errores.
    """
    errores = []

    # 1. Fecha no puede ser futura
    if fecha_prod > date.today():
        errores.append("La fecha de producción no puede ser una fecha futura.")

    # 2. Número de orden no puede estar vacío
    if not orden_id.strip():
        errores.append("El número de orden es obligatorio.")

    # 3. El ID de orden debe seguir el patrón OP+números (flexible)
    orden_limpio = orden_id.strip().upper()
    if orden_limpio and not orden_limpio.startswith("OP"):
        errores.append("El número de orden debería comenzar con 'OP' (ej: OP008).")

    # 4. Cantidad producida debe ser mayor a cero
    if cantidad <= 0:
        errores.append("La cantidad producida debe ser mayor a cero.")

    # 5. Kg de materia prima debe ser mayor a cero
    if kg_mp <= 0:
        errores.append("Los kg de materia prima deben ser mayores a cero.")

    # 6. Costo unitario MP no puede ser negativo ni cero
    if costo_unit_mp <= 0:
        errores.append("El costo por kg de materia prima debe ser mayor a cero.")

    # 7. Validación de costo por kg absurdamente alto (> S/ 500/kg)
    if costo_unit_mp > 500:
        errores.append("El costo por kg de MP supera S/ 500. ¿Es correcto? Verifica el valor.")

    # 8. Horas de MOD deben ser positivas
    if horas_mod <= 0:
        errores.append("Las horas de mano de obra deben ser mayores a cero.")

    # 9. Costo por hora MOD debe ser positivo
    if costo_hora_mod <= 0:
        errores.append("El costo por hora de mano de obra debe ser mayor a cero.")

    # 10. Costo por hora MOD no debería ser absurdo (> S/ 500/hora)
    if costo_hora_mod > 500:
        errores.append("El costo/hora de MOD supera S/ 500. Verifica el valor.")

    # 11. CIF no puede ser negativo
    if costo_cif < 0:
        errores.append("Los costos indirectos no pueden ser negativos.")

    return errores


def exportar_excel(df: pd.DataFrame) -> bytes:
    """
    Convierte el DataFrame a un archivo Excel en memoria (sin guardar en disco).
    Retorna los bytes del archivo para que Streamlit lo ofrezca como descarga.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Registros", index=False)

        # Ajusta el ancho de las columnas automáticamente
        worksheet = writer.sheets["Registros"]
        for col_idx, column in enumerate(df.columns, 1):
            max_length = max(
                df[column].astype(str).map(len).max(),
                len(column)
            ) + 2
            # Excel tiene un límite de 255 caracteres por ancho de columna
            col_width = min(max_length, 50)
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            worksheet.column_dimensions[col_letter].width = col_width

    return buffer.getvalue()


def generar_orden_sugerida(df: pd.DataFrame) -> str:
    """
    Sugiere el siguiente número de orden basado en los registros existentes.
    Si hay registros, incrementa el último número. Si no hay, comienza en OP008
    (continuando desde los 7 que ya existían en el Excel original).
    """
    if df.empty:
        return "OP008"
    try:
        ordenes_existentes = df["OrdenID"].dropna().str.upper()
        numeros = ordenes_existentes.str.extract(r"OP(\d+)")[0].dropna().astype(int)
        if numeros.empty:
            return "OP008"
        return f"OP{numeros.max() + 1:03d}"
    except Exception:
        return "OP008"


def colorear_estado(val):
    """Función para aplicar color condicional en la tabla de datos."""
    colores = {
        "Eficiente": "background-color: #E1F5EE; color: #085041;",
        "Alerta": "background-color: #FAEEDA; color: #633806;",
        "Crítico": "background-color: #FCEBEB; color: #791F1F;",
    }
    return colores.get(val, "")


# ─────────────────────────────────────────────────────────────
# INICIO DE LA APP — ESTADO DE SESIÓN
# ─────────────────────────────────────────────────────────────

# Carga los datos al iniciar (solo una vez por sesión)
if "df_registros" not in st.session_state:
    st.session_state.df_registros = cargar_datos()

# Controla si se muestra la tabla de registros
if "mostrar_tabla" not in st.session_state:
    st.session_state.mostrar_tabla = False

# Controla si el formulario fue enviado con éxito (para limpiarlo)
if "form_enviado" not in st.session_state:
    st.session_state.form_enviado = False

# ─────────────────────────────────────────────────────────────
# ENCABEZADO DE LA APP
# ─────────────────────────────────────────────────────────────

st.markdown("""
<div style="
    background: linear-gradient(135deg, #1E3A5F 0%, #185FA5 100%);
    padding: 1.5rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
">
    <h1 style="color: white; margin: 0; font-size: 1.6rem;">
         Sistema de Captura de Costos
    </h1>
    <p style="color: #B5D4F4; margin: 4px 0 0 0; font-size: 0.95rem;">
        Plancha de Acero S.A.C. — Registro de Órdenes de Producción
    </p>
</div>
""", unsafe_allow_html=True)

# Métricas rápidas en el encabezado
df_actual = st.session_state.df_registros
total_registros = len(df_actual)
registros_hoy = len(df_actual[df_actual["FechaProduccion"] == str(date.today())]) if total_registros > 0 else 0
criticos_hoy = len(df_actual[df_actual["Estado_Varianza"] == "Crítico"]) if total_registros > 0 else 0

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
col_m1.metric("Total registros", total_registros)
col_m2.metric("Registros hoy", registros_hoy)
col_m3.metric("Alertas críticas", criticos_hoy)
col_m4.metric("Archivo CSV", "Activo ✓" if os.path.exists(CSV_PATH) else "Nuevo")

st.divider()

# ─────────────────────────────────────────────────────────────
# LAYOUT PRINCIPAL: FORMULARIO (izquierda) | PANEL INFO (derecha)
# ─────────────────────────────────────────────────────────────

col_form, col_info = st.columns([3, 2], gap="large")

with col_form:
    st.subheader("Nueva Orden de Producción")

    # ── SECCIÓN 1: Datos generales de la orden ──────────────
    with st.expander("Datos generales", expanded=True):
        c1, c2 = st.columns(2)

        with c1:
            fecha_produccion = st.date_input(
                "Fecha de producción *",
                value=date.today(),
                max_value=date.today(),
                help="No se permiten fechas futuras.",
            )

        with c2:
            orden_sugerida = generar_orden_sugerida(st.session_state.df_registros)
            orden_id = st.text_input(
                "Número de orden *",
                value=orden_sugerida,
                help="Formato: OP001, OP002... El sistema sugiere el siguiente número.",
                placeholder="Ej: OP008",
            )

        c3, c4 = st.columns(2)
        with c3:
            planta = st.selectbox(
                "Planta *",
                options=["Arequipa", "Pisco"],
                help="Selecciona la planta donde se realizó la producción.",
            )

        with c4:
            producto = st.selectbox(
                "Producto *",
                options=list(COSTOS_ESTANDAR.keys()),
                help="Selecciona el producto fabricado en esta orden.",
            )

        cantidad_producida = st.number_input(
            "Cantidad producida (unidades) *",
            min_value=0,
            value=0,
            step=1,
            help="Número total de unidades producidas en esta orden.",
        )

    # ── SECCIÓN 2: Materia Prima ─────────────────────────────
    with st.expander("Consumo de Materia Prima", expanded=True):
        c5, c6 = st.columns(2)
        with c5:
            material_mp = st.selectbox(
                "Tipo de material *",
                options=MATERIALES_MP,
                help="Selecciona el tipo de materia prima utilizada.",
            )
        with c6:
            cantidad_kg = st.number_input(
                "Cantidad (kg) *",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.1f",
                help="Total de kg de materia prima consumidos.",
            )

        c7, c8 = st.columns(2)
        with c7:
            costo_unit_mp = st.number_input(
                "Costo por kg (S/) *",
                min_value=0.0,
                value=0.0,
                step=0.1,
                format="%.2f",
                help="Costo unitario de la materia prima por kilogramo.",
            )
        with c8:
            costo_total_mp = cantidad_kg * costo_unit_mp
            st.metric(
                "Costo total MP (calculado)",
                f"S/ {costo_total_mp:,.2f}",
                help="Calculado automáticamente: kg × costo/kg",
            )

    # ── SECCIÓN 3: Mano de Obra Directa ─────────────────────
    with st.expander("Mano de Obra Directa (MOD)", expanded=True):
        c9, c10 = st.columns(2)
        with c9:
            horas_trabajadas = st.number_input(
                "Horas trabajadas *",
                min_value=0.0,
                value=0.0,
                step=1.0,
                format="%.1f",
                help="Total de horas-hombre empleadas en la orden.",
            )
        with c10:
            costo_hora_mod = st.number_input(
                "Costo por hora (S/) *",
                min_value=0.0,
                value=0.0,
                step=0.5,
                format="%.2f",
                help="Costo por hora de mano de obra directa.",
            )

        costo_total_mod = horas_trabajadas * costo_hora_mod
        st.metric(
            "Costo total MOD (calculado)",
            f"S/ {costo_total_mod:,.2f}",
            help="Calculado automáticamente: horas × costo/hora",
        )

    # ── SECCIÓN 4: Costos Indirectos de Fabricación ─────────
    with st.expander("⚙️ Costos Indirectos de Fabricación (CIF)", expanded=True):
        c11, c12 = st.columns(2)
        with c11:
            tipo_cif = st.selectbox(
                "Tipo de CIF *",
                options=TIPOS_CIF,
                help="Categoría del costo indirecto de fabricación.",
            )
        with c12:
            costo_total_cif = st.number_input(
                "Monto total CIF (S/) *",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.2f",
                help="Monto total de costos indirectos asignados a esta orden.",
            )

    # ── SECCIÓN 5: Observaciones ─────────────────────────────
    with st.expander("💬 Observaciones (opcional)"):
        observaciones = st.text_area(
            "Notas adicionales",
            placeholder="Ej: Parada de máquina de 2 horas por mantenimiento preventivo. Lote de material con mayor impureza.",
            max_chars=500,
            help="Cualquier observación relevante para el análisis posterior.",
        )

    st.divider()

    # ── CÁLCULOS EN TIEMPO REAL ──────────────────────────────
    costo_total_real = costo_total_mp + costo_total_mod + costo_total_cif

    if cantidad_producida > 0:
        costo_real_por_unidad = costo_total_real / cantidad_producida
    else:
        costo_real_por_unidad = 0.0

    costo_std = COSTOS_ESTANDAR[producto]["total"]
    varianza_unit = costo_real_por_unidad - costo_std
    estado_varianza = clasificar_varianza(varianza_unit)

    precio_venta = PRECIO_VENTA[producto]
    ingresos = precio_venta * cantidad_producida
    utilidad_bruta = ingresos - costo_total_real
    margen_pct = (utilidad_bruta / ingresos * 100) if ingresos > 0 else 0.0

    # ── PANEL DE RESUMEN PREVIO AL GUARDADO ─────────────────
    st.subheader("📊 Resumen de la orden")

    col_r1, col_r2, col_r3 = st.columns(3)
    col_r1.metric("Costo total real", f"S/ {costo_total_real:,.2f}")
    col_r2.metric(
        "Costo real/unidad",
        f"S/ {costo_real_por_unidad:,.2f}",
        delta=f"{varianza_unit:+.2f} vs estándar",
        delta_color="inverse",  # negativo (eficiente) se muestra en verde
    )
    col_r3.metric("Margen bruto", f"{margen_pct:.1f}%")

    # Alerta visual de varianza
    if cantidad_producida > 0 and costo_total_real > 0:
        if estado_varianza == "Eficiente":
            st.success(f"Eficiente — el costo real está por debajo del estándar (S/ {costo_std}/u).")
        elif estado_varianza == "Alerta":
            st.warning(f"Alerta — el costo real supera ligeramente el estándar (S/ {costo_std}/u). Revisar.")
        else:
            st.error(
                f"Crítico — el costo real supera significativamente el estándar (S/ {costo_std}/u). Requiere justificación.")

    st.divider()

    # ── BOTÓN DE GUARDADO ────────────────────────────────────
    boton_guardar = st.button(
        "💾 Guardar registro",
        type="primary",
        use_container_width=True,
        help="Valida todos los campos y guarda el registro en el CSV.",
    )

    if boton_guardar:
        # Ejecutar validaciones
        errores = validar_formulario(
            fecha_produccion, orden_id, cantidad_producida,
            cantidad_kg, costo_unit_mp, horas_trabajadas,
            costo_hora_mod, costo_total_cif,
        )

        # Validación adicional: si el estado es Crítico, se requiere observación
        if estado_varianza == "Crítico" and not observaciones.strip():
            errores.append("El estado es Crítico. Debes ingresar una observación que justifique el sobrecosto.")

        # Validación de orden duplicada
        if not st.session_state.df_registros.empty:
            orden_upper = orden_id.strip().upper()
            ordenes_guardadas = st.session_state.df_registros["OrdenID"].str.upper().tolist()
            if orden_upper in ordenes_guardadas:
                errores.append(
                    f"Ya existe un registro con la orden {orden_id.strip().upper()}. Verifica el número de orden.")

        if errores:
            st.error("**Errores encontrados. Por favor corrige antes de guardar:**")
            for err in errores:
                st.write(err)
        else:
            # Construir el diccionario del registro
            nuevo_registro = {
                "FechaRegistro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "FechaProduccion": str(fecha_produccion),
                "OrdenID": orden_id.strip().upper(),
                "Planta": planta,
                "Producto": producto,
                "CantidadProducida": cantidad_producida,
                "Material_MP": material_mp,
                "CantidadKg_MP": round(cantidad_kg, 2),
                "CostoUnitario_MP": round(costo_unit_mp, 2),
                "CostoTotal_MP": round(costo_total_mp, 2),
                "HorasTrabajadas": round(horas_trabajadas, 1),
                "CostoHora_MOD": round(costo_hora_mod, 2),
                "CostoTotal_MOD": round(costo_total_mod, 2),
                "TipoCIF": tipo_cif,
                "CostoTotal_CIF": round(costo_total_cif, 2),
                "CostoTotal_Real": round(costo_total_real, 2),
                "CostoRealPorUnidad": round(costo_real_por_unidad, 2),
                "CostoStd_PorUnidad": costo_std,
                "Varianza_PorUnidad": round(varianza_unit, 2),
                "Estado_Varianza": estado_varianza,
                "Ingresos": round(ingresos, 2),
                "UtilidadBruta": round(utilidad_bruta, 2),
                "MargenBruto_Pct": round(margen_pct, 2),
                "Observaciones": observaciones.strip(),
            }

            # Guardar en CSV
            if guardar_registro(nuevo_registro):
                # Actualizar el DataFrame en sesión (sin recargar la página)
                nueva_fila = pd.DataFrame([nuevo_registro])
                st.session_state.df_registros = pd.concat(
                    [st.session_state.df_registros, nueva_fila],
                    ignore_index=True,
                )
                st.session_state.form_enviado = True
                st.success(f"Registro guardado correctamente — Orden {orden_id.strip().upper()} en {planta}.")
                st.balloons()
            else:
                st.error("No se pudo guardar el registro. Revisa los permisos del directorio.")

# ─────────────────────────────────────────────────────────────
# PANEL DERECHO: INFORMACIÓN Y REFERENCIA
# ─────────────────────────────────────────────────────────────

with col_info:
    st.subheader("📌 Referencia de costos estándar")

    # Tabla de costos estándar para que el operario compare
    df_std = pd.DataFrame([
        {
            "Producto": prod,
            "MP Std (S/u)": vals["mp"],
            "MOD Std (S/u)": vals["mod"],
            "CIF Std (S/u)": vals["cif"],
            "Total Std (S/u)": vals["total"],
        }
        for prod, vals in COSTOS_ESTANDAR.items()
    ])
    st.dataframe(df_std, use_container_width=True, hide_index=True)

    st.divider()

    # Indicador del producto seleccionado
    if producto:
        std_seleccionado = COSTOS_ESTANDAR[producto]
        st.markdown(f"**Producto seleccionado:** {producto}")
        st.caption(f"Precio de venta: S/ {PRECIO_VENTA[producto]}/u")

        col_i1, col_i2 = st.columns(2)
        col_i1.metric("Costo std total", f"S/ {std_seleccionado['total']}/u")
        col_i2.metric("Margen estándar esperado",
                      f"{((PRECIO_VENTA[producto] - std_seleccionado['total']) / PRECIO_VENTA[producto] * 100):.1f}%")

    st.divider()

    # Resumen rápido de la sesión actual
    st.subheader("📈 Últimos registros")

    df_vis = st.session_state.df_registros
    if df_vis.empty:
        st.info("Aún no hay registros guardados. Completa el formulario y presiona 'Guardar registro'.")
    else:
        # Mostrar los últimos 5 registros con columnas clave
        cols_resumen = ["OrdenID", "FechaProduccion", "Producto", "Planta",
                        "CostoTotal_Real", "Estado_Varianza", "MargenBruto_Pct"]
        df_resumen = df_vis[cols_resumen].tail(5).copy()
        df_resumen.columns = ["Orden", "Fecha", "Producto", "Planta",
                              "Costo Total (S/)", "Estado", "Margen %"]
        df_resumen = df_resumen.reset_index(drop=True)

        # Aplicar color condicional a la columna Estado
        st.dataframe(
            df_resumen.style.map(colorear_estado, subset=["Estado"]),
            use_container_width=True,
            hide_index=True,
        )

# ─────────────────────────────────────────────────────────────
# SECCIÓN INFERIOR: VER TODOS LOS DATOS + EXPORTAR
# ─────────────────────────────────────────────────────────────

st.divider()
st.subheader("🗄️ Gestión de datos")

col_btn1, col_btn2, col_btn3 = st.columns([2, 2, 3])

with col_btn1:
    if st.button("Ver todos los registros", use_container_width=True):
        st.session_state.mostrar_tabla = not st.session_state.mostrar_tabla

with col_btn2:
    if st.button("Ver registros de hoy", use_container_width=True):
        st.session_state.mostrar_tabla = True

with col_btn3:
    df_export = st.session_state.df_registros
    if not df_export.empty:
        excel_bytes = exportar_excel(df_export)
        st.download_button(
            label="Exportar a Excel (para Power BI)",
            data=excel_bytes,
            file_name=f"costos_plancha_acero_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Descarga todos los registros como archivo .xlsx listo para Power BI.",
        )
    else:
        st.button(
            "Exportar a Excel",
            disabled=True,
            use_container_width=True,
            help="No hay datos para exportar aún.",
        )

# Tabla completa de datos
if st.session_state.mostrar_tabla:
    df_tabla = st.session_state.df_registros

    if df_tabla.empty:
        st.info("No hay registros guardados todavía.")
    else:
        # Estadísticas agregadas encima de la tabla
        st.markdown("#### Resumen estadístico")
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        col_s1.metric("Total órdenes", len(df_tabla))
        col_s2.metric("Ingresos totales", f"S/ {df_tabla['Ingresos'].sum():,.0f}")
        col_s3.metric("Costo total", f"S/ {df_tabla['CostoTotal_Real'].sum():,.0f}")
        col_s4.metric(
            "Margen promedio",
            f"{df_tabla['MargenBruto_Pct'].mean():.1f}%",
        )

        # Filtros rápidos
        st.markdown("#### Filtros")
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filtro_planta = st.multiselect(
                "Planta",
                options=df_tabla["Planta"].unique().tolist(),
                default=df_tabla["Planta"].unique().tolist(),
            )
        with col_f2:
            filtro_producto = st.multiselect(
                "Producto",
                options=df_tabla["Producto"].unique().tolist(),
                default=df_tabla["Producto"].unique().tolist(),
            )
        with col_f3:
            filtro_estado = st.multiselect(
                "Estado varianza",
                options=df_tabla["Estado_Varianza"].unique().tolist(),
                default=df_tabla["Estado_Varianza"].unique().tolist(),
            )

        # Aplicar filtros
        mask = (
                df_tabla["Planta"].isin(filtro_planta) &
                df_tabla["Producto"].isin(filtro_producto) &
                df_tabla["Estado_Varianza"].isin(filtro_estado)
        )
        df_filtrado = df_tabla[mask].reset_index(drop=True)

        st.markdown(f"**{len(df_filtrado)} registro(s) encontrado(s)**")

        # Tabla con formato condicional
        columnas_mostrar = [
            "OrdenID", "FechaProduccion", "Producto", "Planta",
            "CantidadProducida", "CostoTotal_Real", "CostoRealPorUnidad",
            "CostoStd_PorUnidad", "Varianza_PorUnidad", "Estado_Varianza",
            "MargenBruto_Pct", "Observaciones"
        ]
        # Solo muestra columnas que existen en el DataFrame
        columnas_disponibles = [c for c in columnas_mostrar if c in df_filtrado.columns]

        st.dataframe(
            df_filtrado[columnas_disponibles].style.map(
                colorear_estado, subset=["Estado_Varianza"]
            ),
            use_container_width=True,
            hide_index=True,
        )

# ─────────────────────────────────────────────────────────────
# PIE DE PÁGINA
# ─────────────────────────────────────────────────────────────

st.divider()
st.caption(
    f"Plancha de Acero S.A.C. — Sistema de Captura de Costos v1.0 | "
    f"Datos guardados en: `{os.path.abspath(CSV_PATH)}` | "
    f"Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
)