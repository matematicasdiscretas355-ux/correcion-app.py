import streamlit as st
import pandas as pd
import json
import io
import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ==========================================
# TEMA 1: CONFIGURACIÓN GENERAL
# ==========================================

st.set_page_config(
    page_title="Dulce Momento | Contabilidad II",
    page_icon="📊",
    layout="wide"
)

CARPETA = Path(__file__).resolve().parent

ARCHIVO = CARPETA / "productos_estandar_v2.json"
ARCHIVO_ANTERIOR = CARPETA / "productos_estandar.json"


# ==========================================
# TEMA 2: DISEÑO DE LA PÁGINA
# ==========================================

st.markdown("""
<style>

/* FONDO GENERAL */

.stApp {
    background-color: #E7E8EC;
    color: #30323A;
}

[data-testid="stAppViewContainer"] {
    background-color: #E7E8EC;
}

[data-testid="stHeader"] {
    background-color: #E7E8EC;
}

/* MENÚ LATERAL */

[data-testid="stSidebar"],
[data-testid="stSidebarContent"] {
    background-color: #D9DBE1;
}

/* TIPO DE LETRA */

.stApp h1,
.stApp h2,
.stApp h3,
.stApp h4,
.stApp p,
.stApp label,
.stApp li,
.stApp input,
.stApp textarea {
    font-family: "Segoe UI", Arial, sans-serif;
}

/* TÍTULOS */

.stApp h1,
.stApp h2,
.stApp h3 {
    color: #30323A !important;
    font-weight: 700;
}

/* TEXTOS */

.stApp p,
.stApp label {
    color: #30323A;
}

[data-testid="stWidgetLabel"] p {
    color: #30323A !important;
    font-weight: 600;
    font-size: 15px;
}

[data-testid="stCaptionContainer"] p {
    color: #62646D !important;
}

/* CAMPOS */

.stApp input,
.stApp textarea {
    color: #30323A !important;
    background-color: #FFFFFF !important;
}

/* BOTONES */

.stButton button,
.stDownloadButton button,
[data-testid="stFormSubmitButton"] button {
    background-color: #F4F4F6 !important;
    color: #30323A !important;
    border: 1px solid #B5B6C0 !important;
    border-radius: 10px;
    min-height: 43px;
}

.stButton button p,
.stDownloadButton button p,
[data-testid="stFormSubmitButton"] button p {
    color: #30323A !important;
    font-weight: 600;
}

.stButton button:hover,
.stDownloadButton button:hover {
    background-color: #E5E0EB !important;
    border-color: #766487 !important;
}

.stButton button[kind="primary"],
[data-testid="stFormSubmitButton"] button[kind="primary"] {
    background-color: #766487 !important;
    border-color: #766487 !important;
}

.stButton button[kind="primary"] p,
[data-testid="stFormSubmitButton"] button[kind="primary"] p {
    color: #FFFFFF !important;
}

/* TARJETAS DE RESULTADOS */

[data-testid="stMetric"] {
    background-color: #F4F4F6;
    padding: 16px;
    border-radius: 12px;
    border: 1px solid #D0D1D8;
}

[data-testid="stMetricLabel"] p {
    color: #50525B !important;
}

[data-testid="stMetricValue"] {
    color: #30323A !important;
}

/* FORMULARIOS */

[data-testid="stForm"] {
    background-color: #F4F4F6;
    border: 1px solid #D0D1D8;
    border-radius: 12px;
    padding: 20px;
}

/* NOTA INFORMATIVA */

.nota {
    background-color: #F4F4F6;
    border-left: 4px solid #766487;
    padding: 15px;
    border-radius: 8px;
    color: #30323A;
    margin-bottom: 15px;
}

/* PIE DE PÁGINA */

.pie {
    color: #777983;
    font-size: 14px;
    margin-top: 25px;
}

</style>
""", unsafe_allow_html=True)


# ==========================================
# TEMA 3: FUNCIONES GENERALES
# ==========================================

def numero(valor):
    try:
        resultado = float(valor)

        if math.isfinite(resultado) and resultado >= 0:
            return resultado

    except (TypeError, ValueError):
        pass

    return 0.0


def entero(valor, minimo=0):
    return max(minimo, int(numero(valor)))


def dinero(valor):
    return f"Q{valor:,.2f}"


def clasificar_variacion(valor):

    if abs(valor) < 0.000001:
        return "Sin variación"

    if valor > 0:
        return "Desfavorable"

    return "Favorable"


def limpiar_tabla(tabla, columnas):

    registros = []

    for fila in tabla.to_dict("records"):

        nueva = {}

        for columna in columnas:

            valor = fila.get(columna)

            if pd.isna(valor):
                valor = None

            nueva[columna] = valor

        registros.append(nueva)

    return registros


# ==========================================
# TEMA 4: PRODUCTOS
# ==========================================

def crear_ingrediente(nombre="", cantidad=0, unidad="g"):

    return {
        "Ingrediente": nombre,
        "Unidad": unidad,
        "Cantidad estándar": cantidad,
        "Precio estándar Q": 0.0,
        "Cantidad real": 0.0,
        "Precio real Q": 0.0,
        "Proveedor": ""
    }


def crear_producto(nombre, ejemplo=False):

    ingredientes = []
    descripcion = ""

    if ejemplo:

        descripcion = (
            "Pastel artesanal de vainilla, "
            "relleno de crema y fresas frescas. "
            "Decorado con crema batida y fresas naturales."
        )

        receta = [
            ("Harina", 300, "g"),
            ("Azúcar", 250, "g"),
            ("Huevos", 5, "unidad"),
            ("Leche", 200, "ml"),
            ("Aceite vegetal", 100, "ml"),
            ("Polvo de hornear", 12, "g"),
            ("Vainilla", 10, "ml"),
            ("Fresas frescas", 500, "g"),
            ("Crema para batir", 750, "ml"),
            ("Azúcar glass", 100, "g"),
            ("Base para pastel", 1, "unidad"),
            ("Caja para pastel", 1, "unidad")
        ]

        ingredientes = [
            crear_ingrediente(n, c, u)
            for n, c, u in receta
        ]

    return {
        "nombre": nombre,
        "descripcion": descripcion,
        "cantidad": 20,
        "cantidad_real": 20,
        "precio_venta": 0.0,
        "crecimiento": 10.0,
        "ingredientes": ingredientes,
        "mano_obra": [
            {
                "Actividad": "Preparación y horneado",
                "Horas estándar por unidad": 0.0,
                "Tarifa estándar Q": 0.0,
                "Horas reales totales": 0.0,
                "Tarifa real Q": 0.0
            },
            {
                "Actividad": "Decoración y empaque",
                "Horas estándar por unidad": 0.0,
                "Tarifa estándar Q": 0.0,
                "Horas reales totales": 0.0,
                "Tarifa real Q": 0.0
            }
        ],
        "cif": {
            "Variable estándar por unidad": 0.0,
            "Fijo estándar del período": 0.0,
            "Variable real total": 0.0,
            "Fijo real total": 0.0
        }
    }


# ==========================================
# TEMA 5: RESPALDO JSON
# ==========================================

def cargar_datos():

    for ruta in [ARCHIVO, ARCHIVO_ANTERIOR]:

        if ruta.exists():

            try:

                with open(ruta, "r", encoding="utf-8") as archivo:
                    datos = json.load(archivo)

                if (
                    isinstance(datos, dict)
                    and isinstance(datos.get("productos"), dict)
                    and datos["productos"]
                ):
                    return datos

            except (OSError, ValueError, TypeError):
                pass

    nombre = "Pastel de Fresas con Crema"

    return {
        "productos": {
            nombre: crear_producto(nombre, True)
        }
    }


def guardar_datos():

    contenido = json.dumps(
        st.session_state.catalogo,
        ensure_ascii=False,
        indent=2
    )

    temporal = ARCHIVO.with_suffix(".tmp")

    try:

        temporal.write_text(
            contenido,
            encoding="utf-8"
        )

        temporal.replace(ARCHIVO)

        return True

    except OSError as error:

        st.error(
            f"No se pudo guardar el respaldo JSON: {error}"
        )

        return False


# ==========================================
# TEMA 6: CÁLCULOS DE COSTOS ESTÁNDAR
# ==========================================

def calcular(p):

    cantidad = entero(p["cantidad"], 1)
    cantidad_real = entero(p["cantidad_real"], 1)

    precio_venta = numero(p["precio_venta"])

    # MATERIA PRIMA

    mp_unitaria = 0
    mp_estandar = 0
    mp_real = 0

    variacion_precio_mp = 0
    variacion_cantidad_mp = 0

    for fila in p["ingredientes"]:

        cantidad_estandar = numero(
            fila.get("Cantidad estándar")
        )

        precio_estandar = numero(
            fila.get("Precio estándar Q")
        )

        consumo_real = numero(
            fila.get("Cantidad real")
        )

        precio_real = numero(
            fila.get("Precio real Q")
        )

        cantidad_permitida = (
            cantidad_estandar * cantidad_real
        )

        mp_unitaria += (
            cantidad_estandar * precio_estandar
        )

        mp_estandar += (
            cantidad_permitida * precio_estandar
        )

        mp_real += (
            consumo_real * precio_real
        )

        variacion_precio_mp += (
            consumo_real *
            (precio_real - precio_estandar)
        )

        variacion_cantidad_mp += (
            (consumo_real - cantidad_permitida)
            * precio_estandar
        )

    # MANO DE OBRA

    mo_unitaria = 0
    mo_estandar = 0
    mo_real = 0

    variacion_tarifa_mo = 0
    variacion_eficiencia_mo = 0

    for fila in p["mano_obra"]:

        horas_estandar = numero(
            fila.get("Horas estándar por unidad")
        )

        tarifa_estandar = numero(
            fila.get("Tarifa estándar Q")
        )

        horas_reales = numero(
            fila.get("Horas reales totales")
        )

        tarifa_real = numero(
            fila.get("Tarifa real Q")
        )

        horas_permitidas = (
            horas_estandar * cantidad_real
        )

        mo_unitaria += (
            horas_estandar * tarifa_estandar
        )

        mo_estandar += (
            horas_permitidas * tarifa_estandar
        )

        mo_real += (
            horas_reales * tarifa_real
        )

        variacion_tarifa_mo += (
            horas_reales *
            (tarifa_real - tarifa_estandar)
        )

        variacion_eficiencia_mo += (
            (horas_reales - horas_permitidas)
            * tarifa_estandar
        )

    # COSTOS INDIRECTOS

    cif = p["cif"]

    variable_estandar = numero(
        cif["Variable estándar por unidad"]
    )

    fijo_estandar = numero(
        cif["Fijo estándar del período"]
    )

    variable_real = numero(
        cif["Variable real total"]
    )

    fijo_real = numero(
        cif["Fijo real total"]
    )

    cif_estandar = (
        cantidad_real * variable_estandar
        + fijo_estandar
    )

    cif_real = variable_real + fijo_real

    variacion_cif = cif_real - cif_estandar

    # COSTOS TOTALES

    costo_estandar = (
        mp_estandar
        + mo_estandar
        + cif_estandar
    )

    costo_real = (
        mp_real
        + mo_real
        + cif_real
    )

    costo_variable_unitario = (
        mp_unitaria
        + mo_unitaria
        + variable_estandar
    )

    costo_estandar_unitario = (
        costo_variable_unitario
        + fijo_estandar / cantidad
    )

    costo_real_unitario = (
        costo_real / cantidad_real
    )

    variacion_total = (
        costo_real - costo_estandar
    )

    # RENTABILIDAD

    ventas = cantidad_real * precio_venta

    utilidad = ventas - costo_real

    margen = (
        utilidad / ventas * 100
        if ventas > 0
        else None
    )

    # PUNTO DE EQUILIBRIO

    contribucion = (
        precio_venta - costo_variable_unitario
    )

    if contribucion > 0:

        equilibrio = math.ceil(
            fijo_estandar / contribucion
        )

    else:
        equilibrio = None

    # PRESUPUESTO

    crecimiento = numero(p["crecimiento"])

    cantidad_proyectada = math.ceil(
        cantidad * (1 + crecimiento / 100)
    )

    costo_proyectado = (
        cantidad_proyectada
        * costo_variable_unitario
        + fijo_estandar
    )

    return {
        "cantidad": cantidad,
        "cantidad_real": cantidad_real,
        "precio_venta": precio_venta,

        "mp_unitaria": mp_unitaria,
        "mp_estandar": mp_estandar,
        "mp_real": mp_real,

        "mo_unitaria": mo_unitaria,
        "mo_estandar": mo_estandar,
        "mo_real": mo_real,

        "cif_estandar": cif_estandar,
        "cif_real": cif_real,

        "variable_estandar": variable_estandar,
        "fijo_estandar": fijo_estandar,

        "costo_estandar": costo_estandar,
        "costo_real": costo_real,

        "costo_variable_unitario": costo_variable_unitario,
        "costo_estandar_unitario": costo_estandar_unitario,
        "costo_real_unitario": costo_real_unitario,

        "variacion_precio_mp": variacion_precio_mp,
        "variacion_cantidad_mp": variacion_cantidad_mp,
        "variacion_tarifa_mo": variacion_tarifa_mo,
        "variacion_eficiencia_mo": variacion_eficiencia_mo,
        "variacion_cif": variacion_cif,
        "variacion_total": variacion_total,

        "ventas": ventas,
        "utilidad": utilidad,
        "margen": margen,
        "equilibrio": equilibrio,

        "cantidad_proyectada": cantidad_proyectada,
        "costo_proyectado": costo_proyectado
    }


# ==========================================
# TEMA 7: EXPORTACIÓN A EXCEL
# ==========================================

def crear_excel(catalogo):

    libro = Workbook()

    resumen = libro.active
    resumen.title = "Resumen General"

    resumen.append([
        "Producto",
        "Producción real",
        "Costo estándar permitido",
        "Costo real",
        "Variación total",
        "Ventas",
        "Utilidad",
        "Margen",
        "Punto de equilibrio",
        "Producción proyectada"
    ])

    for indice, (nombre, p) in enumerate(
        catalogo["productos"].items(), 1
    ):

        codigo = f"P{indice:02d}"

        # DATOS GENERALES

        datos = libro.create_sheet(
            codigo + " Datos"
        )

        filas = [
            ["Concepto", "Valor"],
            ["Producto", nombre],
            ["Descripción", p.get("descripcion", "")],
            ["Producción presupuestada", p["cantidad"]],
            ["Producción real", p["cantidad_real"]],
            ["Precio de venta", p["precio_venta"]],
            ["Crecimiento", p["crecimiento"] / 100],
            [
                "CIF variable estándar unitario",
                numero(
                    p["cif"]["Variable estándar por unidad"]
                )
            ],
            [
                "CIF fijo estándar",
                numero(
                    p["cif"]["Fijo estándar del período"]
                )
            ],
            [
                "CIF variable real",
                numero(
                    p["cif"]["Variable real total"]
                )
            ],
            [
                "CIF fijo real",
                numero(
                    p["cif"]["Fijo real total"]
                )
            ]
        ]

        for fila in filas:
            datos.append(fila)

        datos["B7"].number_format = "0.00%"

        # MATERIA PRIMA

        mp = libro.create_sheet(
            codigo + " MP"
        )

        mp.append([
            "Ingrediente",
            "Unidad",
            "Cantidad estándar unitaria",
            "Precio estándar Q",
            "Cantidad real total",
            "Precio real Q",
            "Cantidad permitida",
            "Costo estándar",
            "Costo real",
            "Variación precio",
            "Variación cantidad"
        ])

        for i, fila in enumerate(
            p["ingredientes"], 2
        ):

            mp.append([
                fila.get("Ingrediente", ""),
                fila.get("Unidad", ""),
                numero(fila.get("Cantidad estándar")),
                numero(fila.get("Precio estándar Q")),
                numero(fila.get("Cantidad real")),
                numero(fila.get("Precio real Q")),

                f"=C{i}*'{datos.title}'!B5",
                f"=G{i}*D{i}",
                f"=E{i}*F{i}",
                f"=E{i}*(F{i}-D{i})",
                f"=(E{i}-G{i})*D{i}"
            ])

        fila_mp = mp.max_row + 1

        mp.cell(fila_mp, 1, "TOTALES")

        for columna in range(8, 12):

            letra = get_column_letter(columna)

            mp.cell(
                fila_mp,
                columna,
                f"=SUM({letra}2:{letra}{fila_mp-1})"
            )

        # MANO DE OBRA

        mo = libro.create_sheet(
            codigo + " MO"
        )

        mo.append([
            "Actividad",
            "Horas estándar unitarias",
            "Tarifa estándar Q",
            "Horas reales totales",
            "Tarifa real Q",
            "Horas permitidas",
            "Costo estándar",
            "Costo real",
            "Variación tarifa",
            "Variación eficiencia"
        ])

        for i, fila in enumerate(
            p["mano_obra"], 2
        ):

            mo.append([
                fila.get("Actividad", ""),
                numero(
                    fila.get("Horas estándar por unidad")
                ),
                numero(
                    fila.get("Tarifa estándar Q")
                ),
                numero(
                    fila.get("Horas reales totales")
                ),
                numero(
                    fila.get("Tarifa real Q")
                ),

                f"=B{i}*'{datos.title}'!B5",
                f"=F{i}*C{i}",
                f"=D{i}*E{i}",
                f"=D{i}*(E{i}-C{i})",
                f"=(D{i}-F{i})*C{i}"
            ])

        fila_mo = mo.max_row + 1

        mo.cell(fila_mo, 1, "TOTALES")

        for columna in range(7, 11):

            letra = get_column_letter(columna)

            mo.cell(
                fila_mo,
                columna,
                f"=SUM({letra}2:{letra}{fila_mo-1})"
            )

        # CIF

        cif = libro.create_sheet(
            codigo + " CIF"
        )

        filas_cif = [
            ["Concepto", "Valor Q"],

            [
                "CIF variable estándar",
                f"='{datos.title}'!B5*'{datos.title}'!B8"
            ],

            [
                "CIF fijo estándar",
                f"='{datos.title}'!B9"
            ],

            [
                "CIF estándar permitido",
                "=SUM(B2:B3)"
            ],

            [
                "CIF variable real",
                f"='{datos.title}'!B10"
            ],

            [
                "CIF fijo real",
                f"='{datos.title}'!B11"
            ],

            [
                "CIF real total",
                "=SUM(B5:B6)"
            ],

            [
                "Variación CIF",
                "=B7-B4"
            ]
        ]

        for fila in filas_cif:
            cif.append(fila)

        # RESULTADOS

        resultados = libro.create_sheet(
            codigo + " Resultados"
        )

        filas_resultados = [
            ["Concepto", "Valor"],

            [
                "Producción presupuestada",
                f"='{datos.title}'!B4"
            ],

            [
                "Producción real",
                f"='{datos.title}'!B5"
            ],

            [
                "Precio de venta",
                f"='{datos.title}'!B6"
            ],

            [
                "MP estándar permitido",
                f"='{mp.title}'!H{fila_mp}"
            ],

            [
                "MP real",
                f"='{mp.title}'!I{fila_mp}"
            ],

            [
                "MO estándar permitido",
                f"='{mo.title}'!G{fila_mo}"
            ],

            [
                "MO real",
                f"='{mo.title}'!H{fila_mo}"
            ],

            [
                "CIF estándar permitido",
                f"='{cif.title}'!B4"
            ],

            [
                "CIF real",
                f"='{cif.title}'!B7"
            ],

            [
                "Costo estándar permitido",
                "=B5+B7+B9"
            ],

            [
                "Costo real",
                "=B6+B8+B10"
            ],

            [
                "Variación MP precio",
                f"='{mp.title}'!J{fila_mp}"
            ],

            [
                "Variación MP cantidad",
                f"='{mp.title}'!K{fila_mp}"
            ],

            [
                "Variación MO tarifa",
                f"='{mo.title}'!I{fila_mo}"
            ],

            [
                "Variación MO eficiencia",
                f"='{mo.title}'!J{fila_mo}"
            ],

            [
                "Variación CIF",
                f"='{cif.title}'!B8"
            ],

            [
                "Variación total",
                "=SUM(B13:B17)"
            ],

            [
                "Comprobación de variaciones",
                "=B12-B11-B18"
            ],

            [
                "Costo real unitario",
                "=IF(B3>0,B12/B3,0)"
            ],

            [
                "MP estándar unitario",
                "=IF(B3>0,B5/B3,0)"
            ],

            [
                "MO estándar unitario",
                "=IF(B3>0,B7/B3,0)"
            ],

            [
                "Costo variable estándar unitario",
                f"=B21+B22+'{datos.title}'!B8"
            ],

            [
                "Costo estándar unitario",
                f"=B23+'{datos.title}'!B9/B2"
            ],

            [
                "Ventas reales",
                "=B3*B4"
            ],

            [
                "Utilidad real",
                "=B25-B12"
            ],

            [
                "Margen",
                "=IF(B25>0,B26/B25,0)"
            ],

            [
                "Contribución estándar",
                "=B4-B23"
            ],

            [
                "Punto de equilibrio",
                (
                    f'=IF(B28>0,ROUNDUP('
                    f"'{datos.title}'!B9/B28,0),"
                    '"No calculable")'
                )
            ]
        ]

        for fila in filas_resultados:
            resultados.append(fila)

        resultados["B27"].number_format = "0.00%"

        # PRESUPUESTO

        presupuesto = libro.create_sheet(
            codigo + " Presupuesto"
        )

        filas_presupuesto = [
            ["Concepto", "Inicial", "Proyectado"],

            [
                "Unidades",
                f"='{datos.title}'!B4",
                f"=ROUNDUP(B2*(1+'{datos.title}'!B7),0)"
            ],

            [
                "Materia prima",
                f"=B2*'{resultados.title}'!B21",
                f"=C2*'{resultados.title}'!B21"
            ],

            [
                "Mano de obra",
                f"=B2*'{resultados.title}'!B22",
                f"=C2*'{resultados.title}'!B22"
            ],

            [
                "CIF variable",
                f"=B2*'{datos.title}'!B8",
                f"=C2*'{datos.title}'!B8"
            ],

            [
                "CIF fijo",
                f"='{datos.title}'!B9",
                f"='{datos.title}'!B9"
            ],

            [
                "Costo total",
                "=SUM(B3:B6)",
                "=SUM(C3:C6)"
            ],

            [
                "Ventas",
                f"=B2*'{resultados.title}'!B4",
                f"=C2*'{resultados.title}'!B4"
            ],

            [
                "Utilidad",
                "=B8-B7",
                "=C8-C7"
            ]
        ]

        for fila in filas_presupuesto:
            presupuesto.append(fila)

        # RESUMEN GENERAL

        resumen.append([
            f"='{datos.title}'!B2",
            f"='{resultados.title}'!B3",
            f"='{resultados.title}'!B11",
            f"='{resultados.title}'!B12",
            f"='{resultados.title}'!B18",
            f"='{resultados.title}'!B25",
            f"='{resultados.title}'!B26",
            f"='{resultados.title}'!B27",
            f"='{resultados.title}'!B29",
            f"='{presupuesto.title}'!C2"
        ])

        resumen.cell(
            resumen.max_row, 8
        ).number_format = "0.00%"

    # FORMATO GENERAL DEL EXCEL

    for hoja in libro:

        hoja.freeze_panes = "A2"
        hoja.sheet_view.showGridLines = False
        hoja.row_dimensions[1].height = 35

        for celda in hoja[1]:

            celda.fill = PatternFill(
                "solid",
                fgColor="766487"
            )

            celda.font = Font(
                color="FFFFFF",
                bold=True
            )

            celda.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

        for columna in range(
            1, hoja.max_column + 1
        ):

            letra = get_column_letter(columna)

            hoja.column_dimensions[letra].width = (
                38 if columna == 1 else 25
            )

        for fila in hoja.iter_rows(min_row=2):

            for celda in fila:

                if celda.row % 2 == 0:

                    celda.fill = PatternFill(
                        "solid",
                        fgColor="F0F0F3"
                    )

    libro.calculation.fullCalcOnLoad = True
    libro.calculation.forceFullCalc = True

    salida = io.BytesIO()

    libro.save(salida)

    return salida.getvalue()


# ==========================================
# TEMA 8: INICIAR EL SISTEMA
# ==========================================

if "catalogo" not in st.session_state:

    st.session_state.catalogo = cargar_datos()


productos = st.session_state.catalogo["productos"]

if (
    "seleccion" not in st.session_state
    or st.session_state.seleccion not in productos
):

    st.session_state.seleccion = next(
        iter(productos)
    )


# ==========================================
# TEMA 9: MENÚ LATERAL
# ==========================================

with st.sidebar:

    st.header("Catálogo de productos")

    nombres = list(productos)

    seleccionado = st.selectbox(
        "Seleccionar producto",
        nombres,
        index=nombres.index(
            st.session_state.seleccion
        ),
        key="selector_productos"
    )

    if seleccionado != st.session_state.seleccion:

        st.session_state.seleccion = seleccionado

        st.rerun()

    st.divider()

    # AGREGAR PRODUCTO

    with st.form(
        "formulario_nuevo",
        clear_on_submit=True
    ):

        nuevo_nombre = st.text_input(
            "Nombre del nuevo producto"
        )

        agregar = st.form_submit_button(
            "Agregar producto",
            use_container_width=True
        )

    if agregar:

        nuevo_nombre = nuevo_nombre.strip()

        if not nuevo_nombre:

            st.warning(
                "Debes ingresar el nombre del producto."
            )

        elif nuevo_nombre in productos:

            st.warning(
                "Ese producto ya existe."
            )

        else:

            productos[nuevo_nombre] = crear_producto(
                nuevo_nombre
            )

            st.session_state.seleccion = nuevo_nombre

            if guardar_datos():
                st.rerun()

    st.divider()

    # ELIMINAR PRODUCTO

    st.subheader("Administrar productos")

    confirmar = st.checkbox(
        "Confirmar eliminación",
        key="confirmar_eliminar"
    )

    if st.button(
        "Eliminar producto seleccionado",
        disabled=(
            not confirmar
            or len(productos) <= 1
        ),
        use_container_width=True
    ):

        del productos[st.session_state.seleccion]

        st.session_state.seleccion = next(
            iter(productos)
        )

        if guardar_datos():
            st.rerun()

    st.divider()

    st.caption(
        "Respaldo JSON automático activado."
    )


# ==========================================
# TEMA 10: PRODUCTO SELECCIONADO
# ==========================================

nombre = st.session_state.seleccion

p = productos[nombre]

st.title("Dulce Momento, Gran Sentimiento")

st.caption(
    "Sistema de Gestión de Costos Estándar "
    "y Control de Producción"
)

st.divider()

st.header(nombre)

st.markdown("""
<div class="nota">
Ingresa los datos del producto y presiona
<strong>Guardar datos del producto</strong>
cuando termines. Los resultados se actualizarán
automáticamente y se guardará el respaldo JSON.
</div>
""", unsafe_allow_html=True)


# ==========================================
# TEMA 11: FORMULARIO DE DATOS
# ==========================================

# IMPORTANTE:
# TODOS LOS CAMPOS ESTÁN DENTRO DEL FORMULARIO.
# STREAMLIT NO RECARGARÁ LA PÁGINA
# MIENTRAS SE ESCRIBEN LOS NÚMEROS.

COLUMNAS_MP = [
    "Ingrediente",
    "Unidad",
    "Cantidad estándar",
    "Precio estándar Q",
    "Cantidad real",
    "Precio real Q",
    "Proveedor"
]

COLUMNAS_MO = [
    "Actividad",
    "Horas estándar por unidad",
    "Tarifa estándar Q",
    "Horas reales totales",
    "Tarifa real Q"
]

with st.form(
    key=f"formulario_producto_{nombre}",
    clear_on_submit=False
):

    descripcion = st.text_area(
        "Descripción del producto",
        value=p.get("descripcion", ""),
        height=100
    )

    # CONFIGURACIÓN

    st.subheader("Configuración de producción")

    columna1, columna2 = st.columns(2)

    with columna1:

        cantidad = st.number_input(
            "Producción presupuestada",
            min_value=1,
            value=entero(p["cantidad"], 1),
            step=1
        )

        precio_venta = st.number_input(
            "Precio de venta unitario (Q)",
            min_value=0.0,
            value=numero(p["precio_venta"]),
            step=1.0,
            format="%.2f"
        )

    with columna2:

        cantidad_real = st.number_input(
            "Producción real terminada",
            min_value=1,
            value=entero(p["cantidad_real"], 1),
            step=1
        )

        crecimiento = st.number_input(
            "Crecimiento proyectado (%)",
            min_value=0.0,
            value=numero(p["crecimiento"]),
            step=1.0,
            format="%.2f"
        )

    # MATERIA PRIMA

    st.header("1. Materia prima directa")

    st.write(
        "Ingresa los ingredientes, las cantidades "
        "estándar, los precios y los consumos reales."
    )

    tabla_mp = st.data_editor(
        pd.DataFrame(
            p["ingredientes"],
            columns=COLUMNAS_MP
        ),
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key=f"tabla_mp_{nombre}",
        column_config={
            "Cantidad estándar":
                st.column_config.NumberColumn(
                    "Cantidad estándar",
                    min_value=0.0,
                    format="%.4f"
                ),

            "Precio estándar Q":
                st.column_config.NumberColumn(
                    "Precio estándar Q",
                    min_value=0.0,
                    format="%.4f"
                ),

            "Cantidad real":
                st.column_config.NumberColumn(
                    "Cantidad real",
                    min_value=0.0,
                    format="%.4f"
                ),

            "Precio real Q":
                st.column_config.NumberColumn(
                    "Precio real Q",
                    min_value=0.0,
                    format="%.4f"
                )
        }
    )

    st.caption(
        "Las cantidades estándar corresponden "
        "a una unidad del producto. Las cantidades "
        "reales corresponden al consumo total."
    )

    # MANO DE OBRA

    st.header("2. Mano de obra directa")

    tabla_mo = st.data_editor(
        pd.DataFrame(
            p["mano_obra"],
            columns=COLUMNAS_MO
        ),
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key=f"tabla_mo_{nombre}",
        column_config={
            "Horas estándar por unidad":
                st.column_config.NumberColumn(
                    "Horas estándar por unidad",
                    min_value=0.0,
                    format="%.4f"
                ),

            "Tarifa estándar Q":
                st.column_config.NumberColumn(
                    "Tarifa estándar Q",
                    min_value=0.0,
                    format="%.4f"
                ),

            "Horas reales totales":
                st.column_config.NumberColumn(
                    "Horas reales totales",
                    min_value=0.0,
                    format="%.4f"
                ),

            "Tarifa real Q":
                st.column_config.NumberColumn(
                    "Tarifa real Q",
                    min_value=0.0,
                    format="%.4f"
                )
        }
    )

    # COSTOS INDIRECTOS

    st.header("3. Costos indirectos de fabricación")

    cif = p["cif"]

    columna1, columna2 = st.columns(2)

    with columna1:

        cif_variable_estandar = st.number_input(
            "CIF variable estándar por unidad (Q)",
            min_value=0.0,
            value=numero(
                cif["Variable estándar por unidad"]
            ),
            step=1.0,
            format="%.2f"
        )

        cif_variable_real = st.number_input(
            "CIF variable real total (Q)",
            min_value=0.0,
            value=numero(
                cif["Variable real total"]
            ),
            step=1.0,
            format="%.2f"
        )

    with columna2:

        cif_fijo_estandar = st.number_input(
            "CIF fijo estándar del período (Q)",
            min_value=0.0,
            value=numero(
                cif["Fijo estándar del período"]
            ),
            step=1.0,
            format="%.2f"
        )

        cif_fijo_real = st.number_input(
            "CIF fijo real total (Q)",
            min_value=0.0,
            value=numero(
                cif["Fijo real total"]
            ),
            step=1.0,
            format="%.2f"
        )

    st.divider()

    # BOTÓN DE GUARDAR

    guardar_producto = st.form_submit_button(
        "Guardar datos del producto",
        type="primary",
        use_container_width=True
    )


# ==========================================
# TEMA 12: ACTUALIZAR DATOS
# ==========================================

if guardar_producto:

    # DATOS GENERALES

    p["descripcion"] = descripcion
    p["cantidad"] = cantidad
    p["cantidad_real"] = cantidad_real
    p["precio_venta"] = precio_venta
    p["crecimiento"] = crecimiento

    # MATERIA PRIMA

    p["ingredientes"] = limpiar_tabla(
        tabla_mp,
        COLUMNAS_MP
    )

    # MANO DE OBRA

    p["mano_obra"] = limpiar_tabla(
        tabla_mo,
        COLUMNAS_MO
    )

    # COSTOS INDIRECTOS

    p["cif"] = {
        "Variable estándar por unidad":
            cif_variable_estandar,

        "Fijo estándar del período":
            cif_fijo_estandar,

        "Variable real total":
            cif_variable_real,

        "Fijo real total":
            cif_fijo_real
    }

    # RESPALDO AUTOMÁTICO

    if guardar_datos():

        st.success(
            "Datos guardados correctamente. "
            "El respaldo JSON se actualizó."
        )


# ==========================================
# TEMA 13: RESULTADOS CONTABLES
# ==========================================

r = calcular(p)

st.header("4. Resultados contables")

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "Costo estándar unitario",
    dinero(r["costo_estandar_unitario"])
)

col2.metric(
    "Costo real unitario",
    dinero(r["costo_real_unitario"])
)

col3.metric(
    "Costo estándar permitido",
    dinero(r["costo_estandar"])
)

col4.metric(
    "Costo real total",
    dinero(r["costo_real"])
)

resumen = pd.DataFrame({
    "Concepto": [
        "Materia prima estándar permitida",
        "Materia prima real",
        "Mano de obra estándar permitida",
        "Mano de obra real",
        "CIF estándar permitido",
        "CIF real",
        "Costo estándar permitido",
        "Costo real total"
    ],

    "Monto Q": [
        r["mp_estandar"],
        r["mp_real"],
        r["mo_estandar"],
        r["mo_real"],
        r["cif_estandar"],
        r["cif_real"],
        r["costo_estandar"],
        r["costo_real"]
    ]
})

st.dataframe(
    resumen,
    hide_index=True,
    use_container_width=True
)


# ==========================================
# TEMA 14: ANÁLISIS DE VARIACIONES
# ==========================================

st.header("5. Análisis de variaciones")

variaciones = pd.DataFrame([
    [
        "Materia prima - Precio",
        r["variacion_precio_mp"]
    ],
    [
        "Materia prima - Cantidad",
        r["variacion_cantidad_mp"]
    ],
    [
        "Mano de obra - Tarifa",
        r["variacion_tarifa_mo"]
    ],
    [
        "Mano de obra - Eficiencia",
        r["variacion_eficiencia_mo"]
    ],
    [
        "CIF - Variación global",
        r["variacion_cif"]
    ],
    [
        "VARIACIÓN TOTAL",
        r["variacion_total"]
    ]
], columns=[
    "Concepto",
    "Variación Q"
])

variaciones["Clasificación"] = (
    variaciones["Variación Q"].apply(
        clasificar_variacion
    )
)

st.dataframe(
    variaciones,
    hide_index=True,
    use_container_width=True
)

if r["variacion_total"] > 0:

    st.warning(
        "Variación desfavorable: "
        "el costo real supera al estándar."
    )

elif r["variacion_total"] < 0:

    st.success(
        "Variación favorable: "
        "el costo real es menor al estándar."
    )

else:

    st.info(
        "No existe variación total."
    )


# ==========================================
# TEMA 15: RENTABILIDAD
# ==========================================

st.header("6. Rentabilidad")

col1, col2, col3 = st.columns(3)

col1.metric(
    "Ventas",
    dinero(r["ventas"])
)

col2.metric(
    "Utilidad",
    dinero(r["utilidad"])
)

col3.metric(
    "Margen sobre ventas",
    (
        f'{r["margen"]:.2f}%'
        if r["margen"] is not None
        else "No calculable"
    )
)


# ==========================================
# TEMA 16: PUNTO DE EQUILIBRIO
# ==========================================

st.header("7. Punto de equilibrio")

if r["equilibrio"] is not None:

    st.success(
        f'Se necesitan vender {r["equilibrio"]} '
        "unidades para cubrir los costos estándar."
    )

else:

    st.warning(
        "El precio de venta debe superar "
        "el costo variable estándar unitario."
    )


# ==========================================
# TEMA 17: PRESUPUESTO
# ==========================================

st.header("8. Presupuesto de producción")

costo_inicial = (
    r["cantidad"]
    * r["costo_variable_unitario"]
    + r["fijo_estandar"]
)

ventas_iniciales = (
    r["cantidad"] * r["precio_venta"]
)

ventas_proyectadas = (
    r["cantidad_proyectada"]
    * r["precio_venta"]
)

presupuesto = pd.DataFrame({
    "Concepto": [
        "Unidades",
        "Costo estándar total Q",
        "Ventas Q",
        "Utilidad estándar Q"
    ],

    "Inicial": [
        r["cantidad"],
        costo_inicial,
        ventas_iniciales,
        ventas_iniciales - costo_inicial
    ],

    "Proyectado": [
        r["cantidad_proyectada"],
        r["costo_proyectado"],
        ventas_proyectadas,
        ventas_proyectadas - r["costo_proyectado"]
    ]
})

st.dataframe(
    presupuesto,
    hide_index=True,
    use_container_width=True
)

st.subheader("Comparación de costos")

grafica = pd.DataFrame({
    "Costo Q": {
        "Estándar permitido":
            r["costo_estandar"],

        "Real":
            r["costo_real"]
    }
})

st.bar_chart(grafica)


# ==========================================
# TEMA 18: EXPORTACIÓN
# ==========================================

st.header("9. Exportar información")

st.download_button(
    label="Descargar Excel de todos los productos",

    data=crear_excel(
        st.session_state.catalogo
    ),

    file_name="Dulce_Momento_Costos_Estandar.xlsx",

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),

    use_container_width=True
)

st.caption(
    "Los datos confirmados se guardan "
    "automáticamente en el respaldo JSON."
)

st.markdown("""
<div class="pie">
Dulce Momento, Gran Sentimiento | Contabilidad II
</div>
""", unsafe_allow_html=True)
