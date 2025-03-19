import json
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font


def cargar_datos_asistencia(ruta_json):
    # Carga los datos desde un archivo JSON y retorna la lista de estudiantes
    with open(ruta_json, 'r', encoding='utf-8') as archivo:
        datos = json.load(archivo)
    return datos.get("estudiantes", [])  # Devuelve una lista vacía si no se encuentra la clave "estudiantes"


def crear_reporte_asistencia(nombre_archivo, ruta_json):
    # Inicializa el archivo Excel y la hoja activa
    wb = Workbook()
    ws = wb.active
    ws.title = "Asistencia"

    # Ajuste de columnas
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 50
    for col in range(3, 20):  # Columnas C a T para días de asistencia
        ws.column_dimensions[chr(64 + col)].width = 4
    ws.column_dimensions['T'].width = 8  # Incrementa el ancho de la columna T para que se lea claramente "TOTAL"

    # Estilos y bordes
    bold_font = Font(bold=True, size=13)
    header_font = Font(bold=True, size=12)
    normal_font = Font(size=12)
    center_alignment = Alignment(horizontal="center", vertical="center")
    left_alignment = Alignment(horizontal="left", vertical="center")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Encabezados principales
    encabezados_principales = [
        ("A1:T1", "UNIVERSIDAD NACIONAL DE LA AMAZONIA PERUANA"),
        ("A2:T2", "ESCUELA DE POSTGRADO"),
        ("A3:T3", "MAESTRIA EN EDUCACIÓN CON MENCION EN GESTION EDUCATIVA - 2024-I")
    ]

    for rango, texto in encabezados_principales:
        ws.merge_cells(rango)  # Fusiona celdas
        cell = ws[rango.split(":")[0]]
        cell.value = texto
        cell.font = bold_font
        cell.alignment = center_alignment

    # Datos generales (Asignatura, Docente, Fecha)
    datos_generales = [
        ("A4:B4", "ASIGNATURA", "C4:T4", "MDU-101: FILOSOFIA DE LA EDUCACION UNIVERSITARIA"),
        ("A5:B5", "DOCENTE", "C5:T5", "Dra. JULIANA SANCHEZ BABILONIA"),
        ("A6:B6", "FECHA", "C6:T6", "Del 08 al 30 de junio del 2024")
    ]

    for celdas1, texto1, celdas2, texto2 in datos_generales:
        ws.merge_cells(celdas1)
        ws.merge_cells(celdas2)
        ws[celdas1.split(":")[0]].value = texto1
        ws[celdas2.split(":")[0]].value = texto2
        ws[celdas1.split(":")[0]].font = header_font
        ws[celdas2.split(":")[0]].font = normal_font
        ws[celdas1.split(":")[0]].alignment = left_alignment
        ws[celdas2.split(":")[0]].alignment = left_alignment

    # Subtítulo de Control de Asistencia
    ws.merge_cells('C7:T7')
    ws['C7'] = "Control de Asistencia"
    ws['C7'].alignment = center_alignment
    ws['C7'].font = bold_font

    # Encabezado de la tabla de asistencia
    ws['B8'] = "APELLIDOS Y NOMBRES / secciones"
    ws['B8'].alignment = center_alignment
    ws['B8'].font = bold_font
    ws['B8'].border = thin_border

    for i in range(1, 18):  # Columnas de asistencia (1 al 17)
        cell = ws.cell(row=8, column=2 + i)
        cell.value = i
        cell.alignment = center_alignment
        cell.border = thin_border

    ws['T8'] = "TOTAL"
    ws['T8'].alignment = center_alignment
    ws['T8'].font = bold_font
    ws['T8'].border = thin_border

    # Cargar datos de estudiantes
    estudiantes = cargar_datos_asistencia(ruta_json)

    # Llenado de datos de asistencia
    for fila, estudiante in enumerate(estudiantes, start=9):
        # Número de estudiante
        ws.cell(row=fila, column=1, value=f"'{fila - 8:02}").alignment = center_alignment
        ws.cell(row=fila, column=1).border = thin_border

        # Nombre del estudiante
        ws.cell(row=fila, column=2, value=estudiante["nombre"]).border = thin_border

        # Registro de asistencia
        total_asistencias = 0
        for j, asistencia in enumerate(estudiante["asistencia"], start=1):
            cell = ws.cell(row=fila, column=2 + j)
            cell.value = asistencia
            cell.alignment = center_alignment
            cell.border = thin_border
            if asistencia == "P":  # Suma si está presente
                total_asistencias += 1

        # Total de asistencias por estudiante
        ws.cell(row=fila, column=20, value=total_asistencias).alignment = center_alignment
        ws.cell(row=fila, column=20).border = thin_border

    # Guardar el archivo Excel
    wb.save(nombre_archivo)


# Ejemplo de uso
crear_reporte_asistencia("Reporte_Asistencia_Mejorado.xlsx", "datos_asistencia.json")

#python exel.py