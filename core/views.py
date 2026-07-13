import openpyxl
import pandas as pd
from datetime import date, datetime
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db import transaction, models
from django.db.models import Sum 
from django.http import HttpResponse
from django.template.loader import get_template
from xhtml2pdf import pisa  # Requiere: pip install xhtml2pdf

# Importaciones para el diseño y estilizado profesional del Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date,datetime
# Importaciones de tu App (Modelos y Formularios)
from .models import PlanCaptura, ProgramaReal, CursoExcel, Capacitacion, Curso, CargaSTPS
from .forms import CargaSTPSForm, PlanCapturaForm, ProgramaRealForm, CursoExcelForm, CapacitacionForm


# ==========================================
# FUNCIÓN AUXILIAR: ESTILIZAR REPORTES EXCEL (TODOS LOS CAMPOS DE LA STPS)
# ==========================================
def generar_excel_estilizado(registros_queryset):
    """
    Función interna para unificar el diseño visual de las descargas 
    de Excel de la STPS utilizando los 21 campos exactos de tus especificaciones.
    """
    if registros_queryset.exists():
        df_exportar = pd.DataFrame(list(registros_queryset))
        
        # Mapeo y orden exacto de campos del modelo hacia las columnas del Excel final
        columnas_modelo = [
            'curp', 'nombre', 'primer_apellido', 'segundo_apellido', 'clave_estado', 
            'clave_municipio', 'clave_ocupacion', 'clave_niv_estudio', 'clave_doc_probatorio', 
            'clave_institucion', 'clave_curso', 'nombre_curso', 'clave_area_tematica', 
            'duracion', 'fec_inicio', 'fec_termino', 'clave_tip_agent', 'rfc_agente_stps', 
            'clave_modalidad', 'clave_capacitacion', 'clave_establec'
        ]
        
        # Filtrar solo las columnas que existan en el DataFrame para evitar KeyErrors
        df_exportar = df_exportar[[col for col in columnas_modelo if col in df_exportar.columns]]
        
        # Renombrar las cabeceras idéntico a las plantillas oficiales de Excel
        df_exportar.columns = [
            'CURP', 'NOMBRE', 'PRIMER APELLIDO', 'SEGUNDO APELLIDO', 'CLAVE ESTADO', 
            'CLAVE MUNICIPIO', 'CLAVE OCUPACION', 'CLAVE NIV ESTUDIO', 'CLAVE DOC PROBATORI', 
            'CLAVE INSTITUC', 'CLAVE CURSO', 'NOMBRE CURSO', 'CLAVE AREA TEMATICA', 
            'DURACIO', 'FEC INICIO', 'FEC TERMINO', 'CLAVE TIP AGENT', 'RFC AGENTE STPS', 
            'CLAVE MODALIDAD', 'CLAVE CAPACITACIO', 'CLAVE ESTABLEC'
        ]
    else:
        df_exportar = pd.DataFrame(columns=[
            'CURP', 'NOMBRE', 'PRIMER APELLIDO', 'SEGUNDO APELLIDO', 'CLAVE ESTADO', 
            'CLAVE MUNICIPIO', 'CLAVE OCUPACION', 'CLAVE NIV ESTUDIO', 'CLAVE DOC PROBATORI', 
            'CLAVE INSTITUC', 'CLAVE CURSO', 'NOMBRE CURSO', 'CLAVE AREA TEMATICA', 
            'DURACIO', 'FEC INICIO', 'FEC TERMINO', 'CLAVE TIP AGENT', 'RFC AGENTE STPS', 
            'CLAVE MODALIDAD', 'CLAVE CAPACITACIO', 'CLAVE ESTABLEC'
        ])
        
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=Reporte_Consolidado_STPS.xlsx'
    
    with pd.ExcelWriter(response, engine='openpyxl') as writer:
        df_exportar.to_excel(writer, index=False, sheet_name='Base STPS')
        workbook = writer.book
        worksheet = writer.sheets['Base STPS']
        
        # Definición de Estilos Visuales profesionales
        fuente_titulo = Font(name='Segoe UI', size=14, bold=True, color='0891B2')
        fuente_cabecera = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
        fuente_datos = Font(name='Segoe UI', size=10, color='1E293B')
        
        fill_cabecera = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
        fill_cebra = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
        
        alineacion_centro = Alignment(horizontal='center', vertical='center')
        alineacion_izq = Alignment(horizontal='left', vertical='center')
        
        borde_delgado = Border(
            left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
        )
        
        # Insertar renglón de título corporativo
        worksheet.insert_rows(1, 2)
        worksheet['A1'] = "REPORTE CONSOLIDADO DE REGISTROS - CATÁLOGO STPS"
        worksheet['A1'].font = fuente_titulo
        worksheet.row_dimensions[1].height = 25
        
        # Diseñar Fila de Cabecera (Fila 3)
        worksheet.row_dimensions[3].height = 28
        for cell in worksheet[3]:
            cell.font = fuente_cabecera
            cell.fill = fill_cabecera
            cell.alignment = alineacion_centro
            cell.border = borde_delgado
            
        # Diseñar Filas de Datos (Fila 4 en adelante)
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=4, max_row=worksheet.max_row), start=4):
            worksheet.row_dimensions[row_idx].height = 20
            for cell in row:
                cell.font = fuente_datos
                cell.border = borde_delgado
                
                # Columnas de claves, identificadores y fechas van centradas
                if cell.column_letter in ['A', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U']:
                    cell.alignment = alineacion_centro
                else:
                    cell.alignment = alineacion_izq
                
                if row_idx % 2 == 0:
                    cell.fill = fill_cebra
                    
        # Auto-ajustar ancho
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 14)
                
    return response


# ==========================================
# 1. VISTA DE CARGA MASIVA STPS (EXCEL Y CSV MS-DOS)
# ==========================================
@login_required
@transaction.atomic
def cargar_stps(request):
    descargar_url = None
    columnas = None
    datos_filas = None
    filas_count = 0

    if 'descargar' in request.GET:
        descargar_url = "?descargar=true"

    if request.method == 'POST' and request.FILES.get('archivo_excel'):
        excel_file = request.FILES['archivo_excel']
        accion = request.POST.get('accion') 
        nombre_archivo = excel_file.name.lower()
        
        try:
            # Lectura condicional según extensión (Soporta Excel y CSV MS-DOS/Latin-1)
            if nombre_archivo.endswith('.csv'):
                try:
                    df = pd.read_csv(excel_file, dtype=str, encoding='latin-1', sep=None, engine='python')
                except Exception:
                    excel_file.seek(0)
                    df = pd.read_csv(excel_file, dtype=str, encoding='utf-8', sep=None, engine='python')
            elif nombre_archivo.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(excel_file, dtype=str)
            else:
                messages.error(request, "Formato de archivo no admitido. Sube un archivo .xlsx, .xls o .csv")
                return redirect('cargar_stps')
            
            df.columns = [str(c).strip().upper() for c in df.columns]
            df = df.fillna("")
            
            if accion == 'previsualizar':
                columnas = df.columns.tolist()
                datos_filas = df.head(15).values.tolist()
                filas_count = len(df)
                messages.info(request, "Mostrando una vista previa del formato STPS.")

            elif accion == 'guardar':
                registros_nuevos = 0
                
                if df.empty:
                    messages.error(request, "El archivo está vacío o no se pudo leer ninguna fila.")
                else:
                    for index, row in df.iterrows():
                        def limpiar_dato_stps(nombre_principal, alternativas=[]):
                            val = row.get(nombre_principal, "")
                            if val == "":
                                for alt in alternativas:
                                    val = row.get(alt, "")
                                    if val != "":
                                        break
                            
                            val_str = str(val).strip()
                            if val_str.endswith('.0'):
                                val_str = val_str[:-2]
                            return val_str

                        CargaSTPS.objects.create(
                            curp=limpiar_dato_stps('CURP'),
                            nombre=limpiar_dato_stps('NOMBRE'),
                            primer_apellido=limpiar_dato_stps('PRIMER APELLIDO'),
                            segundo_apellido=limpiar_dato_stps('SEGUNDO APELLIDO'),
                            clave_estado=limpiar_dato_stps('CLAVE ESTADO'),
                            clave_municipio=limpiar_dato_stps('CLAVE MUNICIPIO'),
                            clave_ocupacion=limpiar_dato_stps('CLAVE OCUPACION'),
                            clave_curso=limpiar_dato_stps('CLAVE CURSO'),
                            nombre_curso=limpiar_dato_stps('NOMBRE CURSO'),
                            clave_area_tematica=limpiar_dato_stps('CLAVE AREA TEMATICA'),
                            duracion=limpiar_dato_stps('DURACION', ['DURACIO']), 
                            fec_inicio=limpiar_dato_stps('FEC INICIO', ['FECHA INICIO']),
                            fec_termino=limpiar_dato_stps('FEC TERMINO', ['FECHA TERMINO']),
                            rfc_agente_stps=limpiar_dato_stps('RFC AGENTE STPS'),
                            clave_modalidad=limpiar_dato_stps('CLAVE MODALIDAD'),
                            clave_niv_estudio=limpiar_dato_stps('CLAVE NIV ESTUDIOS', ['CLAVE NIV ESTUDIO']),
                            clave_doc_probatorio=limpiar_dato_stps('CLAVE DOC PROBATORIO', ['CLAVE DOC PROBATORI']),
                            clave_institucion=limpiar_dato_stps('CLAVE INSTITUCION', ['CLAVE INSTITUC']),
                            clave_tip_agent=limpiar_dato_stps('CLAVE TIP AGENTE', ['CLAVE TIP AGENT']),
                            clave_capacitacion=limpiar_dato_stps('CLAVE CAPACITACION', ['CLAVE CAPACITACIO']),
                            clave_establec=limpiar_dato_stps('CLAVE ESTABLECIMIENTO', ['CLAVE ESTABLEC']),
                        )
                        registros_nuevos += 1
                    
                    messages.success(request, f"¡Éxito! Se cargaron {registros_nuevos} registros. Archivo validado correctamente.")
                    return redirect(f"{request.path}?descargar=true")

        except Exception as e:
            messages.error(request, f"Error al procesar el archivo STPS: {e}")

    if 'descargar' in request.GET:
        registros = CargaSTPS.objects.all().values(
            'curp', 'nombre', 'primer_apellido', 'segundo_apellido', 'clave_estado', 
            'clave_municipio', 'clave_ocupacion', 'clave_niv_estudio', 'clave_doc_probatorio', 
            'clave_institucion', 'clave_curso', 'nombre_curso', 'clave_area_tematica', 
            'duracion', 'fec_inicio', 'fec_termino', 'clave_tip_agent', 'rfc_agente_stps', 
            'clave_modalidad', 'clave_capacitacion', 'clave_establec'
        )
        return generar_excel_estilizado(registros)

    return render(request, 'cargar_stps.html', {
        'descargar_url': descargar_url, 
        'columnas': columnas, 
        'datos_filas': datos_filas, 
        'filas_count': filas_count
    })


# ==========================================
# 2. VISTA DE CAPTURA MANUAL STPS
# ==========================================
@login_required
@transaction.atomic
def captura_manual_stps(request):
    if 'descargar' in request.GET:
        registros = CargaSTPS.objects.all().values(
            'curp', 'nombre', 'primer_apellido', 'segundo_apellido', 'clave_estado', 
            'clave_municipio', 'clave_ocupacion', 'clave_niv_estudio', 'clave_doc_probatorio', 
            'clave_institucion', 'clave_curso', 'nombre_curso', 'clave_area_tematica', 
            'duracion', 'fec_inicio', 'fec_termino', 'clave_tip_agent', 'rfc_agente_stps', 
            'clave_modalidad', 'clave_capacitacion', 'clave_establec'
        )
        return generar_excel_estilizado(registros)

    if request.method == 'POST':
        form = CargaSTPSForm(request.POST)
        if form.is_valid():
            nuevo_registro = form.save()
            messages.success(request, f"¡Éxito! Se registró manualmente a {nuevo_registro.nombre} en el catálogo STPS.")
            return redirect('captura_manual_stps')
    else:
        form = CargaSTPSForm()

    return render(request, 'captura_manual_stps.html', {'form': form})


# ==========================================
# 3. VISTAS DE EDICIÓN Y LISTADO STPS
# ==========================================
@login_required
def lista_modificar_stps(request):
    registros = CargaSTPS.objects.all().order_by('-id')
    return render(request, 'lista_modificar_stps.html', {'registros': registros})


@login_required
@transaction.atomic
def editar_registro_stps(request, pk):
    registro = get_object_or_404(CargaSTPS, pk=pk)
    if request.method == 'POST':
        form = CargaSTPSForm(request.POST, instance=registro)
        if form.is_valid():
            form.save()
            messages.success(request, f"¡Registro de {registro.nombre} actualizado con éxito!")
            return redirect('lista_modificar_stps')
    else:
        form = CargaSTPSForm(instance=registro)
    return render(request, 'editar_registro_stps.html', {'form': form, 'registro': registro})


# ==========================================
# 4. COMPLEMENTOS Y DASHBOARD GENERAL
# ==========================================
@login_required
def dashboard(request):
    return render(request, 'dashboard.html')


@login_required
def lista_modificar_plan(request):
    registros = PlanCaptura.objects.all()
    return render(request, 'lista_modificar_plan.html', {'registros': registros})


@login_required
@transaction.atomic
def editar_registro_plan(request, pk):
    registro = get_object_or_404(PlanCaptura, pk=pk)
    if request.method == 'POST':
        form = PlanCapturaForm(request.POST, instance=registro)
        if form.is_valid():
            form.save()
            messages.success(request, "¡Registro actualizado correctamente!")
            return redirect('lista_modificar_plan')
    else:
        form = PlanCapturaForm(instance=registro)
    return render(request, 'editar_registro_plan.html', {'form': form, 'registro': registro})


@login_required
@transaction.atomic
def plan_captura(request):
    form = PlanCapturaForm(request.POST or None)
    registros = PlanCaptura.objects.all()
    if form.is_valid():
        form.save()
        return redirect('plan_captura')
    return render(request, 'plan_captura.html', {'form': form, 'registros': registros})


@login_required
@transaction.atomic
def programa_real(request):
    form = ProgramaRealForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        form.save()
        return redirect('programa_real')

    registros = ProgramaReal.objects.all().order_by('no')
    totales = registros.aggregate(
        total_importe=Sum('importe'), total_operativo=Sum('operativo'),
        total_promotores=Sum('promotores'), total_administrative=Sum('administrativo'),
        total_confianza=Sum('confianza')
    )
    total_general_participantes = (
        (totales['total_operativo'] or 0) + (totales['total_promotores'] or 0) + 
        (totales['total_administrative'] or 0) + (totales['total_confianza'] or 0)
    )
    return render(request, 'programa_real.html', {
        'form': form, 'registros': registros, 'totales': totales,
        'total_general_participantes': total_general_participantes
    })


@login_required
@transaction.atomic
def cursos_nuevos(request):
    mes_seleccionado = request.GET.get('mes', None)
    if not mes_seleccionado:
        mes_seleccionado = str(datetime.now().month).zfill(2)
        
    meses_dict = {
        '01': 'ENERO', '02': 'FEBRERO', '03': 'MARZO', '04': 'ABRIL',
        '05': 'MAYO', '06': 'JUNIO', '07': 'JULIO', '08': 'AGOSTO',
        '09': 'SEPTIEMBRE', '10': 'OCTUBRE', '11': 'NOVIEMBRE', '12': 'DICIEMBRE'
    }
    nombre_mes = meses_dict.get(mes_seleccionado, 'MES NO VALIDO')
    instance_id = request.GET.get('editar_id', None)
    instance = get_object_or_404(CursoExcel, id=instance_id) if instance_id else None
    
    if request.method == 'POST':
        form = CursoExcelForm(request.POST, instance=instance)
        if form.is_valid():
            nuevo_registro = form.save(commit=False)
            if not instance:
                nuevo_registro.mes = mes_seleccionado 
            nuevo_registro.save()
            messages.success(request, "Registro guardado correctamente.")
            return redirect(f"{request.path}?mes={mes_seleccionado}")
    else:
        form = CursoExcelForm(instance=instance)

    registros = CursoExcel.objects.filter(mes=mes_seleccionado).order_by('no')
    totales = registros.aggregate(
        total_cantidad=Sum('cantidad'), total_constancia=Sum('constancia'),
        total_operativo=Sum('operativo'), total_promotores=Sum('promotores'),
        total_administrative=Sum('administrativo'), total_confianza=Sum('confianza'),
        total_hombres=Sum('hombres'), total_mujeres=Sum('mujeres')
    )
    total_general = (
        (totales['total_operativo'] or 0) + (totales['total_promotores'] or 0) + 
        (totales['total_administrative'] or 0) + (totales['total_confianza'] or 0)
    )
    return render(request, 'cursos_nuevos.html', {
        'form': form, 'registros': registros, 'totales': totales, 'total_general': total_general,
        'mes_seleccionado': mes_seleccionado, 'nombre_mes': nombre_mes, 'meses_dict': meses_dict, 'editando': bool(instance)
    })


@login_required
def capacitaciones(request):
    return render(request, 'placeholder.html', {'titulo': 'Capacitaciones'})


@login_required
@transaction.atomic
def gestion_capacitaciones(request):
    if request.method == 'POST':
        form = CapacitacionForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('gestion_capacitaciones')
    else:
        form = CapacitacionForm()
    registros = Capacitacion.objects.all()
    return render(request, 'capacitaciones.html', {'form': form, 'registros': registros})


@login_required
def buscar_reporte(request):
    query = request.GET.get('q', '').strip()
    registros = PlanCaptura.objects.filter(nomina__icontains=query) if query else PlanCaptura.objects.all()
    return render(request, 'buscar_reporte.html', {'query': query, 'registros': registros})


@login_required
def descargar_pdf(request):
    query = request.GET.get('q', '').strip()
    registros = PlanCaptura.objects.filter(nomina__icontains=query) if query else PlanCaptura.objects.all()
    context = {'registros': registros, 'query': query or 'Todos los registros'}
    template = get_template('pdf_template.html')
    html = template.render(context)
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="Reporte_Nomina_{query or "General"}.pdf"'
    pisa_status = pisa.CreatePDF(html, dest=response)
    if pisa_status.err:
        return HttpResponse('Ocurrió un error al generar el PDF', status=500)
    return response
from datetime import date, datetime
import openpyxl
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import render, redirect
from .models import Curso, PlanCaptura

@login_required
def cargar_excel(request):
    if request.method == "POST" and request.FILES.get('archivo_excel'):
        excel_file = request.FILES['archivo_excel']
        if not excel_file.name.endswith('.xlsx') and not excel_file.name.endswith('.xls'):
            messages.error(request, 'Formato inválido. Por favor sube un archivo de Excel (.xlsx o .xls).')
            return redirect('cargar_excel')

        try:
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            worksheet = wb.active
            
            # 1. MAPEO FORZADO / SEGURO DE CABECERAS
            idx_no, idx_cc, idx_puesto, idx_nomina, idx_nombre, idx_area, idx_curp, idx_induccion = 0, 1, 2, 3, 4, 5, 6, 7
            idx_curso, idx_fi, idx_ft, idx_duracion, idx_horas = 13, 14, 15, 16, 18

            # Calibración dinámica por lectura de texto real
            for fila_head in range(16, 19):
                row_head = [str(cell.value).strip().upper() if cell.value is not None else "" for cell in worksheet[fila_head]]
                for i, text in enumerate(row_head):
                    if "NÓMINA" in text or "NOMINA" in text: idx_nomina = i
                    elif "NOMBRE" in text and "CURSO" not in text: idx_nombre = i
                    elif "PUESTO" in text: idx_puesto = i
                    elif "CENTRO" in text or "C.C" in text: idx_cc = i
                    elif "CURP" in text: idx_curp = i
                    elif "ÁREA" in text or "AREA" in text: idx_area = i
                    elif "INDUC" in text: idx_induccion = i
                    elif "NOMBRE DEL CURSO" in text or "CURSO" in text: idx_curso = i
                    elif "FI" == text or "INICIO" in text: idx_fi = i
                    elif "FT" == text or "TÉRMINO" in text or "TERMINO" in text: idx_ft = i
                    elif "D" == text and len(text) == 1: idx_duracion = i
                    elif "TOTAL HORAS" in text or "TOTAL DE HORAS" in text: idx_horas = i

            print(f"--> Índices de Mapeo Activos: Nómina={idx_nomina}, Nombre={idx_nombre}, Curso={idx_curso}")

            # ==========================================================
            # 2. PROCESAMIENTO SIN FILTROS DE EXCLUSIÓN
            # ==========================================================
            registros_creados = 0
            
            # Variables de memoria histórica con respaldos genéricos para evitar bloqueos iniciales
            ultimo_no = "1"
            ultimo_cc = "GENERAL"
            ultimo_puesto = "OPERATIVO"
            ultimo_nomina = "0000"
            ultimo_nombre = "TRABAJADOR EN PROCESO"  # <-- Inicializado con texto para evitar el 'continue' erróneo
            ultimo_area = "PRODUCCION"
            ultimo_curp = "XAXX010101X"
            ultimo_induccion = False
            ultimo_curso = "CURSO DE CAPACITACIÓN"
            ultima_fi = date.today()
            ultima_ft = date.today()
            ultima_duracion = 0
            ultima_horas = 0

            for idx_fila, row_tuple in enumerate(worksheet.iter_rows(values_only=True), start=1):
                if idx_fila <= 18:
                    continue

                if not row_tuple or all(cell is None for cell in row_tuple):
                    continue

                row = list(row_tuple)
                while len(row) < 30:
                    row.append(None)

                # 1. Extracción e inspección de valores string
                val_no = str(row[idx_no]).strip() if row[idx_no] is not None else ""
                val_cc = str(row[idx_cc]).strip() if row[idx_cc] is not None else ""
                val_puesto = str(row[idx_puesto]).strip() if row[idx_puesto] is not None else ""
                val_nomina = str(row[idx_nomina]).strip() if row[idx_nomina] is not None else ""
                val_nombre = str(row[idx_nombre]).strip() if row[idx_nombre] is not None else ""
                val_area = str(row[idx_area]).strip() if row[idx_area] is not None else ""
                val_curp = str(row[idx_curp]).strip() if row[idx_curp] is not None else ""

                val_nombre_upper = val_nombre.upper()
                val_nomina_upper = val_nomina.upper()

                # Ignorar firmas y totales finales
                if "TOTAL" in val_nombre_upper or "TOTAL" in val_nomina_upper or "ELABORÓ" in val_nombre_upper:
                    continue

                # --- ARRASTRE DE DATOS DEL TRABAJADOR ---
                if val_no and val_no not in ["", "0", "0.0", "None", "NONE"]: ultimo_no = val_no
                if val_cc and val_cc not in ["", "0", "0.0", "None", "NONE"]: ultimo_cc = val_cc
                if val_puesto and val_puesto not in ["", "0", "0.0", "None", "NONE"]: ultimo_puesto = val_puesto
                if val_nomina and val_nomina not in ["", "0", "0.0", "None", "NONE"]: ultimo_nomina = val_nomina
                if val_nombre and val_nombre not in ["", "0", "0.0", "None", "NONE"]: ultimo_nombre = val_nombre
                if val_area and val_area not in ["", "0", "0.0", "None", "NONE"]: ultimo_area = val_area
                if val_curp and val_curp not in ["", "0", "0.0", "None", "NONE"]: ultimo_curp = val_curp

                if not ultimo_nombre or ultimo_nombre.upper() in ["", "0", "0.0", "NONE", "TRABAJADOR EN PROCESO"]:
                    if val_nombre:  # Intento de rescate inmediato
                        ultimo_nombre = val_nombre
                    else:
                        continue

                # --- ARRASTRE DE CAPACITACIÓN Y FECHAS (EVITA LOS NULLS) ---
                val_curso = str(row[idx_curso]).strip() if row[idx_curso] is not None else ""
                if not val_curso or val_curso.upper() in ["", "0", "0.0", "NONE"]:
                    val_curso = str(row[idx_curso + 1]).strip() if row[idx_curso + 1] is not None else ""

                if val_curso and val_curso.upper() not in ["", "0", "0.0", "NONE"]:
                    ultimo_curso = val_curso

                # Procesamiento tolerante de fechas con memoria
                def parsear_fecha(campo):
                    if not campo: return None
                    if isinstance(campo, (date, datetime)): return campo.date() if isinstance(campo, datetime) else campo
                    str_f = str(campo).strip()
                    for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"]:
                        try: return datetime.strptime(str_f[:10], fmt).date()
                        except ValueError: continue
                    return None

                val_fi = parsear_fecha(row[idx_fi])
                val_ft = parsear_fecha(row[idx_ft])

                if val_fi: ultima_fi = val_fi
                if val_ft: ultima_ft = val_ft
                elif val_fi: ultima_ft = val_fi

                # Duraciones y Horas con memoria
                val_duracion = row[idx_duracion]
                val_horas = row[idx_horas] if idx_horas < len(row) and row[idx_horas] is not None else val_duracion

                try: 
                    if val_duracion is not None: ultima_duracion = int(float(val_duracion))
                except: pass

                try: 
                    if val_horas is not None: ultima_horas = int(float(val_horas))
                    elif val_duracion is not None: ultima_horas = ultima_duracion
                except: pass

                # Inducción con memoria
                val_induccion_raw = str(row[idx_induccion]).strip().upper() if row[idx_induccion] is not None else ""
                if val_induccion_raw in ['X', 'SI', 'SÍ', '1', 'TRUE']:
                    ultimo_induccion = True
                elif row[idx_induccion] is not None:
                    ultimo_induccion = False

                # Formatear nómina
                if ultimo_nomina.endswith('.0'):
                    ultimo_nomina = ultimo_nomina[:-2]

                try: numero_registro = int(float(ultimo_no))
                except: numero_registro = registros_creados + 1

                # Obtener objeto Curso
                curso_obj, _ = Curso.objects.get_or_create(nombre=ultimo_curso)

                # --- INSERCIÓN EN BASE DE DATOS SIN CAMPOS NULL ---
                PlanCaptura.objects.create(
                    no=numero_registro, 
                    centro_costos=ultimo_cc, 
                    puesto=ultimo_puesto,
                    nomina=ultimo_nomina, 
                    nombre_trabajador=ultimo_nombre, 
                    area=ultimo_area,
                    curp=ultimo_curp, 
                    induccion=ultimo_induccion, 
                    columna_ind1="", 
                    columna_ind2="",
                    fecha_inicio=ultima_fi, 
                    fecha_termino=ultima_ft, 
                    duracion=ultima_duracion,
                    total_horas=ultima_horas, 
                    curso=curso_obj
                )
                registros_creados += 1

            print(f"--> [EXITO] SE HAN VOLCADO A LA BASE DE DATOS: {registros_creados} REGISTROS.")
            messages.success(request, f'¡Excelente! Se han capturado {registros_creados} registros directamente en PlanCaptura.')
            return redirect('dashboard')
            
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            messages.error(request, f'Error crítico en la inserción: {str(e)}')
            return redirect('cargar_excel')

    return render(request, 'core/cargar_excel.html')
import os
from io import BytesIO
from django.conf import settings
from django.shortcuts import render, get_object_or_404
from django.http import HttpResponse, JsonResponse
from django.contrib.auth.decorators import login_required
from django.db.models import Q
from .models import CargaSTPS

# Paquete oficial correcto para renderizar plantillas Word
from docxtpl import DocxTemplate

@login_required
def visor_dc3_excel(request):
    """
    Maneja la interfaz web de búsqueda del personal (Autocomplete AJAX).
    """
    if request.headers.get('x-requested-with') == 'XMLHttpRequest' or request.GET.get('term'):
        q = request.GET.get('term', '').strip()
        q = q.replace('\t', ' ')
        palabras = [p for p in q.split(' ') if p]

        if palabras:
            filtro = Q()
            for palabra in palabras:
                filtro &= (
                    Q(nombre__icontains=palabra) |
                    Q(primer_apellido__icontains=palabra) |
                    Q(segundo_apellido__icontains=palabra)
                )
            registros = CargaSTPS.objects.filter(filtro)[:10]
        else:
            registros = CargaSTPS.objects.none()

        resultados = []
        for reg in registros:
            nombre = f"{reg.nombre or ''} {reg.primer_apellido or ''} {reg.segundo_apellido or ''}"
            nombre = " ".join(nombre.split())

            resultados.append({
                "id": reg.id,
                "label": f"{nombre} — [{reg.nombre_curso or 'SIN CURSO'}]",
                "value": nombre
            })

        return JsonResponse(resultados, safe=False)

    context = {}
    empleado_id = request.GET.get('empleado_id')

    if empleado_id:
        registro = get_object_or_404(CargaSTPS, id=empleado_id)

        try:
            horas = int(registro.duracion)
        except (ValueError, TypeError):
            horas = 0

        nombre = f"{registro.nombre or ''} {registro.primer_apellido or ''} {registro.segundo_apellido or ''}"
        nombre = " ".join(nombre.split())

        context = {
            "registro": registro,
            "nombre_completo": nombre,
            "error_horas": horas < 0
        }

    return render(request, "capacitacion/visor_dc3_excel.html", context)


# =========================================================================
# 📥 DESCARGAR DC3 - ACTUALIZADO CON HORAS EN CUADRITOS INDIVIDUALES
# =========================================================================

import os
import pandas as pd
from io import BytesIO
from django.conf import settings
from django.shortcuts import get_object_or_404, render, redirect
from django.http import HttpResponse
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from docxtpl import DocxTemplate

# Importamos todos tus modelos oficiales
from .models import CargaSTPS, PlanCaptura, AreaTematica

@login_required
def cargar_areas_tematicas(request):
    """
    Interfaz para subir el Excel independiente de Áreas Temáticas y guardarlo en la Base de Datos
    """
    if not request.user.is_staff:
        return HttpResponse("Acceso denegado", status=403)

    if request.method == 'POST' and request.FILES.get('archivo_excel'):
        excel_file = request.FILES['archivo_excel']
        
        try:
            # 1. Leer todas las pestañas del Excel
            dict_excel = pd.read_excel(excel_file, sheet_name=None)
            
            # 2. Buscar primero si existe una pestaña que se parezca a 'area' o 'tematica'
            nombre_pestana = next((sheet for sheet in dict_excel.keys() if 'area' in sheet.lower() or 'tematica' in sheet.lower()), None)
            
            # 3. 💡 FLEXIBILIDAD: Si no la encuentra por nombre, agarra la primera pestaña que tenga el archivo
            if not nombre_pestana:
                nombre_pestana = list(dict_excel.keys())[0]
            
            # Cargar el DataFrame de la pestaña seleccionada
            df = dict_excel[nombre_pestana]
            
            # Limpiar nombres de columnas (quitar espacios y pasar a minúsculas)
            df.columns = [str(col).strip().lower() for col in df.columns]
            
            # Detectar dinámicamente las columnas de Clave y Descripción
            col_clave = next((c for c in df.columns if 'clave' in c or 'id' in c or 'cod' in c), None)
            col_desc = next((c for c in df.columns if 'desc' in c or 'nombre' in c or 'area' in c), None)
            
            if not col_clave or not col_desc:
                messages.error(
                    request, 
                    f"El Excel (Hoja: '{nombre_pestana}') debe tener una columna para la 'Clave' y otra para la 'Descripción'. Verifique los encabezados."
                )
                return render(request, 'core/cargar_areas.html')
            
            contador_creados = 0
            
            # 4. Recorrer las filas del Excel e impactar la BD
            for _, fila in df.iterrows():
                clave_valor = str(fila[col_clave]).strip()
                descripcion_valor = str(fila[col_desc]).strip()
                
                # Ignorar registros vacíos o nulos
                if pd.isna(fila[col_clave]) or clave_valor.lower() == 'nan' or not clave_valor:
                    continue
                
                # Quitar decimales flotantes molestos si la clave se lee como número (ej. "95.0" -> "95")
                if clave_valor.endswith('.0'):
                    clave_valor = clave_valor[:-2]
                
                # Guarda o actualiza en el modelo oficial
                AreaTematica.objects.update_or_create(
                    clave=clave_valor,
                    defaults={'descripcion': descripcion_valor}
                )
                contador_creados += 1
                
            messages.success(request, f"¡Catálogo procesado con éxito! Se sincronizaron {contador_creados} áreas temáticas desde la hoja '{nombre_pestana}'.")
            return redirect('cargar_areas_tematicas')
            
        except Exception as e:
            messages.error(request, f"Error al procesar el archivo Excel: {str(e)}")
            
    return render(request, 'core/cargar_areas.html')

@login_required
def descargar_dc3_relleno(request, empleado_id):
    """
    Genera y descarga el archivo Word DC-3 cruzando Puesto y Área Temática oficial
    """
    if not request.user.is_staff:
        return HttpResponse("Acceso denegado", status=403)

    # 1. Obtener datos del trabajador
    registro = get_object_or_404(CargaSTPS, id=empleado_id)

    nombre_completo = f"{registro.primer_apellido or ''} {registro.segundo_apellido or ''} {registro.nombre or ''}".strip()
    nombre_completo = " ".join(nombre_completo.split()).upper()

    ruta_plantilla = os.path.abspath(
        os.path.join(settings.BASE_DIR, "media", "plantillas", "DC-3-Plantilla.docx")
    )

    try:
        doc = DocxTemplate(ruta_plantilla)
    except FileNotFoundError:
        return HttpResponse(f"No se encontró la plantilla en: {ruta_plantilla}", status=404)

    # --- BUSCAR EL PUESTO EN PLAN_CAPTURA ---
    puesto_trabajador = ""
    curp_registro = (registro.curp or "").strip()

    if curp_registro:
        plan_empleado = PlanCaptura.objects.filter(curp__iexact=curp_registro).first()
        if plan_empleado:
            puesto_trabajador = getattr(plan_empleado, 'puesto', "")

    if not puesto_trabajador and nombre_completo:
        plan_empleado = PlanCaptura.objects.filter(nombre_trabajador__icontains=nombre_completo).first()
        if plan_empleado:
            puesto_trabajador = getattr(plan_empleado, 'puesto', "")

    puesto_trabajador = (puesto_trabajador or "").upper().strip()

    # --- BUSCAR EL ÁREA TEMÁTICA EN EL NUEVO MODELO ---
    area_tematica_curso = ""
    clave_area = str(getattr(registro, 'clave_area_tematica', "") or "").strip()

    if clave_area:
        # Buscamos en el modelo oficial que acabamos de crear
        area_obj = AreaTematica.objects.filter(clave=clave_area).first()
        if area_obj:
            area_tematica_curso = area_obj.descripcion.upper().strip()

    # Si no se encuentra en la base de datos, dejamos la clave numérica por seguridad
    if not area_tematica_curso:
        area_tematica_curso = clave_area.upper()

    # --- PROCESAR HORAS ---
    valor_duracion = getattr(registro, 'duracion', None) or getattr(registro, 'horas', None) or 0
    try:
        horas_texto = str(int(valor_duracion))
    except (ValueError, TypeError):
        horas_texto = str(valor_duracion).strip()

    # --- PROCESAR FECHAS ---
    f_inicio_obj = getattr(registro, 'fec_inicio', None)
    ano_ini, mes_ini, dia_ini = "", "", ""
    if f_inicio_obj:
        try:
            ano_ini, mes_ini, dia_ini = f_inicio_obj.strftime('%Y'), f_inicio_obj.strftime('%m'), f_inicio_obj.strftime('%d')
        except AttributeError:
            f_str = str(f_inicio_obj).strip()
            if "-" in f_str:
                partes = f_str.split()[0].split('-')
                if len(partes) == 3: ano_ini, mes_ini, dia_ini = partes[0], partes[1], partes[2]

    f_fin_obj = getattr(registro, 'fec_termino', None)
    ano_fin, mes_fin, dia_fin = "", "", ""
    if f_fin_obj:
        try:
            ano_fin, mes_fin, dia_fin = f_fin_obj.strftime('%Y'), f_fin_obj.strftime('%m'), f_fin_obj.strftime('%d')
        except AttributeError:
            f_str = str(f_fin_obj).strip()
            if "-" in f_str:
                partes = f_str.split()[0].split('-')
                if len(partes) == 3: ano_fin, mes_fin, dia_fin = partes[0], partes[1], partes[2]

    # Diccionario final para Word
    contexto_word = {
        "nombre": nombre_completo,
        "ocupacion": (registro.clave_ocupacion or "").strip(),
        "empresa": "LECHE PARA EL BIENESTAR LICONSA S.A. DE C.V.",
        "curso": (registro.nombre_curso or "").upper().strip(),
        "agente": (registro.rfc_agente_stps or "REGISTRO INTERNO LICONSA").upper().strip(),
        
        "puesto": puesto_trabajador,
        "area_tematica": area_tematica_curso,
        
        "duracion_horas": horas_texto,
        "ano_inicio": ano_ini, "mes_inicio": mes_ini, "dia_inicio": dia_ini,
        "ano_fin": ano_fin, "mes_fin": mes_fin, "dia_fin": dia_fin,
    }

    # Cuadritos CURP
    curp_texto = (registro.curp or "").upper().strip().ljust(18)
    for i, letra in enumerate(curp_texto):
        contexto_word[f"c{i}"] = letra

    # Cuadritos RFC Empresa
    rfc_texto = "LIC950821M42".upper().strip().ljust(12)
    for i, letra in enumerate(rfc_texto):
        contexto_word[f"r{i}"] = letra

    doc.render(contexto_word)

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    response = HttpResponse(output.getvalue(), content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    filename = f"DC3_{(registro.curp or registro.id).upper().strip()}.docx"
    response["Content-Disposition"] = f'attachment; filename="{filename}"'

    return response

import os
from io import BytesIO
from datetime import datetime
from django.conf import settings
from django.shortcuts import get_object_or_404
from django.http import HttpResponse
from django.contrib.auth.decorators import login_required
from docxtpl import DocxTemplate
from .models import CatalogoOcupacion
from .models import CargaSTPS, PlanCaptura, AreaTematica


@login_required
def descargar_dc3_relleno(request, empleado_id):

    # 🔒 Seguridad
    if not request.user.is_staff:
        return HttpResponse("Acceso denegado", status=403)

    # 📌 Obtener registro
    registro = get_object_or_404(CargaSTPS, id=empleado_id)

    # =========================
    # 🧍 NOMBRE COMPLETO
    # =========================
    nombre_completo = f"{registro.primer_apellido or ''} {registro.segundo_apellido or ''} {registro.nombre or ''}"
    nombre_completo = " ".join(nombre_completo.split()).upper()

    # =========================
    # 📄 PLANTILLA WORD
    # =========================
    ruta_plantilla = os.path.join(
        settings.BASE_DIR,
        "media",
        "plantillas",
        "DC-3-Plantilla.docx"
    )

    if not os.path.exists(ruta_plantilla):
        return HttpResponse("No existe la plantilla DC-3", status=404)

    doc = DocxTemplate(ruta_plantilla)

    # =========================
    # 🧑‍💼 PUESTO (CRUCE)
    # =========================
    puesto_trabajador = ""
    curp = (registro.curp or "").strip()

    if curp:
        plan = PlanCaptura.objects.filter(curp__iexact=curp).first()
        if plan:
            puesto_trabajador = plan.puesto or ""

    if not puesto_trabajador:
        plan = PlanCaptura.objects.filter(
            nombre_trabajador__icontains=nombre_completo
        ).first()
        if plan:
            puesto_trabajador = plan.puesto or ""

    puesto_trabajador = puesto_trabajador.upper()

    # =========================
    # 📚 ÁREA TEMÁTICA
    # =========================
    clave_area = str(registro.clave_area_tematica or "").strip()

    if clave_area.endswith(".0"):
        clave_area = clave_area[:-2]

    area = AreaTematica.objects.filter(clave=clave_area).first()
    area_tematica = area.descripcion.upper() if area else clave_area

    # =========================
    # ⏱ DURACIÓN
    # =========================
    duracion = registro.duracion or registro.horas or 0

    try:
        horas = str(int(duracion))
    except:
        horas = str(duracion)

    # =========================
    # 📅 FUNCIÓN DE FECHA ROBUSTA
    # =========================
    def formatear_fecha(valor):
        if not valor:
            return "", "", ""

        # Si es fecha tipo Django
        if hasattr(valor, "strftime"):
            return (
                valor.strftime("%Y"),
                valor.strftime("%m"),
                valor.strftime("%d")
            )

        texto = str(valor).strip().split()[0]

        formatos = [
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%Y/%m/%d",
            "%d/%m/%y",
            "%d-%m-%y",
        ]

        for formato in formatos:
            try:
                fecha = datetime.strptime(texto, formato)
                return (
                    fecha.strftime("%Y"),
                    fecha.strftime("%m"),
                    fecha.strftime("%d")
                )
            except:
                continue

        return "", "", ""

    # =========================
    # 📅 OBTENER FECHAS
    # =========================
    inicio = (
        getattr(registro, 'fec_inicio', None)
        or getattr(registro, 'fecha_inicio', None)
    )

    fin = (
        getattr(registro, 'fec_termino', None)
        or getattr(registro, 'fecha_termino', None)
        or getattr(registro, 'fec_fin', None)
    )

    ano_ini, mes_ini, dia_ini = formatear_fecha(inicio)
    ano_fin, mes_fin, dia_fin = formatear_fecha(fin)

    # =========================
    # 📦 CONTEXTO WORD
    # =========================
    contexto = {
        "nombre": nombre_completo,
        "ocupacion": (registro.clave_ocupacion or ""),
        "empresa": "LECHE PARA EL BIENESTAR LICONSA S.A. DE C.V.",
        "curso": (registro.nombre_curso or "").upper(),
        "agente": (registro.rfc_agente_stps or "REGISTRO INTERNO").upper(),

        "puesto": puesto_trabajador,
        "area_tematica": area_tematica,

        "duracion_horas": horas,

        "ano_inicio": ano_ini,
        "mes_inicio": mes_ini,
        "dia_inicio": dia_ini,

        "ano_fin": ano_fin,
        "mes_fin": mes_fin,
        "dia_fin": dia_fin,
    }

    # =========================
    # 🆔 CURP (LETRA POR LETRA)
    # =========================
    curp_texto = (registro.curp or "").upper().ljust(18)

    for i, letra in enumerate(curp_texto):
        contexto[f"c{i}"] = letra

    # =========================
    # 🏢 RFC EMPRESA
    # =========================
    rfc = "LIC950821M42".ljust(12)

    for i, letra in enumerate(rfc):
        contexto[f"r{i}"] = letra

    # =========================
    # 🧾 RENDER WORD
    # =========================
    doc.render(contexto)

    # =========================
    # 💾 GUARDAR EN SERVIDOR
    # =========================
    ruta_guardado = os.path.join(settings.MEDIA_ROOT, "dc3_generados")
    os.makedirs(ruta_guardado, exist_ok=True)

    curp_archivo = (registro.curp or f"ID_{registro.id}").upper().strip()
    nombre_archivo = f"DC3_{curp_archivo}.docx"
    ruta_completa = os.path.join(ruta_guardado, nombre_archivo)

    doc.save(ruta_completa)

    # =========================
    # 📥 DESCARGA AUTOMÁTICA
    # =========================
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    response["Content-Disposition"] = f'attachment; filename="{nombre_archivo}"'

    return response
import pandas as pd
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from .models import CatalogoOcupacion


@login_required
def cargar_ocupaciones(request):

    if not request.user.is_staff:
        return HttpResponse("Acceso denegado", status=403)

    if request.method == 'POST' and request.FILES.get('archivo_excel'):
        excel_file = request.FILES['archivo_excel']

        try:
            dict_excel = pd.read_excel(excel_file, sheet_name=None)

            # Buscar hoja
            nombre_pestana = next(
                (sheet for sheet in dict_excel.keys() if 'ocupa' in sheet.lower()),
                list(dict_excel.keys())[0]
            )

            df = dict_excel[nombre_pestana]
            df.columns = [str(col).strip().lower() for col in df.columns]

            # Detectar columnas
            col_clave = next((c for c in df.columns if 'clave' in c or 'id' in c or 'cod' in c), None)
            col_desc = next((c for c in df.columns if 'desc' in c or 'nombre' in c or 'ocupa' in c), None)

            if not col_clave or not col_desc:
                messages.error(request, f"El Excel (Hoja: '{nombre_pestana}') debe tener columna clave y descripción.")
                return render(request, 'core/cargar_ocupaciones.html')

            contador = 0

            for _, fila in df.iterrows():

                clave_valor = str(fila[col_clave]).strip().upper()
                descripcion_valor = str(fila[col_desc]).strip().upper()

                if pd.isna(fila[col_clave]) or clave_valor.lower() == 'nan' or not clave_valor:
                    continue

                # 🔥 LIMPIEZA CLAVE (CRÍTICO)
                clave_valor = clave_valor.replace(" ", "").replace("\t", "")

                if clave_valor.endswith(".0"):
                    clave_valor = clave_valor[:-2]

                print("GUARDANDO:", clave_valor, "->", descripcion_valor)

                CatalogoOcupacion.objects.update_or_create(
                    clave=clave_valor,
                    defaults={'descripcion': descripcion_valor}
                )

                contador += 1

            messages.success(request, f"✔ {contador} ocupaciones cargadas correctamente.")
            return redirect('cargar_ocupaciones')

        except Exception as e:
            messages.error(request, f"Error al procesar ocupaciones: {str(e)}")

    return render(request, 'core/cargar_ocupaciones.html')