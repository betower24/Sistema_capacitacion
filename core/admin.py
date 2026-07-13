from django.contrib import admin
from .models import Curso, PlanCaptura, ProgramaReal

# 1. Configuración para el Catálogo de Cursos
@admin.register(Curso)
class CursoAdmin(admin.ModelAdmin):
    list_display = ('id', 'nombre')  # Columnas visibles
    search_fields = ('nombre',)     # Barra de búsqueda

# 2. Configuración para el Plan de Captura
@admin.register(PlanCaptura)
class PlanCapturaAdmin(admin.ModelAdmin):
    # Columnas principales en la lista (Actualizadas)
    list_display = ('no', 'nombre_trabajador', 'centro_costos', 'puesto', 'curso', 'fecha_inicio', 'fecha_termino')
    
    # Filtros laterales (Cambiado 'modalidad' por 'modalidad_curso')
    list_filter = ('centro_costos', 'area', 'induccion', 'capacitacion', 'modalidad_curso')
    
    # Buscador por nombre, nómina o CURP
    search_fields = ('nombre_trabajador', 'nomina', 'curp')
    
    # Ordenar por el número de registro por defecto
    ordering = ('no',)

# 3. Configuración para el Programa Real
@admin.register(ProgramaReal)
class ProgramaRealAdmin(admin.ModelAdmin):
    # Columnas visibles (incluye tu propiedad calculada 'total_participantes')
    list_display = ('nombre', 'importe', 'instructor', 'fecha', 'total_participantes')
    
    # Filtros laterales
    list_filter = ('tipo_accion', 'modalidad')
    
    # Buscador por nombre o instructor
    search_fields = ('nombre', 'instructor')
