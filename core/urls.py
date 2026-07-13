from django.urls import path
from . import views

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
    path('plan/', views.plan_captura, name='plan_captura'),
    path('programa/', views.programa_real, name='programa_real'),
    path('cursos/', views.cursos_nuevos, name='cursos_nuevos'),
    path('capacitaciones/', views.gestion_capacitaciones, name='gestion_capacitaciones'),
    path('capacitaciones/', views.capacitaciones, name='capacitaciones'),
    path('reportes/', views.buscar_reporte, name='buscar_reporte'),
    path('reportes/pdf/', views.descargar_pdf, name='descargar_pdf'),
    path('cargar-excel/', views.cargar_excel, name='cargar_excel'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('plan-captura/', views.plan_captura, name='plan_captura'),
    
    # NUEVAS RUTAS DE MODIFICACIÓN DE DATOS
    path('plan-captura/modificar/', views.lista_modificar_plan, name='lista_modificar_plan'),
    path('plan-captura/modificar/<int:pk>/editar/', views.editar_registro_plan, name='editar_registro_plan'),
    # ... Tus rutas anteriores
    path('panel/cargar-stps/', views.cargar_stps, name='cargar_stps'),
    # Rutas restantes del sistema
    path('cargar-excel/', views.cargar_excel, name='cargar_excel'),
    path('stps/registros/', views.lista_modificar_stps, name='lista_modificar_stps'),
    path('stps/editar/<int:pk>/', views.editar_registro_stps, name='editar_registro_stps'),
path('stps/captura-manual/', views.captura_manual_stps, name='captura_manual_stps'),
path('plan-captura/visor-dc3-excel/', views.visor_dc3_excel, name='visor_dc3_excel'),
    path('plan-captura/descargar-dc3-relleno/<int:empleado_id>/', views.descargar_dc3_relleno, name='descargar_dc3_relleno'),
path('cargar-areas/', views.cargar_areas_tematicas, name='cargar_areas_tematicas'),
    path('descargar-dc3/<int:empleado_id>/', views.descargar_dc3_relleno, name='descargar_dc3_relleno'),
    path('cargar-ocupaciones/', views.cargar_ocupaciones, name='cargar_ocupaciones'),
]



