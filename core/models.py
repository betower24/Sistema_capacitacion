from django.db import models

# =========================
# CATÁLOGO DE CURSOS
# =========================
# cursos/models.py


from django.db import models
from django.db import models

class CargaSTPS(models.Model):
    curp = models.CharField(max_length=18, null=True, blank=True)
    nombre = models.CharField(max_length=100, null=True, blank=True)
    primer_apellido = models.CharField(max_length=100, null=True, blank=True)
    segundo_apellido = models.CharField(max_length=100, null=True, blank=True)
    clave_estado = models.CharField(max_length=10, null=True, blank=True)
    clave_municipio = models.CharField(max_length=10, null=True, blank=True)
    clave_ocupacion = models.CharField(max_length=20, null=True, blank=True)
    clave_niv_estudio = models.CharField(max_length=10, null=True, blank=True)
    clave_doc_probatorio = models.CharField(max_length=10, null=True, blank=True)
    clave_institucion = models.CharField(max_length=50, null=True, blank=True)
    clave_curso = models.CharField(max_length=50, null=True, blank=True)
    nombre_curso = models.CharField(max_length=200, null=True, blank=True)
    clave_area_tematica = models.CharField(max_length=20, null=True, blank=True)
    duracion = models.CharField(max_length=10, null=True, blank=True)
    fec_inicio = models.CharField(max_length=20, null=True, blank=True)
    fec_termino = models.CharField(max_length=20, null=True, blank=True)
    clave_tip_agent = models.CharField(max_length=10, null=True, blank=True)
    rfc_agente_stps = models.CharField(max_length=15, null=True, blank=True)
    clave_modalidad = models.CharField(max_length=10, null=True, blank=True)
    clave_capacitacion = models.CharField(max_length=10, null=True, blank=True)
    clave_establec = models.CharField(max_length=10, null=True, blank=True)
    fecha_creacion = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.curp} - {self.nombre_curso}"

class Curso(models.Model):
    """
    Modelo base para almacenar los nombres únicos de los cursos.
    Se utiliza principalmente para relacionarlo con el Plan de Captura.
    """
    nombre = models.CharField(max_length=255, unique=True)

    def __str__(self):
        return self.nombre

from django.db import models
from django.utils import timezone


class CursoExcel(models.Model):
    no = models.IntegerField(verbose_name="Número", default=1)
    nombre = models.CharField(max_length=255, verbose_name="Nombre del Curso")
    cantidad = models.IntegerField(default=0, verbose_name="Cantidad")
    constancia = models.IntegerField(default=0, verbose_name="Constancias")

    operativo = models.IntegerField(default=0, verbose_name="Operativo")
    promotores = models.IntegerField(default=0, verbose_name="Promotores")
    administrativo = models.IntegerField(default=0, verbose_name="Administrativo")
    confianza = models.IntegerField(default=0, verbose_name="Confianza")

    hombres = models.IntegerField(default=0, verbose_name="Hombres")
    mujeres = models.IntegerField(default=0, verbose_name="Mujeres")

    mes = models.CharField(max_length=2, default="01", verbose_name="Mes del Registro")
    anio = models.IntegerField(default=2026, verbose_name="Año del Registro")
    fecha_registro = models.DateField(default=timezone.now, verbose_name="Fecha de registro")

    class Meta:
        verbose_name = "Curso Excel"
        verbose_name_plural = "Cursos Excel"
        unique_together = [('nombre', 'mes', 'anio')]

    def __str__(self):
        return f"{self.no} - {self.nombre} ({self.mes}/{self.anio})"

    @property
    def total_participantes(self):
        return (
            (self.operativo or 0) +
            (self.promotores or 0) +
            (self.administrativo or 0) +
            (self.confianza or 0)
        )    
# =========================
# PLAN DE CAPTURA
# =========================# =========================
# PLAN DE CAPTURA (Actualizado según tu Excel)
# =========================
class PlanCaptura(models.Model):
    # Columnas iniciales
    no = models.IntegerField(default=1, verbose_name="No.")
    centro_costos = models.CharField(max_length=100, verbose_name="Centro de Costos")
    puesto = models.CharField(max_length=100, verbose_name="Puesto")
    nomina = models.CharField(max_length=50, verbose_name="Nómina")
    nombre_trabajador = models.CharField(max_length=255, null=True, blank=True)
    area = models.CharField(max_length=100, verbose_name="Área")
    curp = models.CharField(max_length=18, verbose_name="CURP")

    # Bloque de Inducción
    induccion = models.BooleanField(default=False, verbose_name="Inducción")
    columna_ind1 = models.CharField(max_length=100, blank=True, null=True, verbose_name="Columna1 (Inducción)")
    columna_ind2 = models.CharField(max_length=100, blank=True, null=True, verbose_name="Columna2 (Inducción)")

    # Bloque de Capacitación
    capacitacion = models.BooleanField(default=False, verbose_name="Capacitación")
    columna_cap1 = models.CharField(max_length=100, blank=True, null=True, verbose_name="Columna1 (Capacitación)")
    columna_cap2 = models.CharField(max_length=100, blank=True, null=True, verbose_name="Columna2 (Capacitación)")

    # Curso
    modalidad_curso = models.CharField(max_length=100, default='', verbose_name="Modalidad del Curso")
    curso = models.ForeignKey(Curso, on_delete=models.CASCADE, verbose_name="Nombre del Curso")

    # Fechas y Duración (FI, FT, D)
    fecha_inicio = models.DateField(verbose_name="FI")
    fecha_termino = models.DateField(verbose_name="FT")
    duracion = models.IntegerField(verbose_name="D")

    # Datos finales (G, Género, Costo, Horas)
    
    genero = models.CharField(max_length=10, verbose_name="Género")
    costo = models.CharField(max_length=20, verbose_name="Costo")
    total_horas = models.IntegerField(verbose_name="Total de Horas")

    def __str__(self):
        return f"{self.no} - {self.nombre_trabajador}"



# =========================
# PROGRAMA REAL
# =========================
class ProgramaReal(models.Model):
    # Campo numérico inicial
    no = models.IntegerField(default=1, verbose_name="No.")
    
    # Bloque de datos generales
    nombre = models.CharField(max_length=255, verbose_name="Nombre")
    importe = models.FloatField(default=0, verbose_name="Importe")
    tipo_accion = models.CharField(max_length=100, verbose_name="Tipo de Acción")
    modalidad = models.CharField(max_length=100, verbose_name="Modalidad")
    fecha = models.CharField(max_length=100, verbose_name="Fecha")
    
    # Instructor y Constancia
    instructor = models.CharField(max_length=100, verbose_name="Instructor Interno ó Externo")
    constancia = models.IntegerField(verbose_name="Constancia")

    # Bloque de participantes (Sindicatos y Confianza)
    operativo = models.IntegerField(default=0, verbose_name="Operativo Sindicalizado")
    promotores = models.IntegerField(default=0, verbose_name="Promotores Sociales")
    administrativo = models.IntegerField(default=0, verbose_name="Administrativo Sindicalizado")
    confianza = models.IntegerField(default=0, verbose_name="Confianza")

    # Propiedad dinámica para el total por fila
    @property
    def total_participantes(self):
        return self.operativo + self.promotores + self.administrativo + self.confianza

    def __str__(self):
        return f"{self.no} - {self.nombre}"

class Capacitacion(models.Model):
    consecutivo = models.IntegerField(
        verbose_name="Consecutivo"
    )
    nombre_tipo_capacitacion = models.CharField(
        max_length=255, 
        verbose_name="Nombre Tipo Capacitación"
    )
    participantes_operativos = models.PositiveIntegerField(
        default=0, 
        verbose_name="Participantes Operativos"
    )
    hombres = models.PositiveIntegerField(
        default=0, 
        verbose_name="Hombres"
    )
    mujeres = models.PositiveIntegerField(
        default=0, 
        verbose_name="Mujeres"
    )
    fortalecimiento_desempenio = models.PositiveIntegerField(
        default=0, 
        verbose_name="Fortalecimiento del Desempeño"
    )

    @property
    def total_participantes(self):
        """Calcula dinámicamente el total sumando hombres y mujeres"""
        return self.hombres + self.mujeres

    class Meta:
        verbose_name = "Capacitación"
        verbose_name_plural = "Capacitaciones"
        ordering = ['consecutivo']

    def __str__(self):
        return f"{self.consecutivo} - {self.nombre_tipo_capacitacion}"
    
    from django.db import models

# ... Tus otros modelos existentes (como CargaSTPS y PlanCaptura) ...

class AreaTematica(models.Model):
    clave = models.CharField(max_length=50, unique=True, verbose_name="Clave")
    descripcion = models.CharField(max_length=255, verbose_name="Descripción")

    def __str__(self):
        return f"{self.clave} - {self.descripcion}"

    class Meta:
        verbose_name = "Área Temática"
        verbose_name_plural = "Áreas Temáticas"
        db_table = "core_areatematica" # Forzamos el nombre exacto

class AreaLaboral(models.Model):
    codigo = models.CharField(max_length=10, unique=True, verbose_name="Código")
    nombre = models.CharField(max_length=255, verbose_name="Nombre del Área")

    class Meta:
        verbose_name = "Área Laboral"
        verbose_name_plural = "Áreas Laborales"
        ordering = ['codigo']

    def __str__(self):
        return f"{self.codigo} {self.nombre}"

class SubareaLaboral(models.Model):
    area = models.ForeignKey(AreaLaboral, on_delete=models.CASCADE, related_name='subareas')
    codigo = models.CharField(max_length=10, unique=True, verbose_name="Código")
    nombre = models.CharField(max_length=255, verbose_name="Nombre de la Subárea")

    class Meta:
        verbose_name = "Subárea Laboral"
        verbose_name_plural = "Subáreas Laborales"
        ordering = ['codigo']

    def __str__(self):
        return f"{self.codigo} {self.nombre}"
class CatalogoAgente(models.Model):
    nombre = models.CharField(max_length=255, unique=True, verbose_name="Nombre / Institución")

    class Meta:
        verbose_name = "Agente / Institución"
        verbose_name_plural = "Catálogo de Agentes e Instituciones"
        ordering = ['nombre']

    def __str__(self):
        return self.nombre