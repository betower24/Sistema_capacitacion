from django import forms
from .models import PlanCaptura, ProgramaReal
from .models import CursoExcel  # Asegúrate de que importe el modelo correcto
from django import forms
from .models import CargaSTPS

class CargaSTPSForm(forms.ModelForm):
    class Meta:
        model = CargaSTPS
        fields = '__all__'  # Esto hará editables todos los campos de la STPS
        # Si prefieres que las fechas tengan un calendario visual en el navegador:
        widgets = {
            'fec_inicio': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'fec_termino': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Añadimos la clase form-control de Bootstrap automáticamente a cada campo
        for field in self.fields.values():
            if not isinstance(field.widget, forms.DateInput):
                field.widget.attrs['class'] = 'form-control'
from django import forms
from .models import CursoExcel, Curso
from django import forms
from .models import CursoExcel, Curso
from django import forms
from .models import CursoExcel, Curso


from django import forms
from .models import CursoExcel, Curso


class CursoExcelForm(forms.ModelForm):
    nombre = forms.ChoiceField(
        label="Nombre del Curso",
        widget=forms.Select(attrs={'class': 'form-select'}),
        required=True
    )

    class Meta:
        model = CursoExcel
        fields = [
            'no', 'nombre', 'fecha', 'cantidad', 'constancia',
            'operativo', 'promotores', 'administrativo', 'confianza',
            'hombres', 'mujeres'
        ]
        widgets = {
            'no': forms.NumberInput(attrs={'class': 'form-control'}),
            'fecha': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'cantidad': forms.NumberInput(attrs={'class': 'form-control'}),
            'constancia': forms.NumberInput(attrs={'class': 'form-control'}),
            'operativo': forms.NumberInput(attrs={'class': 'form-control'}),
            'promotores': forms.NumberInput(attrs={'class': 'form-control'}),
            'administrativo': forms.NumberInput(attrs={'class': 'form-control'}),
            'confianza': forms.NumberInput(attrs={'class': 'form-control'}),
            'hombres': forms.NumberInput(attrs={'class': 'form-control'}),
            'mujeres': forms.NumberInput(attrs={'class': 'form-control'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        cursos = Curso.objects.all().order_by('nombre')
        opciones = [('', '-- Selecciona un curso --')]
        opciones += [(c.nombre, c.nombre) for c in cursos]
        self.fields['nombre'].choices = opciones
class PlanCapturaForm(forms.ModelForm):
    class Meta:
        model = PlanCaptura
        fields = '__all__'
        widgets = {
            'fecha_inicio': forms.DateInput(attrs={'type': 'date'}),
            'fecha_termino': forms.DateInput(attrs={'type': 'date'}),
        }


from django import forms
from .models import ProgramaReal


class ProgramaRealForm(forms.ModelForm):
    class Meta:
        model = ProgramaReal
        fields = [
            'no', 'nombre', 'importe', 'tipo_accion', 'modalidad', 'fecha',
            'instructor', 'constancia', 'operativo', 'promotores',
            'administrativo', 'confianza',
        ]
        widgets = {
            'no': forms.NumberInput(attrs={'class': 'form-control'}),
            'nombre': forms.TextInput(attrs={'class': 'form-control'}),
            'importe': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01'}),
            'tipo_accion': forms.TextInput(attrs={'class': 'form-control'}),
            'modalidad': forms.TextInput(attrs={'class': 'form-control'}),
            'fecha': forms.TextInput(attrs={'class': 'form-control'}),
            'instructor': forms.TextInput(attrs={'class': 'form-control'}),
            'constancia': forms.NumberInput(attrs={'class': 'form-control'}),
            'operativo': forms.NumberInput(attrs={'class': 'form-control'}),
            'promotores': forms.NumberInput(attrs={'class': 'form-control'}),
            'administrativo': forms.NumberInput(attrs={'class': 'form-control'}),
            'confianza': forms.NumberInput(attrs={'class': 'form-control'}),
        }
from django import forms
from .models import Capacitacion

from django import forms
from .models import Capacitacion

class CapacitacionForm(forms.ModelForm):
    class Meta:
        model = Capacitacion
        fields = [
            'consecutivo', 'nombre_tipo_capacitacion', 'tipo_capacitacion', 'modalidad', 'numero_acciones',
            'participantes_operativos', 'participantes_enlace', 'participantes_mandos_medios',
            'participantes_mandos_superiores', 'participantes_categorias_especiales',
            'induccion', 'fortalecimiento_desempenio', 'actualizacion', 'desarrollo',
            'certificacion', 'sensibilizacion',
            'hombres', 'mujeres',
        ]
        widgets = {
            field: forms.NumberInput(attrs={'class': 'form-control'})
            for field in [
                'consecutivo', 'numero_acciones',
                'participantes_operativos', 'participantes_enlace', 'participantes_mandos_medios',
                'participantes_mandos_superiores', 'participantes_categorias_especiales',
                'induccion', 'fortalecimiento_desempenio', 'actualizacion', 'desarrollo',
                'certificacion', 'sensibilizacion', 'hombres', 'mujeres',
            ]
        }
        widgets.update({
            'nombre_tipo_capacitacion': forms.TextInput(attrs={'class': 'form-control'}),
            'tipo_capacitacion': forms.TextInput(attrs={'class': 'form-control'}),
            'modalidad': forms.TextInput(attrs={'class': 'form-control'}),
        })

class PlanCapturaForm(forms.ModelForm):
    class Meta:
        model = PlanCaptura
        fields = '__all__'  # O define una lista exacta de tus campos: ['nomina', 'nombre', 'total_horas', etc.]
        
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Añadimos las clases de Bootstrap de forma automática a cada campo para que se vea premium
        for field in self.fields.values():
            field.widget.attrs.update({'class': 'form-control rounded-3'})
        #from core.models import Curso
#Curso.objects.all().delete()