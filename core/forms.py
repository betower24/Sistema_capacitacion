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
class CursoExcelForm(forms.ModelForm):
    class Meta:
        model = CursoExcel
        # Esto incluirá automáticamente todos los campos del modelo en tu HTML
        fields = [
            'no', 'nombre', 'cantidad', 'constancia', 
            'operativo', 'promotores', 'administrativo', 'confianza', 
            'hombres', 'mujeres'
        ]
        # Opcional: Puedes agregar clases de Bootstrap para que se vea bien
        widgets = {
            'nombre': forms.TextInput(attrs={'class': 'form-control'}),
            'no': forms.NumberInput(attrs={'class': 'form-control'}),
            'cantidad': forms.NumberInput(attrs={'class': 'form-control'}),
            'constancia': forms.NumberInput(attrs={'class': 'form-control'}),
            'operativo': forms.NumberInput(attrs={'class': 'form-control'}),
            'promotores': forms.NumberInput(attrs={'class': 'form-control'}),
            'administrativo': forms.NumberInput(attrs={'class': 'form-control'}),
            'confianza': forms.NumberInput(attrs={'class': 'form-control'}),
            'hombres': forms.NumberInput(attrs={'class': 'form-control'}),
            'mujeres': forms.NumberInput(attrs={'class': 'form-control'}),
        }
class PlanCapturaForm(forms.ModelForm):
    class Meta:
        model = PlanCaptura
        fields = '__all__'
        widgets = {
            'fecha_inicio': forms.DateInput(attrs={'type': 'date'}),
            'fecha_termino': forms.DateInput(attrs={'type': 'date'}),
        }


class ProgramaRealForm(forms.ModelForm):
    class Meta:
        model = ProgramaReal
        fields = '__all__'
        
from .models import Capacitacion

class CapacitacionForm(forms.ModelForm):
    class Meta:
        model = Capacitacion
        fields = [
            'consecutivo', 
            'nombre_tipo_capacitacion', 
            'participantes_operativos', 
            'hombres', 
            'mujeres', 
            'fortalecimiento_desempenio'
        ]
        from .models import PlanCaptura

class PlanCapturaForm(forms.ModelForm):
    class Meta:
        model = PlanCaptura
        fields = '__all__'  # O define una lista exacta de tus campos: ['nomina', 'nombre', 'total_horas', etc.]
        
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Añadimos las clases de Bootstrap de forma automática a cada campo para que se vea premium
        for field in self.fields.values():
            field.widget.attrs.update({'class': 'form-control rounded-3'})
        