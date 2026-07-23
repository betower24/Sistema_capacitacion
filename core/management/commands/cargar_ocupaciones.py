from django.core.management.base import BaseCommand
from core.models import CatalogoOcupacion

class Command(BaseCommand):
    help = 'Carga el catálogo completo de áreas y subáreas'

    def handle(self, *args, **options):
        datos = [
            ("01", "01.1", "Agricultura y silvicultura"),
            ("01", "01.2", "Ganadería"),
            ("01", "01.3", "Pesca y acuacultura"),

            ("02", "02.1", "Exploración"),
            ("02", "02.2", "Extracción"),
            ("02", "02.3", "Refinación y beneficio"),
            ("02", "02.4", "Provisión de energía"),
            ("02", "02.5", "Provisión de agua"),

            ("03", "03.1", "Planeación y dirección de obras"),
            ("03", "03.2", "Edificación y urbanización"),
            ("03", "03.3", "Acabado"),
            ("03", "03.4", "Instalación y mantenimiento"),

            ("04", "04.1", "Mecánica"),
            ("04", "04.2", "Electricidad"),
            ("04", "04.3", "Electrónica"),
            ("04", "04.4", "Informática"),
            ("04", "04.5", "Telecomunicaciones"),
            ("04", "04.6", "Procesos industriales"),

            ("05", "05.1", "Minerales no metálicos"),
            ("05", "05.2", "Metales"),
            ("05", "05.3", "Alimentos y bebidas"),
            ("05", "05.4", "Textiles y prendas de vestir"),
            ("05", "05.5", "Materia orgánica"),
            ("05", "05.6", "Productos químicos"),
            ("05", "05.7", "Productos metálicos y de hule y plástico"),
            ("05", "05.8", "Productos eléctricos y electrónicos"),
            ("05", "05.9", "Productos impresos"),

            ("06", "06.1", "Ferroviario"),
            ("06", "06.2", "Autotransporte"),
            ("06", "06.3", "Aéreo"),
            ("06", "06.4", "Marítimo y fluvial"),
            ("06", "06.5", "Servicios de apoyo"),

            ("07", "07.1", "Comercio"),
            ("07", "07.2", "Alimentación y hospedaje"),
            ("07", "07.3", "Turismo"),
            ("07", "07.4", "Deporte y esparcimiento"),
            ("07", "07.5", "Servicios personales"),
            ("07", "07.6", "Reparación de artículos de uso doméstico y personal"),
            ("07", "07.7", "Limpieza"),
            ("07", "07.8", "Servicio postal y mensajería"),

            ("08", "08.1", "Bolsa, banca y seguros"),
            ("08", "08.2", "Administración"),
            ("08", "08.3", "Servicios legales"),

            ("09", "09.1", "Servicios médicos"),
            ("09", "09.2", "Inspección sanitaria y del medio ambiente"),
            ("09", "09.3", "Seguridad social"),
            ("09", "09.4", "Protección de bienes y/o personas"),

            ("10", "10.1", "Publicación"),
            ("10", "10.2", "Radio, cine, televisión y teatro"),
            ("10", "10.3", "Interpretación artística"),
            ("10", "10.4", "Traducción e interpretación lingüística"),
            ("10", "10.5", "Publicidad, propaganda y relaciones públicas"),

            ("11", "11.1", "Investigación"),
            ("11", "11.2", "Enseñanza"),
            ("11", "11.3", "Difusión cultural"),
        ]

        creados = 0
        for area, subarea, desc in datos:
            obj, created = CatalogoOcupacion.objects.get_or_create(
                area=area,
                subarea=subarea,
                defaults={'descripcion': desc}
            )
            if created:
                creados += 1

        self.stdout.write(self.style.SUCCESS(f'¡Listo! Se cargaron {creados} registros en el catálogo.'))