from locust import HttpUser, task, between, events
import re

# ====== CONFIGURA ESTO ======
USUARIO = "admin"          # tu usuario staff
PASSWORD = "tu_password"   # tu contraseña
# ============================

class UsuarioSistemaCapacitacion(HttpUser):
    wait_time = between(1, 2)  # espera 1–2 s entre peticiones

    def on_start(self):
        """Login una vez al iniciar cada usuario virtual"""
        # 1) Obtener CSRF del login
        r = self.client.get("/login/", name="login_get")
        csrf = self._csrf(r)

        # 2) Enviar login
        self.client.post(
            "/login/",
            data={
                "username": USUARIO,
                "password": PASSWORD,
                "csrfmiddlewaretoken": csrf,
            },
            headers={"Referer": self.client.base_url + "/login/"},
            name="login_post",
        )

    def _csrf(self, response):
        """Extrae csrfmiddlewaretoken del HTML"""
        if not response or not response.text:
            return ""
        m = re.search(r'name=["\']csrfmiddlewaretoken["\'] value=["\']([^"\']+)', response.text)
        return m.group(1) if m else ""

    # ---- Vistas principales (más peso = se prueban más) ----

    @task(4)
    def dashboard(self):
        self.client.get("/dashboard/", name="dashboard")
        # Si tu URL del panel es otra, cámbiala:
        # self.client.get("/", name="dashboard")

    @task(3)
    def capacitaciones(self):
        self.client.get("/capacitaciones/", name="capacitaciones")

    @task(3)
    def cursos_nuevos(self):
        self.client.get("/cursos/", name="cursos_nuevos")
        # Probar un mes concreto
        self.client.get("/cursos/?mes=01&anio=2026", name="cursos_nuevos_mes")

    @task(2)
    def programa_real(self):
        self.client.get("/programa-real/", name="programa_real")

    @task(2)
    def plan_captura(self):
        self.client.get("/plan-captura/", name="plan_captura")
        # Ajusta si tu URL es distinta

    @task(1)
    def visor_dc3(self):
        self.client.get("/plan-captura/visor-dc3-excel/", name="visor_dc3")
        # Ajusta según tu urls.py

    @task(1)
    def cargar_capacitaciones_page(self):
        """Solo abre la pantalla de carga (no sube archivo)"""
        self.client.get("/cargar-capacitaciones/", name="cargar_capacitaciones")