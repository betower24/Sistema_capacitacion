from django.contrib import admin
from django.urls import include, path
from django.contrib.auth import views as auth_views  # 1. Importa las vistas de autenticación

urlpatterns = [
    path('admin/', admin.site.urls),
    # 2. Define tu propia ruta de login apuntando a tu plantilla
    path('', include('core.urls')),
    path('login/', auth_views.LoginView.as_view(template_name='login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(), name='logout'),
]