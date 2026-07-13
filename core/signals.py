from django.contrib.auth.signals import user_logged_in, user_login_failed
from django.dispatch import receiver
from django.contrib import messages

@receiver(user_logged_in)
def alerta_inicio_sesion_exitoso(sender, request, user, **kwargs):
    """
    Se ejecuta automáticamente cuando un usuario inicia sesión con éxito.
    """
    nombre_completo = f"{user.first_name} {user.last_name}".strip()
    usuario_identificador = nombre_completo if nombre_completo else user.username
    
    # Mandamos un mensaje de tipo 'success'
    messages.success(
        request, 
        f"¡Bienvenido de vuelta, {usuario_identificador}!"
    )

@receiver(user_login_failed)
def alerta_inicio_sesion_fallido(sender, request, credentials, **kwargs):
    """
    Se ejecuta automáticamente cuando falla un intento de inicio de sesión.
    """
    usuario_intentado = credentials.get('username', 'Desconocido')
    
    # Mandamos un mensaje de tipo 'error'
    messages.error(
        request, 
        f"Acceso denegado. El usuario '{usuario_intentado}' o la contraseña son incorrectos."
    )