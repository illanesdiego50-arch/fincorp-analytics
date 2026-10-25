"""
Módulo de Facturación y Acceso (Compatibilidad)
Todos los usuarios registrados cuentan con acceso ilimitado a todas las funcionalidades.
"""

def has_feature_access(feature: str = "", plan: str = "") -> bool:
    """Todos los usuarios registrados tienen acceso completo al 100% de las funciones."""
    return True


def render_paywall_card(feature_name: str = "", benefit_description: str = ""):
    """No-op: Los paywalls han sido deshabilitados."""
    pass


def render_stripe_checkout():
    """No-op: La pasarela de pago simulada ha sido deshabilitada."""
    pass
