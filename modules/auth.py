"""
Módulo de Autenticación y Gestión de Usuarios
Implementa la gestión de usuarios, almacenamiento seguro en SQLite,
hashing criptográfico con salt (SHA-256) y control de sesiones en Streamlit.
"""

import sqlite3
import hashlib
import secrets
import os
from typing import Optional, Dict, Any, Tuple
import streamlit as st

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "users.db")


def get_db_connection() -> sqlite3.Connection:
    """Obtiene o crea una conexión con la base de datos SQLite."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Inicializa la tabla de usuarios y carga las cuentas demo si no existen."""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        salt TEXT NOT NULL,
        plan TEXT NOT NULL DEFAULT 'free',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.commit()
    
    # Sembrar usuario demo por defecto si la tabla está vacía
    cursor.execute("SELECT COUNT(*) FROM users")
    count = cursor.fetchone()[0]
    
    if count == 0:
        _create_seeded_user("demo", "demo@fincorp.com", "demo123", "user")
        _create_seeded_user("admin", "admin@fincorp.com", "admin123", "user")
        
    conn.close()


def _hash_password(password: str, salt: str) -> str:
    """Genera hash SHA-256 de la contraseña con salt."""
    return hashlib.sha256((password + salt).encode("utf-8")).hexdigest()


def _create_seeded_user(username: str, email: str, password: str, plan: str = "user"):
    """Crea un usuario inicial con salt seguro."""
    salt = secrets.token_hex(16)
    pwd_hash = _hash_password(password, salt)
    conn = get_db_connection()
    try:
        conn.execute(
            "INSERT INTO users (username, email, password_hash, salt, plan) VALUES (?, ?, ?, ?, ?)",
            (username, email, pwd_hash, salt, plan)
        )
        conn.commit()
    except sqlite3.IntegrityError:
        pass
    finally:
        conn.close()


def register_user(username: str, email: str, password: str) -> Tuple[bool, str]:
    """Registra una nueva cuenta de usuario en la base de datos."""
    username = username.strip().lower()
    email = email.strip().lower()
    
    if len(username) < 3:
        return False, "El nombre de usuario debe tener al menos 3 caracteres."
    if "@" not in email or "." not in email:
        return False, "El correo electrónico ingresado no es válido."
    if len(password) < 6:
        return False, "La contraseña debe tener al menos 6 caracteres."
        
    salt = secrets.token_hex(16)
    pwd_hash = _hash_password(password, salt)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            "INSERT INTO users (username, email, password_hash, salt, plan) VALUES (?, ?, ?, ?, ?)",
            (username, email, pwd_hash, salt, "user")
        )
        conn.commit()
        return True, "¡Cuenta creada exitosamente! Ya puedes iniciar sesión con tus credenciales."
    except sqlite3.IntegrityError as e:
        err_msg = str(e).lower()
        if "username" in err_msg:
            return False, "El nombre de usuario ya está en uso."
        elif "email" in err_msg:
            return False, "El correo electrónico ya se encuentra registrado."
        return False, "El usuario o correo ya existen."
    finally:
        conn.close()


def authenticate_user(username: str, password: str) -> Optional[Dict[str, Any]]:
    """Autentica a un usuario y retorna su perfil si las credenciales son correctas."""
    username = username.strip().lower()
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username = ? OR email = ?", (username, username))
    user_row = cursor.fetchone()
    conn.close()
    
    if not user_row:
        return None
        
    salt = user_row["salt"]
    expected_hash = user_row["password_hash"]
    given_hash = _hash_password(password, salt)
    
    if secrets.compare_digest(expected_hash, given_hash):
        return {
            "id": user_row["id"],
            "username": user_row["username"],
            "email": user_row["email"]
        }
    return None


def update_user_plan(username: str, new_plan: str = "pro") -> bool:
    """Actualiza el plan de suscripción de un usuario en SQLite."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE users SET plan = ? WHERE username = ?", (new_plan.lower(), username.lower()))
    conn.commit()
    updated = cursor.rowcount > 0
    conn.close()
    return updated


def get_user_plan(username: str) -> str:
    """Obtiene el plan actual del usuario desde la base de datos."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT plan FROM users WHERE username = ?", (username.lower(),))
    row = cursor.fetchone()
    conn.close()
    return row["plan"] if row else "free"


# Inicializar base de datos al importar el módulo
init_db()
