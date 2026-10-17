# conexion.py — Conexión a Supabase (proyecto Gestión de Viáticos, esquema mesa_entradas).
# Lee el .env de esta carpeta; no usa el .env compartido de viáticos.
import os
from pathlib import Path

from dotenv import load_dotenv
from supabase import create_client

load_dotenv(Path(__file__).resolve().parent / ".env")


def cliente():
    """Cliente de Supabase apuntado al esquema mesa_entradas."""
    url = os.getenv("SUPABASE_URL", "").strip().strip('"')
    key = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "").strip().strip('"')
    if not url or not key:
        raise ValueError("Faltan SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY en el .env de MesaEntradas")
    return create_client(url, key).schema(os.getenv("SUPABASE_SCHEMA", "mesa_entradas"))
