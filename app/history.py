"""
app/history.py – Persistência do histórico de schedules gerados.
"""
import os
import sys
import json
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

if getattr(sys, "frozen", False):
    # Executável PyInstaller: usa o diretório do .exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Modo desenvolvimento: sobe um nível a partir de app/
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

HISTORY_FILE = os.path.join(BASE_DIR, "history.json")
MAX_HISTORY_ENTRIES = 50



def _load_raw() -> list:
    """Carrega a lista bruta de entradas do arquivo history.json."""
    try:
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
    except Exception as e:
        logger.error(f"Error loading history: {e}")
    return []


def _save_raw(entries: list):
    """Persiste a lista de entradas no arquivo history.json."""
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(entries, f, ensure_ascii=False, indent=2, default=str)
    except Exception as e:
        logger.error(f"Error saving history: {e}")


def add_entry(excel_path: str, num_waves: int, num_devices: int, rfc: str,
              country_code: str = "BR", wave_labels: list = None):
    """
    Adiciona uma entrada ao histórico de schedules.

    Args:
        excel_path: Caminho do arquivo Excel gerado
        num_waves: Número de waves geradas
        num_devices: Total de dispositivos processados
        rfc: Número do RFC
        country_code: Código do país
        wave_labels: Lista de labels das waves (ex: ["Wave 1 - 01/04/2026", ...])
    """
    entry = {
        "timestamp": datetime.now().isoformat(),
        "excel_path": excel_path,
        "filename": os.path.basename(excel_path),
        "num_waves": num_waves,
        "num_devices": num_devices,
        "rfc": rfc or "N/A",
        "country_code": country_code,
        "wave_labels": wave_labels or [],
    }
    entries = _load_raw()
    entries.insert(0, entry)  # Mais recente primeiro
    # Limitar tamanho do histórico
    entries = entries[:MAX_HISTORY_ENTRIES]
    _save_raw(entries)
    logger.info(f"History entry added: {entry['filename']}")
    return entry


def get_history() -> list:
    """
    Retorna a lista de entradas de histórico, mais recente primeiro.

    Returns:
        list: Lista de dicionários com os campos de cada schedule
    """
    return _load_raw()


def clear_history():
    """Remove todas as entradas do histórico."""
    _save_raw([])
    logger.info("History cleared")


def remove_entry(timestamp: str):
    """Remove uma entrada específica pelo timestamp."""
    entries = _load_raw()
    entries = [e for e in entries if e.get("timestamp") != timestamp]
    _save_raw(entries)
