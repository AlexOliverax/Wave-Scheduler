"""
app/updater.py — Auto-update via GitHub Releases

Fluxo:
  1. Lê a versão atual de VERSION (na raiz do projeto ou ao lado do executável)
  2. Chama a API pública do GitHub para obter a última release
  3. Compara versões usando packaging.version (ou fallback manual)
  4. Se houver atualização, retorna as informações da release
  5. Baixa o instalador .exe da release e executa com subprocess

Configuração necessária:
  Defina GITHUB_OWNER e GITHUB_REPO abaixo com os dados do seu repositório.
"""

import os
import sys
import logging
import threading
import urllib.request
import urllib.error
import json
import subprocess
import tempfile
from typing import Optional, Tuple, Dict

logger = logging.getLogger(__name__)

# ── Configuração — edite com os dados do seu repositório ──────────────────────
GITHUB_OWNER = "AlexOliverax"
GITHUB_REPO  = "Wave-Scheduler"
GITHUB_API   = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
# ─────────────────────────────────────────────────────────────────────────────


def _get_version_file_path() -> str:
    """Retorna o caminho do arquivo VERSION (suporta exe compilado e dev)."""
    if getattr(sys, "frozen", False):
        # Executável PyInstaller
        base = os.path.dirname(sys.executable)
    else:
        # Modo desenvolvimento
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "VERSION")


def get_current_version() -> str:
    """Lê a versão atual do arquivo VERSION."""
    try:
        path = _get_version_file_path()
        with open(path, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return "0.0.0"


def _parse_version(version_str: str) -> Tuple[int, ...]:
    """Converte 'v2.1.0' ou '2.1.0' em (2, 1, 0) para comparação numérica."""
    clean = version_str.lstrip("vV").strip()
    parts = clean.split(".")
    result = []
    for p in parts:
        try:
            result.append(int(p))
        except ValueError:
            result.append(0)
    return tuple(result)


def check_for_update(timeout: int = 8) -> Optional[Dict]:
    """
    Verifica se há uma versão mais nova no GitHub.

    Returns:
        dict com 'tag_name', 'name', 'body', 'html_url', 'download_url' se houver update.
        None se estiver na versão mais recente ou se não conseguir verificar.
    """
    try:
        req = urllib.request.Request(
            GITHUB_API,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": f"WavesScheduler/{get_current_version()}",
            }
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        latest_tag = data.get("tag_name", "0.0.0")
        current    = get_current_version()

        if _parse_version(latest_tag) > _parse_version(current):
            # Encontrar o asset .exe do instalador
            download_url = None
            for asset in data.get("assets", []):
                name = asset.get("name", "").lower()
                if name.endswith(".exe") and "setup" in name:
                    download_url = asset.get("browser_download_url")
                    break

            return {
                "tag_name":    latest_tag,
                "name":        data.get("name", latest_tag),
                "body":        data.get("body", ""),
                "html_url":    data.get("html_url", ""),
                "download_url": download_url,
                "current_version": current,
            }
        return None

    except urllib.error.URLError as e:
        logger.warning(f"Updater: sem conexão ou timeout — {e}")
        return None
    except Exception as e:
        logger.warning(f"Updater: erro inesperado — {e}")
        return None


def check_for_update_async(callback) -> threading.Thread:
    """
    Verifica atualizações em segundo plano.
    
    Args:
        callback: função chamada com o resultado (dict ou None) na thread
    Returns:
        Thread iniciada
    """
    def _worker():
        result = check_for_update()
        try:
            callback(result)
        except Exception as e:
            logger.error(f"Updater callback error: {e}")

    t = threading.Thread(target=_worker, daemon=True, name="updater-check")
    t.start()
    return t


def download_and_install(download_url: str, progress_callback=None) -> bool:
    """
    Baixa o instalador do GitHub e o executa.

    Args:
        download_url: URL do asset .exe no GitHub
        progress_callback: função(bytes_baixados, total_bytes) ou None

    Returns:
        True se o download e a execução do instalador iniciaram com sucesso
    """
    tmp_path = ""
    try:
        with tempfile.NamedTemporaryFile(
            suffix="_WavesScheduler_Setup.exe",
            delete=False,
            dir=tempfile.gettempdir()
        ) as tmp:
            tmp_path = tmp.name

        logger.info(f"Updater: baixando de {download_url} → {tmp_path}")

        def _reporthook(count, block_size, total_size):
            if progress_callback and total_size > 0:
                downloaded = min(count * block_size, total_size)
                try:
                    progress_callback(downloaded, total_size)
                except Exception:
                    pass

        urllib.request.urlretrieve(download_url, tmp_path, reporthook=_reporthook)
        logger.info(f"Updater: download concluído — {os.path.getsize(tmp_path):,} bytes")

        # Executa o instalador em modo silencioso (o app fecha automaticamente)
        subprocess.Popen([tmp_path], close_fds=True)
        logger.info("Updater: instalador iniciado")
        return True

    except Exception as e:
        logger.error(f"Updater: erro no download/instalação — {e}")
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        return False
