"""Script para integrar o auto-updater no ui.py"""
import re

UI_PATH = r"e:\Projetos - Antigravity\Project - Waves Scheduler\app\ui.py"

with open(UI_PATH, "r", encoding="utf-8") as f:
    content = f.read()

# ── 1. Adicionar import do updater ────────────────────────────────────────────
OLD_IMPORT = "from app.history import add_entry"
NEW_IMPORT = """from app.history import add_entry
from app.updater import check_for_update_async, download_and_install, get_current_version"""

if OLD_IMPORT in content and "from app.updater" not in content:
    content = content.replace(OLD_IMPORT, NEW_IMPORT, 1)
    print("[OK] Import do updater adicionado")
else:
    print("[SKIP] Import já existe ou marcador não encontrado")

# ── 2. Adicionar botão Update no header (antes do botão Histórico) ─────────────
OLD_HEADER = '''        # Botão Histórico
        history_btn = QPushButton("📋  " + get_translation("history", self.current_language))'''

NEW_HEADER = '''        # Botão Atualizar
        self.update_btn = QPushButton("🔄  " + get_translation("check_updates", self.current_language))
        self.update_btn.setMaximumWidth(135)
        self.update_btn.setToolTip(get_translation("check_updates_tip", self.current_language))
        self.update_btn.clicked.connect(self.check_for_updates)
        header_layout.addWidget(self.update_btn)

        # Botão Histórico
        history_btn = QPushButton("📋  " + get_translation("history", self.current_language))'''

if OLD_HEADER in content:
    content = content.replace(OLD_HEADER, NEW_HEADER, 1)
    print("[OK] Botão Update adicionado ao header")
else:
    print("[SKIP] Header marker não encontrado")

# ── 3. Adicionar verificação automática no final do init_ui ───────────────────
OLD_RECS = "        self.update_recommendations()"
NEW_RECS = """        self.update_recommendations()

        # Verificar atualizações em segundo plano ao iniciar
        check_for_update_async(self._on_update_check_done)"""

# Substituir apenas a primeira ocorrência (dentro de init_ui)
content = content.replace(OLD_RECS, NEW_RECS, 1)
print("[OK] Verificação automática ao iniciar adicionada")

# ── 4. Adicionar métodos check_for_updates e _on_update_check_done ────────────
OLD_SHOW_HISTORY = "    def show_history(self):"

NEW_UPDATE_METHODS = '''    def check_for_updates(self, silent=False):
        """Verifica atualizações manualmente (com feedback visual)."""
        self.update_btn.setEnabled(False)
        self.update_btn.setText("🔄  ...")
        self.statusBar().showMessage("🔍 " + get_translation("checking_updates", self.current_language))

        def _done(info):
            # Callback rodando na thread — usar QTimer para voltar à UI thread
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(0, lambda: self._on_update_check_done(info, silent=silent))

        check_for_update_async(_done)

    def _on_update_check_done(self, update_info, silent=False):
        """Callback chamado quando a verificação de updates termina."""
        lang = self.current_language
        self.update_btn.setEnabled(True)
        self.update_btn.setText("🔄  " + get_translation("check_updates", lang))

        if update_info is None:
            # Sem update
            if not silent:
                self.statusBar().showMessage("✅ " + get_translation("up_to_date", lang))
                QMessageBox.information(
                    self,
                    get_translation("check_updates", lang),
                    get_translation("up_to_date", lang)
                )
            else:
                self.statusBar().showMessage("✅ v" + get_current_version())
            return

        # Update disponível
        current = update_info.get("current_version", "?")
        latest  = update_info.get("tag_name", "?")
        name    = update_info.get("name", latest)
        notes   = update_info.get("body", "")[:400] or get_translation("no_release_notes", lang)

        msg = (
            f"🆕 {get_translation('update_available', lang)}\n\n"
            f"  {get_translation('current_version_label', lang)}: v{current}\n"
            f"  {get_translation('new_version_label', lang)}: {latest}\n\n"
            f"📋 {notes}\n\n"
            f"{get_translation('update_prompt', lang)}"
        )

        reply = QMessageBox.question(
            self,
            f"🌊 {name}",
            msg,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )

        if reply != QMessageBox.Yes:
            self.statusBar().showMessage(
                f"⏭️ {get_translation('update_skipped', lang)}"
            )
            return

        download_url = update_info.get("download_url")
        if not download_url:
            QMessageBox.warning(
                self,
                "Download",
                get_translation("update_no_asset", lang) + "\\n" + update_info.get("html_url", "")
            )
            return

        self.statusBar().showMessage("⬇️ " + get_translation("downloading_update", lang))
        self.update_btn.setEnabled(False)

        def _download_done():
            ok = download_and_install(
                download_url,
                progress_callback=lambda done, total: self.statusBar().showMessage(
                    f"⬇️ {done//1024:,} KB / {total//1024:,} KB"
                )
            )
            from PyQt5.QtCore import QTimer
            if ok:
                QTimer.singleShot(0, lambda: (
                    QMessageBox.information(
                        self,
                        get_translation("success", lang),
                        get_translation("update_installing", lang)
                    ),
                    self.close()
                ))
            else:
                QTimer.singleShot(0, lambda: (
                    self.update_btn.setEnabled(True),
                    QMessageBox.critical(
                        self,
                        get_translation("error", lang),
                        get_translation("update_download_error", lang)
                    )
                ))

        import threading
        threading.Thread(target=_download_done, daemon=True, name="updater-download").start()

    def show_history(self):'''

if OLD_SHOW_HISTORY in content:
    content = content.replace(OLD_SHOW_HISTORY, NEW_UPDATE_METHODS, 1)
    print("[OK] Métodos check_for_updates e _on_update_check_done adicionados")
else:
    print("[SKIP] show_history marker não encontrado")

with open(UI_PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("[DONE] ui.py integrado com auto-updater")
