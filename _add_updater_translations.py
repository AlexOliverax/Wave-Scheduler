"""Adiciona traduções do updater ao utils.py"""

UI_PATH = r"e:\Projetos - Antigravity\Project - Waves Scheduler\app\utils.py"

with open(UI_PATH, "r", encoding="utf-8") as f:
    content = f.read()

UPDATER_TRANSLATIONS = {
    "pt-BR": {
        "check_updates": "Verificar Atualizações",
        "check_updates_tip": "Verificar se há uma versão mais nova disponível no GitHub",
        "checking_updates": "Verificando atualizações...",
        "up_to_date": "O aplicativo está atualizado!",
        "update_available": "Atualização disponível",
        "current_version_label": "Versão atual",
        "new_version_label": "Nova versão",
        "update_prompt": "Deseja baixar e instalar agora?",
        "update_skipped": "Atualização adiada",
        "update_no_asset": "Instalador não encontrado na release. Acesse manualmente:",
        "downloading_update": "Baixando atualização...",
        "update_installing": "Instalador iniciado! O aplicativo será fechado para concluir a instalação.",
        "update_download_error": "Falha ao baixar a atualização. Verifique sua conexão e tente novamente.",
        "no_release_notes": "Sem notas de versão.",
    },
    "en-US": {
        "check_updates": "Check Updates",
        "check_updates_tip": "Check for a newer version on GitHub",
        "checking_updates": "Checking for updates...",
        "up_to_date": "The application is up to date!",
        "update_available": "Update available",
        "current_version_label": "Current version",
        "new_version_label": "New version",
        "update_prompt": "Do you want to download and install now?",
        "update_skipped": "Update postponed",
        "update_no_asset": "Installer not found in the release. Visit manually:",
        "downloading_update": "Downloading update...",
        "update_installing": "Installer launched! The application will close to complete installation.",
        "update_download_error": "Failed to download update. Check your connection and try again.",
        "no_release_notes": "No release notes.",
    },
    "es-US": {
        "check_updates": "Buscar actualizaciones",
        "check_updates_tip": "Verificar si hay una versión más nueva en GitHub",
        "checking_updates": "Verificando actualizaciones...",
        "up_to_date": "¡La aplicación está actualizada!",
        "update_available": "Actualización disponible",
        "current_version_label": "Versión actual",
        "new_version_label": "Nueva versión",
        "update_prompt": "¿Desea descargar e instalar ahora?",
        "update_skipped": "Actualización pospuesta",
        "update_no_asset": "Instalador no encontrado en la versión. Acceda manualmente:",
        "downloading_update": "Descargando actualización...",
        "update_installing": "¡Instalador iniciado! La aplicación se cerrará para completar la instalación.",
        "update_download_error": "Error al descargar la actualización. Verifique su conexión.",
        "no_release_notes": "Sin notas de versión.",
    },
    "fr-FR": {
        "check_updates": "Vérifier les mises à jour",
        "check_updates_tip": "Vérifier s'il existe une version plus récente sur GitHub",
        "checking_updates": "Vérification des mises à jour...",
        "up_to_date": "L'application est à jour !",
        "update_available": "Mise à jour disponible",
        "current_version_label": "Version actuelle",
        "new_version_label": "Nouvelle version",
        "update_prompt": "Voulez-vous télécharger et installer maintenant ?",
        "update_skipped": "Mise à jour reportée",
        "update_no_asset": "Installateur non trouvé dans la version. Accédez manuellement :",
        "downloading_update": "Téléchargement de la mise à jour...",
        "update_installing": "Installateur lancé ! L'application va se fermer pour terminer l'installation.",
        "update_download_error": "Échec du téléchargement. Vérifiez votre connexion.",
        "no_release_notes": "Pas de notes de version.",
    },
    "de-DE": {
        "check_updates": "Updates prüfen",
        "check_updates_tip": "Auf GitHub nach einer neueren Version suchen",
        "checking_updates": "Auf Updates prüfen...",
        "up_to_date": "Die Anwendung ist aktuell!",
        "update_available": "Update verfügbar",
        "current_version_label": "Aktuelle Version",
        "new_version_label": "Neue Version",
        "update_prompt": "Möchten Sie jetzt herunterladen und installieren?",
        "update_skipped": "Update verschoben",
        "update_no_asset": "Installer in der Version nicht gefunden. Manuell aufrufen:",
        "downloading_update": "Update wird heruntergeladen...",
        "update_installing": "Installer gestartet! Die Anwendung wird geschlossen.",
        "update_download_error": "Download fehlgeschlagen. Verbindung prüfen.",
        "no_release_notes": "Keine Versionshinweise.",
    },
}

# Inserir as traduções em cada bloco de idioma
for lang, keys in UPDATER_TRANSLATIONS.items():
    # Encontrar o marcador de fim do bloco de cada idioma
    # Inserir antes do fechamento do dicionário do idioma
    marker = f'"language_name"'  # chave que sabemos que existe em todos os idiomas
    
    for key, value in keys.items():
        key_line = f'        "{key}": "{value}",'
        if f'"{key}":' not in content or key_line not in content:
            # Procurar o bloco do idioma e inserir as chaves
            # Encontra o bloco específico do idioma usando um marcador próximo
            lang_marker = f'"{lang}"' + ": {"
            if lang_marker not in content:
                print(f"[SKIP] Language block for {lang} not found")
                continue
            
            # Inserir após "language_name"
            insert_after = f'"language_name"'
            lang_pos = content.find(lang_marker)
            if lang_pos < 0:
                continue
            next_lang_name = content.find(insert_after, lang_pos)
            if next_lang_name < 0:
                continue
            
            # Pular até o fim da linha do language_name
            eol = content.find('\n', next_lang_name)
            if eol < 0:
                continue
            
            if f'"{key}":' not in content[lang_pos:lang_pos+5000]:
                content = content[:eol+1] + f'        "{key}": "{value}",\n' + content[eol+1:]
                print(f"[ADD] {lang}: {key}")

with open(UI_PATH, "w", encoding="utf-8") as f:
    f.write(content)
print("[DONE] Translations added to utils.py")
