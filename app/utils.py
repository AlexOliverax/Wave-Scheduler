import os
import json
import logging
import pytz
import holidays
from datetime import datetime, timedelta
from functools import lru_cache

logger = logging.getLogger(__name__)

DEFAULT_CONFIG = {
    "language": "pt-BR",
    "timezone": "UTC",
    "last_directory": "",
    "country": "BR"
}

# Dicionário de traduções expandido
TRANSLATIONS = {
    "pt-BR": {
        # Interface principal
        "app_title": "Waves Scheduler",
        "data_file_selection": "Seleção de Arquivo de Dados",
        "select_data_file": "Selecionar Arquivo de Dados (CSV/Excel)",
        "no_file_selected": "Nenhum arquivo selecionado",
        "device_information": "Informações do Dispositivo",
        "total_devices": "Total de Dispositivos: {}",
        "recommendations": "Recomendações",
        "ideal_bandwidth": "Largura de Banda Ideal (MB):",
        "ideal_waves": "Ondas Ideais:",
        "devices_per_wave": "Dispositivos por Onda:",
        "wave_configuration": "Configuração de Ondas",
        "rfc": "RFC:",
        "bandwidth": "Largura de Banda (MB):",
        "number_of_waves": "Número de Ondas:",
        "start_date": "Data de Início:",
        "outlook_integration": "Adicionar Ondas ao Outlook 365",
        "timezone": "Fuso Horário:",
        "timezone_select": "🌍 Selecionar",
        "country": "País:",
        "language": "Idioma:",
        "avoid_holidays": "Evitar feriados e fins de semana",
        "view_holidays": "📅 Ver Feriados do Ano",
        "start_time": "Hora de Início:",
        "end_time": "Hora de Fim:",
        "all_day_event": "Evento de Dia Inteiro:",
        "location": "Local:",
        "location_placeholder": "Ex: Sala de Reuniões, Online, etc.",
        "required_participants": "Participantes Obrigatórios*:",
        "optional_participants": "Participantes Opcionais:",
        "participants_placeholder": "email@dhl.com; email2@dhl.com",
        "generate_waves": "Gerar Ondas",
        "mb_per_device": "MB por Dispositivo:",
        "ready": "Pronto",
        
        # Diálogos
        "select_columns": "Selecionar Colunas",
        "column_instructions": "Selecione as colunas para incluir na saída. Colunas obrigatórias estão pré-selecionadas.",
        "required_columns": "Colunas obrigatórias: {}",
        "select_all": "Selecionar Todas",
        "clear_selection": "Limpar Seleção",
        "ok": "OK",
        "cancel": "Cancelar",
        "timezone_selection": "Seleção de Fuso Horário",
        "search_timezone": "Pesquisar fuso horário...",
        "holidays_view": "Feriados de {} - {}",
        "national_holidays": "Feriados Nacionais",
        "close": "Fechar",
        "no_holidays_found": "Nenhum feriado encontrado para este ano.",
        
        # Mensagens
        "supported_formats": "Formatos de Arquivo Suportados",
        "formats_info": "Você pode selecionar os seguintes formatos:\n\n- Arquivos CSV (.csv)\n- Arquivos Excel (.xlsx, .xls)",
        "file_error": "Erro ao carregar arquivo",
        "processing_file": "Processando arquivo...",
        "waves_generated": "Ondas geradas com sucesso!",
        "generation_error": "Erro ao gerar ondas",
        "holiday_warning": "⚠️ Atenção: A data selecionada ({}) é um feriado: {}",
        "information": "Informação",
        "restart_required": "Por favor, reinicie a aplicação para aplicar as mudanças de idioma.",
        # Geração e progresso
        "generating": "Gerando...",
        "step_holidays": "Carregando feriados...",
        "step_labels": "Gerando rótulos de waves...",
        "step_devices": "Distribuindo dispositivos...",
        "step_excel": "Gerando arquivo Excel...",
        "step_done": "Concluído!",
        "waves_generated_msg": "{} waves geradas com sucesso!",
        "file_label": "Arquivo:",
        "outlook_success": "Eventos do Outlook criados com sucesso.",
        "outlook_failed": "Eventos do Outlook não puderam ser criados.",
        "outlook_unavailable": "Microsoft Outlook não está disponível. O arquivo Excel será gerado, mas os eventos não serão criados.",
        "outlook_no_participants": "É necessário informar pelo menos um participante obrigatório.",
        "invalid_emails": "Os seguintes emails têm formato inválido:",
        "file_load_error": "Falha ao carregar arquivo. Verifique o formato e tente novamente.",
        "col_load_error": "Falha ao carregar arquivo com as colunas selecionadas.",
        "excel_error": "Falha ao gerar arquivo Excel. Verifique o log.",
        "error_occurred": "Ocorreu um erro:",
        "save_excel_title": "Salvar Arquivo Excel",
        "save_csv_title": "Salvar Arquivo CSV",
        "export_csv": "Exportar CSV também",
        "open_file": "Abrir Arquivo",
        "open_folder": "Abrir Pasta",
        "success": "Sucesso",
        "error": "Erro",
        "warning": "Aviso",
        "holiday_warning_title": "Aviso de Feriado",
        # Histórico
        "history": "Histórico",
        "history_title": "Histórico de Schedules Gerados",
        "history_empty": "Nenhum schedule gerado ainda.",
        "history_clear": "Limpar Histórico",
        "history_clear_confirm": "Tem certeza que deseja limpar todo o histórico?",
        "history_col_date": "Data/Hora",
        "history_col_rfc": "RFC",
        "history_col_waves": "Waves",
        "history_col_devices": "Dispositivos",
        "history_col_file": "Arquivo",
        "history_open": "Abrir Arquivo",
        # Preview de waves
        "wave_preview": "Preview das Waves",
        "wave_preview_title": "Preview — Confirmar antes de salvar",
        "wave_col_wave": "Wave",
        "wave_col_date": "Data",
        "wave_col_weekday": "Dia da Semana",
        "wave_col_devices": "Dispositivos",
        "confirm_save": "Confirmar e Salvar",
        # Validação RFC
        "rfc_invalid": "Texto muito curto (mínimo 3 caracteres)",
        "rfc_valid": "RFC válido",
        # Body do email Outlook
        "email_body_label": "Corpo do Email:",
        "email_body_placeholder": "Olá,\n\nSegue o agendamento da {{wave}} para {{data}}.\n\nAtenciosamente,",
        # Dias da semana (fallback)
        "monday": "Segunda-feira",
        "tuesday": "Terça-feira",
        "wednesday": "Quarta-feira",
        "thursday": "Quinta-feira",
        "friday": "Sexta-feira",
        "saturday": "Sábado",
        "sunday": "Domingo",
        
        # Países e feriados
        "countries": {
            "BR": "Brasil",
            "US": "Estados Unidos", 
            "AR": "Argentina",
            "CL": "Chile",
            "PE": "Peru",
            "CO": "Colômbia",
            "MX": "México",
            "CA": "Canadá"
        },
        "languages": {
            "pt-BR": "Português (Brasil)",
            "en-US": "English (United States)",
            "es-US": "Español (Estados Unidos)",
            "fr-FR": "Français (France)",
            "de-DE": "Deutsch (Deutschland)"
        },
        
        # Dias da semana
        "monday": "Segunda-feira",
        "tuesday": "Terça-feira",
        "wednesday": "Quarta-feira",
        "thursday": "Quinta-feira",
        "friday": "Sexta-feira",
        "saturday": "Sábado",
        "sunday": "Domingo",
        
        # Tradução dos feriados brasileiros
        "holiday_names": {
            "Universal Fraternization Day": "Dia da Confraternização Universal",
            "Good Friday": "Sexta-feira Santa",
            "Tiradentes' Day": "Dia de Tiradentes",
            "Worker's Day": "Dia do Trabalhador",
            "Independence Day": "Dia da Independência",
            "Our Lady of Aparecida": "Nossa Senhora Aparecida",
            "All Souls' Day": "Dia de Finados",
            "Republic Proclamation Day": "Dia da Proclamação da República",
            "National Day of Zumbi and Black Awareness": "Dia Nacional de Zumbi e da Consciência Negra",
            "Christmas Day": "Natal"
        }
    },
    "en-US": {
        # Interface principal
        "app_title": "Waves Scheduler",
        "data_file_selection": "Data File Selection",
        "select_data_file": "Select Data File (CSV/Excel)",
        "no_file_selected": "No file selected",
        "device_information": "Device Information",
        "total_devices": "Total Devices: {}",
        "recommendations": "Recommendations",
        "ideal_bandwidth": "Ideal Bandwidth (MB):",
        "ideal_waves": "Ideal Waves:",
        "devices_per_wave": "Devices per Wave:",
        "wave_configuration": "Wave Configuration",
        "rfc": "RFC:",
        "bandwidth": "Bandwidth (MB):",
        "number_of_waves": "Number of Waves:",
        "start_date": "Start Date:",
        "outlook_integration": "Add Waves to Outlook 365",
        "timezone": "Time Zone:",
        "timezone_select": "🌍 Select",
        "country": "Country:",
        "language": "Language:",
        "avoid_holidays": "Avoid holidays and weekends",
        "view_holidays": "📅 View Year Holidays",
        "start_time": "Start Time:",
        "end_time": "End Time:",
        "all_day_event": "All Day Event:",
        "location": "Location:",
        "location_placeholder": "Ex: Meeting Room, Online, etc.",
        "required_participants": "Required Participants*:",
        "optional_participants": "Optional Participants:",
        "participants_placeholder": "email@company.com; email2@company.com",
        "generate_waves": "Generate Waves",
        "mb_per_device": "MB per Device:",
        "ready": "Ready",
        
        # Diálogos
        "select_columns": "Select Columns",
        "column_instructions": "Select the columns to include in the output. Required columns are pre-selected.",
        "required_columns": "Required columns: {}",
        "select_all": "Select All",
        "clear_selection": "Clear Selection",
        "ok": "OK",
        "cancel": "Cancel",
        "timezone_selection": "Time Zone Selection",
        "search_timezone": "Search timezone...",
        "holidays_view": "Holidays of {} - {}",
        "national_holidays": "National Holidays",
        "close": "Close",
        "no_holidays_found": "No holidays found for this year.",
        
        # Mensagens
        "supported_formats": "Supported File Formats",
        "formats_info": "You can select the following formats:\n\n- CSV files (.csv)\n- Excel files (.xlsx, .xls)",
        "file_error": "Error loading file",
        "processing_file": "Processing file...",
        "waves_generated": "Waves generated successfully!",
        "generation_error": "Error generating waves",
        "holiday_warning": "⚠️ Warning: The selected date ({}) is a holiday: {}",
        "information": "Information",
        "restart_required": "Please restart the application to apply the language changes.",
        # Generation & progress
        "generating": "Generating...",
        "step_holidays": "Loading holidays...",
        "step_labels": "Generating wave labels...",
        "step_devices": "Distributing devices...",
        "step_excel": "Generating Excel file...",
        "step_done": "Done!",
        "waves_generated_msg": "{} waves generated successfully!",
        "file_label": "File:",
        "outlook_success": "Outlook events created successfully.",
        "outlook_failed": "Outlook events could not be created.",
        "outlook_unavailable": "Microsoft Outlook is not available. The Excel file will be generated, but events will not be created.",
        "outlook_no_participants": "At least one required participant must be specified.",
        "invalid_emails": "The following emails have invalid format:",
        "file_load_error": "Failed to load file. Please check the file format and try again.",
        "col_load_error": "Failed to load file with selected columns.",
        "excel_error": "Failed to generate Excel file. Check the log.",
        "error_occurred": "An error occurred:",
        "save_excel_title": "Save Excel File",
        "save_csv_title": "Save CSV File",
        "export_csv": "Also export CSV",
        "open_file": "Open File",
        "open_folder": "Open Folder",
        "success": "Success",
        "error": "Error",
        "warning": "Warning",
        "holiday_warning_title": "Holiday Warning",
        # History
        "history": "History",
        "history_title": "Generated Schedules History",
        "history_empty": "No schedules generated yet.",
        "history_clear": "Clear History",
        "history_clear_confirm": "Are you sure you want to clear all history?",
        "history_col_date": "Date/Time",
        "history_col_rfc": "RFC",
        "history_col_waves": "Waves",
        "history_col_devices": "Devices",
        "history_col_file": "File",
        "history_open": "Open File",
        # Wave preview
        "wave_preview": "Wave Preview",
        "wave_preview_title": "Preview — Confirm before saving",
        "wave_col_wave": "Wave",
        "wave_col_date": "Date",
        "wave_col_weekday": "Weekday",
        "wave_col_devices": "Devices",
        "confirm_save": "Confirm & Save",
        # RFC validation
        "rfc_invalid": "Text too short (minimum 3 chars)",
        "rfc_valid": "Valid RFC",
        # Outlook email body
        "email_body_label": "Email Body:",
        "email_body_placeholder": "Hello,\n\nPlease find the schedule for {{wave}} on {{date}}.\n\nBest regards,",

        # Países e feriados
        "countries": {
            "BR": "Brazil",
            "US": "United States",
            "AR": "Argentina", 
            "CL": "Chile",
            "PE": "Peru",
            "CO": "Colombia",
            "MX": "Mexico",
            "CA": "Canada"
        },
        "languages": {
            "pt-BR": "Português (Brasil)",
            "en-US": "English (United States)",
            "es-US": "Español (Estados Unidos)",
            "fr-FR": "Français (France)",
            "de-DE": "Deutsch (Deutschland)"
        },
        
        # Dias da semana
        "monday": "Monday",
        "tuesday": "Tuesday",
        "wednesday": "Wednesday",
        "thursday": "Thursday",
        "friday": "Friday",
        "saturday": "Saturday",
        "sunday": "Sunday",
    },
    "es-US": {
        # Interface principal
        "app_title": "Waves Scheduler",
        "data_file_selection": "Selección de Archivo de Datos",
        "select_data_file": "Seleccionar Archivo de Datos (CSV/Excel)",
        "no_file_selected": "Ningún archivo seleccionado",
        "device_information": "Información del Dispositivo",
        "total_devices": "Total de Dispositivos: {}",
        "recommendations": "Recomendaciones",
        "ideal_bandwidth": "Ancho de Banda Ideal (MB):",
        "ideal_waves": "Ondas Ideales:",
        "devices_per_wave": "Dispositivos por Onda:",
        "wave_configuration": "Configuración de Ondas",
        "rfc": "RFC:",
        "bandwidth": "Ancho de Banda (MB):",
        "number_of_waves": "Número de Ondas:",
        "start_date": "Fecha de Inicio:",
        "outlook_integration": "Agregar Ondas a Outlook 365",
        "timezone": "Zona Horaria:",
        "timezone_select": "🌍 Seleccionar",
        "country": "País:",
        "language": "Idioma:",
        "avoid_holidays": "Evitar feriados y fines de semana",
        "view_holidays": "📅 Ver Feriados del Año",
        "start_time": "Hora de Inicio:",
        "end_time": "Hora de Fin:",
        "all_day_event": "Evento de Todo el Día:",
        "location": "Ubicación:",
        "location_placeholder": "Ej: Sala de Reuniones, Online, etc.",
        "required_participants": "Participantes Obligatorios*:",
        "optional_participants": "Participantes Opcionales:",
        "participants_placeholder": "email@dhl.com; email2@dhl.com",
        "generate_waves": "Generar Ondas",
        "mb_per_device": "MB por Dispositivo:",
        "ready": "Listo",
        
        # Diálogos
        "select_columns": "Seleccionar Columnas",
        "column_instructions": "Seleccione las columnas para incluir en la salida. Las columnas obligatorias están preseleccionadas.",
        "required_columns": "Columnas obligatorias: {}",
        "select_all": "Seleccionar Todas",
        "clear_selection": "Limpiar Selección",
        "ok": "OK",
        "cancel": "Cancelar",
        "timezone_selection": "Selección de Zona Horaria",
        "search_timezone": "Buscar zona horaria...",
        "holidays_view": "Feriados de {} - {}",
        "national_holidays": "Feriados Nacionales",
        "close": "Cerrar",
        "no_holidays_found": "No se encontraron feriados para este año.",
        
        # Mensagens
        "supported_formats": "Formatos de Archivo Soportados",
        "formats_info": "Puede seleccionar los siguientes formatos:\n\n- Archivos CSV (.csv)\n- Archivos Excel (.xlsx, .xls)",
        "file_error": "Error al cargar archivo",
        "processing_file": "Procesando archivo...",
        "waves_generated": "¡Ondas generadas exitosamente!",
        "generation_error": "Error al generar ondas",
        "information": "Información",
        "restart_required": "Por favor, reinicie la aplicación para aplicar los cambios de idioma.",
        "generating": "Generando...",
        "step_holidays": "Cargando festivos...",
        "step_labels": "Generando etiquetas de waves...",
        "step_devices": "Distribuyendo dispositivos...",
        "step_excel": "Generando archivo Excel...",
        "step_done": "¡Completado!",
        "waves_generated_msg": "¡{} waves generadas con éxito!",
        "file_label": "Archivo:",
        "outlook_success": "Eventos de Outlook creados con éxito.",
        "outlook_failed": "Los eventos de Outlook no pudieron ser creados.",
        "outlook_unavailable": "Microsoft Outlook no está disponible. Se generará el archivo Excel, pero no se crearán los eventos.",
        "outlook_no_participants": "Se debe especificar al menos un participante obligatorio.",
        "invalid_emails": "Los siguientes correos tienen formato inválido:",
        "file_load_error": "Error al cargar el archivo. Verifique el formato e intente de nuevo.",
        "col_load_error": "Error al cargar el archivo con las columnas seleccionadas.",
        "excel_error": "Error al generar el archivo Excel. Verifique el registro.",
        "error_occurred": "Ocurrió un error:",
        "save_excel_title": "Guardar Archivo Excel",
        "save_csv_title": "Guardar Archivo CSV",
        "export_csv": "También exportar CSV",
        "open_file": "Abrir Archivo",
        "open_folder": "Abrir Carpeta",
        "success": "Éxito",
        "error": "Error",
        "warning": "Advertencia",
        "holiday_warning_title": "Aviso de Festivo",
        "history": "Historial",
        "history_title": "Historial de Schedules Generados",
        "history_empty": "No hay schedules generados aún.",
        "history_clear": "Limpiar Historial",
        "history_clear_confirm": "¿Está seguro de que desea borrar todo el historial?",
        "history_col_date": "Fecha/Hora",
        "history_col_rfc": "RFC",
        "history_col_waves": "Waves",
        "history_col_devices": "Dispositivos",
        "history_col_file": "Archivo",
        "history_open": "Abrir Archivo",
        "wave_preview": "Vista Previa de Waves",
        "wave_preview_title": "Vista Previa — Confirmar antes de guardar",
        "wave_col_wave": "Wave",
        "wave_col_date": "Fecha",
        "wave_col_weekday": "Día de la Semana",
        "wave_col_devices": "Dispositivos",
        "confirm_save": "Confirmar y Guardar",
        "rfc_invalid": "Texto muito curto (mínimo 3 caracteres)",
        "rfc_valid": "RFC válido",
        "email_body_label": "Cuerpo del Email:",
        "email_body_placeholder": "Hola,\n\nAdjunto el cronograma de {{wave}} para {{fecha}}.\n\nAtentamente,",
        "holiday_warning": "⚠️ Advertencia: La fecha seleccionada ({}) es un feriado: {}",
        "information": "Información",
        "restart_required": "Por favor, reinicie la aplicación para aplicar los cambios de idioma.",
        
        # Países e feriados
        "countries": {
            "BR": "Brasil",
            "US": "Estados Unidos",
            "AR": "Argentina",
            "CL": "Chile", 
            "PE": "Perú",
            "CO": "Colombia",
            "MX": "México",
            "CA": "Canadá"
        },
        "languages": {
            "pt-BR": "Português (Brasil)",
            "en-US": "English (United States)",
            "es-US": "Español (Estados Unidos)",
            "fr-FR": "Français (France)",
            "de-DE": "Deutsch (Deutschland)"
        },
        
        # Dias da semana
        "monday": "Lunes",
        "tuesday": "Martes",
        "wednesday": "Miércoles",
        "thursday": "Jueves",
        "friday": "Viernes",
        "saturday": "Sábado",
        "sunday": "Domingo",
        
        # Tradução dos feriados brasileiros para espanhol
        "holiday_names": {
            "Universal Fraternization Day": "Día de la Confraternización Universal",
            "Good Friday": "Viernes Santo",
            "Tiradentes' Day": "Día de Tiradentes",
            "Worker's Day": "Día del Trabajador",
            "Independence Day": "Día de la Independencia",
            "Our Lady of Aparecida": "Nuestra Señora Aparecida",
            "All Souls' Day": "Día de los Difuntos",
            "Republic Proclamation Day": "Día de la Proclamación de la República",
            "National Day of Zumbi and Black Awareness": "Día Nacional de Zumbi y la Conciencia Negra",
            "Christmas Day": "Navidad",
            # Feriados argentinos
            "New Year's Day": "Año Nuevo",
            "Carnival Monday": "Lunes de Carnaval",
            "Carnival Tuesday": "Martes de Carnaval",
            "National Day of Remembrance for Truth and Justice": "Día Nacional de la Memoria por la Verdad y la Justicia",
            "Veteran's Day and the Fallen in the Malvinas War": "Día del Veterano y de los Caídos en la Guerra de Malvinas",
            "Maundy Thursday": "Jueves Santo",
            "Labor Day": "Día del Trabajador",
            "Bridge Public Holiday": "Feriado con Fines Turísticos",
            "May Revolution Day": "Día de la Revolución de Mayo",
            "Pass to the Immortality of General Don Martín Miguel de Güemes": "Paso a la Inmortalidad del General Don Martín Miguel de Güemes",
            "Pass to the Immortality of General Don Manuel Belgrano": "Paso a la Inmortalidad del General Don Manuel Belgrano",
            "Pass to the Immortality of General Don José de San Martín": "Paso a la Inmortalidad del General Don José de San Martín",
            "Respect for Cultural Diversity Day": "Día del Respeto a la Diversidad Cultural",
            "National Sovereignty Day": "Día de la Soberanía Nacional",
            "Immaculate Conception": "Día de la Inmaculada Concepción",
            # Feriados chilenos
            "Holy Saturday": "Sábado Santo",
            "Navy Day": "Día de la Armada",
            "National Day of Indigenous Peoples": "Día Nacional de los Pueblos Indígenas",
            "Saint Peter and Saint Paul's Day": "Día de San Pedro y San Pablo",
            "Our Lady of Mount Carmel": "Nuestra Señora del Carmen",
            "Assumption Day": "Día de la Asunción",
            "Army Day": "Día del Ejército",
            "Meeting of Two Worlds' Day": "Día del Encuentro de Dos Mundos",
            "Reformation Day": "Día de la Reforma",
            "All Saints' Day": "Día de Todos los Santos",
            # Feriados peruanos
            "Easter Sunday": "Domingo de Pascua",
            "Battle of Arica and Flag Day": "Día de la Batalla de Arica y Día de la Bandera",
            "Peruvian Air Force Day": "Día de la Fuerza Aérea del Perú",
            "Great Military Parade Day": "Día de la Gran Parada Militar",
            "Battle of Junín Day": "Día de la Batalla de Junín",
            "Rose of Lima Day": "Día de Santa Rosa de Lima",
            "Battle of Angamos Day": "Día de la Batalla de Angamos",
            "Battle of Ayacucho Day": "Día de la Batalla de Ayacucho"
        }
    },
    "fr-FR": {
        # Interface principale
        "app_title": "Planificateur de Vagues",
        "data_file_selection": "Sélection de Fichier de Données",
        "select_data_file": "Sélectionner un Fichier de Données (CSV/Excel)",
        "no_file_selected": "Aucun fichier sélectionné",
        "device_information": "Informations sur les Appareils",
        "total_devices": "Total d'Appareils: {}",
        "recommendations": "Recommandations",
        "ideal_bandwidth": "Bande Passante Idéale (MB):",
        "ideal_waves": "Vagues Idéales:",
        "devices_per_wave": "Appareils par Vague:",
        "wave_configuration": "Configuration des Vagues",
        "rfc": "RFC:",
        "bandwidth": "Bande Passante (MB):",
        "number_of_waves": "Nombre de Vagues:",
        "start_date": "Date de Début:",
        "outlook_integration": "Ajouter les Vagues à Outlook 365",
        "timezone": "Fuseau Horaire:",
        "timezone_select": "🌍 Sélectionner",
        "country": "Pays:",
        "language": "Langue:",
        "avoid_holidays": "Éviter les jours fériés et les week-ends",
        "view_holidays": "📅 Voir les Jours Fériés de l'Année",
        "start_time": "Heure de Début:",
        "end_time": "Heure de Fin:",
        "all_day_event": "Événement Toute la Journée:",
        "location": "Lieu:",
        "location_placeholder": "Ex: Salle de Réunion, En ligne, etc.",
        "required_participants": "Participants Obligatoires*:",
        "optional_participants": "Participants Optionnels:",
        "participants_placeholder": "email@entreprise.com; email2@entreprise.com",
        "generate_waves": "Générer les Vagues",
        "mb_per_device": "Mo par Appareil:",
        "ready": "Prêt",
        
        # Dialogues
        "select_columns": "Sélectionner les Colonnes",
        "column_instructions": "Sélectionnez les colonnes à inclure dans la sortie. Les colonnes obligatoires sont présélectionnées.",
        "required_columns": "Colonnes obligatoires: {}",
        "select_all": "Tout Sélectionner",
        "clear_selection": "Effacer la Sélection",
        "ok": "OK",
        "cancel": "Annuler",
        "timezone_selection": "Sélection du Fuseau Horaire",
        "search_timezone": "Rechercher un fuseau horaire...",
        "holidays_view": "Jours Fériés de {} - {}",
        "national_holidays": "Jours Fériés Nationaux",
        "close": "Fermer",
        "no_holidays_found": "Aucun jour férié trouvé pour cette année.",
        "information": "Information",
        "restart_required": "Veuillez redémarrer l'application pour appliquer les changements de langue.",
        
        # Messages
        "supported_formats": "Formats de Fichier Pris en Charge",
        "formats_info": "Vous pouvez sélectionner les formats suivants:\\n\\n- Fichiers CSV (.csv)\\n- Fichiers Excel (.xlsx, .xls)",
        "file_error": "Erreur lors du chargement du fichier",
        "processing_file": "Traitement du fichier...",
        "waves_generated": "Vagues générées avec succès!",
        "generation_error": "Erreur lors de la génération des vagues",
        "holiday_warning": "⚠️ Attention: La date sélectionnée ({}) est un jour férié: {}",
        "information": "Information",
        "restart_required": "Veuillez redémarrer l'application pour appliquer les modifications de langue.",
        "generating": "Génération en cours...",
        "step_holidays": "Chargement des jours fériés...",
        "step_labels": "Génération des étiquettes de vagues...",
        "step_devices": "Distribution des appareils...",
        "step_excel": "Génération du fichier Excel...",
        "step_done": "Terminé!",
        "waves_generated_msg": "{} vagues générées avec succès!",
        "file_label": "Fichier:",
        "outlook_success": "Événements Outlook créés avec succès.",
        "outlook_failed": "Les événements Outlook n'ont pas pu être créés.",
        "outlook_unavailable": "Microsoft Outlook n'est pas disponible. Le fichier Excel sera généré, mais les événements ne seront pas créés.",
        "outlook_no_participants": "Au moins un participant obligatoire doit être spécifié.",
        "invalid_emails": "Les emails suivants ont un format invalide:",
        "file_load_error": "Échec du chargement du fichier. Vérifiez le format et réessayez.",
        "col_load_error": "Échec du chargement avec les colonnes sélectionnées.",
        "excel_error": "Échec de la génération du fichier Excel. Vérifiez le journal.",
        "error_occurred": "Une erreur s'est produite:",
        "save_excel_title": "Enregistrer le fichier Excel",
        "save_csv_title": "Enregistrer le fichier CSV",
        "export_csv": "Exporter aussi en CSV",
        "open_file": "Ouvrir le Fichier",
        "open_folder": "Ouvrir le Dossier",
        "success": "Succès",
        "error": "Erreur",
        "warning": "Avertissement",
        "holiday_warning_title": "Avertissement de Jour Férié",
        "history": "Historique",
        "history_title": "Historique des Plannings Générés",
        "history_empty": "Aucun planning généré.",
        "history_clear": "Effacer l'historique",
        "history_clear_confirm": "Êtes-vous sûr de vouloir effacer tout l'historique?",
        "history_col_date": "Date/Heure",
        "history_col_rfc": "RFC",
        "history_col_waves": "Vagues",
        "history_col_devices": "Appareils",
        "history_col_file": "Fichier",
        "history_open": "Ouvrir le Fichier",
        "wave_preview": "Aperçu des Vagues",
        "wave_preview_title": "Aperçu — Confirmer avant d'enregistrer",
        "wave_col_wave": "Vague",
        "wave_col_date": "Date",
        "wave_col_weekday": "Jour de la Semaine",
        "wave_col_devices": "Appareils",
        "confirm_save": "Confirmer et Enregistrer",
        "rfc_invalid": "Texte trop court (min. 3 caractères)",
        "rfc_valid": "RFC valide",
        "email_body_label": "Corps du Email:",
        "email_body_placeholder": "Bonjour,\n\nVeuillez trouver le calendrier de {{vague}} pour {{date}}.\n\nCordialement,",
        
        # Pays et jours fériés
        "countries": {
            "BR": "Brésil",
            "US": "États-Unis",
            "AR": "Argentine",
            "CL": "Chili",
            "PE": "Pérou",
            "CO": "Colombie",
            "MX": "Mexique",
            "CA": "Canada"
        },
        "languages": {
            "pt-BR": "Português (Brasil)",
            "en-US": "English (United States)",
            "es-US": "Español (Estados Unidos)",
            "fr-FR": "Français (France)",
            "de-DE": "Deutsch (Deutschland)"
        },
        
        # Jours de la semaine
        "monday": "Lundi",
        "tuesday": "Mardi",
        "wednesday": "Mercredi",
        "thursday": "Jeudi",
        "friday": "Vendredi",
        "saturday": "Samedi",
        "sunday": "Dimanche"
    },
    "de-DE": {
        # Hauptschnittstelle
        "app_title": "Wellen-Planer",
        "data_file_selection": "Datendatei-Auswahl",
        "select_data_file": "Datendatei Auswählen (CSV/Excel)",
        "no_file_selected": "Keine Datei ausgewählt",
        "device_information": "Geräteinformationen",
        "total_devices": "Gesamtanzahl der Geräte: {}",
        "recommendations": "Empfehlungen",
        "ideal_bandwidth": "Ideale Bandbreite (MB):",
        "ideal_waves": "Ideale Wellen:",
        "devices_per_wave": "Geräte pro Welle:",
        "wave_configuration": "Wellen-Konfiguration",
        "rfc": "RFC:",
        "bandwidth": "Bandbreite (MB):",
        "number_of_waves": "Anzahl der Wellen:",
        "start_date": "Startdatum:",
        "outlook_integration": "Wellen zu Outlook 365 hinzufügen",
        "timezone": "Zeitzone:",
        "timezone_select": "🌍 Auswählen",
        "country": "Land:",
        "language": "Sprache:",
        "avoid_holidays": "Feiertage und Wochenenden vermeiden",
        "view_holidays": "📅 Feiertage des Jahres Anzeigen",
        "start_time": "Startzeit:",
        "end_time": "Endzeit:",
        "all_day_event": "Ganztägiges Ereignis:",
        "location": "Ort:",
        "location_placeholder": "Z.B.: Besprechungsraum, Online, usw.",
        "required_participants": "Erforderliche Teilnehmer*:",
        "optional_participants": "Optionale Teilnehmer:",
        "participants_placeholder": "email@firma.com; email2@firma.com",
        "generate_waves": "Wellen Generieren",
        "mb_per_device": "MB pro Gerät:",
        "ready": "Bereit",
        
        # Dialoge
        "select_columns": "Spalten Auswählen",
        "column_instructions": "Wählen Sie die Spalten aus, die in der Ausgabe enthalten sein sollen. Erforderliche Spalten sind vorausgewählt.",
        "required_columns": "Erforderliche Spalten: {}",
        "select_all": "Alle Auswählen",
        "clear_selection": "Auswahl Löschen",
        "ok": "OK",
        "cancel": "Abbrechen",
        "timezone_selection": "Zeitzonenauswahl",
        "search_timezone": "Zeitzone suchen...",
        "holidays_view": "Feiertage von {} - {}",
        "national_holidays": "Nationale Feiertage",
        "close": "Schließen",
        "no_holidays_found": "Keine Feiertage für dieses Jahr gefunden.",
        "information": "Information",
        "restart_required": "Bitte starten Sie die Anwendung neu, um die Sprachänderungen anzuwenden.",
        
        # Nachrichten
        "supported_formats": "Unterstützte Dateiformate",
        "formats_info": "Sie können die folgenden Formate auswählen:\\n\\n- CSV-Dateien (.csv)\\n- Excel-Dateien (.xlsx, .xls)",
        "file_error": "Fehler beim Laden der Datei",
        "processing_file": "Datei wird verarbeitet...",
        "waves_generated": "Wellen erfolgreich generiert!",
        "generation_error": "Fehler beim Generieren der Wellen",
        "information": "Information",
        "restart_required": "Bitte starten Sie die Anwendung neu, um die Sprachänderungen zu übernehmen.",
        "generating": "Wird generiert...",
        "step_holidays": "Feiertage werden geladen...",
        "step_labels": "Wave-Bezeichnungen werden generiert...",
        "step_devices": "Geräte werden verteilt...",
        "step_excel": "Excel-Datei wird generiert...",
        "step_done": "Fertig!",
        "waves_generated_msg": "{} Waves erfolgreich generiert!",
        "file_label": "Datei:",
        "outlook_success": "Outlook-Ereignisse erfolgreich erstellt.",
        "outlook_failed": "Outlook-Ereignisse konnten nicht erstellt werden.",
        "outlook_unavailable": "Microsoft Outlook ist nicht verfügbar. Die Excel-Datei wird generiert, aber keine Ereignisse erstellt.",
        "outlook_no_participants": "Mindestens ein erforderlicher Teilnehmer muss angegeben werden.",
        "invalid_emails": "Folgende E-Mails haben ein ungültiges Format:",
        "file_load_error": "Datei konnte nicht geladen werden. Bitte Format überprüfen.",
        "col_load_error": "Datei mit ausgewählten Spalten konnte nicht geladen werden.",
        "excel_error": "Excel-Datei konnte nicht generiert werden. Protokoll prüfen.",
        "error_occurred": "Ein Fehler ist aufgetreten:",
        "save_excel_title": "Excel-Datei speichern",
        "save_csv_title": "CSV-Datei speichern",
        "export_csv": "Auch als CSV exportieren",
        "open_file": "Datei öffnen",
        "open_folder": "Ordner öffnen",
        "success": "Erfolg",
        "error": "Fehler",
        "warning": "Warnung",
        "holiday_warning_title": "Feiertagswarnung",
        "history": "Verlauf",
        "history_title": "Verlauf der generierten Zeitpläne",
        "history_empty": "Noch keine Zeitpläne generiert.",
        "history_clear": "Verlauf löschen",
        "history_clear_confirm": "Sind Sie sicher, dass Sie den gesamten Verlauf löschen möchten?",
        "history_col_date": "Datum/Uhrzeit",
        "history_col_rfc": "RFC",
        "history_col_waves": "Wellen",
        "history_col_devices": "Geräte",
        "history_col_file": "Datei",
        "history_open": "Datei öffnen",
        "wave_preview": "Wellenvorschau",
        "wave_preview_title": "Vorschau — Vor dem Speichern bestätigen",
        "wave_col_wave": "Welle",
        "wave_col_date": "Datum",
        "wave_col_weekday": "Wochentag",
        "wave_col_devices": "Geräte",
        "confirm_save": "Bestätigen & Speichern",
        "rfc_invalid": "Text zu kurz (min. 3 Zeichen)",
        "rfc_valid": "Gültiger RFC",
        "email_body_label": "E-Mail-Text:",
        "email_body_placeholder": "Hallo,\n\nbitte beachten Sie den Zeitplan für {{welle}} am {{datum}}.\n\nMit freundlichen Grüßen,",
        "holiday_warning": "⚠️ Warnung: Das ausgewählte Datum ({}) ist ein Feiertag: {}",
        
        # Länder und Feiertage
        "countries": {
            "BR": "Brasilien",
            "US": "Vereinigte Staaten",
            "AR": "Argentinien",
            "CL": "Chile",
            "PE": "Peru",
            "CO": "Kolumbien",
            "MX": "Mexiko",
            "CA": "Kanada"
        },
        "languages": {
            "pt-BR": "Português (Brasil)",
            "en-US": "English (United States)",
            "es-US": "Español (Estados Unidos)",
            "fr-FR": "Français (France)",
            "de-DE": "Deutsch (Deutschland)"
        },
        
        # Wochentage
        "monday": "Montag",
        "tuesday": "Dienstag",
        "wednesday": "Mittwoch",
        "thursday": "Donnerstag",
        "friday": "Freitag",
        "saturday": "Samstag",
        "sunday": "Sonntag"
    }
}

CONFIG_FILE = "config.json"

def load_config():
    """
    Load configuration from config.json file.
    If the file doesn't exist, create it with default values.
    
    Returns:
        dict: Configuration dictionary
    """
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                logger.info("Configuration loaded successfully")
                return config
        else:
            save_config(DEFAULT_CONFIG)
            logger.info("Created default configuration file")
            return DEFAULT_CONFIG
    except Exception as e:
        logger.error(f"Error loading configuration: {str(e)}")
        return DEFAULT_CONFIG

def save_config(config):
    """
    Save configuration to config.json file.
    
    Args:
        config (dict): Configuration dictionary to save
    """
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
        logger.info("Configuration saved successfully")
    except Exception as e:
        logger.error(f"Error saving configuration: {str(e)}")

def get_timezone_list():
    """
    Get a list of common timezones.
    
    Returns:
        list: List of timezone strings
    """
    # Return a subset of common timezones or all timezones
    return sorted(pytz.all_timezones)

@lru_cache(maxsize=64)
def get_country_holidays(year, country_code):
    """
    Get holidays for a specific country and year (national holidays only).
    Resultados são cacheados para evitar recargas repetidas.

    Args:
        year (int): Year to get holidays for
        country_code (str): Country code (BR, US, AR, CL, PE, CO, MX, CA)

    Returns:
        tuple: Tuple of (date, name) pairs (hashable for lru_cache)
    """
    try:
        country_map = {
            "BR": holidays.Brazil,
            "US": holidays.UnitedStates,
            "AR": holidays.Argentina,
            "CL": holidays.Chile,
            "PE": holidays.Peru,
            "CO": holidays.Colombia,
            "MX": holidays.Mexico,
            "CA": holidays.Canada,
        }
        cls = country_map.get(country_code)
        if cls is None:
            return ()
        h = cls(years=year)
        return tuple(sorted(h.items()))
    except Exception as e:
        logger.error(f"Error getting holidays for {country_code}: {str(e)}")
        return ()


def get_country_holidays_dict(year, country_code):
    """
    Wrapper que retorna um dict a partir do cache (compatibilidade).

    Returns:
        dict: Dictionary with date as key and holiday name as value
    """
    return dict(get_country_holidays(year, country_code))

def get_brazilian_holidays(year):
    """
    Get Brazilian holidays for a specific year (kept for compatibility).

    Args:
        year (int): Year to get holidays for

    Returns:
        dict: Dictionary with date as key and holiday name as value
    """
    return get_country_holidays_dict(year, "BR")

def is_holiday(date, country_code):
    """
    Check if a date is a holiday in the specified country.

    Args:
        date (datetime.date): Date to check
        country_code (str): Country code

    Returns:
        tuple: (is_holiday: bool, holiday_name: str or None)
    """
    try:
        year = date.year
        holidays_dict = get_country_holidays_dict(year, country_code)

        if date in holidays_dict:
            return True, holidays_dict[date]
        return False, None
    except Exception as e:
        logger.error(f"Error checking holiday: {str(e)}")
        return False, None

def get_translation(key, language="pt-BR", *args, **kwargs):
    """
    Get translated text for the specified language.

    Args:
        key (str): Translation key
        language (str): Language code
        *args: Positional format arguments
        **kwargs: Keyword format arguments

    Returns:
        str: Translated text
    """
    try:
        if language not in TRANSLATIONS:
            language = "pt-BR"

        translation = TRANSLATIONS[language].get(key, key)
        if args:
            return translation.format(*args)
        if kwargs:
            return translation.format(**kwargs)
        return translation
    except Exception as e:
        logger.error(f"Error getting translation: {str(e)}")
        return key

def get_supported_languages():
    """
    Get list of supported languages.
    
    Returns:
        dict: Dictionary with language codes and names
    """
    current_lang = DEFAULT_CONFIG.get("language", "pt-BR")
    return TRANSLATIONS.get(current_lang, TRANSLATIONS["pt-BR"]).get("languages", {
        "pt-BR": "Português (Brasil)",
        "en-US": "English (United States)",
        "es-US": "Español (Estados Unidos)"
    })

def get_supported_countries_translated(language="pt-BR"):
    """
    Get list of supported countries translated to the specified language.
    
    Args:
        language (str): Language code
    
    Returns:
        dict: Dictionary with country codes and translated names
    """
    return TRANSLATIONS.get(language, TRANSLATIONS["pt-BR"]).get("countries", {
        "BR": "Brasil",
        "US": "Estados Unidos",
        "AR": "Argentina",
        "CL": "Chile",
        "PE": "Peru"
    })

def get_supported_countries():
    """
    Get list of supported countries with their codes.
    
    Returns:
        dict: Dictionary with country codes and names in Portuguese (default)
    """
    return {
        "BR": "Brasil",
        "US": "Estados Unidos",
        "AR": "Argentina",
        "CL": "Chile",
        "PE": "Peru",
        "CO": "Colômbia",
        "MX": "México",
        "CA": "Canadá"
    }

def is_business_day(date, country_code="BR"):
    """
    Check if a date is a business day (not weekend or holiday).

    Args:
        date (datetime.date): Date to check
        country_code (str): Country code to check holidays for (default: 'BR')

    Returns:
        bool: True if it's a business day
    """
    try:
        # Check if it's weekend
        if date.weekday() >= 5:  # Saturday = 5, Sunday = 6
            return False

        # Check holiday for the configured country
        holidays_dict = get_country_holidays_dict(date.year, country_code)
        return date not in holidays_dict
    except Exception as e:
        logger.error(f"Error checking business day: {str(e)}")
        return True  # Default to business day if error

def get_next_business_day(date, country_code="BR"):
    """
    Get the next business day after the given date.

    Args:
        date (datetime.date): Starting date
        country_code (str): Country code to check holidays for (default: 'BR')

    Returns:
        datetime.date: Next business day
    """
    next_date = date + timedelta(days=1)
    while not is_business_day(next_date, country_code):
        next_date += timedelta(days=1)
    return next_date

def get_country_states():
    """
    Get list of Brazilian states for holiday selection.
    
    Returns:
        dict: Dictionary with state codes and names
    """
    return {
        'AC': 'Acre',
        'AL': 'Alagoas', 
        'AP': 'Amapá',
        'AM': 'Amazonas',
        'BA': 'Bahia',
        'CE': 'Ceará',
        'DF': 'Distrito Federal',
        'ES': 'Espírito Santo',
        'GO': 'Goiás',
        'MA': 'Maranhão',
        'MT': 'Mato Grosso',
        'MS': 'Mato Grosso do Sul',
        'MG': 'Minas Gerais',
        'PA': 'Pará',
        'PB': 'Paraíba',
        'PR': 'Paraná',
        'PE': 'Pernambuco',
        'PI': 'Piauí',
        'RJ': 'Rio de Janeiro',
        'RN': 'Rio Grande do Norte',
        'RS': 'Rio Grande do Sul',
        'RO': 'Rondônia',
        'RR': 'Roraima',
        'SC': 'Santa Catarina',
        'SP': 'São Paulo',
        'SE': 'Sergipe',
        'TO': 'Tocantins'
    }

def translate_holiday_name(holiday_name, language="pt-BR"):
    """
    Translate holiday name to the specified language.
    
    Args:
        holiday_name (str): Original holiday name in English
        language (str): Target language code
    
    Returns:
        str: Translated holiday name or original if translation not found
    """
    try:
        holiday_translations = TRANSLATIONS.get(language, TRANSLATIONS["pt-BR"]).get("holiday_names", {})
        return holiday_translations.get(holiday_name, holiday_name)
    except Exception as e:
        logger.error(f"Error translating holiday name '{holiday_name}': {str(e)}")