import logging
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

def _to_com_datetime(dt):
    """
    Converte um objeto datetime do Python para um formato compatível com o
    COM do Outlook (pywintypes.DateTime), garantindo que as datas sejam
    aceitas corretamente pela interface COM.
    """
    try:
        import pywintypes
        return pywintypes.Time(dt)
    except Exception:
        # Fallback: retornar como string no formato ISO para o COM interpretar
        return dt.strftime("%Y-%m-%d %H:%M:%S")

def is_outlook_available():
    """
    Verifica se o Outlook COM está disponível no sistema.

    Returns:
        bool: True se o Outlook pode ser iniciado, False caso contrário
    """
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook is not None
    except Exception:
        return False

def create_outlook_events(wave_labels, start_time, end_time, rfc, all_day=False, location="",
                          required_participants=None, optional_participants=None, email_body=None):
    """
    Cria eventos no Outlook para cada wave com participantes.

    Args:
        wave_labels (list): Lista de rótulos das waves
        start_time (time): Horário de início
        end_time (time): Horário de término
        rfc (str): Número do RFC
        all_day (bool): Se é um evento de dia inteiro
        location (str): Local do evento
        required_participants (list): Lista de emails dos participantes obrigatórios
        optional_participants (list): Lista de emails dos participantes opcionais
        email_body (str): Corpo do email personalizado. Suporta {{wave}}, {{rfc}}, {{local}},
                          {{participantes}} como placeholders que serão substituídos automaticamente.
                          Se None, usa o corpo padrão.

    Returns:
        bool: True se os eventos foram criados com sucesso, False caso contrário
    """
    try:
        import win32com.client
        # Verificar disponibilidade e conectar ao Outlook
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            logger.error(f"Outlook não está disponível: {str(e)}")
            return False
        
        # Definir valores padrão
        if required_participants is None:
            required_participants = []
        if optional_participants is None:
            optional_participants = []
        
        for i, wave_label in enumerate(wave_labels, 1):
            # Extrair a data do rótulo da wave (formato: "Wave X - dd/mm/yyyy")
            date_str = wave_label.split(" - ")[1]
            wave_date = datetime.strptime(date_str, "%d/%m/%Y")
            
            # Criar o evento
            appointment = outlook.CreateItem(0)  # 0 = olAppointmentItem
            
            # Formato do título: "Wave 1 - RFC", "Wave 2 - RFC", etc.
            appointment.Subject = f"Wave {i} - {rfc}"
            appointment.Location = location or "A definir"
            
            if all_day:
                appointment.AllDayEvent = True
                appointment.Start = wave_date.strftime("%Y-%m-%d")
            else:
                # Combinar data e hora
                start_datetime = datetime.combine(wave_date.date(), start_time)
                end_datetime = datetime.combine(wave_date.date(), end_time)
                
                appointment.Start = _to_com_datetime(start_datetime)
                appointment.End = _to_com_datetime(end_datetime)
            
            # Adicionar participantes obrigatórios
            for email in required_participants:
                recipient = appointment.Recipients.Add(email)
                recipient.Type = 1  # 1 = olRequired (obrigatório)
                
            # Adicionar participantes opcionais
            for email in optional_participants:
                recipient = appointment.Recipients.Add(email)
                recipient.Type = 2  # 2 = olOptional (opcional)
            
            # Resolver os destinatários
            appointment.Recipients.ResolveAll()
            
            # Corpo do email: personalizado ou padrão
            if email_body:
                body = email_body.replace("{{wave}}", wave_label)
                body = body.replace("{{rfc}}", rfc or "")
                body = body.replace("{{local}}", location or "A definir")
                body = body.replace("{{data}}", wave_date.strftime("%d/%m/%Y"))
                body = body.replace("{{participantes}}", "; ".join(required_participants) if required_participants else "Nenhum")
            else:
                body = f"""Wave: {wave_label}
RFC: {rfc}
Local: {location or 'A definir'}

Participantes Obrigatórios: {'; '.join(required_participants) if required_participants else 'Nenhum'}
Participantes Opcionais: {'; '.join(optional_participants) if optional_participants else 'Nenhum'}

Este é um evento automático criado pelo Waves Scheduler."""
            appointment.Body = body
            appointment.ReminderSet = True
            appointment.ReminderMinutesBeforeStart = 15

            # Salvar o evento (não envia convites automaticamente)
            # Para enviar convites, o usuário pode abrir o evento no Outlook e enviar manualmente
            appointment.Save()
            
            logger.info("Evento criado: Wave %d - %s em %s", i, rfc, wave_date.strftime('%d/%m/%Y'))
            
        logger.info(f"Eventos do Outlook criados com sucesso para {len(wave_labels)} waves")
        logger.info(f"Participantes obrigatórios: {len(required_participants)}")
        logger.info(f"Participantes opcionais: {len(optional_participants)}")
        return True
        
    except Exception as e:
        logger.error(f"Erro ao criar eventos no Outlook: {str(e)}", exc_info=True)
        return False