import logging
from datetime import datetime

logger = logging.getLogger(__name__)

# ── Constantes COM do Outlook ──────────────────────────────────────────────────
_OL_APPOINTMENT_ITEM = 0
_OL_MEETING         = 1   # MeetingStatus: transforma em convite de reunião
_OL_REQUIRED        = 1   # Recipient.Type
_OL_OPTIONAL        = 2   # Recipient.Type


def _to_com_datetime(dt: datetime):
    """
    Converte datetime Python → pywintypes.Time (formato nativo COM).
    O Outlook 365 moderno rejeita strings e objetos datetime Python puros.
    """
    try:
        import pywintypes
        return pywintypes.Time(dt.timetuple())
    except Exception:
        # Fallback: string ISO — Outlook ainda consegue interpretar
        return dt.strftime("%Y-%m-%d %H:%M:%S")


def is_outlook_available() -> bool:
    """
    Verifica se o Outlook COM (clássico) está disponível no sistema.
    Retorna False se o 'Novo Outlook' (baseado em web) estiver ativo,
    pois ele não suporta automação COM.
    """
    try:
        import win32com.client
        app = win32com.client.Dispatch("Outlook.Application")
        # Verificar se é uma instância válida com MAPI
        ns = app.GetNamespace("MAPI")
        return ns is not None
    except Exception:
        return False


def create_outlook_events(
    wave_labels,
    start_time,
    end_time,
    rfc,
    all_day=False,
    location="",
    required_participants=None,
    optional_participants=None,
    email_body=None,
) -> bool:
    """
    Cria eventos/reuniões no Outlook para cada wave.

    Compatível com Outlook 365 (clássico) v2408+.
    Não funciona com o 'Novo Outlook' (web-based), que não suporta COM.

    Returns:
        bool: True se TODOS os eventos foram criados, False se algum falhou.
    """
    if required_participants is None:
        required_participants = []
    if optional_participants is None:
        optional_participants = []

    has_participants = bool(required_participants or optional_participants)

    try:
        import win32com.client
    except ImportError:
        logger.error("pywin32 não está instalado — necessário para integração COM.")
        return False

    # ── Conectar ao Outlook ────────────────────────────────────────────────────
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        outlook.GetNamespace("MAPI")          # Valida que o perfil MAPI existe
    except Exception as e:
        logger.error("Outlook não disponível (ou Novo Outlook ativo): %s", e)
        return False

    created = 0
    failed  = 0

    for i, wave_label in enumerate(wave_labels, 1):
        try:
            # ── Extrair data ───────────────────────────────────────────────────
            parts = wave_label.split(" - ")
            if len(parts) < 2:
                raise ValueError(f"Formato inválido de wave_label: '{wave_label}'")
            wave_date = datetime.strptime(parts[1].strip(), "%d/%m/%Y")

            # ── Criar AppointmentItem ──────────────────────────────────────────
            appt = outlook.CreateItem(_OL_APPOINTMENT_ITEM)

            appt.Subject  = f"Wave {i} - {rfc}" if rfc else f"Wave {i}"
            appt.Location = location or ""

            # ── Datas ─────────────────────────────────────────────────────────
            if all_day:
                appt.AllDayEvent = True
                # Data como string no formato MM/DD/YYYY (padrão EUA do COM)
                appt.Start = wave_date.strftime("%m/%d/%Y")
            else:
                start_dt = datetime.combine(wave_date.date(), start_time)
                end_dt   = datetime.combine(wave_date.date(), end_time)
                appt.Start = _to_com_datetime(start_dt)
                appt.End   = _to_com_datetime(end_dt)

            # ── Participantes → converte em Reunião ───────────────────────────
            if has_participants:
                # MeetingStatus = 1 (olMeeting): obrigatório para convites
                appt.MeetingStatus = _OL_MEETING

                for email in required_participants:
                    recip = appt.Recipients.Add(email.strip())
                    recip.Type = _OL_REQUIRED

                for email in optional_participants:
                    recip = appt.Recipients.Add(email.strip())
                    recip.Type = _OL_OPTIONAL

                # ResolveAll() pode retornar False para emails externos — não é fatal
                try:
                    appt.Recipients.ResolveAll()
                except Exception as e_res:
                    logger.warning("ResolveAll() falhou (emails externos?) — continuando: %s", e_res)

            # ── Corpo do email ─────────────────────────────────────────────────
            if email_body:
                body = email_body
                body = body.replace("{{wave}}", wave_label)
                body = body.replace("{{rfc}}", rfc or "")
                body = body.replace("{{local}}", location or "")
                body = body.replace("{{data}}", wave_date.strftime("%d/%m/%Y"))
                body = body.replace(
                    "{{participantes}}",
                    "; ".join(required_participants) if required_participants else "Nenhum",
                )
            else:
                body = (
                    f"Wave: {wave_label}\n"
                    f"RFC: {rfc or 'N/A'}\n"
                    f"Local: {location or 'A definir'}\n\n"
                    f"Participantes Obrigatórios: "
                    f"{'; '.join(required_participants) if required_participants else 'Nenhum'}\n"
                    f"Participantes Opcionais: "
                    f"{'; '.join(optional_participants) if optional_participants else 'Nenhum'}\n\n"
                    f"Este é um evento automático criado pelo Waves Scheduler."
                )

            appt.Body = body
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = 15

            # ── Salvar (não envia convites — usuário envia manualmente) ────────
            appt.Save()
            created += 1
            logger.info("Evento criado: Wave %d (%s)", i, wave_date.strftime("%d/%m/%Y"))

        except Exception as e:
            failed += 1
            logger.error("Falha ao criar Wave %d: %s", i, e, exc_info=True)

    logger.info(
        "Outlook: %d eventos criados, %d falhas (total: %d waves)",
        created, failed, len(wave_labels),
    )
    return failed == 0