"""
app/outlook.py — Integração com Outlook via COM ou .ics

Estratégia:
  1. Tenta usar automação COM (Outlook clássico instalado)
  2. Se COM não disponível → gera arquivo .ics e abre no Outlook
     O Outlook (qualquer versão, incluindo web/novo) importa .ics automaticamente.
"""
import os
import logging
import tempfile
import subprocess
from datetime import datetime
from uuid import uuid4

logger = logging.getLogger(__name__)

# ── Constantes COM ─────────────────────────────────────────────────────────────
_OL_APPOINTMENT_ITEM = 0
_OL_MEETING          = 1
_OL_REQUIRED         = 1
_OL_OPTIONAL         = 2


# ─────────────────────────────────────────────────────────────────────────────
# Helpers internos
# ─────────────────────────────────────────────────────────────────────────────

def _to_com_datetime(dt: datetime):
    """datetime Python → pywintypes.Time (COM-safe)."""
    try:
        import pywintypes
        return pywintypes.Time(dt.timetuple())
    except Exception:
        return dt.strftime("%Y-%m-%d %H:%M:%S")


def _is_com_available() -> bool:
    """Retorna True só se o Outlook clássico COM estiver registrado."""
    try:
        import win32com.client
        win32com.client.Dispatch("Outlook.Application")
        return True
    except Exception:
        return False


# ─────────────────────────────────────────────────────────────────────────────
# API pública
# ─────────────────────────────────────────────────────────────────────────────

def is_outlook_available() -> bool:
    """
    Verifica se a integração Outlook está disponível de QUALQUER forma:
    - Outlook clássico COM, OU
    - Arquivo .ics pode ser aberto (qualquer Outlook)
    Sempre retorna True em Windows, pois .ics sempre pode ser gerado.
    """
    return True   # .ics funciona em qualquer ambiente


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
    Cria eventos no Outlook para cada wave.

    Tenta COM primeiro; se não disponível, gera .ics e abre no Outlook.
    """
    if required_participants is None:
        required_participants = []
    if optional_participants is None:
        optional_participants = []

    if _is_com_available():
        logger.info("Outlook COM disponível — usando automação COM.")
        return _create_via_com(
            wave_labels, start_time, end_time, rfc, all_day,
            location, required_participants, optional_participants, email_body,
        )
    else:
        logger.info("Outlook COM indisponível (Novo Outlook?) — usando .ics.")
        return _create_via_ics(
            wave_labels, start_time, end_time, rfc, all_day,
            location, required_participants, optional_participants, email_body,
        )


# ─────────────────────────────────────────────────────────────────────────────
# Implementação COM (Outlook clássico)
# ─────────────────────────────────────────────────────────────────────────────

def _create_via_com(
    wave_labels, start_time, end_time, rfc, all_day,
    location, required_participants, optional_participants, email_body,
) -> bool:
    try:
        import win32com.client
        # EnsureDispatch = early binding: gera type info e resolve propriedades
        # corretamente, incluindo Location que falha no late binding (Dispatch).
        try:
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        except Exception:
            # Fallback: late binding se o gencache falhar
            outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        logger.error("Falha ao conectar ao Outlook COM: %s", e)
        return False

    has_participants = bool(required_participants or optional_participants)
    created = failed = 0

    for i, wave_label in enumerate(wave_labels, 1):
        try:
            parts = wave_label.split(" - ")
            wave_date = datetime.strptime(parts[1].strip(), "%d/%m/%Y")

            appt = outlook.CreateItem(_OL_APPOINTMENT_ITEM)
            appt.Subject = f"Wave {i} - {rfc}" if rfc else f"Wave {i}"
            # Location pode falhar no late binding — protegido
            if location:
                try:
                    appt.Location = location
                except Exception:
                    pass

            if all_day:
                appt.AllDayEvent = True
                appt.Start = wave_date.strftime("%m/%d/%Y")
            else:
                appt.Start = _to_com_datetime(datetime.combine(wave_date.date(), start_time))
                appt.End   = _to_com_datetime(datetime.combine(wave_date.date(), end_time))

            if has_participants:
                appt.MeetingStatus = _OL_MEETING
                for email in required_participants:
                    r = appt.Recipients.Add(email.strip())
                    r.Type = _OL_REQUIRED
                for email in optional_participants:
                    r = appt.Recipients.Add(email.strip())
                    r.Type = _OL_OPTIONAL
                try:
                    appt.Recipients.ResolveAll()
                except Exception:
                    pass  # Emails externos não resolvem — não é fatal

            appt.Body = _build_body(wave_label, rfc, location, wave_date,
                                    required_participants, optional_participants, email_body)
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = 15
            appt.Save()
            created += 1
            logger.info("COM: Wave %d criada (%s)", i, wave_date.strftime("%d/%m/%Y"))

        except Exception as e:
            failed += 1
            logger.error("COM: Falha Wave %d: %s", i, e, exc_info=True)

    logger.info("COM: %d criados, %d falhas", created, failed)
    return failed == 0


# ─────────────────────────────────────────────────────────────────────────────
# Implementação .ics (Novo Outlook / qualquer versão)
# ─────────────────────────────────────────────────────────────────────────────

def _ics_datetime(dt: datetime, all_day: bool = False) -> str:
    """Formata datetime para o formato iCalendar."""
    if all_day:
        return dt.strftime("%Y%m%d")
    return dt.strftime("%Y%m%dT%H%M%S")


def _ics_escape(text: str) -> str:
    """Escapa caracteres especiais para iCalendar."""
    return (text or "")                \
        .replace("\\", "\\\\")         \
        .replace(";", "\\;")           \
        .replace(",", "\\,")           \
        .replace("\n", "\\n")          \
        .replace("\r", "")


def _create_via_ics(
    wave_labels, start_time, end_time, rfc, all_day,
    location, required_participants, optional_participants, email_body,
) -> bool:
    """
    Gera um arquivo .ics com todos os eventos e abre no Outlook.
    O Outlook importa o arquivo automaticamente.
    """
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Waves Scheduler//PT",
        "CALSCALE:GREGORIAN",
        "METHOD:REQUEST" if (required_participants or optional_participants) else "METHOD:PUBLISH",
    ]

    for i, wave_label in enumerate(wave_labels, 1):
        try:
            parts = wave_label.split(" - ")
            wave_date = datetime.strptime(parts[1].strip(), "%d/%m/%Y")

            uid = f"{uuid4()}@waves-scheduler"
            subject = _ics_escape(f"Wave {i} - {rfc}" if rfc else f"Wave {i}")
            loc     = _ics_escape(location or "")
            body    = _ics_escape(
                _build_body(wave_label, rfc, location, wave_date,
                            required_participants, optional_participants, email_body)
            )

            lines.append("BEGIN:VEVENT")
            lines.append(f"UID:{uid}")
            lines.append(f"SUMMARY:{subject}")
            lines.append(f"LOCATION:{loc}")
            lines.append(f"DESCRIPTION:{body}")

            if all_day:
                lines.append(f"DTSTART;VALUE=DATE:{_ics_datetime(wave_date, all_day=True)}")
                lines.append(f"DTEND;VALUE=DATE:{_ics_datetime(wave_date, all_day=True)}")
            else:
                start_dt = datetime.combine(wave_date.date(), start_time)
                end_dt   = datetime.combine(wave_date.date(), end_time)
                lines.append(f"DTSTART:{_ics_datetime(start_dt)}")
                lines.append(f"DTEND:{_ics_datetime(end_dt)}")

            # Participantes
            for email in required_participants:
                lines.append(f"ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:{email.strip()}")
            for email in optional_participants:
                lines.append(f"ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:{email.strip()}")

            lines.append("BEGIN:VALARM")
            lines.append("TRIGGER:-PT15M")
            lines.append("ACTION:DISPLAY")
            lines.append("DESCRIPTION:Lembrete")
            lines.append("END:VALARM")
            lines.append("END:VEVENT")
            logger.info(".ics: Wave %d adicionada (%s)", i, wave_date.strftime("%d/%m/%Y"))

        except Exception as e:
            logger.error(".ics: Falha Wave %d: %s", i, e, exc_info=True)

    lines.append("END:VCALENDAR")

    # Salvar e abrir o arquivo .ics
    try:
        ics_content = "\r\n".join(lines) + "\r\n"
        tmp = tempfile.NamedTemporaryFile(
            suffix=f"_waves_{rfc or 'schedule'}.ics",
            delete=False,
            mode="w",
            encoding="utf-8",
            dir=tempfile.gettempdir(),
        )
        tmp.write(ics_content)
        tmp_path = tmp.name
        tmp.close()

        logger.info(".ics: Arquivo gerado em %s", tmp_path)

        # Abrir com o programa padrão (deve ser o Outlook)
        os.startfile(tmp_path)
        logger.info(".ics: Arquivo aberto no Outlook")
        return True

    except Exception as e:
        logger.error(".ics: Falha ao gerar/abrir .ics: %s", e, exc_info=True)
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Helper compartilhado
# ─────────────────────────────────────────────────────────────────────────────

def _build_body(wave_label, rfc, location, wave_date, req, opt, email_body) -> str:
    if email_body:
        body = email_body
        body = body.replace("{{wave}}", wave_label)
        body = body.replace("{{rfc}}", rfc or "")
        body = body.replace("{{local}}", location or "")
        body = body.replace("{{data}}", wave_date.strftime("%d/%m/%Y"))
        body = body.replace("{{participantes}}", "; ".join(req) if req else "Nenhum")
        return body
    return (
        f"Wave: {wave_label}\n"
        f"RFC: {rfc or 'N/A'}\n"
        f"Local: {location or 'A definir'}\n\n"
        f"Participantes Obrigatórios: {'; '.join(req) if req else 'Nenhum'}\n"
        f"Participantes Opcionais: {'; '.join(opt) if opt else 'Nenhum'}\n\n"
        f"Este é um evento automático criado pelo Waves Scheduler."
    )