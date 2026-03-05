"""
Diagnóstico de integração Outlook COM
Execute: python diagnostico_outlook.py
"""
import sys
import traceback
from datetime import datetime, time as dtime

print("=" * 60)
print("  DIAGNÓSTICO WAVES SCHEDULER — OUTLOOK COM")
print("=" * 60)

# 1. pywin32 disponível?
print("\n[1] Verificando pywin32...")
try:
    import win32com.client
    import pywintypes
    print("    OK: pywin32 instalado")
except ImportError as e:
    print(f"    ERRO: pywin32 não encontrado — {e}")
    sys.exit(1)

# 2. Conectar ao Outlook
print("\n[2] Conectando ao Outlook...")
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    print(f"    OK: Dispatch OK — {outlook}")
except Exception as e:
    print(f"    ERRO no Dispatch: {e}")
    traceback.print_exc()
    sys.exit(1)

# 3. MAPI namespace
print("\n[3] Abrindo MAPI namespace...")
try:
    ns = outlook.GetNamespace("MAPI")
    print(f"    OK: MAPI OK")
except Exception as e:
    print(f"    AVISO: GetNamespace falhou: {e}")

# 4. Criar AppointmentItem simples (sem participantes)
print("\n[4] Criando AppointmentItem simples (sem participantes)...")
try:
    appt = outlook.CreateItem(0)  # olAppointmentItem
    appt.Subject = "TESTE Waves Scheduler — SEM participantes"
    start_dt = datetime(2026, 3, 10, 9, 0, 0)
    end_dt   = datetime(2026, 3, 10, 10, 0, 0)
    appt.Start = pywintypes.Time(start_dt.timetuple())
    appt.End   = pywintypes.Time(end_dt.timetuple())
    appt.Body  = "Teste diagnóstico"
    appt.Save()
    print("    OK: Evento simples salvo no Outlook!")
except Exception as e:
    print(f"    ERRO: {e}")
    traceback.print_exc()

# 5. Criar AppointmentItem COM participantes (MeetingStatus=1)
print("\n[5] Criando AppointmentItem COM participante...")
TEST_EMAIL = input("    Email para teste (Enter para pular): ").strip()
if TEST_EMAIL:
    try:
        appt2 = outlook.CreateItem(0)
        appt2.Subject = "TESTE Waves Scheduler — COM participante"
        appt2.MeetingStatus = 1  # olMeeting
        appt2.Start = pywintypes.Time(start_dt.timetuple())
        appt2.End   = pywintypes.Time(end_dt.timetuple())
        recip = appt2.Recipients.Add(TEST_EMAIL)
        recip.Type = 1  # olRequired
        print(f"    Resolvendo destinatário {TEST_EMAIL}...")
        resolved = appt2.Recipients.ResolveAll()
        print(f"    ResolveAll() = {resolved}")
        appt2.Body = "Teste diagnóstico com participante"
        appt2.Save()
        print("    OK: Evento com participante salvo!")
    except Exception as e:
        print(f"    ERRO: {e}")
        traceback.print_exc()
else:
    print("    Pulado.")

print("\n" + "=" * 60)
print("  DIAGNÓSTICO CONCLUÍDO")
print("=" * 60)
input("\nPressione Enter para sair...")
