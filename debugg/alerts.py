import os
import ssl
import smtplib
from email.message import EmailMessage

GMAIL_USER = 'arnold.morales@scapital.mx'
GMAIL_PASSWORD = 'lpnsujvfmnzsxzwb'

def send_report_email(
    recipients: list,
    subject: str,
    body: str,
    smtp_server: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str
):
    """
    Envía un correo electrónico con múltiples adjuntos (PDF, Excel, Imágenes).
    Incluye detección automática de tipos MIME para evitar filtros de spam.
    """
    # if not send_email:
    #     print("📭 Envío de emails desactivado por configuración.")
    #     return

    msg = EmailMessage()
    msg["From"] = smtp_user
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    # --- Envío Seguro ---
    context = ssl.create_default_context()
    # try:
    # Usamos SMTP_SSL si el puerto es 465, o SMTP + STARTTLS si es 587
    # with smtplib.SMTP(smtp_server, smtp_port) as server:
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        # server.starttls(context=context)
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
        
    print(f"✅ Email enviado exitosamente a {len(recipients)} destinatarios.")
    # except Exception as e:
    #     print(f"❌ Error crítico en el servidor SMTP: {e}")

mails = ['arnold.morales@scapital.mx', 'jose.gonzalez@scapital.mx', 'juan.corona@scapital.mx']
flag = True
if flag:
    err = 'Divios por 0'
    send_report_email(mails, 'Flag Error', err, 'smtp.gmail.com', 465, GMAIL_USER, GMAIL_PASSWORD)