# save as test_email.py, run with: python test_email.py
import smtplib
from email.message import EmailMessage

msg = EmailMessage()
msg["From"] = "d.ansong@ucl.ac.uk"
msg["To"] = "d.ansong@ucl.ac.uk"   # send to yourself to verify
msg["Subject"] = "Switch Capacity API — SMTP test"
msg.set_content("If you're reading this, SMTP is working correctly.")

with smtplib.SMTP("smtp-server.ucl.ac.uk", 25, timeout=10) as smtp:
    smtp.send_message(msg)
    print("Email sent OK")