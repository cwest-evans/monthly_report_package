# graph_mailer.py
import base64
import os
import requests
from typing import Iterable, Optional

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = TOKEN_URL.format(tenant_id=tenant_id)
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": GRAPH_SCOPE,
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def send_mail(
    token: str,
    sender_upn: str,
    to_addrs: Iterable[str],
    subject: str,
    html_body: str,
    attachment_path: Optional[str] = None,
    save_to_sent: bool = True,
):
    url = f"https://graph.microsoft.com/v1.0/users/{sender_upn}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    message = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": [{"emailAddress": {"address": a}} for a in to_addrs],
    }

    if attachment_path:
        size = os.path.getsize(attachment_path)
        if size >= 3 * 1024 * 1024:
            raise ValueError(
                f"Attachment is {size/1024/1024:.2f} MB; "
                "simple sendMail attachments must be under ~3 MB. "
                "Use upload-session approach for larger files."
            )

        with open(attachment_path, "rb") as f:
            content = base64.b64encode(f.read()).decode("utf-8")

        message["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": os.path.basename(attachment_path),
            "contentBytes": content
        }]

    payload = {"message": message, "saveToSentItems": save_to_sent}
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    r.raise_for_status()
