from functions_real import *
import requests


def parse_graph_email(msg: dict, access_token: str = None, user_email: str = None) -> dict:
    """Extracts email fields from a Microsoft Graph message object into a structured dict."""
    
    # Extract attachments if available
    attachments_list = []
    if msg.get("hasAttachments", False) and access_token and user_email:
        attachment_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages/{msg['id']}/attachments"
        headers = {"Authorization": f"Bearer {access_token}"}
        res = requests.get(attachment_url, headers=headers)
        if res.status_code == 200:
            attachments_list = [att.get("name") for att in res.json().get("value", [])]

    # Map directly to your target schema
    return {
        "mail_id": str(msg.get("id", "")),
        "local_mail_id": hash(msg.get("id", "")) & 0x7FFFFFFF,  # Safe positive integer ID
        "From": msg.get("sender", {}).get("emailAddress", {}).get("address", ""),
        "To": None,
        "Subject": msg.get("subject", ""),
        "Date": msg.get("receivedDateTime", ""),
        "Body": msg.get("body", {}).get("content", ""),
        "Attachments": attachments_list
    }

# Copy one email object directly from Graph Explorer's JSON response
sample_msg = {
    "id": "AQMkADAwATM3ZmYAZS0zMmI3LTczZTItMDACLTAwCgBGAAADi-LbB56eANROgFgQfIYHDt4HAF7Kwr-eIwhLov4KUa81QBUAAAIBDAAAAF7Kwr-eIwhLov4KUa81QBUABaqkrB4AAAA=",
    "sender": {
                "emailAddress": {
                    "name": "Microsoft account team",
                    "address": "account-security-noreply@accountprotection.microsoft.com"
                }
            },
    "toRecipients": [
                {
                    "emailAddress": {
                        "name": "elyesghouma@outlook.com",
                        "address": "elyesghouma@outlook.com"
                    }
                }
            ],
    "subject": "New app(s) connected to your Microsoft account",
    "receivedDateTime": "2026-08-25T11:33:33Z",
    "body": {
                "contentType": "html",
                "content": "<html dir=\"ltr\"><head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><style type=\"text/css\">\r\n<!--\r\n-->\r\n</style></head><body><table dir=\"ltr\"><tbody><tr><td id=\"i1\" style=\"padding:0; font-family:'Segoe UI Semibold','Segoe UI Bold','Segoe UI','Helvetica Neue Medium',Arial,sans-serif; font-size:17px; color:#707070\">Microsoft account</td></tr><tr><td id=\"i2\" style=\"padding:0; font-family:'Segoe UI Light','Segoe UI','Helvetica Neue Medium',Arial,sans-serif; font-size:41px; color:#2672ec\">New app(s) have access to your data</td></tr><tr><td id=\"i3\" style=\"padding:0; padding-top:25px; font-family:'Segoe UI',Tahoma,Verdana,Arial,sans-serif; font-size:14px; color:#2a2a2a\">Graph Explorer connected to the Microsoft account <a dir=\"ltr\" id=\"iAccount\" class=\"link\" href=\"mailto:el**a@outlook.com\" style=\"color:#2672ec; text-decoration:none\">el**a@outlook.com</a>.</td></tr><tr><td id=\"i4\" style=\"padding:0; padding-top:25px; font-family:'Segoe UI',Tahoma,Verdana,Arial,sans-serif; font-size:14px; color:#2a2a2a\">If you didn't grant this access, please remove the app(s) from your account.</td></tr><tr><td style=\"padding:0; padding-top:25px; font-family:'Segoe UI',Tahoma,Verdana,Arial,sans-serif; font-size:14px; color:#2a2a2a\"><table border=\"0\" cellspacing=\"0\"><tbody><tr><td bgcolor=\"#2672ec\" style=\"background-color:#2672ec; padding-top:5px; padding-right:20px; padding-bottom:5px; padding-left:20px; min-width:50px\"><a id=\"i5\" href=\"https://account.live.com/consent/Manage?fn=email\" style=\"font-family:'Segoe UI Semibold','Segoe UI Bold','Segoe UI','Helvetica Neue Medium',Arial,sans-serif; font-size:14px; text-align:center; text-decoration:none; font-weight:600; letter-spacing:0.02em; color:#fff\">Manage your apps</a></td></tr></tbody></table></td></tr><tr><td id=\"i6\" style=\"padding:0; padding-top:25px; font-family:'Segoe UI',Tahoma,Verdana,Arial,sans-serif; font-size:14px; color:#2a2a2a\">You can also <a id=\"iLink3\" class=\"link\" href=\"https://account.live.com/SecurityNotifications/Update\" style=\"color:#2672ec; text-decoration:none\">opt out</a> or change where you receive security notifications.</td></tr><tr><td id=\"i7\" style=\"padding:0; padding-top:25px; font-family:'Segoe UI',Tahoma,Verdana,Arial,sans-serif; font-size:14px; color:#2a2a2a\">Thanks,</td></tr><tr><td id=\"i8\" style=\"padding:0; font-family:'Segoe UI',Tahoma,Verdana,Arial,sans-serif; font-size:14px; color:#2a2a2a\">The Microsoft account team</td></tr></tbody></table><div lang=\"en\" style=\"margin-top:20px; margin-bottom:10px\"><a class=\"link\" href=\"https://go.microsoft.com/fwlink/?LinkId=521839\">Privacy Statement</a><div style=\"margin-top:10px\">Microsoft Corporation, One Microsoft Way, Redmond, WA 98052</div></div></body></html>"
            },
    "hasAttachments": False,
}

parsed = parse_graph_email(sample_msg)
print(parsed)