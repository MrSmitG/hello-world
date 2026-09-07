"""Placeholder for Outlook Classic integration — implement your send logic here."""

from typing import Optional


def send_outlook_mail(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None,
    from_group_mail: Optional[str] = None,
    signature: Optional[str] = None,
    attachments: Optional[list] = None,
) -> bool:
    """
    Send mail via Outlook Classic from a group mailbox with signature.

    TODO — implement using win32com.client:

        import win32com.client

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = to
        if cc:
            mail.CC = cc
        mail.Subject = subject
        mail.Body = body + "\n\n" + (signature or "")
        # Set SentOnBehalfOfName or account for group mailbox
        # mail.SentOnBehalfOfName = from_group_mail
        for path in attachments or []:
            mail.Attachments.Add(path)
        mail.Send()
        return True
    """
    raise NotImplementedError("Wire up Outlook Classic COM send in this module.")
