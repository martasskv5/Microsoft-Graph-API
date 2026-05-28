import argparse
import requests
from ms_graph import get_access_token, MS_GRAPH_URL


def send_email(
    sender_email: str,
    recipient_emails: str | list[str],
    subject: str,
    body: str,
    access_token: str | None = None,
) -> bool:
    """
    Send an email using Microsoft Graph API.

    Args:
        access_token (str): Valid Microsoft Graph access token
        sender_email (str): Email address of the sender
        recipient_emails (str | list[str]): Recipient email address(es)
        subject (str): Email subject
        body (str): Email body content

    Returns:
        bool: True if email was sent successfully, False otherwise
    """

    if access_token is None:
        access_token = get_access_token(scopes=["User.Read", "Files.ReadWrite.All", "Sites.Manage.All", "Sites.ReadWrite.All", "Mail.Send", "Mail.Send.Shared"], 
                                        # user=sender_email
                                        )

    endpoint = f'{MS_GRAPH_URL}/users/{sender_email}/sendMail'

    recipients = (
        recipient_emails
        if isinstance(recipient_emails, list)
        else [recipient_emails]
    )

    email_data = {
        'message': {
            'subject': subject,
            'body': {
                'contentType': 'HTML',
                'content': body
            },
            'toRecipients': [{'emailAddress': {'address': recipient}} for recipient in recipients]
        }
    }

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    response = requests.post(endpoint, headers=headers, json=email_data)

    return response.ok


if __name__ == "__main__":
    # Example usage
    # send_email("logs@martassk.onmicrosoft.com",
    #            "admin@0z6sw.onmicrosoft.com", "Test Subject", "Test Body")

    parser = argparse.ArgumentParser(description="Send an email using Microsoft Graph API.")
    parser.add_argument("sender_email", type=str, help="Sender's email address")
    parser.add_argument("recipient_emails", type=str, help="Recipient's email address")
    parser.add_argument("subject", type=str, help="Email subject")
    parser.add_argument("body", type=str, help="Email body content")
    args = parser.parse_args()
    
    # Send the email
    # print(f"Sending email from {args.sender_email} to {args.recipient_emails} with subject '{args.subject}'")
    # print(f"Email body: {args.body}")
    success = send_email(args.sender_email, args.recipient_emails, args.subject, args.body)