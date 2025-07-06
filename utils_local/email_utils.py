# import the packages
import re, os, logging, platform
from typing import List, Dict, Union, Optional

if platform.system() == 'Windows':
    import win32com.client as win32

logging.basicConfig(
    level=logging.INFO,
    format='[%(levelname)s] %(asctime)s — %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)


# define the function of validating the email address
def validate_email(
        email_address: str) -> bool:
    """
    Validates whether the input string is a properly formatted email address.

    Parameters:
        email_address (str): The email address to validate.

    Returns:
        bool: True if the email address is valid, False otherwise.

    Raises:
        TypeError: If the input type is invalid.
    """

    # check the error for input type
    if isinstance(email_address, str) == False:
        raise TypeError("email_address should be a string.")

    # return the result
    return re.match(r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$', email_address) is not None

#---

# define the function of validating a single email address or a list of email addresses and return emails combined together
def format_valid_emails(
        email_addresses: Union[str, list[str]]) -> str:
    """
    Validates a single email address or a list of email addresses and return adress combination.

    Parameters:
        email_address_list (str or list[str]): A single email address as a string, or a list of email addresses.

    Returns:
        str: A single valid email address, or a semicolon-separated string of valid email addresses.

    Raises:
        TypeError: If the input type is invalid.
        ValueError: If any email in the input is not properly formatted.
    """

    # check the error for input type
    if not isinstance(email_addresses, (str, list)):
        raise TypeError("email_address_list should be a string or a list.")

    # check whether the string is a valid email address
    elif isinstance(email_addresses, str):
        if not validate_email(email_addresses):
            raise ValueError(f"Invalid email format: {email_addresses}")
        return email_addresses

    # check whether all the strings in the list are valid email addresses
    else:
        for email_address in email_addresses:
            if not validate_email(email_address):
                raise ValueError(f"Invalid email format: {email_address}")
        return ";".join(email_addresses)

#---

def windows_outlook_send_email(
        email_sender: str,
        email_receiver: Union[str, list[str]],
        email_subject: str,
        email_body: str,
        email_attached: Optional[Union[str, list[str]]] = None,
        email_cc: Optional[Union[str, list[str]]] = None,
        email_bcc: Optional[Union[str, list[str]]] = None
) -> None:

    """
    Send an email using Outlook via COM automation in Windows operation system.

    Parameters:
        email_sender (str): The email address to send on behalf of.
        email_receiver (str or list[str]): One or more recipient email addresses.
        email_subject (str): Subject line of the email.
        email_body (str): Body text of the email.
        email_attached (str or list[str], optional): Path(s) to file(s) to attach.
        email_cc (str or list[str], optional): CC recipients.
        email_bcc (str or list[str], optional): BCC recipients.

    Raises:
        ValueError: If email format is invalid.
        FileNotFoundError: If attachment file(s) cannot be found.

    Returns:
        None
    """

    # create a windows outlook item
    mail = win32.Dispatch('outlook.application').CreateItem(0)

    # validate email addresses for sender, receiver, cc and bcc
    if not isinstance(email_sender, str):
        raise ValueError(f"Invalid email format: {email_sender}")
    else:
        mail.SentOnBehalfOfName = format_valid_emails(email_sender)
    mail.To = format_valid_emails(email_receiver)
    if email_cc:
        mail.CC = format_valid_emails(email_cc)
    if email_bcc:
        mail.BCC = format_valid_emails(email_bcc)

    # check email attachments
    if email_attached:
        if isinstance(email_attached, list):
            for f in email_attached:
                file_path = os.path.abspath(f)
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"Attachment not found: {f}")
                mail.Attachments.Add(file_path)
        else:
            file_path = os.path.abspath(email_attached)
            if not os.path.exists(file_path):
                    raise FileNotFoundError(f"Attachment not found: {file_path}")
            mail.Attachments.Add(file_path)

    # add subject and body for email
    mail.Subject = email_subject
    mail.Body = email_body

    # send the email
    mail.Send()

    # log reminder
    logging.info(f"Email sent — Subject: {email_subject}")
