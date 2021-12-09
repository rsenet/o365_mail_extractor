from rich.console import Console
from imaplib import IMAP4
import imaplib
import email
import sys
import re

# AUTHENTICATION
email_user = ""
email_pass = ""


def imap_cnx(username, password):
    """ Connexion on the specified IMAP server
            input: username: o365 username
            input: password: o365 password
            output: mail object
    """
    host = 'outlook.office365.com'
    port = 993
    mail = imaplib.IMAP4_SSL(host,port)
    mail.login(username, password)

    return mail


def get_imap_folder(mail):
    """ Get all reachable folder 
            input: mail object (imap_cnx)
            output: list of folder
    """
    mailFolder = []

    for i in mail.list()[1]:
        list_mailbox = i.decode().split(' "/" ')
        mailFolder.append(list_mailbox[1])

    return mailFolder


def print_mail_list(mailing_list):
    """ Get all reachable folder 
            input: mailing list (array)
            output: list of mail
    """
    # Delete doublons
    mailList = list(dict.fromkeys(mailing_list))

    for element in mailList:
        print(element)


# Variables
mailAddr = []

# Connexion
try:
    imapCnx = imap_cnx(email_user, email_pass)

except:
    print("Error during authentication")
    sys.exit(0)

# Get folder
folderList = get_imap_folder(imapCnx)

console = Console()
tasks = [f"{element}" for element in folderList]

try:
    with console.status("[bold green]Extracting data...") as status:
        for element in tasks:
            try:
                mailCnx   = imap_cnx(email_user, email_pass)
                mailCnx.select(element, readonly=True)
                typ, data = mailCnx.search(None, 'ALL')
                ids       = data[0]
                id_list   = ids.split()

                try:
                    latest_id = int(id_list[-1])

                    # GET ALL INFORMATION ON EACH MESSAGE
                    for index in range(latest_id, 0, -1):
                        typ, msg_data = mailCnx.fetch("%s" % index, '(RFC822)')
                        
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                try:
                                    msg = email.message_from_string(response_part[1].decode())

                                    for header in ['from']:
                                        try:
                                            fromMail   = re.findall('\S+@\S+', msg[header])
                                            returnMail = (fromMail[0].replace('<', '').replace('>', ''))
                                            mailAddr.append(returnMail)

                                        except IndexError:
                                            pass

                                except UnicodeDecodeError as error:
                                    pass

                except IndexError:
                    pass

                mailCnx.close()
                mailCnx.logout()

            except IMAP4.abort:
                imapCnx = imap_cnx(email_user, email_pass)

            except ConnectionResetError:
                print("Connexion reseted")
                print_mail_list(mailAddr)
                sys.exit(0)


            console.log(f"Extraction done for {element}")

        # Display the final list
        print_mail_list(mailAddr)

except KeyboardInterrupt:
    sys.exit(0)