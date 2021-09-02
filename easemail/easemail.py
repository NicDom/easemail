import json
import mimetypes
import re

import yagmail
from yagmail.oauth2 import get_oauth2_info

import win32com.client

import shutil
import os
import sys
from os.path import isfile, join

from tabulate import tabulate

from scraper import get_databases
from message import prepare_contents, list_to_html_table, get_message_html_and_str


PY3 = sys.version_info[0] > 2


class client:
    """The main class of the project.
    """

    def __init__(
        self,
        mail_account=None,
        port=None,
        host=None,
        smtp_starttls=None,
        smtp_ssl=True,
        smtp_set_debuglevel=0,
        smtp_skip_login=False,
        encoding="utf-8",
        soft_email_validation=True,
        file_dir=None,
        data_dir=None,
        **kwargs,
    ):

        self.soft_email_validation = soft_email_validation
        self.smtp_starttls = smtp_starttls
        self.ssl = smtp_ssl
        self.encoding = encoding
        self.smtp_skip_login = smtp_skip_login
        self.smtp_set_debuglevel = smtp_set_debuglevel
        self.host = host
        self.port = port
        self.kwargs = kwargs

        self.home_dir = os.path.dirname(os.path.realpath(__file__))

        if file_dir is None:
            self.file_dir = join(self.home_dir, ".profiles")
        else:
            self.file_dir = file_dir
        if data_dir is None:
            self.data_dir = join(self.home_dir, "data")
        else:
            self.data_dir = data_dir

        self.server_database = join(self.data_dir, "SMTP-Server.json")
        self.security_protocol_database = join(self.data_dir, "PortsAndProtocols.json")

        if mail_account == None:
            mail_account = input("Please enter you email-address here:\n")
            if soft_email_validation and "@" in mail_account:
                yagmail.validate.validate_email_with_regex(mail_account)
        if isinstance(mail_account, dict):
            if len(mail_account) == 1:
                for key, value in mail_account.items():
                    self.useralias = key
                    self.user = value
                else:
                    print("Mail account has to many entries.")
        else:
            self.user = mail_account
        filename = self.user_file_correspoding_to_alias(self.user)
        if filename is not None:
            self.login_using_user_file(filename)
            self.useralias = self.user.split("--")[0]
            return
        if "@gmail.com" in mail_account:
            self.user = mail_account
            self.host = "smtp.gmail.com"
            if isfile(self._default_credential_filename(mail_account)):
                mail_account = self._default_credential_filename(mail_account)
            else:
                text = """It is highly recommended to use some OAuth Client ID to send mails via your gmail account.
Thus, please enter the full path of your Credentials file below, if you allready have some.
If you do not have an app registered for your email sending purposes, visit:
https://console.developers.google.com
and create a new project.
If you do have a project but no Credentials, visit:
https://console.cloud.google.com/projectselector2/apis/credentials?supportedpurview=project
and create some desctop client Credentials.
For an explanation on how to do so, visit:
https://developers.google.com/workspace/guides/create-credentials
Enter the full path of your credentials file here:
"""
                mail_account = input(text)
                try:
                    self.mail_server = self._yagmail(mail_account)
                except:
                    print(
                        "Couldn't connect to the OAuth Client ID. Proceeding without a Credential file. (not recommended)"
                    )
                    self.mail_server = self._yagmail()
            self.mode = "gmail"
        elif isfile(mail_account):
            self.mode = "gmail"
            oauth2_info = get_oauth2_info(mail_account)
            self.user = oauth2_info["email_address"]
            self.mail_server = self._yagmail(mail_account)
        elif mail_account == "outlook":
            self.mode = "outlook"
            self.user = str(self.get_user_from_outlook())
            self.useralias = self.user.split("@")[0]
        elif bool([i for i in ["@outlook", "@msn", "@hotmail"] if i in mail_account]):
            user = self.get_user_from_outlook(mail_account)
            if user:
                self.user = user
                self.useralias = self.user.split("@")[0]
                self.mode = "outlook"
            else:
                self._set_smtp_mode(mail_account)
        else:
            self._set_smtp_mode(mail_account)
        self.write_user_file()

    @property
    def user(self):
        return self._user

    @user.setter
    def user(self, value):
        if self.soft_email_validation and "@" in value:
            yagmail.validate.validate_email_with_regex(value)
        self._user = value

    def _yagmail(self, oauth2_file=None):
        return yagmail.SMTP(
            user=self.user,
            oauth2_file=oauth2_file,
            host=self.host,
            smtp_starttls=self.smtp_starttls,
            smtp_ssl=self.ssl,
            smtp_set_debuglevel=self.smtp_set_debuglevel,
            smtp_skip_login=self.smtp_skip_login,
            encoding=self.encoding,
            soft_email_validation=self.soft_email_validation,
            port=self.port,
        )

    def _set_smtp_mode(self, mail_account):
        self.mode = "smtp"
        self.host, self.port, self.ssl, self.smtp_starttls = self.get_hcpp(mail_account)
        self.user = mail_account
        self.mail_server = self._yagmail()

    def send(
        self,
        to=None,
        subject=None,
        contents=None,
        attachments=None,
        cc=None,
        bcc=None,
        display_only=False,
        prettify_html=True,
        print_only=False,  # required for testing on outlook
    ):
        """This function is used to send emails.

        Args:
            to (list or str, optional): The recipients for the mail. Defaults to None.
            subject (str, optional): The subject header for the mail. Defaults to None.
            contents (list or str, optional): The contents of the mail. If it is a list it will determine, which of the contents it has to treat as text body, tables, inline images or attachments. Defaults to None.
            attachments (list or str, optional): A path to a single attachment or a list of paths to files. Defaults to None.
            cc (list or str, optional): CC of the mail. Defaults to None.
            bcc (list or str, optional): CC of the mail. Defaults to None.
            display_only (bool, optional): If set to true, it will only display the mail instead of sending it. Defaults to False.
            prettify_html (bool, optional): If set to true, it will use the premailer package to prettify the html code of the mail. Defaults to True.
        """
        if self.mode == "gmail" or self.mode == "smtp":
            contents, attachments = prepare_contents(contents, attachments)
            to, cc, bcc = self.prepare_recipients_for_yagmail(to, cc, bcc)
            for i in range(len(contents)):
                content = contents[i]
                if type(content) == list:
                    contents[i] = list_to_html_table(content)
            return self.mail_server.send(
                to=to,
                subject=subject,
                contents=contents,
                attachments=attachments,
                cc=cc,
                bcc=bcc,
                preview_only=display_only,
                prettify_html=prettify_html,
            )
        else:
            return self.send_via_outlook(
                to=to,
                subject=subject,
                contents=contents,
                attachments=attachments,
                cc=cc,
                bcc=bcc,
                display_only=display_only,
                print_only=print_only,
            )

    def send_via_outlook(
        self,
        to=None,
        subject=None,
        contents=None,
        attachments=None,
        cc=None,
        bcc=None,
        display_only=False,
        print_only=False,  # required only for testing on outlook
        mail=None,
    ):
        """Sends emails via outlook

        Args:
            to (str or list, optional): Recipients. Defaults to None.
            subject (str, optional): Subject header. Defaults to None.
            contents (str or list, optional): The contents (see DocString for send()). Defaults to None.
            attachments (str or list, optional): attachments. Defaults to None.
            cc (str or list, optional): CC. Defaults to None.
            bcc (str or list, optional): BCC. Defaults to None.
            display_only (bool, optional): If set to True, it will open outlook and only display the message. Defaults to False.
            mail (<class 'win32com.client.CDispatch'>, optional): The email object. Defaults to None.
        """
        if mail == None:
            mail = self.prepare_outlook_mail(to, subject, contents, attachments, cc, bcc)
        if print_only:
            return mail.To, mail.Body
        else:
            if display_only:
                mail.Display()
                return mail.To, mail.Body
            else:
                mail.Send()
                return mail.To, mail.Body

    def prepare_recipients_for_yagmail(self, to, cc, bcc):
        def _right_format(recipients):
            copy = {}
            if recipients is None:
                return None
            if not isinstance(recipients, list):
                recipients = [recipients]
            if isinstance(recipients, list):
                for i in range(len(recipients)):
                    recipient = recipients[i]
                    if isinstance(recipient, dict):
                        for key, value in recipient.items():
                            if not "@" in key:
                                copy[value] = f"{key} <{value}>"
                    if isinstance(recipient, str):
                        copy[recipient] = f"{recipient.split('@')[0]} <{recipient}>"
            return copy

        return _right_format(to), _right_format(cc), _right_format(bcc)

    def prepare_outlook_mail(self, to=None, subject=None, contents=None, attachments=None, cc=None, bcc=None):
        """Creates the <class 'win32com.client.CDispatch'> object representing the mail for the Outlook Application.

        Returns:
            <class 'win32com.client.CDispatch'>: The mail object for the Outlook Application.
        """
        if subject == None:
            subject = ""
        if isinstance(subject, list):
            subject = " ".join(subject)
        html = ""
        contents, attachments = prepare_contents(contents, attachments)
        body, html = get_message_html_and_str(contents)
        to, cc, bcc = self.prepare_recipients_for_outlook(to, cc, bcc)
        account = self.mail_server.Session.Accounts[self.user]
        mail = self.mail_server.CreateItem(0)  # Alternatively use olMailItem instead of 0 as the argument
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        mail.HTMLBody = html
        if attachments is not None:
            for a in attachments:
                mail.Attachments.Add(a)
        mail.CC = cc
        mail.BCC = bcc
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
        return mail

    def prepare_recipients_for_outlook(self, to, cc, bcc):
        def _list_to_str(recipients):
            if recipients == None:
                recipients = ""
            result = ""
            if not isinstance(recipients, (list, tuple)):
                recipients = [recipients]
            if isinstance(recipients, (list, tuple)):
                for recipient in recipients:
                    if isinstance(recipient, dict):
                        for key, value in recipient.items():
                            if "@" in key:
                                result += value + ";"
                            else:
                                result += key + " " + value + ";"
                    elif isinstance(recipient, str):
                        result += recipient + ";"
                    else:
                        print(
                            "Recipient {} is of type {}, but only {} and {} are supported.".format(
                                recipient, type(recipient), str, dict
                            )
                        )
            else:
                result = recipients
            return result

        return _list_to_str(to), _list_to_str(cc), _list_to_str(bcc)

    def get_hcpp(self, mail):
        """Tries to determine the host, the cryptographic protocol and the port for the given the email address using a database and asks for the input if the method fails.

        Args:
            mail (str): The email address

        Returns:
            str, str, bool, bool or None: The host, True if ssl ist used (False otherwise), True if starttls is used (None otherwise)
        """
        host = None
        port = None
        protocol = None
        mail = "@" + mail.split("@")[1]
        # print(mail)
        if not isfile(self.server_database) or not isfile(self.security_protocol_database):
            get_databases(
                server_database=self.server_database, security_protocol_database=self.security_protocol_database
            )
        with open(self.server_database, "r") as file:
            server = json.loads(file.read())
        with open(self.security_protocol_database, "r") as file:
            pap = json.loads(file.read())
        for item in server:
            if item[0] is not None:
                if mail == item[0]:
                    host = item[1]
                    print("SMTP-Server: {}".format(host))
                    break
        if host is None:
            host = input("Your email provider is not in the database. Please enter your host:\n")
        for item in pap:
            if host == item[0]:
                port = item[2]
                protocol = item[1]
                print("Security Protocol: {}".format(protocol))
                print("Port: {}".format(port))
                break
        if port == None:
            protocol = input(
                "There is no security protocol for your SMTP-Server in the database. Please enter your security protocol (ssl,starttls):\n"
            )
            while not protocol.upper() in ["SSL", "STARTTLS"]:
                protocol = input(
                    f"{protocol.upper()} is no a valid security protocol. Please enter your security protocol (ssl,starttls):\n"
                )
            port = input("Please enter your port:\n")
        if protocol.upper() == "SSL":
            ssl = True
            starttls = None
        elif protocol.upper() == "STARTTLS":
            ssl = False
            starttls = True
        return host, port, ssl, starttls

    def get_user_from_outlook(self, email_adr=None):
        """Is used to determine the Outlook email-account that the user wants to send mails from.

        Args:
            email_adr (str, optional): If None, the function lists all the users that are registered in that Outlook Application. If it is str (the email address) it will check, if the email address
                                       is registered in the Application. Defaults to None.

        Returns:
            bool or str: If email_adr was None returns the chosen account. If email_adr was a string returns email_adr, if email_adr is an account in your Outlook Application. Otherwise it will return False.
        """
        try:
            self.mail_server = win32com.client.Dispatch("Outlook.Application")
        except:
            print("Couldn't connect to Outlook. I will send emails via SMTP.")
            return False
        accounts = self.mail_server.Session.Accounts  # ;
        if email_adr != None:
            for account in accounts:
                if email_adr == str(account):
                    return email_adr
            return False
        if len(accounts) == 1:
            account = accounts[0]
        else:
            options = []
            for i in range(len(accounts)):
                options.append([str(accounts[i])])
            account = options[self.determine_mail_account(options, headers=["email-address"])][0]
        print(f"Setting {account} to the active email account.")
        return account

    def determine_mail_account(self, items, headers):
        """Functions that displays the email-accounts in your Outlook Application together with a generated ID in a table and lets you choose one by either typing the ID or typing (parts of) the email-address.

        Args:
            items (list): The Outlook accounts
            headers (list): The headers for the output of the accounts in a table.

        Returns:
            int: The index of the chosen item from items
        """
        self._print_items(items, headers=headers)
        item_to_delete = input("Please type in the email ID or (parts) of the email account you are willing to use:\n")
        if self._valid_id(item_to_delete, items):
            return int(item_to_delete) - 1
        else:
            items, indices = self.matches(item_to_delete, items)
            if len(items) == 1:
                return indices[0]
            if len(items) > 0:
                self._print_items(items, headers=headers)
                item_to_delete = input("Please type in the ID of the email account you are willing to use:")
                if self._valid_id(item_to_delete, items):
                    return indices[int(item_to_delete)] - 1
                else:
                    print("Your email account is not in the list.")
                    return None

    def matches(self, pattern, items):
        """Check if the pattern is found in one of the first column entries of items. It thus enables the user to not only enter the given email-account ID in 'determine_mail_account' but also parts of the email-address

        Args:
            description (str):
            items (lsit): [description]

        Returns:
            list, int: The list of matches and their indices in items
        """
        matches = []
        indices = []
        for item in items:
            if pattern in item[0]:
                matches.append(item)
                indices.append(items.index(item))
        return matches, indices

    def _valid_id(self, text, items):
        try:
            if 1 <= int(text) < len(items) + 1:
                return True
        except:
            return False

    def _print_items(self, items, headers, format="psql"):
        """Used in 'determine-mail-account' to display all email-accounts registered in the local Outlook Application with a generated ID in a table.

        Args:
            items (list): The items to displayed
            headers (list): The headers for the table
            format (str, optional): The tabulate style for the table. Defaults to "psql".
        """
        i = 0
        for item in items:
            item.insert(0, i + 1)
            i += 1
        print("\n", tabulate(items, ["ID"] + headers, tablefmt=format))
        for item in items:
            del item[0]

    def userfile_exists(self, filename):
        if "@" in filename:
            pattern = re.compile(r"(.*)\@[.]*")
            filename = join(self.file_dir, pattern.findall(self.user)[0] + ".json")
        else:
            filename = join(self.file_dir, filename + ".json")
        return isfile(join(self.file_dir, filename))

    def write_user_file(self):
        filename = self.prepare_filename(self.user)
        email_data = {
            "email_address": self.user,
            "host": self.host,
            "port": self.port,
            "starttls": self.smtp_starttls,
            "ssl": self.ssl,
            "mode": self.mode,
        }
        file = open(filename, "w")
        file.write(json.dumps(email_data))
        file.close()

    def list_files(self, dir):
        """Lists all files in dir

        Args:
            dir (str): Path to a directory

        Returns:
            list: All files in the directory
        """
        try:
            return [f for f in os.listdir(dir) if os.path.isfile(os.path.join(dir, f))]
        except os.error as e:
            print("Error : {}".format(e))

    def user_files(self):
        """Returns a list with all user files

        Returns:
            list: List of user files
        """
        return self.list_files(self.file_dir)

    def user_file_correspoding_to_alias(self, alias):
        for item in self.user_files():
            if alias in item:
                if isfile(join(self.file_dir, item)):
                    return item
        return None

    def prepare_filename(self, filename):
        if "@" in filename and not "--" in filename:
            filename_parts = filename.split("@")
            filename = join(
                self.file_dir, filename_parts[0] + "--" + filename_parts[0] + "@" + filename_parts[1] + ".json"
            )
        else:
            filename = join(self.file_dir, self.user_file_correspoding_to_alias(filename))
        return filename

    def read_user_file(self, filename):
        try:
            with open(filename, "r") as file:
                data = json.loads(file.read())
        except FileNotFoundError:
            return
        return data

    def login_using_user_file(self, filename):
        filename = self.prepare_filename(filename)
        data = self.read_user_file(filename)
        if "installed" in data.keys():
            data = data["installed"]
        if not "google_client_id" in data.keys():
            [self.user, self.host, self.port, self.smtp_starttls, self.ssl, self.mode] = list(data.values())
            if self.mode == "outlook":
                self.user = self.get_user_from_outlook(self.user)
            else:
                self.mail_server = self._yagmail()

        else:
            if "email_address" in data.keys():
                self.user = data["email_address"]
            self.mode = "gmail"
            self.host = "smtp.gmail.com"
            self.mail_server = self._yagmail(filename)

    def _default_credential_filename(self, mail):
        if isfile(mail):
            return mail
        else:
            dir = os.path.dirname(os.path.realpath(__file__))
            filename = mail.replace("@gmail.com", ".json")
            return os.path.join(dir, filename)

    def move_file(self, src, target):
        path = os.path.join(target, src)
        if os.path.exists(path):
            self.delete_file(path)
        try:
            shutil.move(src, target)
        except shutil.Error as e:  ## if failed, re        port it back to the user ##
            print("Error: %s - %s." % (e.filename, e.strerror))

    def copy_file(self, src, target):
        try:
            shutil.copyfile(src, target)
        except shutil.Error as e:  ## if failed, re        port it back to the user ##
            print("Error: %s - %s." % (e.filename, e.strerror))
