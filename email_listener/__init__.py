"""email_listener: Listen in an email folder and process incoming emails.

Example:

    # Create the listener
    listener = EmailListener("example@email.com", "badpassword", "Inbox", "./files/")
    # Log the listener into the IMAP server
    listener.login()
    # Scrape emails from the folder without moving them
    listener.scrape()
    # Scrape emails from the folder, and move them to the "email_listener" folder
    listener.scrape("email_listener")
    # Listen in the folder for 5 minutes, without moving the emails, and not
    # calling any process function on the emails.
    listener.listen(5)
    # Listen in the folder until 1:30pm, moving each new email to the "email_listener"
    # folder, and calling the processing function 'send_reply()'
    listener.listen([13, 30], "email_listener", send_reply)
    # Log the listener out of the IMAP server
    listener.logout()

"""

# Imports from other packages
import email
from imapclient import IMAPClient, SEEN
import os
from dateparser import parse
import pytz
# Imports from this package
from .helpers import (
    calc_timeout,
    get_time,
)
from .email_processing import write_txt_file


class EmailListener:
    """EmailListener object for listening to an email folder and processing emails.

    Attributes:
        email (str): The email to listen to.
        app_password (str): The password for the email.
        folder (str): The email folder to listen in.
        attachment_dir (str): The file path to the folder to save scraped
            emails and attachments to.
        server (IMAPClient): The IMAP server to log into. Defaults to None.

    """

    def __init__(self, email, app_password, folder, attachment_dir=None, search_criteria="UNSEEN", mark_with_flags=[]):
        """Initialize an EmailListener instance.

        Args:
            email (str): The email to listen to.
            app_password (str): The password for the email.
            folder (str): The email folder to listen in. Can be 'INBOX' or
                one of the constants from IMAPClient (SEEN, ALL, etc)
            attachment_dir (str or None): The file path to folder to save scraped
                emails and attachments to.
            search_criteria (str or list): Criteria to use to search emails.
                Defaults to unseen emails.
            mark_with_flags (list): Flags to mark emails after being processed.
                Defaults to empty list.

        Returns:
            None

        """

        self.email = email
        self.app_password = app_password
        self.folder = folder
        self.attachment_dir = attachment_dir
        self.server = None
        self.search_criteria = search_criteria
        self.mark_with_flags = mark_with_flags

    def login(self):
        """Logs in the EmailListener to the IMAP server.

        Args:
            None

        Returns:
            None

        """

        self.server = IMAPClient('imap.gmail.com')
        self.server.login(self.email, self.app_password)

        if self.folder.lower() == 'inbox':
            folder = 'INBOX'
        else:
            folder = None
            for folder_data in self.server.list_folders():
                if self.folder in folder_data[0]:
                    folder = folder_data[2]
            if not folder:
                raise TypeError(f'Invalid folder: {self.folder}')

        self.server.select_folder(folder, readonly=False)

    def logout(self):
        """Logs out the EmailListener from the IMAP server.

        Args:
            None

        Returns:
            None

        """

        self.server.logout()
        self.server = None

    def scrape(self, move=None, mark_unread=False, delete=False, process_func=lambda x, y: None):
        """Scrape unread emails from the current folder.

        Args:
            move (str): The folder to move the emails to. If None, the emails
                are not moved. Defaults to None.
            mark_unread (bool): Whether the emails should be marked as unread.
                Defaults to False.
            delete (bool): Whether the emails should be deleted. Defaults to
                False.

        Returns:
            A list of the file paths to each scraped email.

        """

        # Ensure server is connected
        if type(self.server) is not IMAPClient:
            raise ValueError("server attribute must be type IMAPClient")

        # Search for unseen messages
        messages = self.server.search(self.search_criteria)
        # For each unseen message
        for uid, message_data in self.server.fetch(messages, 'RFC822').items():
            # Get the message
            email_message = email.message_from_bytes(message_data[b'RFC822'])
            # Get who the message is from
            from_email, from_name = self.__get_from(email_message)
            # Get who the message is for
            to_email, to_name = self.__get_to(email_message)

            # Generate the value dictionary to be filled later
            val_dict = {'filters': {}}

            # Display notice
            print("PROCESSING: Email UID = {} from {}".format(uid, from_email))

            # Add the subject
            val_dict["subject"] = self.__get_subject(email_message).strip()
            val_dict["from_email"] = from_email
            val_dict["from_name"] = from_name
            val_dict["to_email"] = to_email
            val_dict["to_name"] = to_name
            val_dict['date'] = self.__get_date(email_message)
            val_dict['id'] = uid
            val_dict['filters']['autoreply'] = self.__is_autoreply(
                email_message)
            val_dict['filters']['bounced'] = self.__is_bounced(email_message)

            # If the email has multiple parts
            if email_message.is_multipart():
                val_dict = self.__parse_multipart_message(
                    email_message, val_dict)

            # If the message isn't multipart
            else:
                val_dict = self.__parse_singlepart_message(
                    email_message, val_dict)

            # Process message first, then execute the options
            process_func(self, val_dict)

            # If required, move the email, mark it as unread, or delete it
            self.__execute_options(uid, move, mark_unread, delete)

    def __is_autoreply(self, email_message):
        header = email_message.get('Auto-Submitted')  # Access the raw headers
        if not header:
            return False
        header = header.lower()
        return 'auto-replied' in header or 'auto-submitted' in header or 'auto-generated' in header

    def __is_bounced(self, email_message):
        header = email_message.get('X-Failed-Recipients')
        return bool(header)

    def __get_date(self, email_message):
        date_str: str | None = email_message.get("Date")
        if not date_str:
            return None
        date_str = date_str.replace('(UTC)', '').strip()
        try:
            date = parse(date_str).astimezone(pytz.utc)
            return date
        except:
            print(f'ERROR PARSING DATE: {date_str}')
            return date_str

    def __get_from(self, email_message):
        """Helper function for getting who an email message is from.

        Args:
            email_message (email.message): The email message to get sender of.

        Returns:
            A string containing the from email address.

        """

        from_raw = email_message.get_all('From', [])
        from_list = email.utils.getaddresses(from_raw)
        if len(from_list[0]) == 1:
            from_email = from_list[0][0]
            from_name = None
        elif len(from_list[0]) == 2:
            from_email = from_list[0][1]
            from_name = from_list[0][0]
        else:
            from_email = "UnknownEmail"
            from_name = None

        return from_email, from_name

    def __get_subject(self, email_message):
        """

        """

        # Get the subject
        subject = email_message.get("Subject")
        # If there isn't a subject
        if subject is None:
            return "No Subject"
        return subject

    def __get_to(self, email_message):
        """

        """

        to_raw = email_message.get_all('To', [])
        to_list = email.utils.getaddresses(to_raw)
        if len(to_list[0]) == 1:
            to_email = to_list[0][0]
            to_name = None
        elif len(to_list[0]) == 2:
            to_email = to_list[0][1]
            to_name = to_list[0][0]
        else:
            to_email = None
            to_name = None
        return to_email, to_name

    def __parse_multipart_message(self, email_message, val_dict):
        """Helper function for parsing multipart email messages.

        Args:
            email_message (email.message): The email message to parse.
            val_dict (dict): A dictionary containing the message data from each
                part of the message. Will be returned after it is updated.

        Returns:
            The dictionary containing the message data for each part of the
            message.

        """

        # For each part
        for part in email_message.walk():
            # If the part is an attachment
            file_name = part.get_filename()
            if self.attachment_dir and bool(file_name):
                # Generate file path
                file_path = os.path.join(self.attachment_dir, file_name)
                file = open(file_path, 'wb')
                file.write(part.get_payload(decode=True))
                file.close()
                # Get the list of attachments, or initialize it if there isn't one
                attachment_list = val_dict.get("attachments") or []
                attachment_list.append("{}".format(file_path))
                val_dict["attachments"] = attachment_list

            # If the part is html text
            elif part.get_content_type() == 'text/html':
                # Convert the body from html to plain text
                val_dict["html"] = part.get_payload()

            # If the part is plain text
            elif part.get_content_type() == 'text/plain':
                # Get the body
                val_dict["text"] = part.get_payload()

        return val_dict

    def __parse_singlepart_message(self, email_message, val_dict):
        """Helper function for parsing singlepart email messages.

        Args:
            email_message (email.message): The email message to parse.
            val_dict (dict): A dictionary containing the message data from each
                part of the message. Will be returned after it is updated.

        Returns:
            The dictionary containing the message data for each part of the
            message.

        """

        # Get the message body, which is plain text
        val_dict["text"] = email_message.get_payload()
        return val_dict

    def __execute_options(self, uid, move, unread, delete):
        """Loop through optional arguments and execute any required processing.

        Args:
            uid (int): The email ID to process.
            move (str): The folder to move the emails to. If None, the emails
                are not moved. Defaults to None.
            unread (bool): Whether the emails should be marked as unread.
                Defaults to False.
            delete (bool): Whether the emails should be deleted. Defaults to
                False.

        Returns:
            None

        """

        # If the message should be marked as unread
        if bool(unread):
            self.server.remove_flags(uid, [SEEN])

        # Mark email with flags passed on creation
        self.server.add_flags(uid, self.mark_with_flags)

        # If a move folder is specified
        if move is not None:
            try:
                # Move the message to another folder
                self.server.move(uid, move)
            except:
                # Create the folder and move the message to the folder
                self.server.create_folder(move)
                self.server.move(uid, move)
        # If the message should be deleted
        elif bool(delete):
            # Move the email to the trash
            self.server.set_gmail_labels(uid, "\\Trash")
        return

    def listen(self, timeout=None, process_func=write_txt_file, **kwargs):
        """Listen in an email folder for incoming emails, and process them.

        Args:
            timeout (int, list or None): Either an integer representing the number
                of minutes to timeout in, a list, formatted as [hour, minute]
                of the local time to timeout at, or None. If None, will listen
                forever.
            process_func (function): A function called to further process the
                emails. The function must take the EmailListener and the dict
                containing the data from the email. Defaults to the example
                function write_txt_file in the email_processing module.
            **kwargs (dict): Additional arguments for processing the email.
                Optional arguments include:
                    move (str): The folder to move emails to. If not set, the
                        emails will not be moved.
                    unread (bool): Whether the emails should be marked as unread.
                        If not set, emails are kept as read.
                    delete (bool): Whether the emails should be deleted. If not
                        set, emails are not deleted.

        Returns:
            None

        """

        # Ensure server is connected
        if type(self.server) is not IMAPClient:
            raise ValueError("server attribute must be type IMAPClient")

        # Get the timeout value
        def should_continue(timeout):
            'If no timeout is defined, listen forever, else return'
            if timeout is None:
                return True
            else:
                outer_timeout = calc_timeout(timeout)
                return get_time() < outer_timeout

        move = kwargs.get('move')
        unread = bool(kwargs.get('unread'))
        delete = bool(kwargs.get('delete'))

        # Run until the timeout is reached
        while (should_continue(timeout)):
            self.__process_new_emails(move, unread, delete, process_func)
            self.__idle(process_func=process_func, **kwargs)
        return

    def __idle(self, process_func=write_txt_file, **kwargs):
        """Helper function, idles in an email folder processing incoming emails.

        Args:
            process_func (function): A function called to further process the
                emails. The function must take only the list of file paths
                returned by the scrape function as an argument. Defaults to the
                example function write_txt_file in the email_processing module.
            **kwargs (dict): Additional arguments for processing the email.
                Optional arguments include:
                    move (str): The folder to move emails to. If not set, the
                        emails will not be moved.
                    unread (bool): Whether the emails should be marked as unread.
                        If not set, emails are kept as read.
                    delete (bool): Whether the emails should be deleted. If not
                        set, emails are not deleted.

        Returns:
            None

        """

        # Set the relevant kwarg variables
        move = kwargs.get('move')
        unread = bool(kwargs.get('unread'))
        delete = bool(kwargs.get('delete'))

        # Start idling
        self.server.idle()
        print("Connection is now in IDLE mode.")
        # Set idle timeout to 5 minutes
        inner_timeout = get_time() + 60*5
        # Until idle times out
        while (get_time() < inner_timeout):
            # Check for a new response every 30 seconds
            responses = self.server.idle_check(timeout=30)
            print("Server sent:", responses if responses else "nothing")
            # If there is a response
            if (responses):
                # Suspend the idling
                self.server.idle_done()
                self.__process_new_emails(move, unread, delete, process_func)
                # Restart idling
                self.server.idle()
        # Stop idling
        self.server.idle_done()
        return

    def __process_new_emails(self, move, unread, delete, process_func):
        # Process the new emails
        self.scrape(
            move=move, mark_unread=unread, delete=delete, process_func=process_func)
