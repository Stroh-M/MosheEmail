import imaplib, smtplib

class Email():
    def __init__(self, emailAdress, emailPassword):
        self.email_address = emailAdress
        self.email_password = emailPassword
        self.imap_mail = imaplib.IMAP4_SSL(f'imap.gmail.com')
        self.smtp_mail = smtplib.SMTP_SSL(f'smtp.gmail.com', 465)

    def login_select_mailbox(self, mailbox, email_from_1, email_from_2):
        self.imap_mail.login(self.email_address, self.email_password)
        self.imap_mail.select(mailbox=mailbox)
        status, self.result = self.imap_mail.search(None, f'OR FROM {email_from_1} FROM {email_from_2}')
        return  self.result[0].split()
        
class ScrapeEmail(Email):
    def __init__(self, emailAdress, emailPassword):
        super().__init__(emailAdress, emailPassword)
        
    def find_pattern(self, email_id, pattern):
        pass
        
