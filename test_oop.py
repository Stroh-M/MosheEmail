import oop

mail = oop.Email("amzupld@gmail.com", "ixwlgxbvekhisqwr")

result = mail.login_select_mailbox('INBOX', 'ebay@ebay.com', 'keurig@em.keurig.com')

print(result)

for i in result:
    oop.ScrapeEmail()