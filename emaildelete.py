# This script removes emails that has specific words found in the body, I made it because I was lazy and didn't want to do it manually

import imaplib
import email
from email.header import decode_header
import concurrent.futures

# Set your email server and credentials 
username = "xyz" # e.g. "test123@test.com"
password = "xyz" # e.g. "test123"

# Connect to the IMAP server
imap = imaplib.IMAP4_SSL("outlook.office365.com") # Change this as needed

# Log in to your email account
imap.login(username, password)

# Select the inbox folder
imap.select("Inbox")

# Search for all emails in the inbox
status, email_ids = imap.search(None, "ALL")

# Parse the email with all case combinations
status, messages1 = imap.search(None, '(BODY "unsubscribe")')
status, messages2 = imap.search(None, '(BODY "Unsubscribe")')
status, messages3 = imap.search(None, '(BODY "UNSUBSCRIBE")')
status, messages4 = imap.search(None, '(BODY "opt-out")')
status, messages5 = imap.search(None, '(BODY "Opt-Out")')
status, messages6 = imap.search(None, '(BODY "OPT-OUT")')
status, messages7 = imap.search(None, '(BODY "newsletter")')
status, messages8 = imap.search(None, '(BODY "Newsletter")')
status, messages9 = imap.search(None, '(BODY "preferences")')
status, messages10 = imap.search(None, '(BODY "Preferences")')

# Combine the messages that meet either condition
messages = list(set(messages1[0].decode().split() + messages2[0].decode().split() + messages3[0].decode().split() + messages4[0].decode().split() + messages5[0].decode().split() + messages6[0].decode().split() + messages7[0].decode().split() + messages8[0].decode().split() + messages9[0].decode().split() + messages10[0].decode().split()))

def process_mail(mail):
    _, msg = imap.fetch(mail, "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            print("Deleting", subject)
    imap.store(mail, "+FLAGS", "\\Deleted")

# Create a ThreadPoolExecutor
with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor: # Changing the thread to higher than 1 will result in Gmail blocking it, not experimented with other email providers
    # Submit tasks to thread pool
    futures = {executor.submit(process_mail, mail) for mail in messages}

# Wait for all tasks to complete
concurrent.futures.wait(futures)
# permanently remove mails that are marked as deleted
imap.expunge()
# close the mailbox
imap.close()
# logout from the account
imap.logout()
