import os
import imaplib
import email
import cups
import time

# Server address and port (example for outlook)
server = "outlook.office365.com"
port = 993
username = "email"
password  = "email.password"

attachment_directory = "/Path/to/save/attachments"
accepted_sender = "permitted@email.com"

# cups init
conn_printer = cups.Connection()
printer = conn_printer.getPrinters()
printer_name = str(list(printer.keys())[0])

# Standard print options
print_options = {
    "media": "A4",
    "fit-to-page": "True"
}

# Extract the email of the sender
def extractMailFROM(input):
    start = input.find("<")
    end = input.find(">")

    if start == -1 or end == -1:
        return None
    
    return input[start + 1:end]

# Connection to the mail server
imap = imaplib.IMAP4_SSL(server)
imap.login(username, password)

while True:
    imap.select("Inbox")

    _, msgnums = imap.search(None, "ALL")

    for msnum in msgnums[0].split():
        _, data = imap.fetch(msnum, "(RFC822)")
        
        message = email.message_from_bytes(data[0][1])

        # Checks if email is authorized
        sender_email = extractMailFROM(message.get('From'))
        if sender_email != accepted_sender:
            continue
        
        # Get attachments from the email
        for part in message.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            filename = part.get_filename()
            if filename:
                filepath = os.path.join(attachment_directory, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Attachment {filename} was saved.")

            file_path = str(attachment_directory + '/' + filename)
            job_id = conn_printer.printFile(printer_name, file_path, "Print Job", print_options)

        # Delete received and printed mail
        imap.store(msnum, "+FLAGS", "\\Deleted")
        time.sleep(10)

imap.expunge()
imap.close()