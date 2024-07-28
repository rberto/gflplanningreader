# How to use 

Create a file named glob.py containing the following lines and replace the content of both variable accordingly

    MAIL_FILTER = "from:abc@gmail.com has:attachment"
    PDFS_PATH = "/home/john/pdfs/"
    CALENDAR_NAME = "calendar"
    NAME = john"

Launch 

    $ python ./plannningreader.py

# What it does

- Connects to gmail and google calendar services using the downloaded credential.json
- Look for existing [CALENDAR_NAME]
- Downloads all the attachements present in the mails filtered by the [MAIL_FILTER] in the [PDFS_PATH] if not already downloaded
- For each new schedule file (attachement)
- Extract shifts for [NAME]
- Extract startdate of shifts per file from filename and start date year from mail reception date
- Parse shifts info
- Upload shift to google calendar 

# Todo
- Add detection of duplicates, to prevent upload of two of the same calendar event
