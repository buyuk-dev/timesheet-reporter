Simple gui tool for tracking work hours and sending weekly reports.
-------------------------------------------------------------------

# Installation

1. python3 gui.py

2. Outlook (app uses your outlook account to send emails)    

3. Install required python packages

    pip install -r requirements.txt

4. Proper configuration file

# Configuration

In order to use the app you have to create configuration file.
Template exists in *config.py.dist* file, you should rename it to *config.py* and change accordingly.

1. receipients: list of email addresses to which report should be send,
2. sender_address: your email address (corresponding to your active outlook account),
3. xls_path: path and filename format to the generated xlsx report file,
4. message_title: raport email message title,
5. message_body: raport email message body.

# Interface

1. save: saves current state in app save file,
2. load: loads previously saved app state,
3. generate: generates xlsx file with raport,
4. send: send generated xlsx file to defined list of receipients,
5. \+ recp: add new receipient field,
6. reset: reset fields to default state.
