Simple gui tool for tracking work hours and sending weekly reports.
-------------------------------------------------------------------

# Usage

1. python3 .

2. Email config: currently supports sending emails with outlook api or directly via microsoft exchange server.
Use *use_outlook* option to choose which mode should be used.
For sending emails with outlook you need to have installed and configured outlook account matching sender_address.
For sending emails with microsoft exchange server you need to have valid credentials to sender_address account. 

3. Install required python packages

    pip install -r requirements.txt

4. Make sure you filled all important fields in configuration file

# Interface

1. save: saves current state in app save file,
2. load: loads previously saved app state,
3. generate: generates xlsx file with raport,
4. send: send generated xlsx file to defined list of receipients,
5. recp: add new receipient field,
6. reset: reset fields to default state.

# Configuration

In order to use the app you have to create configuration file.
Template exists in *config.py.dist* file, you should rename it to *config.py* and change accordingly.

1. receipients: list of email addresses to which report should be send,
2. sender_address: your email address (corresponding to your active outlook account),
3. xls_path: path and filename format to the generated xlsx report file,
4. message_title: raport email message title,
5. message_body: raport email message body.
6. use_outlook: if set to True emails will be send using outlook if sender_address is properly set. Otherwise exchange server will be used directly (requires credentials).
7. credentials: username and password to authenticate with exchange account if use_outlook is False.
8. employee_name: name to put in raport sheet and in message body signature.



