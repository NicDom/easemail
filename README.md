
__*Work in progress...*__
# Easemail
[![pre-commit](https://img.shields.io/badge/pre--commit-enabled-brightgreen?logo=pre-commit&logoColor=white)](https://github.com/pre-commit/pre-commit)
[![PyPI](https://img.shields.io/pypi/v/easemail)](https://pypi.python.org/pypi/easemail/)
[![PyPI](https://img.shields.io/pypi/pyversions/easemail)](https://pypi.python.org/pypi/easemail/)

------
The package has been created gradually while trying to make sending automatic mails in Python as easy as possible. I ended up with a package that basically just adds some features to [yagmail](https://github.com/kootenpv/yagmail). These are:

- Adding the possibilty to send emails via Outlook
- Adding list as tables to emails
- Adding the possibility to include inline images by putting this in an extra line inside of the text block, e.g.
  ```python
  """This is some text, followed by an inline picture
  /path/to/som/image.png
  followed by more text.
  """
  ```
- Determine the SMTP-server, security protocol and port using a database
- Added user files where the SMTP-Server, Security-Protocol and Port are stored, to make once determined data easier to reuse.

Therefore, it heavily depends on the [yagmail](https://github.com/kootenpv/yagmail) and [pywin32](https://github.com/mhammond/pywin32) project.

So, make sure to check out their project pages on Github and leave a star if you like this one, as the actual work happens there!!!

Sending emails can be done in the same way as with [yagmail](https://github.com/kootenpv/yagmail):
```python
import easemail
ema = easemail('mail_account')
hmtl = """
<html>
    <body>
        <p>Hi,<br>
        How are you?<br>
        You can find the easemail project
        <a href="https://github.com/NicDom/easemail">here.</a>
        </p>
    </body>
    </html>
"""
contents = ['Some body including text. There is also hmtl in this mail. A document is attached', html,  'document.pdf']
easemail.send('to@domain.com','subject', contents)
```
but `mail_account` can be either a valid mail address, a credential file for some OAuth2 authentication using a gmail address, just `'outlook'` to use Outlook or an alias for an user file (see below). See further instructions below.


## Installation

```python
pip install easemail
```

## Usage

There are three different usage cases:
### Case 1: Outlook

If you are using Windows and want to use your Outlook Application to send mails, there are two ways to tell easemail. The first is to call easemail via
```python
ema = easemail('outlook')
```
easemail will then give you a list of the in your Outlook Application registered email-addresses and lets aou choose one. If there is only one, easemail will use that one.

Alternatively, you may intialize easemail using the email address in your Outlook Application you are intending to use.

### Case 2: Gmail

If you do want to use a gmail adress, there are again two ways to tell easemail. Either by giving it the path to your credential file for the OAuth2 authentification, i.e. via
```python
ema = easemail('path/to/credential_file.json')
```
 or by giving your gmail address
 ```python
 ema = easemail('yourgmailname@gmail.com')
 ```
 However, as it is strongly recommended to use the authentification via OAuth 2 due to the possibility to revoke tokens, easemail will ask for the path to a credential file. If None is given it then first asks for the client ID and the client secret before asking for a password. You are then offered the option to store your password in a keyring.

 Of course, you can also give your password right from the beginning by running (NOT RECOMMENDED):
 ```python
 ema = easemail('yourgmailname@gmail.com', 'your_password')
 ```



### Case 3: Any email address

If you want to use any email address, just initialize easemail via
```python
ema = easemail('yourmailaddress@domain.com')
```
easemail will then try determine the SMTP-server, security protocol and port using a database. If this fails, you will be ask for the missing information.

As before, you you will be asked for a password, which can be stored in some keyring if you desire. Of course, the password can again be given as the second argument, when initializing easemail.

## User file

If easemail is able to connect to SMTP-Server using the given Server, Protocol and Port it, stores the gathered information in the file `EMAIL_ADDRESS(without @...)--EMAIL_ADDRESS.json`.
Thereby simplifying the reuse of the gathered information in future projects.


## Email contents

The email contents can follow the same syntax as proposed by the [yagmail](https://github.com/kootenpv/yagmail) package. However, I added two features:

- If contents contains a list, this list will be translated to html code (and plain text) and included into the mail.
- Inlnine images can now just be put inside of the text body. Therefore one has just to put the path of the image in some extra line inside of the body, e.g.
  ```python
  """This is some text, followed by an inline picture
  /path/to/som/image.png
  followed by more text.
  """
  ```



## Feedback

I do appreciate every kind of feedback and will try to respond to issues in 24 hours at Github.
