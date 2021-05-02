# Python-Win32Com-MAPI-Email

Manipulation of MAPI using Python to retrieve information from Outlook client. I have implemented some scenarios where Outlook client can be scanned to automate read functionality.

Scenarios Implemented in this project:
1. Scanning through an Outlook email client (Inbox, Outbox, Draft, all other manually created folders)
2. Searching for keywords in the emails.
3. Looking for responses to a poll created using the email.
4. Retrieving the poll responses and looking for desired keywords in the exchange server.

To-do Items:
1. Automate sending, receiving and capturing polls on the email client.

More about the package below:
**MAPI**

mapi (Metadata API) is a python library which provides a high-level interface for media database providers, allowing users to efficiently search for television and movie metadata using a simple interface.

Installation
$ pip install mapi

Running Tests
$ pip install -r requirements-testing.txt
$ python -m pytest

