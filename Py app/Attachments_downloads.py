from zenpy import Zenpy
import datetime


creds = {
    'email' : 'dmitry.dmytrenko@envusa.design',
    'token' : 'zgXYhMmdxGTT0gMPVOaFfRf5xUNue4UPNsE4MaGs',
    'subdomain': 'https://envusa.zendesk.com'}

zenpy_client = Zenpy(**creds)

yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
today = datetime.datetime.now()
for ticket in zenpy_client.search("zenpy", created_between=[yesterday, today], type='ticket', minus='negated'):
    print(ticket)