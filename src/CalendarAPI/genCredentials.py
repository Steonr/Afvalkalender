import os
import sys
from google_auth_oauthlib.flow import InstalledAppFlow
THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
THIS_FOLDER = os.path.dirname(THIS_FOLDER)
client_secret = str(THIS_FOLDER)+"/Authentication/Calendar/client_secret.json"
def authorization():  
    scopes = ["https://www.googleapis.com/auth/calendar"]
    #client_secret.json is a file that contains the key to get acces to the calendar API of a specific account.
    flow = InstalledAppFlow.from_client_secrets_file(client_secret, scopes = scopes)
    credentials = flow.run_local_server(port=0)       
    return credentials


