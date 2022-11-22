from flask import Flask, request
from skpy import Skype
# json
import json
with open('config.json') as config_file:
    config_data = json.load(config_file)

app = Flask(__name__)
@app.route("/send-message", methods=['POST'])
def sendMessage():
    user_data = getGoogleSheetUser();
    print(user_data)
    data = request.json
    # connect Skype
    ch = connectSkype()
    # get skype message
    message = getSkypeMessage(user_data,data)

    # check mentioned user exist
    if "<at id=" in message: 
      #send skype message
      ch.sendMsg(message, rich=True)
    return 'Message Sent';

def connectSkype():
    skype_settings = config_data['skype_settings']
    sk = Skype(skype_settings['SKYPE_USERNAME'], skype_settings['SKYPE_PASSWORD'])
    ch = sk.chats[skype_settings['SKYPE_GROUP_ID']]
    return ch

def getSkypeMessage(user_data,data):
    invitationText = ''
    finalMessage = ''
    for user in user_data:
        if(user['GoogleEmail'] == data['email']):
           invitationText = '<at id="{b1}">{b2}</at> '.format(b1=user['SkypeId'], b2=user['Name'].split()[0])
        if data['attendees']:
            for attendee in data['attendees']:
                 if(user['GoogleEmail'] == attendee['email']):
                    invitationText += '<at id="{b1}">{b2}</at> '.format(b1=user['SkypeId'], b2=user['Name'].split()[0])
    finalMessage = '{b1} Coming Events : {b2}'.format(b1=invitationText, b2=data['summary'])
    return finalMessage;

import gspread
from oauth2client.service_account import ServiceAccountCredentials
def getGoogleSheetUser():
    google_settings = config_data['google_settings']
    gc = gspread.service_account(filename='mycredentials.json')
    url = google_settings['SHEET_URL']
    gsheet = gc.open_by_url(url)
    udata = gsheet.sheet1.get_all_records()
    return udata;

if __name__ == "__main__":
    app.run(debug=True)