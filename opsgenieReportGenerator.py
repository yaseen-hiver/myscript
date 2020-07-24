#!/usr/bin/python

import requests
import json
import csv
import argparse
from csv import writer
import gspread
from datetime import datetime
from datetime import timedelta
from oauth2client.service_account import ServiceAccountCredentials
print("hello")

parser = argparse.ArgumentParser(description='Enter start date, end date and team name')
parser.add_argument("-s", action="store", dest="startDate")
parser.add_argument("-e", action="store", dest="endDate")
parser.add_argument("-t", "--team", action="store")
parser.add_argument("-f", action="store", dest="file")
parser.add_argument("-o", action="store", dest="key")
args = parser.parse_args()
OpsCredential = open(args.key,"r")
OpsCredentials = json.load(OpsCredential)


def get_ops_genie_alert_data(startDate, endDate, team, keys):
    sdate = datetime.strptime(args.startDate, "%d-%m-%Y").strftime("%d%b")
    edate = datetime.strptime(args.endDate, "%d-%m-%Y").strftime("%d%b")
    edate1 = datetime.strptime(args.endDate, "%d-%m-%Y")+timedelta(days=1)
    print(startDate,endDate,team)
    edate1 = edate1.strftime("%d-%m-%Y")
    try:
        response = requests.get("https://api.opsgenie.com/v2/alerts?query=teams:{team} AND createdAt>{startDate} AND createdAt<{endDate}".format(team=team, startDate=startDate, endDate=edate1), headers=keys)
    except:
        raise Exception("Unable to connect to OpsGenie to Fetch Data")
    if response.status_code == 200:
        json_response = response.json()
        print("data fetched")
        fi = sdate + ' - ' + edate
        create_csv(fi, json_response, team)
    else:
        print("unable to fetch data")

def create_csv(fi, json_response, team, keys):
    try:
        data_file = open(fi + '.csv', 'a')
    except:
        raise Exception("unable to open csv file")
    header = ["tinyId", "message", "createdAt", "owner", "team", "notes"]
    data = json_response['data']
    csv_writer = csv.writer(data_file)
    csv_writer.writerow(header)

    writer = csv.DictWriter(data_file, fieldnames=header)
    if(data):
        for record in data:
            notes_request = requests.get("https://api.opsgenie.com/v2/alerts/{id}/notes?identifierType=id".format(id=record["id"]), headers=keys)
            if notes_request.status_code == 200:
                notes_json = notes_request.json()
                if len(notes_json['data']) != 0:
                    note = notes_json['data'][0]['note']
                else:
                    note = 'None'
            # print note
            note_dict = {
                "tinyId": record["tinyId"],
                "message": record["message"],
                "createdAt": record["createdAt"],
                "owner": record["owner"],
                "team": team,
                "notes": note
            }
            # print note_dict
            if(note_dict):
                writer.writerow(note_dict)
        data_file.write('\n')

        data_file.close()
        print("CSV file created successfully")

        pasteCSVToGoogleSheet(fi)
        print("sheets updated")


def pasteCSVToGoogleSheet(fi, credential=args.file):
    scope = ['https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credential, scope)  # alertdata.json = file containing all your credentials for querying google apis
    client = gspread.authorize(credentials)
    google_sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/16Fo6l89vuksZnk5adXpGXlmJRzs8KA2IPeU36rnvHTU/edit#gid=855259435')  # url of google sheet

    try:
        worksheet = google_sheet.worksheet(fi)
    except:
        worksheet = google_sheet.add_worksheet(title=fi, rows="100", cols="20",index=0)

    (firstRow, firstColumn) = gspread.utils.a1_to_rowcol('A1')
    with open(fi+'.csv', 'r') as f:
        csvContents = f.read()
    body = {
        'requests': [{
            'pasteData': {
                "coordinate": {
                    "sheetId": worksheet.id,
                    "rowIndex": firstRow-1,
                    "columnIndex": firstColumn-1,
                },
                "data": csvContents,
                "type": 'PASTE_NORMAL',
                "delimiter": ',',
            }
        }]
    }
    return google_sheet.batch_update(body)



get_ops_genie_alert_data(args.startDate, args.endDate, args.team, OpsCredentials)
