from __future__ import print_function
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import os.path
import datetime
import locale

locale.setlocale(locale.LC_TIME, "ru_RU.UTF-8")  # Setting locale to Russian

# Importing data
filename = 'data.pkl'
with open(filename, 'rb') as f:
    imported_data = pickle.load(f)

# Authorization
SCOPES = ['https://www.googleapis.com/auth/calendar']
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pkl'):
    creds = pickle.load(open("token.pkl", "rb"))
    service = build("calendar", "v3", credentials=creds)
else:
    flow = InstalledAppFlow.from_client_secrets_file("client_secret.json",
                                                     scopes=SCOPES)
    creds = flow.run_local_server(port=0)
    pickle.dump(creds, open("token.pkl", "wb"))
    creds = pickle.load(open("token.pkl", "rb"))
    service = build("calendar", "v3", credentials=creds)


def create_event(start_time_format, end_time_format, summary, description,
                 location):
    event = {
        'summary': summary,
        'location': location,
        'description': description,
        'start': {
            'dateTime': start_time_format,
            'timeZone': 'Europe/Moscow',
        },
        'end': {
            'dateTime': end_time_format,
            'timeZone': 'Europe/Moscow',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': 60},
            ],
        },
    }
    return service.events().insert(
        calendarId='tot3a4vhmmo3jup9rbko7aqqas@group.calendar.google''.com',
        body=event).execute()


# Creating events in Google Calendar
lesson_day = 0
while lesson_day < len(imported_data):
    lesson_id = 1  # Not 0 to avoid copying a date
    if imported_data[lesson_day][0] != "None":
        while lesson_id <= len(imported_data[lesson_day]) - 1:
            print(f'День недели: {lesson_day + 1}-й, занятие: {lesson_id}-е')
            d = imported_data[lesson_day][
                0]  # Retrieving date and lesson start time and formatting it
            t_start = imported_data[lesson_day][lesson_id][0][0:2] + " " + \
                      imported_data[lesson_day][lesson_id][0][3:5]
            # print(t_start)
            t_end = imported_data[lesson_day][lesson_id][0][6:8] + " " + \
                    imported_data[lesson_day][lesson_id][0][9:]
            # print(t_end)
            start_time = datetime.datetime.strptime(d[:-6] + " " + t_start,
                                                    "%d %H %M")
            start_time = start_time.replace(year=2022, month=9)
            start_time_ready = start_time.strftime("%Y-%m-%dT%H:%M:%S")
            end_time = datetime.datetime.strptime(d[:-6] + " " + t_end,
                                                  "%d %H %M")
            end_time = end_time.replace(year=2022, month=9)
            end_time_ready = end_time.strftime("%Y-%m-%dT%H:%M:%S")

            lesson_name = imported_data[lesson_day][lesson_id][
                1]  # Retrieving lesson name

            lesson_description = imported_data[lesson_day][lesson_id][
                2]  # Retrieving professor name
            if lesson_name == "Физическая культура и спорт ":
                lesson_location = "Волейбольный зал"
            elif lesson_name == "Business English" and lesson_day == 3:
                lesson_location = "Zoom 204 856 2717"
                lesson_description = "Губанова Елена Евгеньевна (пароль на " \
                                     "zoom 125834)"
            elif lesson_name == "Испанский язык" and lesson_day == 4:
                lesson_location = "Дистанционно"
            else:
                lesson_location = imported_data[lesson_day][lesson_id][3]  #
                # Retrieving lesson_location
            create_event(start_time_ready, end_time_ready, lesson_name,
                         lesson_description,
                         lesson_location)  # Creating event in the calendar
            lesson_id += 1
        else:
            lesson_day += 1
    else:
        lesson_day += 1
