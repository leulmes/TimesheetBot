import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timezone
import calendar
import re
from dotenv import load_dotenv

load_dotenv()  # take environment variables

class TimesheetEvent:
  def __init__(self, shift_date, hours, location, employee_name, position):
    self.shift_date = shift_date
    self.hours = hours
    self.location = location
    self.employee_name = employee_name
    self.position = position

def grab_location(raw_loc_str):
  return (raw_loc_str[:-1]).strip()

def grab_hours(employee_hours_str):
  hours_regex = re.search(r'\d\.\d', employee_hours_str)
  hours = hours_regex.group(0)
  return float(hours)

def format_position(position):
    match position:
        case 'S':
            return "Summer Teacher"
        case 'M':
            return "Summer Manager"
        case 'L':
            return "Teacher - Lead"
        case 'A':
            return "Teacher - Assistant"
        case 'E':
            return "Special Event"
        case _:
            return "Unknown Event"

# formats an iso string of the format: 2025-08-04 to 08/04/2025
def date_formatter(iso_date):
  split_date_str = iso_date.split('-')
  new_str = split_date_str[1] + "/" + split_date_str[2] + "/" + split_date_str[0]
  return new_str

def grab_calendar_events(f_name, position, credentials):
    target_email = (os.getenv('TARGET_EMAIL'))
    
    timesheetEvents = []
    # Call the Calendar API
    calendar_service = build("calendar", "v3", credentials=credentials)
    # Get first day of the month in ISO8601 String format
    now = datetime.now(timezone.utc)
   
    first_day_month = now.replace(day=1).isoformat() #datetime(2025, 7 % 12, 1, tzinfo=timezone.utc).isoformat()#now.replace(day=1).isoformat()
    first_day_next_month = datetime(now.year, (now.month + 1) % 12, 1, tzinfo=timezone.utc).isoformat()

    print("Getting the upcoming events")
    events_result = (
        calendar_service.events()
        .list(
            calendarId="primary",
            timeMin=first_day_month,
            timeMax=first_day_next_month,
            maxResults=50,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events = events_result.get("items", [])
    if not events:
      print("No upcoming events found.")
      return
    
    # Prints the start and name of the next 50 events
    for event in events:
      # first we try to grab 'dateTime' for timed events, if that fails(we have an all day event), then we grab the 'date' field
      # the 'date' field is present for all day events according to Google Calendar api docs
      start = event["start"].get("dateTime", event["start"].get("date"))
      event_summary = event["summary"]
      email_sender = event["creator"].get("email").lower()
      

      # only process the events that are sent from the target email
      if email_sender == target_email:
        date_regex = re.search(r'\d{4}-\d{2}-\d{2}', start)
        
        # this grabs the name, pos, and hours. i.e "Leul M. (S 8.0)"
        hours_regex = re.search(fr'{f_name}\s[A-Z]{{1}}\.\s\((?:M|S)\s\d\.\d\)', event_summary)
        # grabs everything before the first comma
        location_regex = re.search('^(.+?),', event['location']) 
        location = grab_location(location_regex.group(0))
        employee_hours = grab_hours(hours_regex.group(0))
    
        if date_regex:
            date = date_formatter(date_regex.group(0))
            timesheetEvent = TimesheetEvent(date, employee_hours, location, f_name, format_position(position))
            timesheetEvents.append(timesheetEvent)
    
    return timesheetEvents
        
def create(full_name, position, credentials):
    try:
        event_vals = []
        cal_events = grab_calendar_events((full_name.split(" "))[0], position, credentials)
        sheets_service = build("sheets", "v4", credentials=credentials)
        drive_service = build("drive", "v3", credentials=credentials)

        current_month = datetime.now().month
        current_year = datetime.now().year
        sheet_title = f"{full_name} Timesheet " + monthNumToStr(current_month) + " " + str(current_year)
        
        results = drive_service.files().list(pageSize=1, fields="files(id, name)", q="name='" + sheet_title + "' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",).execute()
        files = results.get("files", [])
       
        if not files:
            # spreadsheet doesn't exist: create it
            spreadsheet = {"properties": {"title": sheet_title}}
            spreadsheet = (
                sheets_service.spreadsheets()
                .create(body=spreadsheet, fields="spreadsheetId")
                .execute()
            )
            spreadsheet_id = (spreadsheet.get('spreadsheetId'))
            sheet_name = "Sheet1" # find way to change the Sheet1 to 'Timesheet' when a sheet is created

            for idx, event in enumerate(cal_events, start=4):
                event_vals.append(
                    {
                        "range": f"{sheet_name}!A{idx}:{idx}",
                        "majorDimension": "ROWS",
                        "values": [[f"{event.shift_date}", f"{event.hours}", f"{event.location}", f"Summer Teacher"]]
                    }
                )
                
            # populate spreadsheet with boilerplate columns
            boilerplate_cols = [
                {
                    "range": f"{sheet_name}!B2",
                    "majorDimension": "COLUMNS",
                    "values": [["Staff Member:"]]
                },
                {
                    "range": f"{sheet_name}!C2",
                    "majorDimension": "COLUMNS",
                    "values": [[f"{full_name}"]]
                },
                {
                    "range": f"{sheet_name}!A3:D4",
                    "majorDimension": "COLUMNS",
                    "values": [["DATE"], ["HOURS"], ["LOCATION"], ["POSITION"]]
                },
                {
                    "range": f"{sheet_name}!F3",
                    "majorDimension": "COLUMNS",
                    "values": [["Totals"]]
                }, 
                {
                    "range": f"{sheet_name}!F9:G14",
                    "majorDimension": "ROWS",
                    "values": [["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G9, $B:B)), 1, 1)", "Summer Manager"], ["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G10, $B:B)), 1, 1)", "Summer Teacher"], ["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G11, $B:B)), 1, 1)", "Teacher - Assistant"], ["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G12, $B:B)), 1, 1)", "Teacher - Lead"], ["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G13, $B:B)), 1, 1)", "Teacher - Online Class"]]
                },
                {
                    "range": f"{sheet_name}!F4:G6",
                    "majorDimension": "ROWS",
                    "values": [["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G4, $B:B)), 1, 1)", "Back Office"], ["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G5, $B:B)), 1, 1)", "ISFT Assistant"]]
                },
                {
                    "range": f"{sheet_name}!F6:G8",
                    "majorDimension": "ROWS",
                    "values": [["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G6, $B:B)), 1, 1)", "ISFT Lead"], ["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G7, $B:B)), 1, 1)", "PSS"]]
                },
                {
                    "range": f"{sheet_name}!F8:G9",
                    "majorDimension": "ROWS",
                    "values": [["=ARRAY_CONSTRAIN(ARRAYFORMULA(SUMIF($D:D, $G8, $B:B)), 1, 1)", "Special Event"]]
                },
            ]
            body = {
                "valueInputOption": "USER_ENTERED",
                "data": boilerplate_cols
            }
            event_vals_body = {
                "valueInputOption": "USER_ENTERED",
                "data": event_vals
            }

            result = (
                sheets_service.spreadsheets()
                .values()
                .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
                .execute()
            )
            result = (
                sheets_service.spreadsheets()
                .values()
                .batchUpdate(spreadsheetId=spreadsheet_id, body=event_vals_body)
                .execute()
            )
            # formatting requests: change font size, bold text, change border color for calculation box
            reqs = {
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 0,
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "horizontalAlignment": "CENTER"
                                }
                            },
                            "fields": "userEnteredFormat.horizontalAlignment"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startColumnIndex": 6,
                                "endColumnIndex": 29
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "horizontalAlignment": "LEFT"
                                }
                            },
                            "fields": "userEnteredFormat.horizontalAlignment"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 1,
                                "endRowIndex": 3,
                                "startColumnIndex": 0,
                                "endColumnIndex": 6
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "fontFamily": "Calibri",
                                        "fontSize": 12,
                                    },
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 3,
                                "startColumnIndex": 5,
                                "endColumnIndex": 7
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "fontFamily": "Calibri",
                                        "fontSize": 12,
                                    },
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 3,
                                "startColumnIndex": 0,
                                "endColumnIndex": 4
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "fontFamily": "Calibri",
                                        "fontSize": 11,
                                    },
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 1,
                                "endRowIndex": 2,
                                "startColumnIndex": 1,
                                "endColumnIndex": 2
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "bold": True,
                                    },
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.bold"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 2,
                                "endRowIndex": 3,
                                "startColumnIndex": 0,
                                "endColumnIndex": 6
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "bold": True
                                    },
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.bold"
                        }
                    },
                    { # update width of location & position col
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": 0,
                                "dimension": "COLUMNS",
                                "startIndex": 2,
                                "endIndex": 4
                            },
                            "properties": {
                            "pixelSize": 194
                            },
                            "fields": "pixelSize"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 3,
                                "startColumnIndex": 1,
                                "endColumnIndex": 2
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "numberFormat": {
                                        "type": "NUMBER",
                                        "pattern": "#.0#"
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.numberFormat"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 3,
                                "startColumnIndex": 0,
                                "endColumnIndex": 1
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "numberFormat": {
                                        "type": "DATE",
                                        "pattern": "m/d/yyy"
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.numberFormat"
                        }
                    },
                    { # color the background for calculations box. rgb(191, 191, 191)
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 2,
                                "endRowIndex": 13,
                                "startColumnIndex": 5,
                                "endColumnIndex": 8
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": {
                                        "red": 191 / 255.0,
                                        "green": 191 / 255.0,
                                        "blue": 191 / 255.0
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.backgroundColor"
                        }
                    },
                    { # make the borders the same color as above to get 'merged cell' look
                        "repeatCell": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 2,
                                "endRowIndex": 13,
                                "startColumnIndex": 5,
                                "endColumnIndex": 8
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "borders": {
                                        "top": {
                                            "style": "SOLID",
                                            "color": {
                                                "red": 191 / 255.0,
                                                "green": 191 / 255.0,
                                                "blue": 191 / 255.0
                                            }
                                        },
                                        "bottom": {
                                            "style": "SOLID",
                                            "color": {
                                                "red": 191 / 255.0,
                                                "green": 191 / 255.0,
                                                "blue": 191 / 255.0
                                            }
                                        },
                                        "left": {
                                            "style": "SOLID",
                                            "color": {
                                                "red": 191 / 255.0,
                                                "green": 191 / 255.0,
                                                "blue": 191 / 255.0
                                            }
                                        },
                                        "right": {
                                            "style": "SOLID",
                                            "color": {
                                                "red": 191 / 255.0,
                                                "green": 191 / 255.0,
                                                "blue": 191 / 255.0
                                            }
                                        }
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.borders"
                        }
                    },
                    {
                        "setDataValidation": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": 3,
                                "startColumnIndex": 3,
                                "endColumnIndex": 4
                            }, 
                            "rule": {
                                "condition": {
                                    "type": "ONE_OF_LIST",
                                    "values": [
                                        {
                                            "userEnteredValue": "Back Office",
                                        },
                                        {
                                            "userEnteredValue": "ISFT Assistant",
                                        },
                                        {
                                            "userEnteredValue": "ISFT Lead",
                                        },
                                        {
                                            "userEnteredValue": "PSS",
                                        },
                                        {
                                            "userEnteredValue": "Special Event",
                                        },
                                        {
                                            "userEnteredValue": "Summer Manager",
                                        },
                                        {
                                            "userEnteredValue": "Summer Teacher",
                                        },
                                        {
                                            "userEnteredValue": "Teacher - Assistant",
                                        },
                                        {
                                            "userEnteredValue": "Teacher - Lead",
                                        },
                                        {
                                            "userEnteredValue": "Teacher - Online Class",
                                        }
                                    ]
                                },
                                "showCustomUi": True,
                                "strict": True
                            }
                        }
                    },
                ]
            }
            result = (
                sheets_service.spreadsheets()
                .batchUpdate(spreadsheetId=spreadsheet_id, body=reqs)
                .execute()
            )
            
            print("A new spreadsheet created :O")
            print(f"Spreadsheet ID: {spreadsheet_id}")

        
        else:
            # spreadsheet already exists: don't create a new one
            print("I like kiwi :)")
            pass
        
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def monthNumToStr(month_num):
    match month_num:
        case 1:
            return "January"
        case 2:
            return "February"
        case 3: 
            return "March"
        case 4:
            return "April"
        case 5:
            return "May"
        case 6:
            return "June"
        case 7:
            return "July"
        case 8:
            return "August"
        case 9:
            return "September"
        case 10:
            return "October"
        case 11:
            return "November"
        case 12:
            return "December"
        case 13:
            return "Oops, something went wrong with the month :("
