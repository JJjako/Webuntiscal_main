import webuntis
import requests
import datetime
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
import re

from typing import List, TypedDict
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time

from typing import List, TypedDict

from config_key import WEBUNTIS_USERNAME, WEBUNTIS_PASSWORD, WEBUNTIS_SCHOOL, NOTION_TOKEN, NOTION_DATABASE_ID

class ExamEntry(TypedDict):
    Fach: str
    Lehrer: str
    Datum: str
    Zeit: str
    Raum: str
    Beschreibung: str

def fetch_homework():
    hw = []
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--silent')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    try:
        driver.get("https://webuntis.com/")
        
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        
        if True:
            if True:
                search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input")))
        
        search_input.clear()
        search_input.send_keys(WEBUNTIS_SCHOOL)
        
        search_input.send_keys(Keys.RETURN)
        
        short_wait = WebDriverWait(driver, 10)
        try:
            first_result = short_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[role='option']")))
        except:
            try:
                first_result = short_wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".school-list-item")))
            except:
                first_result = short_wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '" + WEBUNTIS_SCHOOL + "')]")))
        
        first_result.click()
        
        wait.until(EC.url_contains("webuntis.com"))
        
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        
        try:
            username = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']")))
        except:
            try:
                username = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='username']")))
            except:
                username = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='username']")))
        
        try:
            password = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
        except:
            password = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='password']")))
        
        username.clear()
        password.clear()
        username.send_keys(WEBUNTIS_USERNAME)
        password.send_keys(WEBUNTIS_PASSWORD)
        
        try:
            login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))
        except:
            try:
                login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.login-button")))
            except:
                login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button")))
        
        login_button.click()
        
        wait.until(EC.url_changes(driver.current_url))
        
        driver.get("https://borys.webuntis.com/student-homework")
        
        iframe = wait.until(EC.presence_of_element_located((By.ID, "embedded-webuntis")))
        driver.switch_to.frame(iframe)
        try:
            time.sleep(1)
            dropdown_control = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/section/div/div/div[1]/form/div[1]/div/div/span/div[2]")))
            dropdown_control.click()

            options_container = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@class='Select-menu-outer']")))
            time.sleep(1)
            year_option = options_container.find_element(By.XPATH, ".//*[text()='2024/2025']")
            time.sleep(1)
            year_option.click()
        except Exception as e:
            pass
        wait.until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'un-table')]/tbody/tr")))
        
        tbody_element = driver.find_element(By.XPATH, "//table[contains(@class, 'un-table')]/tbody")
        
        all_table_rows = tbody_element.find_elements(By.XPATH, ".//tr")
        
        due_soon_homework = []
        not_yet_completed_homework = []
        current_section = None
        
        for row in all_table_rows:
            if "un-table__group-header" in row.get_attribute("class"):
                current_section = row.text.strip()
            elif "un-table__row un-homeworklist-table__row" in row.get_attribute("class"):
                if current_section in ["Due soon", "Not yet completed"]:
                    try:
                        cols = row.find_elements(By.TAG_NAME, "td")
                        if len(cols) >= 5:
                            subject = cols[0].text.strip()
                            teacher = cols[2].text.strip()
                            date_of_assignment = cols[3].text.strip()
                            
                            detail_table = cols[4].find_element(By.TAG_NAME, "table")
                            detail_rows = detail_table.find_elements(By.TAG_NAME, "tr")
                            
                            date_due = ""
                            description = ""
                            
                            if len(detail_rows) >= 1:
                                try:
                                    due_date_cell = detail_rows[0].find_element(By.CLASS_NAME, "un-homeworklist-table__detail__due")
                                    date_due = due_date_cell.text.strip()
                                except:
                                    pass
                                
                            if len(detail_rows) >= 2:
                                try:
                                    homework_row = detail_rows[1]
                                    description_spans = homework_row.find_elements(By.TAG_NAME, "span")
                                    description = " ".join([span.text.strip() for span in description_spans]).strip()
                                except:
                                    pass
                            
                            homework_entry = {
                                'subject': subject,
                                'teacher': teacher,
                                'date_of_assignment': date_of_assignment,
                                'date_due': date_due,
                                'description': description
                            }
                            
                            if current_section == "Due soon":
                                due_soon_homework.append(homework_entry)
                            elif current_section == "Not yet completed":
                                not_yet_completed_homework.append(homework_entry)
                            
                    except Exception as e:
                        continue
        
        if due_soon_homework:
            for entry in due_soon_homework:
                hw.append([entry['subject'],entry['teacher'],entry['date_of_assignment'],entry['date_due'],entry['description']])
        if not_yet_completed_homework:
            for entry in not_yet_completed_homework:
                hw.append([entry['subject'],entry['teacher'],entry['date_of_assignment'],entry['date_due'],entry['description']])

        driver.get("https://borys.webuntis.com/student-exams")
        exams = []
        iframe = wait.until(EC.presence_of_element_located((By.ID, "embedded-webuntis")))
        driver.switch_to.frame(iframe)

        try:
            time.sleep(1)
            dropdown_control = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/section/div/div/div[1]/form/div[1]/div/div/span/div[2]")))
            dropdown_control.click()

            options_container = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@class='Select-menu-outer']")))
            time.sleep(1)
            year_option = options_container.find_element(By.XPATH, ".//*[text()='2024/2025']")
            time.sleep(1)
            year_option.click()

            time.sleep(2)

            wait.until(EC.visibility_of_element_located((By.XPATH, "//table[contains(@class, 'un-table')]/tbody")))

        except Exception as e:
            pass

        wait.until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'un-table')]/tbody/tr")))
        
        tbody_element = driver.find_element(By.XPATH, "//table[contains(@class, 'un-table')]/tbody")
        all_table_rows = tbody_element.find_elements(By.XPATH, ".//tr")
        
        for row in all_table_rows:
            if "un-table__row" in row.get_attribute("class"):
                try:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) >= 5:
                        subject = cols[0].text.strip()
                        teacher = cols[1].text.strip()
                        date = cols[2].text.strip()
                        exam_time = cols[3].text.strip()
                        room = cols[4].text.strip()
                        
                        description = ""
                        if len(cols) > 5:
                            description = cols[5].text.strip()
                        
                        exam_entry = {
                            'subject': subject,
                            'teacher': teacher,
                            'date': date,
                            'time': exam_time,
                            'room': room,
                            'description': description
                        }
                        
                        exams.append([exam_entry['subject'], exam_entry['teacher'], 
                                    exam_entry['date'], exam_entry['time'], 
                                    exam_entry['room'], exam_entry['description']])
                        
                except Exception as e:
                    continue
        
        driver.get("https://borys.webuntis.com/student-class-services")
        iframe = wait.until(EC.presence_of_element_located((By.ID, "embedded-webuntis")))
        driver.switch_to.frame(iframe)
        try:
            time.sleep(1)
            dropdown_control = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/section/div/div/div[1]/form/div[1]/div/div/span/div[2]")))
            dropdown_control.click()

            options_container = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@class='Select-menu-outer']")))
            time.sleep(1)
            year_option = options_container.find_element(By.XPATH, ".//*[text()='2024/2025']")
            time.sleep(1)
            year_option.click()
        except Exception as e:
            pass
        wait.until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'un-table')]/tbody/tr")))
        tbody_element = driver.find_element(By.XPATH, "//table[contains(@class, 'un-table')]/tbody")
        all_table_rows = tbody_element.find_elements(By.XPATH, ".//tr")
        
        class_services = []
        
        for row in all_table_rows:
            try:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 6:
                    service = cols[0].text.strip()
                    student = cols[1].text.strip()
                    class_name = cols[2].text.strip()
                    date_from = cols[3].text.strip()
                    date_to = cols[4].text.strip()
                    text_col = cols[5].text.strip()
                    
                    class_services.append([service, student, class_name, date_from, date_to, text_col])
                    
            except Exception as e:
                continue
        
        return [hw,exams, class_services]

    except Exception as e:
        # print(f"An error occurred: {e}")
        pass
    finally:
        try:
            driver.switch_to.default_content()
        except:
            pass
        driver.quit()
        return [hw,exams, class_services]

SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/calendar.events"]

def add_event_to_notion(event_data, notion_token, database_id):
    url = f"https://api.notion.com/v1/pages"
    headers = {
        "Authorization": f"Bearer {notion_token}",
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28"
    }
    
    payload = {
        "parent": {"database_id": database_id},
        "properties": {
            "Title": {"title": [{"text": {"content": event_data.get("title", "")}}]},
            "Date": {"date": {"start": event_data.get("date", "")}},
            "Location": {"rich_text": [{"text": {"content": event_data.get("location", "")}}]},
            "Description": {"rich_text": [{"text": {"content": event_data.get("description", "")}}]}
        }
    }
    
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 200:
        print("Event added to Notion database successfully!")
    else:
        print(f"Failed to add event: {response.text}")

daysintofuture = 3

def authenticate():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def delete_events_in_date_range(service, start_date, end_date):
    if True:
        start_time = datetime.datetime.combine(start_date, datetime.time.min).isoformat() + 'Z'
        end_time = datetime.datetime.combine(end_date, datetime.time.max).isoformat() + 'Z'
        events_result = service.events().list(
            calendarId='primary',
            timeMin=start_time,
            timeMax=end_time,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        
        events = events_result.get('items', [])
        if not events:
            # print("No events found in the specified date range.")
            return
        
        return events

def create_event(service, start_time, end_time, summary="Sample Event", description="", color_id=None,Fach=None,events_da=None):
    # Always define formatted_string based on the start_time
    if isinstance(start_time, str):
        dt_for_format = datetime.datetime.fromisoformat(start_time)
    else:
        dt_for_format = start_time
    formatted_string = dt_for_format.strftime("%m/%d/%Y")

    # Notch start time by 1 minute if color_id == 11
    if color_id == 11:
        # Parse as naive datetime (strip timezone if present)
        if isinstance(start_time, str):
            start_dt = datetime.datetime.fromisoformat(start_time)
            if start_dt.tzinfo is not None:
                start_dt = start_dt.replace(tzinfo=None)
        else:
            start_dt = start_time
        new_start_dt = start_dt + datetime.timedelta(minutes=1)
        new_start_time = new_start_dt.strftime("%Y-%m-%dT%H:%M:%S")
        print(f"Original start: {start_time}, Notched start: {new_start_time}")
        event = {
            "summary": summary,
            "description": description,
            "start": {"dateTime": new_start_time, "timeZone": "Europe/Brussels"},
            "end": {"dateTime": end_time, "timeZone": "Europe/Brussels"},
            "location": Fach,
            "colorId": color_id,
        }
    else:
        event = {
            "summary": summary,
            "description": description,
            "start": {"dateTime": start_time, "timeZone": "Europe/Brussels"},
            "end": {"dateTime": end_time, "timeZone": "Europe/Brussels"},
            "location": Fach,
            "colorId": color_id,
        }
    for haus in hausis:
        dateobj = datetime.datetime.strptime(haus[3], "%A, %m/%d/%Y")
        if(dateobj.strftime("%m/%d/%Y") == formatted_string and summary == haus[0]):
            event["colorId"] = 3
            event["description"] = haus[4]
            event["reminders"] = {
            "useDefault": False}
            if events_da is not None:
                for eventt in events_da:
                    try:
                        if eventt["end"].get("dateTime") == end_time:
                            service.events().delete(calendarId="primary", eventId=eventt["id"]).execute()
                            print(f"Deleted event with same end time: {eventt.get('summary', '')}")
                    except Exception as e:
                        print(f"Error deleting event with same end time: {e}")
        
    for ka in klassenas:
        if datetime.datetime.strftime(datetime.datetime.strptime(re.search(r'\d{2}/\d{2}/\d{4}', ka[5]).group(0), '%m/%d/%Y').date(),'%m/%d/%Y') == formatted_string and summary == ka[0]:
            print("obacht 2")
            event["colorId"] = "5"
            event["reminders"] = {
            "useDefault": False}
    eventstart = event["start"]
    eventend = event["end"]
    schon_da = False
    if events_da is not None:
        for eventt in events_da:
            try:
                if all(k in eventt for k in ["summary", "location", "start", "end", "colorId"]) and \
                   all(k in event for k in ["summary", "location", "start", "end", "colorId"]):
                    if eventt["summary"] == event["summary"] and \
                       event["location"] == eventt["location"] and \
                       eventstart["dateTime"] == eventt["start"]["dateTime"] and \
                       eventt["end"]["dateTime"] == eventend["dateTime"] and \
                       event["colorId"] == eventt["colorId"]:
                        schon_da = True
            except Exception as e:
                # print(e)
                pass
    
    
    if schon_da:
        pass
    else:
        
        try:
            event = service.events().insert(calendarId="primary", body=event).execute()
            print(f"Event created: {event['htmlLink']}")
            time.sleep(1)
        except HttpError as error:
            # print(f"An error occurred: {error}")
            pass
        

def main():
    creds = authenticate()
    if True:
        with webuntis.Session(
                username=WEBUNTIS_USERNAME,
                password=WEBUNTIS_PASSWORD,
                server='borys.webuntis.com',
                school=WEBUNTIS_SCHOOL,
                useragent='Test'
        ).login() as s:
            elements = []
            service = build("calendar", "v3", credentials=creds)
            klasse = s.klassen().filter(name="9b")[0]
            endd_date = datetime.date(2025, 7, 30)
            events_da = delete_events_in_date_range(service, datetime.date.today(), (datetime.date.today() + datetime.timedelta(days=400)))
            timetables = s.timetable_extended(klasse=klasse, start=(datetime.date.today()).strftime("%Y%m%d"), end=endd_date.strftime("%Y%m%d"))
            i = 0
            Widerholenfertig = False
            while Widerholenfertig != True:
                try:
                    elements.append(timetables[i])
                except:
                    print(i, "nicht da")
                    Widerholenfertig = True
                i = i + 1

            notion_token = NOTION_TOKEN
            database_id = NOTION_DATABASE_ID
        
            tomorrow = (datetime.date.today() + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
            
            for item in elements:
                
                try:
                    if item.code == "irregular":
                        create_event(service, item.start.isoformat(), item.end.isoformat(), item.subjects[0].name, item.bkText,"2",item.rooms[0].name,events_da)
                    else:
                        if item.code == "cancelled":
                            create_event(service, item.start.isoformat(), item.end.isoformat(), item.subjects[0].name,"","11",item.rooms[0].name,events_da)
                        else:
                            
                            create_event(service, item.start.isoformat(), item.end.isoformat(), item.subjects[0].name, "", "9",item.rooms[0].name,events_da)
                except:
                    try:
                        create_event(service, item.start.isoformat(), item.end.isoformat(), item.lstext, "", "9","Kein Code",events_da)


                    except Exception as e:
                        print(item)
                        print(e)   
                        try:
                            create_event(service, item.start.isoformat(), item.end.isoformat(), "Bestonderes Event", "", "9",item.rooms[0].name,events_da)
                        except:
                            create_event(service, item.start.isoformat(), item.end.isoformat(), "Bestonderes Event", "", "9", "",events_da)
                    
    if False:
        for mensch in class_services_data:
            try:
                start_date = datetime.datetime.strptime(mensch[3], "%a, %m/%d/%Y")
                end_date = datetime.datetime.strptime(mensch[4], "%a, %m/%d/%Y")
                event_title = f"{mensch[0]} - {mensch[1]}"
                event_title = f"{mensch[0]} - {mensch[1]}"
                
                start_time = start_date.replace(hour=8, minute=0).isoformat()
                end_time = end_date.replace(hour=8, minute=0).isoformat()
                
                create_event(service, start_time, end_time, event_title, "", "4", "", events_da)
                time.sleep(10)
            except Exception as e:
                # print(f"Error processing event for {mensch[0]} - {mensch[1]}: {e}")
                continue

if __name__ == "__main__":
    alles = fetch_homework()
    hausis = alles[0]
    klassenas  = alles[1]
    class_services_data = alles[2]
    print(class_services_data)
    
    main()
def authenticate():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def event_key(event):
    # Use summary, location, start, and end as the key
    summary = event.get("summary", "")
    location = event.get("location", "")
    start = event["start"].get("dateTime") or event["start"].get("date")
    end = event["end"].get("dateTime") or event["end"].get("date")
    return (summary, location, start, end)

def mainn():
    
    page_token = None
    seen = set()
    duplicates = []
    try:
        while True:
            events_result = service.events().list(
                calendarId='primary',
                pageToken=page_token,
                maxResults=2500,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            events = events_result.get('items', [])
            for event in events:
                key = event_key(event)
                if key in seen:
                    duplicates.append(event)
                else:
                    seen.add(key)
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
        print(f"Found {len(duplicates)} duplicate events.")
        if len(duplicates) == 0:
            return
        for event in duplicates:
            try:
                service.events().delete(calendarId='primary', eventId=event['id']).execute()
                print(f"Deleted duplicate: {event.get('summary', '')} on {event['start']}")
            except HttpError as error:
                print(f"An error occurred: {error}")
        mainn()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    creds = authenticate()
    service = build("calendar", "v3", credentials=creds)
    mainn() 