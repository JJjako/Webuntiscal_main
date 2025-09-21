
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

from config_key import WEBUNTIS_USERNAME, WEBUNTIS_PASSWORD, WEBUNTIS_SCHOOL

with webuntis.Session(
                username=WEBUNTIS_USERNAME,
                password=WEBUNTIS_PASSWORD,
                server='borys.webuntis.com',
                school=WEBUNTIS_SCHOOL,
                useragent='Test'
        ).login() as s:
        elements = []
        klasse = s.klassen().filter(name="10b")[0]
        endd_date = datetime.date(2025, 12, 30)
        timetables = s.timetable_extended(klasse=klasse, start=(datetime.date.today()).strftime("%Y%m%d"), end=endd_date.strftime("%Y%m%d"))
        print(s.teachers())