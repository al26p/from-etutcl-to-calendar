from __future__ import print_function
import hashlib
import re
from datetime import datetime, timedelta
import pytz
import json
import os
import shutil
import requests
from bs4 import BeautifulSoup

import xlrd

import httplib2

from apiclient import discovery
from oauth2client import client, tools
from oauth2client.file import Storage


regex = r"((([0-9]{3,4})(a|sa|sd|DL)?)[ \/])?([0-9]{3,4}) x ((([0-9]{3,4})(a|sa|sd|DL)?)[ \/])?([0-9]{3,4})"


with open('data/config/config.json', 'r') as outfile:
    config = json.load(outfile)

SCOPES = config['google-calendar']['scopes']
APPLICATION_NAME = config['google-calendar']['app_name']
CLIENT_SECRET_FILE = config['google-calendar']['client_secret_file']
CALENDAR_ID = config['google-calendar']['calendar_id']
ETUTCL_USER = config['etutcl']['user']
ETUTCL_PWD = config['etutcl']['password']
API_KEY = config['google-calendar']['api_key']
ETUTCL_URL = config['etutcl']['url']
flags = []


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    credential_dir = 'data/credentials'
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials


class Mission:
    num_mission = None
    prise_mission = None
    lieu_mission = None
    debut_service = None
    fin_service = None
    lieu_fin_service = None
    fin_mission = None
    definition = ''

    def __init__(self, num_mission, definition, jour=datetime.today()):
        self.num_mission = num_mission
        self.definition = definition
        self.__parse_definition_horraire(definition, jour)
        self.__suppose_lieux_depuis_mission()

    def __parse_definition_horraire(self, definition, jour):
        m = re.fullmatch(regex, definition).groups()
        self.prise_mission = parse_heure(m[2] or m[4], jour)
        self.lieu_mission = m[3]
        self.debut_service = parse_heure(m[4], jour)
        self.fin_service = parse_heure(m[7] or m[9], jour)
        self.lieu_fin_service = m[8]
        self.fin_mission = parse_heure(m[9], jour)

    def __suppose_lieux_depuis_mission(self):
        if self.lieu_mission is None:
            if self.num_mission.startswith('A') or self.num_mission.startswith('Dis'):
                self.lieu_mission = 'S'
            elif self.num_mission.startswith('B'):
                self.lieu_mission = 'C'
        if self.lieu_fin_service is None:
            if self.num_mission.startswith('A') or self.num_mission.startswith('Dis'):
                self.lieu_fin_service = 'S'
            elif self.num_mission.startswith('B'):
                self.lieu_fin_service = 'C'

    @property
    def lieu_mission_humain(self):
        return self.__lieu_a_humain(self.lieu_mission)

    @property
    def lieu_fin_mission_humain(self):
        return self.__lieu_a_humain(self.lieu_fin_service)

    def __lieu_a_humain(self, lieu):
        if lieu == 'S':
            return 'Soie'
        elif lieu == 'C':
            return 'CH'
        elif lieu == 'a':
            return 'ATE'
        elif lieu == 'sa':
            return 'SA'
        elif lieu == 'sd':
            return 'SD'
        elif lieu == 'DL':
            return 'DL'
        else:
            return lieu

    def to_row_value(self):
        return [
            self.num_mission,
            self.definition,
            self.prise_mission.strftime('%d/%m/%Y %H:%M'),
            self.lieu_mission_humain,
            self.debut_service.strftime('%d/%m/%Y %H:%M'),
            self.fin_service.strftime('%d/%m/%Y %H:%M'),
            self.lieu_fin_mission_humain,
            self.fin_mission.strftime('%d/%m/%Y %H:%M'),
        ]

    def event_to_google(self):
        """Shows basic usage of the Google Calendar API.

            Creates a Google Calendar API service object and outputs a list of the next
            10 events on the user's calendar.
            """
        credentials = get_credentials()
        http = credentials.authorize(httplib2.Http())
        credentials.refresh(http)
        service = discovery.build('calendar', 'v3', http=http)

        summary = self.num_mission + ' : ' + self.lieu_mission_humain + ' - ' + self.lieu_fin_mission_humain

        ### Test if event is already there
        paris = pytz.timezone('Europe/Paris')
        time = paris.localize(self.prise_mission)
        time = str(time)
        time = time.replace(' ', 'T')
        isThere = False
        print(time)

        eventsResult = service.events().list(
            calendarId=CALENDAR_ID, timeMin=time, maxResults=3, singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])
        for event in events:
            if event['summary'] == summary:
                isThere = True

        if not isThere:
            event = {
                'summary': summary,
                'location': self.definition,
                'description': self.num_mission + ' : ' + self.definition,
                'start': {
                    'dateTime': self.prise_mission.isoformat(),
                    'timeZone': 'Europe/Paris',
                },
                'end': {
                    'dateTime': self.fin_mission.isoformat(),
                    'timeZone': 'Europe/Paris',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'popup', 'minutes': 60},
                        {'method': 'popup', 'minutes': 30},
                        {'method': 'popup', 'minutes': 15},
                    ],
                },
            }

            event = service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
            print(self.__str__())
            print('Event created: %s' % (event.get('htmlLink')))

    def __str__(self):
        return ' '.join(self.to_row_value())


class Crawler:

    folder = os.path.join(os.getcwd(), 'data/xls')

    def __init__(self, login, password, url):
        try:
            os.mkdir(self.folder)
        except FileExistsError:
            pass
        self.password = password
        self.login = login
        self.url = url
        self.session = requests.Session()
        self.cookies = requests.cookies.RequestsCookieJar()

    def login_on_site(self):
        r = self.session.post(
            self.url + '/login/request',
            data={'username': self.login, 'password': self.password}, cookies=self.cookies)
        content = str(r.content, 'utf-8')
        if r.status_code > 300 or 'Mot de passe incorrect' in content:
            raise Exception('Wrong login!')
        if r.history:
            for historicResponse in r.history:
                self.cookies.update(historicResponse.cookies)
        self.cookies.update(r.cookies)

    def load_dispos(self, semaine):
        print(semaine)
        r = self.session.get(self.url + '/dispos/' + semaine, cookies=self.cookies)
        content = str(r.content, 'utf-8')
        soup = BeautifulSoup(content, 'html.parser')
        tables = soup.find_all('table', "data")
        works = []
        for table in tables:
            first_line = table.find('tbody').findAll('tr')[0].findAll('td')
            day = first_line[0]
            day = str(day.find(text=re.compile("[0-9]{2}/[0-9]{2}/[0-9]{2}"))).strip()
            data = first_line[-1]
            a_tag = data.find('a')
            if a_tag is None:
                continue
            group = int(data.find(text=re.compile("Groupe attribué")).nextSibling.string)
            link = a_tag.attrs['href']
            day = datetime.strptime(day, "%d/%m/%y")
            self.download(link)
            works += list_horaires(import_excel(self.path_for(link)), group, day)
        return works

    def download(self, file):
        path = self.path_for(file)
        r = self.session.get(self.url + file, stream=True, cookies=self.cookies)
        if r.status_code == 200:
            with open(path, 'wb') as f:
                for chunk in r:
                    f.write(chunk)

    def path_for(self, file):
        return os.path.join(self.folder, hashlib.md5(file.encode('utf-8')).hexdigest())


def import_excel(excel):
    excelHoraire = xlrd.open_workbook(excel)
    mySheet = excelHoraire.sheet_by_index(0)
    return mySheet


def parse_heure(h, jour):
    if h is None:
        return None
    time = datetime.strptime(h, '%H%M').time()
    if time < datetime.strptime('300', '%H%M').time():
        jour += timedelta(days=1)
    return datetime.combine(jour, time)


def list_horaires(sheet, groupe, jour):
    row = None
    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if row_value[0] == groupe:
            row_value.remove(groupe)
            row = row_value
            break
    if row is None:
        raise Exception('Groupe '+ groupe + ' introuvable!')
    for x in range(row.count('')):
        row.remove('')
    works = [Mission(row[i], row[i+1], jour) for i in range(0, len(row), 2)]
    return works


def clean():
    folder = 'data/xls'
    for the_file in os.listdir(folder):
        file_path = os.path.join(folder, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(e)
    os.rmdir('data/xls')


if __name__ == "__main__":
    c = Crawler(ETUTCL_USER, ETUTCL_PWD, ETUTCL_URL)
    c.login_on_site()
    annee = 2015
    semaine = 1
    parisTz = pytz.timezone('Europe/Paris')
    now = parisTz.localize(datetime.now())
    annee = now.strftime('%Y')
    semaine = now.strftime('%W')
    for x in range(0, 2):
        semaine = int(semaine) + x
        annee = str(annee)
        semaine = str(semaine)
        urlSemaine = annee + '/' + semaine
        missions = c.load_dispos(urlSemaine)
        [mission.event_to_google() for mission in missions]
    clean()