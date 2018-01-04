import hashlib
import re
from datetime import datetime, timedelta

import os
import requests
from bs4 import BeautifulSoup

import xlrd

regex = r"((([0-9]{3,4})(a|sa|sd|DL)?)[ \/])?([0-9]{3,4}) x ((([0-9]{3,4})(a|sa|sd|DL)?)[ \/])?([0-9]{3,4})"


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
        self.prise_mission = parseHeure(m[2] or m[4], jour)
        self.lieu_mission = m[3]
        self.debut_service = parseHeure(m[4], jour)
        self.fin_service = parseHeure(m[7] or m[9], jour)
        self.lieu_fin_service = m[8]
        self.fin_mission = parseHeure(m[9], jour)

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

    def to_row_value(self):
        return [
            self.num_mission,
            self.definition,
            self.prise_mission.strftime('%d/%m/%Y %H:%M'),
            self.lieu_mission,
            self.debut_service.strftime('%d/%m/%Y %H:%M'),
            self.fin_service.strftime('%d/%m/%Y %H:%M'),
            self.lieu_fin_service,
            self.fin_mission.strftime('%d/%m/%Y %H:%M'),
        ]

    def __str__(self):
        return ' '.join(self.to_row_value())


class Crawler:

    folder = os.path.join(os.getcwd(), 'data')

    def __init__(self, login, password, url="http://etutcl.fr"):
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

    def load_dispos(self):
        r = self.session.get(self.url + '/dispos', cookies=self.cookies)
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


def parseHeure(h, jour):
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
    works = [Mission(row[i], row[i + 1], jour) for i in range(0, len(row), 2)]
    return works


if __name__ == "__main__":
    c = Crawler(os.getenv('LOGIN'), os.getenv('PASSWORD'))
    c.login_on_site()
    works = c.load_dispos()
    for work in works:
        print(work)
    # list_horaires(excelHoraires, gpe, datetime(2018, 1, 3))
