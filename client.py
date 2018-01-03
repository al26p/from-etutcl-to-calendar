import re
from datetime import datetime

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

    def __init__(self, num_mission, definition):
        self.num_mission = num_mission
        self.definition = definition
        m = re.fullmatch(regex, definition).groups()
        self.prise_mission = parseHeure(m[2] or m[4])
        self.lieu_mission = m[3]
        self.debut_service = parseHeure(m[4])
        self.fin_service = parseHeure(m[7])
        self.lieu_fin_service = m[8]
        self.fin_mission = parseHeure(m[9])
        self.__suppose_lieux_depuis_mission()

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



def import_excel(excel):
    excelHoraire = xlrd.open_workbook(excel)
    mySheet = excelHoraire.sheet_by_index(0)
    return mySheet


def parseHeure(h):
    if h is None:
        return None
    return datetime.strptime(h, '%H%M').time()


def list_horaires(sheet, groupe):
    row = None
    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if row_value[0] == groupe:
            row_value.remove(groupe)
            row = row_value
            break
    if row is None:
        raise Exception('Groupe introuvable!')
    for x in range(row.count('')):
        row.remove('')
    works = [Mission(row[i], row[i+1]) for i in range(0, len(row), 2)]
    print(works)


if __name__ == "__main__":
    path = 'data-test/test 1.xls'
    gpe = 11
    excelHoraires = import_excel(path)
    list_horaires(excelHoraires, gpe, datetime(2018,1,3))
