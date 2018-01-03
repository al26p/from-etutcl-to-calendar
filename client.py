import xlrd


def import_excel(excel):
    excelHoraire = xlrd.open_workbook(excel)
    mySheet = excelHoraire.sheet_by_index(0)
    return mySheet

def export_horaires(sheet, groupe):
    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if row_value[0] == groupe:
            row_value.remove(groupe)
            for x in range(row_value.count('')):
                row_value.remove('')
            print(row_value)


if __name__ == "__main__":
    path = 'data-test/test 1.xls'
    excelHoraires = import_excel(path)
    export_horaires(excelHoraires, 11)
