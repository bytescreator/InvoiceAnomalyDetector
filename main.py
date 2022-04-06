import os
import openpyxl
import matplotlib.pyplot as plt

INPUT_SHEETS_FOLDER="./xlses"
REPORT_OUTPUT_FOLDER="./output"

class XLSXData:
    def __init__(self, forFile: str) -> None:
        self.xlsFile = forFile

        self.load()

    def load(self) -> None:
        self.book = openpyxl.load_workbook(self.xlsFile)
        sheets = [self.book[sh] for sh in self.book.sheetnames]

        self.rows = []
        for sheet in sheets:
            for row in sheet:
                self.rows.append([i.value for i in row])

        self.header = self.rows[0]
        self.rows = self.rows[1:] # Throw away headers

        naming = os.path.basename(os.path.splitext(self.xlsFile)[0]).split('-')

        self.year = int(naming[0])
        self.month = int(naming[1])-1
        self.tariff = naming[2]

def matchByIndID(datas: list) -> dict:
    output = {}
    allIds = []
    for rows in datas:
        if rows == None:
            continue

        allIds += [int(row[0]) for row in rows.rows]
    
    [allIds.remove(i) for i in allIds if allIds.count(i) > 1]

    for Id in allIds:
        for index, rows in enumerate(datas):
            if rows == None:
                continue

            for row in rows.rows:
                if int(row[0]) == Id:
                    if not (Id in output.keys()):
                        output.update({Id: {"indexes": [index], "row_datas": [row[1:]]}})
                    elif not (index in output[Id]["indexes"]):
                        output[Id]["indexes"].append(index)
                        output[Id]["row_datas"].append(row[1:])
    
    return output

def load_xlses() -> dict:
    output = {}

    for f in os.listdir(INPUT_SHEETS_FOLDER):
        x=XLSXData(os.path.join(INPUT_SHEETS_FOLDER, f))

        if not (x.tariff in output.keys()):
            output.update({x.tariff: {x.year: [None for i in range(12)]}})

        else:
            if not (x.year in output[x.tariff].keys()):
                output[x.tariff].update({x.year: [None for i in range(12)]})

        output[x.tariff][x.year][x.month] = x

    return output

file_structure = load_xlses()

for tariff in file_structure:
    print(f"{tariff} tarifesi işleniyor...")
    for year in file_structure[tariff]:
        print(f"{tariff} tarifesinin {year} yılı işleniyor...")
        a = matchByIndID(file_structure[tariff][year])
        for id in a:
            print(id, [(i[13],a[id]['indexes'][c]) for c, i in enumerate(a[id]["row_datas"])])
