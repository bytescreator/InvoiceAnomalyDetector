from functools import lru_cache

import os
import json

import xlrd

INPUT_SHEETS_FOLDER="./xlses"
REPORT_OUTPUT_FOLDER="./output"
SUBS_JSON_LOC="./subs.json"

ELECTRIC_SUBSCRIBER_CODE="1"

class XLSXData:
    def __init__(self, forFile: str) -> None:
        self.xlsFile = forFile

        self.load()

    def load(self) -> None:
        self.book = xlrd.open_workbook(self.xlsFile)
        sheets = [sh for sh in self.book.sheets()]

        self.rows = []
        for sheet in sheets:
            self.rows += [[x.value for x in sheet.row(i)] for i in range(sheet.nrows)]

        self.header = self.rows[0]
        self.rows = self.rows[1:] # Throw away headers

        naming = os.path.basename(os.path.splitext(self.xlsFile)[0]).split('-')

        self.year = int(naming[0])
        self.month = int(naming[1])-1
        self.tariff = naming[2]

def get_mapping():
    map = []

    with open(SUBS_JSON_LOC, mode='r') as f:
        subs=json.load(f)
        createList = lambda x: [
             x["SAYAC_NO"]
            ,x["ISLETME_KODU"]
            ,x["ABN_ID"]
            ,x["ABN_TUR_ID"]
            ,x["ILCE_ADI"]
        ]
        for i in subs:
            if (i["ABN_TUR_ID"] == ELECTRIC_SUBSCRIBER_CODE) and (i['AKTIF_PASIF'] == 'E'):
                map.append(createList(i))

    return map

@lru_cache(500)
def get_ind_id(sub):
    all_scores = []
    _map = get_mapping()
    for saved_sub in _map:
        if saved_sub[1] == '':
            saved_sub[1] = "00.00.00.00.00"
        if saved_sub[1].count(".") != 4:
            saved_sub[1] = "00.00.00.00.00"

        all_scores.append(get_match_score(sub[0], saved_sub[0], sub[1], saved_sub[1]))

    _max = (0, 0)
    _last_max_index = None

    for index, score in enumerate(all_scores):
        _avg_score = (score[0]*3+score[1])/4

        if (score[0] > _max[0] and score[1] > _max[1]) or (_avg_score > (_max[0]*3+_max[1])/4):
            _max = score
            _last_max_index = index

    return _map[_last_max_index][2]

def get_match_score(abn1, abn2, ist1, ist2):
    first_op_num = [int(i) for i in ist1.split(".")]
    second_op_num = [int(i) for i in ist2.split(".")]
    first_sub_num=abn1.zfill(20)
    second_sub_num=abn2.zfill(20)

    ordered_sub_match_count=0
    for index,part in enumerate(first_sub_num):
        if part == second_sub_num[index]:
            ordered_sub_match_count+=1

    unordered_sub_match_count=0
    for part in first_sub_num:
        if part in second_sub_num:
            unordered_sub_match_count+=1

    sub_match_score = ((ordered_sub_match_count+unordered_sub_match_count))/2

    ordered_op_match_count=0
    for index, part in enumerate(first_op_num):
        if part == second_op_num[index]:
            ordered_op_match_count+=1

    unordered_op_match_count=0
    for part in first_op_num:
        if part in second_op_num:
            unordered_op_match_count+=1

    op_match_score = ((ordered_op_match_count+unordered_op_match_count))/2

    return sub_match_score, op_match_score

def match(datas: list) -> dict:
    output = {}
    allSubs = []

    for rows in datas:
        if rows == None:
            continue
        allSubs+=[(row[2], row[1], row[3]) for row in rows.rows]
        # Abn No, isletme kod, sayac no (optimizasyon için kullanılacak)
    
    [allSubs.remove(i) for i in allSubs if allSubs.count(i) > 1]

    for sub in allSubs:
        for index, rows in enumerate(datas):
            if rows == None:
                continue

            for row in rows.rows:
                if sub[2] == row[3]:
                    Id = get_ind_id(sub[:-1])
                    if not (Id in output.keys()):
                        output.update({Id: {"indexes": [index], "row_datas": [row[1:]]}})
                    elif not (index in output[Id]["indexes"]):
                        output[Id]["indexes"].append(index)
                        output[Id]["row_datas"].append(row[1:])

                    print(f"get_ind_id({sub[:-1]}) -> {Id}")
    
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
        a = match(file_structure[tariff][year])
        for id in a:
            print(id, [(i[13],a[id]['indexes'][c]) for c, i in enumerate(a[id]["row_datas"])])
        print(len(a))

