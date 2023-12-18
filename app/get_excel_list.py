import xlrd
import time
import json


def read_excel(file_address):
    try:
        data = xlrd.open_workbook(file_address, formatting_info=True)
        sheet = data.sheet_by_index(0)
        raw_res = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        for i in range(len(raw_res)):
            raw_res[i] = raw_res[i][0:4]
        raw_res.pop(0)
        return raw_res
    except:
        return [["Нет данных"]]


def fetch_list_classes(filepath):
    try:
        unsorted_data = read_excel(filepath)
        grades_id = set()
        for i in unsorted_data:
            grades_id.add(i[-1])
        grades_id_list = list(grades_id)
        grades_id_list.sort()
        return grades_id_list
    except:
        return ["Нет данных"]


def fetch_data(file_address):
    unsorted_data = read_excel(file_address)
    classes = fetch_list_classes(file_address)
    print(classes, "\n\n\n\n\n")
    data = fetch_dict_classes(classes, file_address)
    print(data, "\n\n\n\n\n")

    return data


def fetch_dict_classes(filepath):
    unsorted_data = read_excel(filepath)
    classes = fetch_list_classes(filepath)
    res = dict()
    for j in classes:
        res[j] = []
    for i in range(len(unsorted_data)):
        res[unsorted_data[i][-1]].append(list(unsorted_data[i][0:4]))
    
    return res

if __name__ == "__main__":
    with open('config.json') as json_config:
        config = json.load(json_config)

    start = time.time()

    fetch_data(config["filepath_xls"])

    end = time.time()
    
    print(f'Done! Completed in {round(end-start, 3)} sec')

