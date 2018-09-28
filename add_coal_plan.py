import xlrd
import os
import re
import csv
import re
from datetime import datetime
import xlsxwriter 
 
def main():
    INPUT_FOLDER = "oper"
    STATIONS_TITLES_FILE = "stations_ids.csv"
    with open(STATIONS_TITLES_FILE, "r") as stf:
        lines = csv.reader(stf)
        lines = [l for l in lines]
        lines = [l for l in lines[1:]]
        stations_dict = {}
        for l in lines:
            stations_dict[l[0]] = l[1]
    files = os.listdir(INPUT_FOLDER)
    files = [f for f in files if f.endswith(".xls") or f.endswith(".xlsx")] 
    coal_plan_dict = {}
    for f in files:
        print(f)
        wb = xlrd.open_workbook(os.path.join(INPUT_FOLDER, f) )
        file_dict = load_workbook(wb, stations_dict)
        coal_plan_dict.update(file_dict)
    return coal_plan_dict

def is_blank(cell):
    NOT_SPACE = re.compile("\S+")
    return NOT_SPACE.search(str(cell.value)) == None

def coal_type_refine(s, coal_type_string):
    NOT_SPACE = re.compile("\S+")
    if type(s) == type(float()) or NOT_SPACE.search(str(s)) == None:
        s = coal_type_string.strip()
    if s.startswith("А"):
        s = "АШ+П"
    elif s.startswith("Г"):
        s = "ГД"
    elif s.startswith("П"):
        s = "АШ+П"
    return s

def load_workbook(wb, stations_dict):
    SHEET_NUMBER = 2
    sheet = wb.sheet_by_index(SHEET_NUMBER)
    ncols = sheet.ncols
    nrows = sheet.nrows
    date = xlrd.xldate_as_tuple(sheet.cell(0, 0).value,0)
    date = datetime(*date[0:6]).strftime("%d.%m.%Y")
    print(date)
    plan_dict = {}
    plan_dict[date] = {}
    plan_id = None
    coal_type_string = None
    for i in range(1, nrows):
        row = [sheet.cell(i, c_number) for c_number in range(ncols)]
        if plan_id:
            if not is_blank(row[0]):
                if row[0].value != row[0].value.upper():
                    station = row[0].value.replace("Ө", "").strip()
                    if "вугілля" in station.lower():
                        coal_type_string = station
                    else:
                        if station in stations_dict.keys():
                            if row[plan_id].value != "":
                                plan_coal = round(float(row[plan_id].value), 1)
                            else:
                                plan_coal = ""
                            coal_type = coal_type_refine(row[1].value, coal_type_string)
                            station = stations_dict[station]
                            if not station in plan_dict[date].keys():
                                plan_dict[date][station] = {}
                            plan_dict[date][station][coal_type] = plan_coal        
        else:
            for i in range(len(row)):
                if str(row[i].value).strip().startswith("Запас"):
                    plan_id = i-1
    return plan_dict
