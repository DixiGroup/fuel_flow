import csv
import xlrd
import xlsxwriter
import datetime
import os
import sys
import re
import add_coal_plan

FIELD_FILE_1 = "fields_t012331.csv"
STAIONS_FILE = "pp_correspondence.csv"
MONTHES_DICT = {"01":"січня" , "02":"лютого", "03":"березня", "04":"квітня", "05":"травня", "06":"червня", "07":"липня", "08":"серпня", "09":"вересня", "10":"жовтня", "11":"листопада", "12":"грудня"}
INPUT_FOLDER_1 = "t012331"
HEADERS = ["date", "company", "company_code", "plant_name", "plant_code", "fuel_type", "income", "spend", "reserve_plan", "reserve_fact"]
OUTPUT_FOLDER = "opendata"
FILENAME_TEMPLATE = "fuel_on_stations_{date:s}"

def load_workbook(wb):
    sheet = wb.sheet_by_index(0)
    nrows = sheet.nrows
    sheet_dict = {}
    for h in HEADERS:
        sheet_dict[h] = []
    for i in range(0, nrows):
        if month_in_row(i, sheet):
            cellvalue = sheet.cell(i, 0).value
            parts = cellvalue.split()
            month = month_in_row(i, sheet)
            day = re.search("\d{2}", cellvalue).group()
            year = re.search("\d{4}", cellvalue).group()
            date = ".".join([day, month, year])
        else:
            #print(pp_dict)
            #print(sheet.cell(i, 0).value)
            if sheet.cell(i, 0).value in pp_dict.keys() and sheet.cell(i, 0).value != "":
                pp = sheet.cell(i, 0).value
                print(pp, pp_dict[pp]['plant_name'])
                #print(pp)
                #вирівняти кількість рядків      
                for k in columns_dict.keys():
                    sheet_dict["date"].append(date)
                    sheet_dict['company'].append(pp_dict[pp]['company'])
                    sheet_dict['company_code'].append(pp_dict[pp]['company_code'])
                    sheet_dict['plant_name'].append(pp_dict[pp]['plant_name'])       
                    sheet_dict["fuel_type"].append(k)
                    value_type_dict = columns_dict[k]
                    for v in value_type_dict:
                        sheet_dict[v].append(sheet.cell(i, int(value_type_dict[v])).value)
                    if k == "газ":
                        sheet_dict["income"].append("")
                        sheet_dict["reserve_fact"].append("")
                    
    return sheet_dict

def format_code(v):
    s = str(v).split('.')[0]    
    if len(s) < 8:
        s = "0" * (8 - len(s)) + s
    return s

def month_in_row(row_number, sheet):
    cellvalue = sheet.cell(row_number, 0).value
    for k in MONTHES_DICT.keys():
        if MONTHES_DICT[k] in str(cellvalue):
            return k

def dict_to_list(dict_, headers):
    l = []
    for i in range(len(dict_[headers[0]])):
        new_l = []
        for h in headers:
            #print(h, i)
            new_l.append(dict_[h][i])
        l.append(new_l)
    return l

with open(STAIONS_FILE, "r") as sf:
    pp_reader = csv.reader(sf)
    pp_lines = [line for line in pp_reader]
    pp_lines = [pp_lines[i] for i in range(len(pp_lines)) if i > 0]
pp_dict = {}
for l in pp_lines:
    pp_dict[l[3]] = {}
    pp_dict[l[3]]['company'] = l[0]
    pp_dict[l[3]]['company_code'] = format_code(l[1]) if l[1] != "NA" else ""
    pp_dict[l[3]]['plant_name'] = l[2]


monthes_dict = {}
for k in MONTHES_DICT:
    monthes_dict[MONTHES_DICT[k]] = k

columns_dict = {}
with open(FIELD_FILE_1, "r") as ff:   
    lines = csv.reader(ff)
    lines = [l for l in lines]
    lines = [lines[i] for i in range(len(lines)) if i > 0]
for l in lines:
    if not columns_dict.get(str(l[2])):
        columns_dict[str(l[2])] = {}
    columns_dict[str(l[2])][l[1]] = l[0]    
    """if not columns_dict[str(l[0])].get(l[2]):
        columns_dict[str(l[0])][l[2]] = {}

        columns_dict[str(l[0])][l[2]] = [l[1]]
    else:
        columns_dict[str(l[0])][l[2]].append(l[1])
    columns_dict[str(l[0])]['value_type'] = l[1]
    columns_dict[str(l[0])]['fuel_type'] = l[2]"""
print(columns_dict)
files = os.listdir(INPUT_FOLDER_1)
files = [f for f in files if f.endswith(".xls")]
fuel_dict = {}
for f in files:
    print(f)
    wb = xlrd.open_workbook(os.path.join(INPUT_FOLDER_1, f), formatting_info = True)
    file_dict = load_workbook(wb)
    if fuel_dict == {}:
        fuel_dict = file_dict
    else:
        for k in fuel_dict.keys():
            fuel_dict[k] += file_dict[k]
print(fuel_dict)
#сюди треба додати коди станцій, коли буде існувати довідник
fuel_dict["plant_code"] = [""] * len(fuel_dict[HEADERS[0]])
plan_dict = add_coal_plan.main()
print(plan_dict)
#print(plan_dict["29.07.2018"])
for i in range(len(fuel_dict[HEADERS[0]])):
    #print(fuel_dict['date'][i], fuel_dict['plant_name'][i], fuel_dict['fuel_type'][i])
    plan = plan_dict.get(fuel_dict['date'][i], {}).get(fuel_dict['plant_name'][i], {}).get(fuel_dict['fuel_type'][i],"")
    fuel_dict['reserve_plan'].append(plan)
    #print(plan)
#print(len(fuel_dict['income']), len(fuel_dict['reserve_fact']), len(fuel_dict['spend']))
fuel_dict['date'] = list(map(lambda x: datetime.datetime.strptime(x, "%d.%m.%Y"), fuel_dict["date"]))
date_to_filename = max(fuel_dict['date']).strftime("%d_%m_%Y")
filename = os.path.join(OUTPUT_FOLDER, FILENAME_TEMPLATE.format(date = date_to_filename))
coal_list = dict_to_list(fuel_dict, HEADERS)
coal_list = sorted(coal_list, key=lambda x: x[0],reverse=True)
with open(filename + ".csv", "w", newline="") as cfile:
    csvwriter = csv.writer(cfile)
    csvwriter.writerow(HEADERS)
    for i in range(len(coal_list)):
        l = coal_list[i][:]
        l[0] = datetime.datetime.strftime(l[0],"%d.%m.%Y")
        csvwriter.writerow(l)
out_wb = xlsxwriter.Workbook(filename + ".xlsx")
worksheet = out_wb.add_worksheet()
datef = out_wb.add_format({'num_format':"dd.mm.yyyy"})
numf = out_wb.add_format({'num_format':"0.00"})
headerf = out_wb.add_format({'bold':True})
for i in range(len(HEADERS)):
    worksheet.write(0, i, HEADERS[i], headerf)
for i in range(len(coal_list)):
    for j in range(len(HEADERS)):
        if j == 0:
            worksheet.write(i+1, j, coal_list[i][j], datef)
        elif j >  5:
            worksheet.write(i+1, j, coal_list[i][j], numf)
        else:
            worksheet.write(i+1, j, coal_list[i][j])
out_wb.close()      