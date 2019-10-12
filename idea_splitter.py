import os
import re
import openpyxl
from collections import OrderedDict
TITLE=0
DESCRIPTION=1
ATTACHMENT=2
TEAM_NAME=3
THEME=4 #unused - only 1 theme
STATUS=5
LEADER_NAME=6
LEADER_EMAIL=7

raw_projects =[]
clean_projects = OrderedDict()
raw_projects = openpyxl.load_workbook("hashcode-idea-submissions.xlsx")
sheet = raw_projects.active

final_xlsx = openpyxl.Workbook()
final_sheet = final_xlsx.active
row_count =0
headers =''
for row in sheet.iter_rows(min_row=1, min_col=1, max_col=8): #iterate over all rows
    #first remove disqualified idea and get the latest idea
    dlist= []
    
    for cell in row:
        dlist.append(cell.value)
    if (row_count== 0): #skip header processing
        row_count+=1
        headers = dlist
        continue
    print("Processing idea number ",row_count)
    if (dlist[STATUS] != "Disqualified"): #not disqualified
        clean_projects[dlist[TEAM_NAME]] = (tuple(dlist),row_count) #overwrites data, hence only latest idea remains
    row_count+=1

final_sheet.append(headers)
final_clean_projects = sorted(clean_projects.values(),key = lambda x: x[1])
for item in final_clean_projects:
    final_sheet.append(item[0])
    if (item[0][ATTACHMENT]!=None):
        print("WILL GET", item[0][ATTACHMENT])
        #os.system("wget -O "+"./attachments/"+clean_projects[key][TITLE]+".pdf "+clean_projects[key][ATTACHMENT])
final_xlsx.save("final_ideas.xlsx")
