# выделение из шаблона документа переменных замены, их описания, координаты в ьаблице -> запись в БД таблицы DOCP
#
import os
import re
# from string import Template
# import requests
# from requests.exceptions import HTTPError,ConnectTimeout,ReadTimeout,ConnectionError
# import sys
from sys import argv, exit

import docx
import openpyxl
from sqlalchemy import create_engine, text

#from docx import Document
# import textract
try:
  prname, WorkDoc,file_name = argv
#  WorkDoc = 137
  ss = ''
  sss = ''
  pr_cell = ''
  bef_cell = '' # текст перед переменной
  aft_cell = '' # текс после переменной
  row = 0  # номер строки переменной замены
  sel = 0  # номер ячейки в строке
# =========================================================================
  def DBEngine( dbase='BP', host='localhost'):  # установка соединения c БД
     HOSTNAME = host  # The address of the database server
     PORT = '5432'  # Default port
     USERNAME = 'postgres'  # The user who logs in to the database
     PASSWORD = 'rfn15' if (host == 'localhost') or (host == '127.0.0.1') else 'GTyiii78tre'
     DATABASE = dbase  # Select the database
     DB_URI = 'postgresql+psycopg2://{}:{}@{}:{}/{}'.format(USERNAME, PASSWORD, HOSTNAME, PORT, DATABASE)
     engine = create_engine(DB_URI)
     return engine
#=================поиск параметров в строке и запись их в БД============================================
  def ParAdd(s,doc=False):
     global WorkDoc, row, sel, bef_cell, aft_cell
     n = 0
     j = 0
     m = 0
     ret = -1
     while m != -1:
        m = s.find('${', n);
        if m != -1:
            n = s.find('}', m + 1)
            if n == -1:
                print(' ошибка в шаблоне ${..')
                exit
            ss = "'"+s[m:n + 1]+"'"  # ключ
            if doc:
               sss = ' '
               if aft_cell != '':
                  sss = aft_cell
               else:
                   if bef_cell != '':
                      sss =bef_cell
            else:
                sss = s[j: m-1]
            if (re.search(r"[а-яА-Я]",sss) == None) or (sss.find('${') != -1) :
               sss = ' '
            if s[2] == 'd':
               sss = "текущий день"
            elif s[2] =='m':
                sss = "текущий месяц"
            elif s[2] =='y':
                sss = "текущий год"
            elif s[2] =='s':
                sss = "serial"
            elif s[2] == 'D':
               sss = "день"
            elif s[2] =='M':
                sss = "месяц"
            elif s[2] =='Y':
                sss = "год"
            s = f"'{sss}'"
            sa = 'insert into docp (docp_p,docp_d,docp_r,docp_c,doc_id) values (' + ss + ',' + s + \
                 ',' + str(row)+','+str(sel)+','+str(WorkDoc) + ') on conflict do nothing'
            sq = con.execute(text(sa))
            con.commit()
       #     print (sq)
            ret = 0
        j = n + 1
     return ret
  def GetPL(file_name):
      global WorkDoc, row, sel, bef_cell, aft_cell
      ##=======================textract любой файл===================================================
      #         text = textract.process(file_name)  #,encoding='UTF-8')
      #         text = text.decode('UTF-8')
      #         lines = re.split("\n",text)      #text.splitlines()
      #         per.writelines(lines)
      #         pr_cell=''
      #         for line in lines:
      #                if (len(line) > 1):
      #                   ParAdd(line) #, pr_cell, True)
      #                   pr_cell = line

      ext = os.path.splitext(file_name)
#=================обработка файлов txt=======================================
      if ext[1] in ['.txt', '.html']:
        with open(file_name, 'r', encoding='UTF-8') as fi:
          lines = fi.readlines()
          row = 0
          sel = 0
          for line in lines:
              row += 1
              ParAdd(line)
#=======================python-docx=======================================================
      if ext[1] in ['.doc','.docx']:
         doc = docx.Document(file_name)
#===================абзацы===============================================================
         if len(doc.paragraphs) > 1:
            row = 0
            for par in doc.paragraphs:
                row += 1
                ParAdd(par.text)
#==================таблицы - шаблоны документов==========================================
         n = len(doc.tables)    # кол.таблиц
         if n > 0: #
           for tab in doc.tables:
              # for row in range(len(tab.rows)):
              #    for sel in range(len(tab.columns)):
              #       if tab.cell(row+1,sel+1) != None:
              #         s = tab.cell(row+1,sel+1).text
              #         ce = tab.cell(row + 2, sel)
              #         if ce != None:
              #             pr_cell = ce.text
              #         if len(s) > 3:
              #           print('row=', row, 'sel=', sel, ' ', s, ' ',pr_cell )
              #           ParAdd(s, pr_cell, True)
                n = len(tab.rows)
                row = 0
                for ro in tab.rows:
                   row += 1
#                   n = len(ro.cells)
                   bef_cell = ''
                   aft_cell = ''
                   sel = 0
                   for cell in ro.cells:
                       sel += 1
                       s = cell.text
                       if (len(s) > 3) and (s != bef_cell):   # not in [pr_cell,'.',',']:
                          if row < n:
                            ce = tab.cell(row,sel)
                            if ce != None:
                               aft_cell = ce.text
                          ParAdd(s,True)
                          bef_cell = s
 #                         print(s,row,sel)

#=======================xlrd=======================================================
      if ext[1] =='.xls':
         workbook = xlrd.open_workbook(file_name)
         sheets_name = workbook.sheet_names()
         for names in sheets_name:
             worksheet = workbook.sheet_by_name(names)
             num_rows = worksheet.nrows
             num_cells = worksheet.ncols
             for row in range(num_rows):
                 pr_cell = ''
                 row += 1
                 for sel in range(num_cells):
                     val = worksheet.cell_value(row, sel)
                     sel += 1
                     if val != None:
                       if (len(val) > 2) and (val != pr_cell):
                          ParAdd(val,pr_cell,True)
                     #     print(val)
                          pr_cell = val
#=============================openpyxl==================================================
      if ext[1] =='.xlsx':
         workbook = openpyxl.load_workbook(file_name,data_only=True)
         sheet = workbook.active
         for row in range(sheet.max_row):
             pr_cell = ''
             row += 1
             for sel in range(sheet.max_column):
                 sel += 1
                 val = sheet.cell(row, sel).value
                 if val != None:
                   if (len(val) > 2) and (val != pr_cell):
                      ParAdd(val,pr_cell,True)
                    #  print(val)
                      pr_cell = val
#  ===================преобразование doc  в txt через файл==================================
#             output = pypandoc.convert_file(file_name, 'plain', outputfile=ext[0]+'.txt')
#             print(output)
               # results = re.findall(r'\$\{\w+\}',read_line)
               # for result in results:
               #     descs = re.search(r'\n')
  engine = DBEngine()
  with engine.connect() as con:
     GetPL(file_name)
     exit(0)
        # if ext in ['.doc','.docx']:
        # if ext in ['.xls', '.xlsx']:
#   s = 'select P1,P2 from BPR where BPR_ID='+WorkBPR
# #  P = list(DBFoo(s))
#   d = {
#     'P1': P[0],
#     'P2': P[1]
#   }
#   with open('foo.txt', 'r',encoding='UTF-8') as f,\
#      open('res.txt','w',encoding='UTF-8') as fw:
#      src = f.readlines()
#      for srcl in src:
#         template = Template(srcl)
#         result = template.safe_substitute(d)
#         fw.writelines(result)
#except ConnectTimeout:
#     print('Ошибка подключения к БД BP')
except FileNotFoundError:
     print('file not found-' + file_name)
     exit(-1)