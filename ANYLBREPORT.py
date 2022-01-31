# -*- coding: utf-8 -*-
"""
Created on Fri Jul 10 17:14:22 2020

@author: Nandal
"""

import pandas as pd

import _datetime
from tkinter import *
import tkinter as tk
import matplotlib.pyplot as plt
plt.style.use('seaborn-whitegrid')
import re
import numpy as np

def territory():
  global ter
  ter = entryT.get()

def csv_reader():
  global data
  global dataD
  global data_V
  territory()
  global ter
  #reading drafts
  dataD = pd.read_csv('data.csv',  encoding = 'cp1252', sep = ',', low_memory= False)
  dataD = dataD.loc[dataD['TERRITORYNAME'] == ter]
  dataD = dataD.loc[dataD['LAYOUT_NM'] == 'Draft Complaints']
  dataD = dataD.loc[dataD['STATUSCODE'] != 'Not a Complaint']
#sanitizing the data with proper dates
  try:
    dataD = dataD.drop(dataD[dataD['COMPLAINTCLOSEDON'] == '111'].index)
  except:
    None
  try:
    dataD['CREATEDON'] = pd.to_datetime(dataD['CREATEDON'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    dataD['CREATEDON'] = pd.to_datetime(dataD['CREATEDON'].str.split().str[0], format='%d-%m-%y').dt.date
  try: 
    dataD['LASTMODIFIEDON'] = pd.to_datetime(dataD['LASTMODIFIEDON'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    dataD['LASTMODIFIEDON'] = pd.to_datetime(dataD['LASTMODIFIEDON'].str.split().str[0], format='%d-%m-%y').dt.date
  try:
    dataD['COMPLAINTCLOSEDON'] = pd.to_datetime(dataD['COMPLAINTCLOSEDON'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    dataD['COMPLAINTCLOSEDON'] = pd.to_datetime(dataD['COMPLAINTCLOSEDON'].str.split().str[0], format='%d-%m-%y').dt.date

  dataD = dataD.filter(['CASEID',
                        'STATUSCODE',
                        'CREATEDON',
                        'LASTMODIFIEDBYNAME',
                        'LASTMODIFIEDON',
                        'ASSIGNEDTONAME'])

  dataD = dataD.rename(columns = {'CASEID':'Draft ID',
                        'STATUSCODE':'Draft Status',
                        'CREATEDON':'Date of Origin',
                        'LASTMODIFIEDBYNAME':'Draft Last Modified By',
                        'LASTMODIFIEDON':'Status Changed on',
                        'ASSIGNEDTONAME':'Current Owner'})
  dataD = dataD.set_index('Draft ID')
  td = _datetime.date.today()
  dataD['Pending Days from today'] =  dataD['Date of Origin'].apply(lambda x: np.busday_count(x, td))
  dataD = dataD.sort_values("Pending Days from today", ascending=False)
  
  #reading non drafts
  data = pd.read_csv('data.csv',  encoding = 'cp1252', sep = ',', low_memory= False)
  data = data.loc[data['TERRITORYNAME'] == ter]
  data = data.loc[data['LAYOUT_NM'] != 'Draft Complaints']
  data['DATE_OF_REGISTRATION'].fillna(data['CREATEDON'], inplace=True)
  try:
    data = data.drop(data[data['COMPLAINTCLOSEDON'] == '111'].index)
  except:
    None
  try:
    data['CREATEDON'] = pd.to_datetime(data['CREATEDON'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    data['CREATEDON'] = pd.to_datetime(data['CREATEDON'].str.split().str[0], format='%d-%m-%y').dt.date
  try:
    data['DATE_OF_REGISTRATION'] = pd.to_datetime(data['DATE_OF_REGISTRATION'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    data['DATE_OF_REGISTRATION'] = pd.to_datetime(data['DATE_OF_REGISTRATION'].str.split().str[0], format='%d-%m-%y').dt.date
  try: 
    data['LASTMODIFIEDON'] = pd.to_datetime(data['LASTMODIFIEDON'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    data['LASTMODIFIEDON'] = pd.to_datetime(data['LASTMODIFIEDON'].str.split().str[0], format='%d-%m-%y').dt.date
  try:
    data['COMPLAINTCLOSEDON'] = pd.to_datetime(data['COMPLAINTCLOSEDON'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    data['COMPLAINTCLOSEDON'] = pd.to_datetime(data['COMPLAINTCLOSEDON'].str.split().str[0], format='%d-%m-%y').dt.date
  del data['CREATEDON']
  data = data.rename(columns = {'DATE_OF_REGISTRATION':'CREATEDON'})
  
  global data_no
  data_NO = pd.read_csv('NOPNO.csv',  encoding = 'cp1252', sep = ',', low_memory= False)
  
  data_NO = data_NO.rename(columns = {'COMPLAINNUMBER' : 'COMPLAINT_REF_NO'})
  global data_merge
  data_merge = pd.merge(data, data_NO, on = 'COMPLAINT_REF_NO', how = 'inner')
  try:
    data_merge = data_merge.drop(data_merge[data_merge['COMPLAINTCLOSEDON_x'] == '111'].index)
  except:
    None
  try:
    data_merge = data_merge.drop(data_merge[data_merge['COMPLAINTCLOSEDON_x'] == '19'].index)
  except:
    None
  data_merge.loc[data_merge['STATUSCODE_x'] == 'Assign to other Regulatory bodies', 'COMPLAINTCLOSEDON_x'] = data_merge['LASTMODIFIEDON_x']
  data_merge.loc[data_merge['STATUSCODE_x'] == 'Assign to other Regulatory bodies', 'STATUSCODE_x'] = 'Complaint Closed'
  try:
    data_merge.loc[data_merge['STATUSCODE_x'] == 'Assign to other RBI Department', 'COMPLAINTCLOSEDON_x'] = data_merge['LASTMODIFIEDON_x']
    data_merge.loc[data_merge['STATUSCODE_x'] == 'Assign to other RBI Department', 'STATUSCODE_x'] = 'Complaint Closed'
  except:
    None
  try:
    data_merge.loc[data_merge['STATUSCODE_x'] == 'Appeal Closed', 'STATUSCODE_x'] = 'Complaint Closed'
  except:
    None
  try:
    data_merge.loc[data_merge['STATUSCODE_x'] == 'New Appeal', 'STATUSCODE_x'] = 'Complaint Closed'
  except:
    None
  try:
    data_merge.loc[data_merge['STATUSCODE_x'] == 'Sent to AA Reviewer', 'STATUSCODE_x'] = 'Complaint Closed'
  except:
    None
  try:
    data_merge.loc[data_merge['STATUSCODE_x'] == 'Sent to AA', 'STATUSCODE_x'] = 'Complaint Closed'
  except:
    None
  try:
    data_merge.loc[data_merge['STATUSCODE_x'] == 'Sent back to AA DO', 'STATUSCODE_x'] = 'Complaint Closed'
  except:
    None
  try:
    data_merge.loc[data_merge['STATUSCODE_x'] == 'Sent to AA', 'STATUSCODE_x'] = 'Complaint Closed'
  except:
    None
  try:
    data_merge.loc[data_merge['STATUSCODE_x'] == 'Sent to Other Regulatory Bodies', 'STATUSCODE_x'] = 'Complaint Closed'
  except:
    None
  # try:
  #   data_merge['CREATEDON_x'] = pd.to_datetime(data_merge['CREATEDON_x'].str.split().str[0], format='%d-%m-%Y').dt.date
  # except:
  #   data_merge['CREATEDON_x'] = pd.to_datetime(data_merge['CREATEDON_x'].str.split().str[0], format='%d-%m-%y').dt.date
  # try: 
  #   data_merge['LASTMODIFIEDON_x'] = pd.to_datetime(data_merge['LASTMODIFIEDON_x'].str.split().str[0], format='%d-%m-%Y').dt.date
  # except:
  #   data_merge['LASTMODIFIEDON_x'] = pd.to_datetime(data_merge['LASTMODIFIEDON_x'].str.split().str[0], format='%d-%m-%y').dt.date
  # try:
  #   data_merge['COMPLAINTCLOSEDON_x'] = pd.to_datetime(data_merge['COMPLAINTCLOSEDON_x'].str.split().str[0], format='%d-%m-%Y').dt.date
  # except:
  #   data_merge['COMPLAINTCLOSEDON_x'] = pd.to_datetime(data_merge['COMPLAINTCLOSEDON_x'].str.split().str[0], format='%d-%m-%y').dt.date

  try:  
    data_merge['CREATEDON_y'] = pd.to_datetime(data_merge['CREATEDON_y'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    data_merge['CREATEDON_y'] = pd.to_datetime(data_merge['CREATEDON_y'].str.split().str[0], format='%d-%m-%y').dt.date
  try:
    data_merge['LASTMODIFIEDON_y'] = pd.to_datetime(data_merge['LASTMODIFIEDON_y'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    data_merge['LASTMODIFIEDON_y'] = pd.to_datetime(data_merge['LASTMODIFIEDON_y'].str.split().str[0], format='%d-%m-%y').dt.date
  try:
    data_merge['COMPLAINTCLOSEDON_y'] = pd.to_datetime(data_merge['COMPLAINTCLOSEDON_y'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    data_merge['COMPLAINTCLOSEDON_y'] = pd.to_datetime(data_merge['COMPLAINTCLOSEDON_y'].str.split().str[0], format='%d-%m-%y').dt.date
  try:
    data_merge['STATUSCODECHANGEDON_NO'] = pd.to_datetime(data_merge['STATUSCODECHANGEDON_NO'].str.split().str[0], format='%d-%m-%Y').dt.date
  except:
    data_merge['STATUSCODECHANGEDON_NO'] = pd.to_datetime(data_merge['STATUSCODECHANGEDON_NO'].str.split().str[0], format='%d-%m-%y').dt.date
#reading file for VGS functions
  data_V = data_merge.loc[data_merge['STATUSCODE_x'] != 'Complaint Closed']
  data_V = data_V.filter(['BANK_NAME',
                          'COMPLAINT_REF_NO',
                          'COMPLAINTNAME',
                          'BANKNAME',
                          'CURRENTOWNER',
                          'CASEID_y',
                          'STATUSCODE_NO',
                          'STATUSCODECHANGEDON_NO',
                          'STATUSCODE_x',
                          'ASSIGNEDTONAME',
                          'AGEING_DAYS'])
  
  data_V = data_V.rename(columns = {'BANK_NAME':'Issuer Bank',
                        'COMPLAINT_REF_NO':'Complaint No.',
                        'COMPLAINTNAME':'Complainants Name',
                        'BANKNAME':'Acquirer Bank',
                        'CURRENTOWNER':'NO Record Owner',
                        'CASEID_y':'Case ID(NO)',
                        'STATUSCODE_NO':'NO Record Status',
                        'STATUSCODECHANGEDON_NO':'NO Record Last changed on',
                        'STATUSCODE_x':'RBIO Status Code',
                        'ASSIGNEDTONAME':'RBIO DO Name',
                        'AGEING_DAYS':'Age in Days'})
#sanitizing the data with proper dates
  # try:
  #   data_V['NO Record Last changed on'] = pd.to_datetime(data_V['NO Record Last changed on'].str.split().str[0], format='%d-%m-%Y').dt.date
  # except:
  #   data_V['NO Record Last changed on'] = pd.to_datetime(data_V['NO Record Last changed on'].str.split().str[0], format='%d-%m-%y').dt.date
  data_V = data_V.sort_values("Issuer Bank", ascending=True)
  try:
    data_V.loc[( (data_V["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_V['RBIO DO Name'] == 'Prachi Goel') ) , ["Issuer Bank"] ] = "SBI CARDS"
    data_V.loc[( (data_V["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_V['RBIO DO Name'] == 'Vandana Malhotra') ) , ["Issuer Bank"] ] = "SBI CHD CIRCLE"
    data_V.loc[( (data_V["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_V['RBIO DO Name'] == 'Meenu Arora') ) , ["Issuer Bank"] ] = "SBI DEL CIRCLE"
    data_V.loc[( (data_V["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_V['NO Record Owner'] == 'SBI CARD NO PANKAJ NARULA') ) , ["Acquirer Bank"] ] = "SBI CARDS"
  except:
    None
  try:
    data_V.loc[( (data_V["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_V['NO Record Owner'] == 'SBI NO NEW DEL MA') ) , ["Acquirer Bank"] ] = "SBI DEL CIRCLE"
    data_V.loc[( (data_V["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_V['NO Record Owner'] == 'SBI NO CHD VM') ) , ["Acquirer Bank"] ] = "SBI CHD CIRCLE"
  except:
    None
#sanitizing the data with proper dates for both data and NOPNO
  try:
    data_V = data_V.drop(data_V[data_V['COMPLAINTCLOSEDON'] == '111'].index)
  except:
    None
  
  
  data.loc[data['STATUSCODE'] == 'Assign to other Regulatory bodies', 'COMPLAINTCLOSEDON'] = data['LASTMODIFIEDON']
  data.loc[data['STATUSCODE'] == 'Assign to other Regulatory bodies', 'STATUSCODE'] =  'Complaint Closed'
  try:
    data.loc[data['STATUSCODE'] == 'Assign to other RBI Department', 'COMPLAINTCLOSEDON'] = data['LASTMODIFIEDON']
    data.loc[data['STATUSCODE'] == 'Assign to other RBI Department', 'STATUSCODE'] =  'Complaint Closed'
  except:
    None
  try:
    data.loc[data['STATUSCODE'] == 'Appeal Closed', 'STATUSCODE'] = 'Complaint Closed'
  except:
    None
  try:
    data.loc[data['STATUSCODE'] == 'New Appeal', 'STATUSCODE'] = 'Complaint Closed'
  except:
    None
  try:
    data.loc[data['STATUSCODE'] == 'Sent to AA Reviewer', 'STATUSCODE'] = 'Complaint Closed'
  except:
    None
  try:
    data.loc[data['STATUSCODE'] == 'Sent to AA', 'STATUSCODE'] = 'Complaint Closed'
  except:
    None
  try:
    data.loc[data['STATUSCODE'] == 'Sent back to AA DO', 'STATUSCODE'] = 'Complaint Closed'
  except:
    None
  try:
    data.loc[data['STATUSCODE'] == 'Sent to AA', 'STATUSCODE'] = 'Complaint Closed'
  except:
    None
  try:
    data.loc[data['STATUSCODE'] == 'Sent to Other Regulatory Bodies', 'STATUSCODE'] = 'Complaint Closed'
  except:
    None
  
  
  print("\n\n Data has been read into the program successfully!!")

def PAO():
  global pao_date
  pao_date = entryS.get()
  pao_date = pd.to_datetime(pao_date, format = '%d-%m-%Y').date()
  global data
  global data_PAO
  global data_M_PAO
  global data_merge
  data_PAO = data.loc[data['STATUSCODE'] != 'Complaint Closed']
  data_PAO = data_PAO.loc[data_PAO['CREATEDON'] <= pao_date]
  data_M_PAO = data_PAO.loc[data_PAO['STATUSCODE'] != 'New Complaint']
  data_M_PAO = data_M_PAO.loc[data_M_PAO['STATUSCODE'] != 'Sent to other office']
  data_M_PAO = data_M_PAO.loc[data_M_PAO['STATUSCODE'] != 'Sent to Secretary BO']
  data_M_PAO = data_M_PAO.loc[data_M_PAO['STATUSCODE'] != 'Sent to BO']
  data_M_PAO = data_M_PAO.loc[data_M_PAO['STATUSCODE'] != 'BO Decision']
  global data_M
  data_M = data_merge.loc[data_merge['STATUSCODE_x'] != 'Complaint Closed']
  data_M = data_M.loc[data_M['CREATEDON_x'] <= pao_date]
  data_M = data_M.loc[data_M['STATUSCODE_x'] != 'New Complaint']
  data_M = data_M.loc[data_M['STATUSCODE_x'] != 'Sent to other office']
  data_M = data_M.loc[data_M['STATUSCODE_x'] != 'Sent to Secretary BO']
  data_M = data_M.loc[data_M['STATUSCODE_x'] != 'Sent to BO']
  data_M = data_M.loc[data_M['STATUSCODE_x'] != 'BO Decision']
  data_M = data_M.filter(['BANK_NAME',
                          'COMPLAINT_REF_NO',
                          'COMPLAINTNAME',
                          'BANKNAME',
                          'CURRENTOWNER',
                          'CASEID_y',
                          'STATUSCODE_NO',
                          'STATUSCODECHANGEDON_NO',
                          'STATUSCODE_x',
                          'ASSIGNEDTONAME',
                          'AGEING_DAYS'])
  
  data_M = data_M.rename(columns = {'BANK_NAME':'Issuer Bank',
                        'COMPLAINT_REF_NO':'Complaint No.',
                        'COMPLAINTNAME':'Complainants Name',
                        'BANKNAME':'Acquirer Bank',
                        'CURRENTOWNER':'NO Record Owner',
                        'CASEID_y':'Case ID(NO)',
                        'STATUSCODE_NO':'NO Record Status',
                        'STATUSCODECHANGEDON_NO':'NO Record Last changed on',
                        'STATUSCODE_x':'RBIO Status Code',
                        'ASSIGNEDTONAME':'RBIO DO Name',
                        'AGEING_DAYS':'Age in Days'})
  data_M = data_M.sort_values("Age in Days", ascending=False)
  try:
    data_M.loc[( (data_M["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_M['RBIO DO Name'] == 'Prachi Goel') ) , ["Issuer Bank"] ] = "SBI CARDS"
    data_M.loc[( (data_M["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_M['RBIO DO Name'] == 'Vandana Malhotra') ) , ["Issuer Bank"] ] = "SBI CHD CIRCLE"
    data_M.loc[( (data_M["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_M['RBIO DO Name'] == 'Meenu Arora') ) , ["Issuer Bank"] ] = "SBI DEL CIRCLE"
    data_M.loc[( (data_M["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_M['NO Record Owner'] == 'SBI CARD NO PANKAJ NARULA') ) , ["Acquirer Bank"] ] = "SBI CARDS"
  except:
    None
  try:
    data_M.loc[( (data_M["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_M['NO Record Owner'] == 'SBI NO NEW DEL MA') ) , ["Acquirer Bank"] ] = "SBI DEL CIRCLE"
    data_M.loc[( (data_M["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_M['NO Record Owner'] == 'SBI NO CHD VM') ) , ["Acquirer Bank"] ] = "SBI CHD CIRCLE"
  except:
    None


def counter():
  global data_PAO
  count_all['text'] =  str(len(data_PAO))
  count_all_M['text'] =  str(len(data_M_PAO))
  
def to_xl(df, name):
  
  writer = pd.ExcelWriter("{}.xlsx".format(name), engine = 'xlsxwriter')
  df.to_excel(writer, sheet_name="TOTAL")

  workbook = writer.book
  worksheet = writer.sheets["TOTAL"]
  
  format1 = workbook.add_format({'num_format':'0'})
  format2 = workbook.add_format({'bold': True, 'font_color': 'red'})
  format4 = workbook.add_format({'bold': True})
  
  worksheet.set_column('A:A', 30, None)
  worksheet.set_column('B:B', 16, format1)
  worksheet.set_column('C:C', 17, None)
  worksheet.set_column('D:D', 28, None)
  worksheet.set_column('E:E', 25, format2)
  worksheet.set_column('F:F', 10, None)
  worksheet.set_column('G:G', 14, format2)
  worksheet.set_column('H:H', 14, format4)
  worksheet.set_column('I:I', 14, None)
  worksheet.set_column('J:J', 14, None)
  worksheet.set_column('K:K', 10, format2)
  
  writer.save()
  writer.close()

def to_xl_all():
  global data_PAO
  if not data_PAO.empty:
    temp = data_PAO.filter(['BANK_NAME',
                            'COMPLAINT_REF_NO',
                            'STATUSCODE',
                            'COMPLAINTNAME',
                            'ASSIGNEDTONAME',
                            'AGEING_DAYS'])
    temp = temp.set_index('BANK_NAME')
    temp = temp.sort_values("AGEING_DAYS", ascending=False)
    global pao_date
    to_xl(temp, 'All-Pending-as-on-{}'.format(pao_date))
  
def to_xl_M():
  global data_M_PAO
  if not data_M_PAO.empty:
    temp = data_M_PAO.filter(['BANK_NAME',
                            'COMPLAINT_REF_NO',
                            'STATUSCODE',
                            'COMPLAINTNAME',
                            'ASSIGNEDTONAME',
                            'AGEING_DAYS'])
    temp = temp.set_index('BANK_NAME')
    temp = temp.sort_values("AGEING_DAYS", ascending=False)
    global pao_date
    to_xl(temp, 'Maintainable-Pending-as-on-{}'.format(pao_date))

def to_xl_DOW():
  global data_M_PAO
  if not data_M_PAO.empty:
    temp = data_M_PAO.filter(['COMPLAINT_REF_NO',
                            'BANK_NAME',
                            'STATUSCODE',
                            'COMPLAINTNAME',
                            'ASSIGNEDTONAME',
                            'AGEING_DAYS'])
    # temp = temp.set_index('BANK_NAME')
    temp = temp.sort_values("AGEING_DAYS", ascending=False)
    global pao_date
    # 
    dos = []
    checked = {}
    for index in temp['ASSIGNEDTONAME']:
      if index not in checked:
        checked[index] = True
        dos.append(index)
    dos_new = [re.sub(r'[":\-();*!@#$%^&=`~+,.<>?/\n"]', "_",x) for x in dos]
    temp = temp.set_index('ASSIGNEDTONAME')
    for i in range(len(dos_new)):
      data_do_wise = temp.loc[dos[i]]
      to_xl(data_do_wise, '{}-on-{}'.format(dos_new[i],pao_date))
      
def to_xl_NOW():
  global data_M
  global pao_date
  temp = data_M
  temp = temp.loc[temp['NO Record Status'] != 'Sent to RBI']
  temp = temp.loc[temp['NO Record Status'] != 'New']
  temp = temp.set_index('Issuer Bank')
  banks = []
  checked = {}
  for index in temp.index.to_series():
    if index not in checked:
      checked[index] = True
      banks.append(index)
  for i in range(len(banks)):
    if str(banks[i]) == 'nan':
      del banks[i]
  banks_new = [re.sub(r'[":\-();*!@#$%^&=`~+,.<>?/\n"]', "_",x) for x in banks]
  for i in range(len(banks_new)):
    data_bank_wise = temp.loc[banks[i]]
    to_xl(data_bank_wise, 'I-{}-on-{}'.format(banks_new[i],pao_date)) 


def to_xl_ANOW():
  global data_M
  global pao_date
  temp = data_M
  temp = temp.loc[temp['NO Record Status'] != 'Sent to RBI']
  temp = temp.loc[temp['NO Record Status'] != 'New']
  temp = temp[['Complaint No.',
               'Issuer Bank',
               'Complainants Name',
               'Acquirer Bank',
               'NO Record Owner',
               'Case ID(NO)',
               'NO Record Status',
               'NO Record Last changed on',
               'RBIO Status Code',
               'RBIO DO Name',
               'Age in Days']]
  temp = temp.set_index('Acquirer Bank')
  banks = []
  checked = {}
  for index in temp.index.to_series():
    if index not in checked:
      checked[index] = True
      banks.append(index)
  for i in range(len(banks)):
    if str(banks[i]) == 'nan':
      del banks[i]
  banks_new = [re.sub(r'[":\-();*!@#$%^&=`~+,.<>?/\n"]', "_",x) for x in banks]
  for i in range(len(banks_new)):
    data_bank_wise = temp.loc[banks[i]]
    to_xl(data_bank_wise, 'A-{}-on-{}'.format(banks_new[i],pao_date)) 
  
def STR():
  global fr, to
  global data_R
  global data_merge
  data_R = data_merge.loc[data_merge['STATUSCODE_x'] != 'Complaint Closed']
  data_R = data_R.loc[data_R['STATUSCODE_x'] != 'New Complaint']
  data_R = data_R.loc[data_R['STATUSCODE_x'] != 'Sent to other office']
  data_R = data_R.loc[data_R['STATUSCODE_x'] != 'Sent to Secretary BO']
  data_R = data_R.loc[data_R['STATUSCODE_x'] != 'Sent to BO']
  data_R = data_R.loc[data_R['STATUSCODE_x'] != 'BO Decision']
  data_R = data_R.filter(['BANK_NAME',
                          'COMPLAINT_REF_NO',
                          'COMPLAINTNAME',
                          'BANKNAME',
                          'CURRENTOWNER',
                          'CASEID_y',
                          'STATUSCODE_NO',
                          'STATUSCODECHANGEDON_NO',
                          'STATUSCODE_x',
                          'ASSIGNEDTONAME',
                          'AGEING_DAYS'])
  
  data_R = data_R.rename(columns = {'BANK_NAME':'Issuer Bank',
                        'COMPLAINT_REF_NO':'Complaint No.',
                        'COMPLAINTNAME':'Complainants Name',
                        'BANKNAME':'Acquirer Bank',
                        'CURRENTOWNER':'NO Record Owner',
                        'CASEID_y':'Case ID(NO)',
                        'STATUSCODE_NO':'NO Record Status',
                        'STATUSCODECHANGEDON_NO':'NO Record Last changed on',
                        'STATUSCODE_x':'RBIO Status Code',
                        'ASSIGNEDTONAME':'RBIO DO Name',
                        'AGEING_DAYS':'Age in Days'})
  data_R = data_R.sort_values("Age in Days", ascending=False)
  try:
    data_R.loc[( (data_R["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_R['RBIO DO Name'] == 'Prachi Goel') ) , ["Issuer Bank"] ] = "SBI CARDS"
    data_R.loc[( (data_R["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_R['RBIO DO Name'] == 'Vandana Malhotra') ) , ["Issuer Bank"] ] = "SBI CHD CIRCLE"
    data_R.loc[( (data_R["Issuer Bank"] == 'STATE BANK OF INDIA') &  (data_R['RBIO DO Name'] == 'Meenu Arora') ) , ["Issuer Bank"] ] = "SBI DEL CIRCLE"
    data_R.loc[( (data_R["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_R['NO Record Owner'] == 'SBI CARD NO PANKAJ NARULA') ) , ["Acquirer Bank"] ] = "SBI CARDS"
  except:
    None
  try:
    data_R.loc[( (data_R["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_R['NO Record Owner'] == 'SBI NO NEW DEL MA') ) , ["Acquirer Bank"] ] = "SBI DEL CIRCLE"
    data_R.loc[( (data_R["Acquirer Bank"] == 'STATE BANK OF INDIA') &  (data_R['NO Record Owner'] == 'SBI NO CHD VM') ) , ["Acquirer Bank"] ] = "SBI CHD CIRCLE"
  except:
    None

  fr = entryF.get()
  to = entryTO.get()
  fr = pd.to_datetime(fr, format = '%d-%m-%Y').date()
  to = pd.to_datetime(to, format = '%d-%m-%Y').date()
  temp = data_R
  cond1 = temp['NO Record Status'] == 'Sent to RBI'
  cond2 =  temp['NO Record Status'] == 'Advisory Complied'
  temp = temp[cond1 | cond2]
  temp = temp.loc[temp['NO Record Last changed on'] >= fr]
  temp = temp.loc[temp['NO Record Last changed on'] <= to]
  temp = temp.set_index('Issuer Bank')
  to_xl(temp, 'New-Replies-{}-to-{}'.format(fr,to)) 
  
  
def VGS_DO():
  global data_merge
  global data_V
  periodIR = int(entryIR.get())
  period101 = int(entry101.get())
  periodAI = int(entryAI.get())
  periodNR = int(entryNR.get())
  #temp is all new complaints / sent to other office ones
  temp = data_merge.loc[data_merge['STATUSCODE_x'] != 'Complaint Closed']
  cond1 = temp['STATUSCODE_x'] == 'New Complaint'
  cond2 =  temp['STATUSCODE_x'] == 'Sent to other office'
  temp = temp[cond1 | cond2 ]
  temp = temp.filter(['BANK_NAME',
                      'COMPLAINT_REF_NO',
                      'COMPLAINTNAME',
                      'BANKNAME',
                      'STATUSCODE_x',
                      'ASSIGNEDTONAME',
                      'CREATEDON_x'])
  
  temp = temp.rename(columns = {'BANK_NAME':'Issuer Bank',
                        'COMPLAINT_REF_NO':'Complaint No.',
                        'COMPLAINTNAME':'Complainants Name',
                        'BANKNAME':'Acquirer Bank',
                        'STATUSCODE_x':'RBIO Status Code',
                        'ASSIGNEDTONAME':'RBIO DO Name',
                        'CREATEDON_x':'Origin Date'})
  td = _datetime.date.today()
  temp['Pending Weekdays from today'] =  temp['Origin Date'].apply(lambda x: np.busday_count(x, td))
  temp = temp.set_index('Issuer Bank')
  temp.drop_duplicates(subset='Complaint No.', keep = 'first', inplace = True)
  temp = temp.loc[temp['Pending Weekdays from today'] > periodNR]
  #temp2 is all maintainable complaints pending with DOs only along with pending age of NO Record as on today column added 
  temp2 = data_V
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'New Complaint']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to other office']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to Secretary BO']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to BO']
  temp2 = temp2.set_index('Issuer Bank')
  td = _datetime.date.today()
  temp2['Pending With Bank from today'] =  temp2['NO Record Last changed on'].apply(lambda x: np.busday_count(x, td))
  #temp3 has information required cases past the given period 
  temp3 = temp2
  temp3 = temp3.loc[temp3['NO Record Status'] == 'Information Required']
  temp3 = temp3.loc[temp3['Pending With Bank from today'] > periodIR]
  #temp4 has 10-1 notice issued cases past the given period
  temp4 = temp2
  temp4 = temp4.loc[temp4['NO Record Status'] == '10-1 Notice Issued']
  temp4 = temp4.loc[temp4['Pending With Bank from today'] > period101]
  #temp5 has Issue Advisories issued cases past the given period
  temp5 = temp2
  temp5 = temp5.loc[temp5['NO Record Status'] == ' Advisory Issued']
  temp5 = temp5.loc[temp5['Pending With Bank from today'] > periodAI]
  #temp6 has all which are in sent to rbi, advisory complied stage
  temp6 = temp2
  cond1 = temp6['NO Record Status'] == 'Sent to RBI'
  cond2 =  temp6['NO Record Status'] == 'Advisory Complied'
  temp6 = temp6[cond1 | cond2]
  temp6 = temp6.sort_values("Complaint No.", ascending=True)
  #calling to_xl_VGS_DO with all dataframes and DO Name
  dos = []
  checked = {}
  for index in data_V['RBIO DO Name']:
    if index not in checked:
      checked[index] = True
      dos.append(index)
  for do in dos:
    cnt=0
    templist = [temp,temp3,temp4,temp5]
    for dfs in templist:
      t = len(dfs.loc[dfs['RBIO DO Name'] == do])
      cnt = cnt + t
    if cnt>0:
      to_xl_VGS_DO(temp,temp3,temp4,temp5,temp6, do)

def to_xl_VGS_DO(temp,temp3,temp4,temp5,temp6, name):

  writer = pd.ExcelWriter("{}-MultiCriteria.xlsx".format(name), engine = 'xlsxwriter')
  temp = temp[temp['RBIO DO Name']==name]
  temp3 = temp3[temp3['RBIO DO Name']==name]
  temp4 = temp4[temp4['RBIO DO Name']==name]
  temp5 = temp5[temp5['RBIO DO Name']==name]
  temp6 = temp6[temp6['RBIO DO Name']==name]
  temp.to_excel(writer, sheet_name="New Recpts Past Period")
  temp3.to_excel(writer, sheet_name="Info Reqd Past Period")
  temp4.to_excel(writer, sheet_name="10-1 Past Period")
  temp5.to_excel(writer, sheet_name="Issue Advisory Past Period")
  temp6.to_excel(writer, sheet_name="Bank Has Replied")

  workbook = writer.book
  for sheets in writer.sheets:  
    worksheet = writer.sheets[sheets]
    
    format1 = workbook.add_format({'num_format':'0'})
    format2 = workbook.add_format({'bold': True, 'font_color': 'red'})
    format4 = workbook.add_format({'bold': True})
    
    worksheet.set_column('A:A', 35, None)
    worksheet.set_column('B:B', 35, format1)
    worksheet.set_column('C:C', 35, None)
    worksheet.set_column('D:D', 35, None)
    worksheet.set_column('E:E', 35, format2)
    worksheet.set_column('F:F', 35, None)
    worksheet.set_column('G:G', 35, format2)
    worksheet.set_column('H:H', 35, format4)
    worksheet.set_column('I:I', 35, None)
    worksheet.set_column('J:J', 35, None)
    worksheet.set_column('K:K', 35, None)
    worksheet.set_column('L:L', 35, format2)
  
  writer.save()
  writer.close()


def VGS_DO_ALL():
  global data_merge
  global data_V
  periodIR = int(entryIR.get())
  period101 = int(entry101.get())
  periodAI = int(entryAI.get())
  periodNR = int(entryNR.get())
  #temp is all new complaints / sent to other office ones
  temp = data_merge.loc[data_merge['STATUSCODE_x'] != 'Complaint Closed']
  cond1 = temp['STATUSCODE_x'] == 'New Complaint'
  cond2 =  temp['STATUSCODE_x'] == 'Sent to other office'
  temp = temp[cond1 | cond2 ]
  temp = temp.filter(['BANK_NAME',
                      'COMPLAINT_REF_NO',
                      'COMPLAINTNAME',
                      'BANKNAME',
                      'STATUSCODE_x',
                      'ASSIGNEDTONAME',
                      'CREATEDON_x'])
  
  temp = temp.rename(columns = {'BANK_NAME':'Issuer Bank',
                        'COMPLAINT_REF_NO':'Complaint No.',
                        'COMPLAINTNAME':'Complainants Name',
                        'BANKNAME':'Acquirer Bank',
                        'STATUSCODE_x':'RBIO Status Code',
                        'ASSIGNEDTONAME':'RBIO DO Name',
                        'CREATEDON_x':'Origin Date'})
  td = _datetime.date.today()
  temp['Pending Weekdays from today'] =  temp['Origin Date'].apply(lambda x: np.busday_count(x, td))
  temp = temp.set_index('Issuer Bank')
  temp.drop_duplicates(subset='Complaint No.', keep = 'first', inplace = True)
  temp = temp.loc[temp['Pending Weekdays from today'] > periodNR]
  #temp0 is ALL pending with pendency as on today days added as a column
  temp0 = data_V
  temp0 = temp0.set_index('Issuer Bank')
  temp0 = temp0.drop(['Acquirer Bank','NO Record Owner', 'Case ID(NO)', 'NO Record Status', 'NO Record Last changed on'], axis=1)
  temp0.drop_duplicates(subset='Complaint No.', keep = 'first', inplace = True)
  #temp2 is all maintainable complaints pending with DOs only along with pending age of NO Record as on today column added 
  temp2 = data_V
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'New Complaint']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to other office']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to Secretary BO']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to BO']
  temp2 = temp2.set_index('Issuer Bank')
  td = _datetime.date.today()
  temp2['Pending With Bank from today'] =  temp2['NO Record Last changed on'].apply(lambda x: np.busday_count(x, td))
  #temp3 has information required cases past the given period 
  temp3 = temp2
  temp3 = temp3.loc[temp3['NO Record Status'] == 'Information Required']
  temp3 = temp3.loc[temp3['Pending With Bank from today'] > periodIR]
  #temp4 has 10-1 notice issued cases past the given period
  temp4 = temp2
  temp4 = temp4.loc[temp4['NO Record Status'] == '10-1 Notice Issued']
  temp4 = temp4.loc[temp4['Pending With Bank from today'] > period101]
  #temp5 has Issue Advisories issued cases past the given period
  temp5 = temp2
  temp5 = temp5.loc[temp5['NO Record Status'] == ' Advisory Issued']
  temp5 = temp5.loc[temp5['Pending With Bank from today'] > periodAI]
  #temp6 has all which are in sent to rbi, advisory complied stage
  temp6 = temp2
  cond1 = temp6['NO Record Status'] == 'Sent to RBI'
  cond2 =  temp6['NO Record Status'] == 'Advisory Complied'
  temp6 = temp6[cond1 | cond2]
  temp6 = temp6.sort_values("Complaint No.", ascending=True)
  #calling to_xl_VGS_DO with all dataframes and DO Name
  to_xl_VGS_DO_ALL(temp,temp0,temp3,temp4,temp5,temp6, 'FFC-Total-Pending')

def to_xl_VGS_DO_ALL(temp,temp0,temp3,temp4,temp5,temp6, name):
  
  writer = pd.ExcelWriter("{}-MultiCriteria.xlsx".format(name), engine = 'xlsxwriter')
  
  temp0.to_excel(writer, sheet_name="All Pending")
  temp.to_excel(writer, sheet_name="New Recpts Past Period")
  temp3.to_excel(writer, sheet_name="Info Reqd Past Period")
  temp4.to_excel(writer, sheet_name="10-1 Past Period")
  temp5.to_excel(writer, sheet_name="Issue Advisory Past Period")
  temp6.to_excel(writer, sheet_name="All Where Bank Has Replied")

  workbook = writer.book
  for sheets in writer.sheets:  
    worksheet = writer.sheets[sheets]
    
    format1 = workbook.add_format({'num_format':'0'})
    format2 = workbook.add_format({'bold': True, 'font_color': 'red'})
    format4 = workbook.add_format({'bold': True})
    
    worksheet.set_column('A:A', 35, None)
    worksheet.set_column('B:B', 35, format1)
    worksheet.set_column('C:C', 35, None)
    worksheet.set_column('D:D', 35, None)
    worksheet.set_column('E:E', 35, format2)
    worksheet.set_column('F:F', 35, None)
    worksheet.set_column('G:G', 35, format2)
    worksheet.set_column('H:H', 35, format4)
    worksheet.set_column('I:I', 35, None)
    worksheet.set_column('J:J', 35, None)
    worksheet.set_column('K:K', 35, None)
    worksheet.set_column('L:L', 35, format2)
  
  writer.save()
  writer.close()

def VGS_Drafts():
  territory()
  global ter
  global dataD
  #temp1 has all which are directly recd from mail
  temp1 = dataD
  temp1 = temp1[temp1['Draft Status'] == 'Draft Complaint']
  #temp2 has all which are recd from other office
  temp2 = dataD
  temp2 = temp2[temp2['Draft Status'] == 'Sent to other Ombudsman office']
  #temp3 has all which are pending for approval
  temp3 = dataD
  temp3 = temp3[temp3['Draft Status'] == 'Sent for approval']
  #calling to_xl_VGS_DO with all dataframes and IO Names
  dos = []
  checked = {}
  for index in dataD['Current Owner']:
    if index not in checked:
      checked[index] = True
      dos.append(index)
  for do in dos:
    to_xl_VGS_Drafts(dataD, temp1, temp2, temp3, do)
  
def to_xl_VGS_Drafts(dataD, temp1, temp2, temp3, name):
  writer = pd.ExcelWriter("Total {}-MultiCriteria.xlsx".format(name), engine = 'xlsxwriter')
  
  dataD.to_excel(writer, sheet_name="Total Pending")
  temp1.to_excel(writer, sheet_name="Recd directly in this office")
  temp2.to_excel(writer, sheet_name="Recd From Other Offices")
  temp3.to_excel(writer, sheet_name="Pending for Approval- if RO")

  workbook = writer.book
  for sheets in writer.sheets:  
    worksheet = writer.sheets[sheets]
    
    format1 = workbook.add_format({'num_format':'0'})
    format2 = workbook.add_format({'bold': True, 'font_color': 'red'})
    format4 = workbook.add_format({'bold': True})
    
    worksheet.set_column('A:A', 35, format1)
    worksheet.set_column('B:B', 35, format2)
    worksheet.set_column('C:C', 35, None)
    worksheet.set_column('D:D', 35, None)
    worksheet.set_column('E:E', 35, None)
    worksheet.set_column('F:F', 35, format4)
    worksheet.set_column('G:G', 35, format2)
  
  writer.save()
  writer.close()
  
def VGS_Drafts_ALL():
  territory()
  global ter
  global dataD
  #temp1 has all which are directly recd from mail
  temp1 = dataD
  temp1 = temp1[temp1['Draft Status'] == 'Draft Complaint']
  #temp2 has all which are recd from other office
  temp2 = dataD
  temp2 = temp2[temp2['Draft Status'] == 'Sent to other Ombudsman office']
  #temp3 has all which are pending for approval
  temp3 = dataD
  temp3 = temp3[temp3['Draft Status'] == 'Sent for approval']
  to_xl_VGS_Drafts_ALL(dataD, temp1, temp2, temp3, 'Tot Drafts')
  
def to_xl_VGS_Drafts_ALL(dataD,temp1,temp2,temp3, name):
  
  writer = pd.ExcelWriter("{}-MultiCriteria.xlsx".format(name), engine = 'xlsxwriter')

  dataD.to_excel(writer, sheet_name="Total Pending")
  temp1.to_excel(writer, sheet_name="Recd directly in this office")
  temp2.to_excel(writer, sheet_name="Recd From Other Offices")
  temp3.to_excel(writer, sheet_name="Pending for Approval- if RO")

  workbook = writer.book
  for sheets in writer.sheets:  
    worksheet = writer.sheets[sheets]
    
    format1 = workbook.add_format({'num_format':'0'})
    format2 = workbook.add_format({'bold': True, 'font_color': 'red'})
    format4 = workbook.add_format({'bold': True})
    
    worksheet.set_column('A:A', 35, format1)
    worksheet.set_column('B:B', 35, format2)
    worksheet.set_column('C:C', 35, None)
    worksheet.set_column('D:D', 35, None)
    worksheet.set_column('E:E', 35, None)
    worksheet.set_column('F:F', 35, format4)
    worksheet.set_column('G:G', 35, format2)
  
  writer.save()
  writer.close()

def VGS_BW():
  global data_merge
  global data_V
  periodIR = int(entryIR.get())
  period101 = int(entry101.get())
  periodAI = int(entryAI.get())
  periodNR = int(entryNR.get())
  #temp is all new complaints / sent to other office ones
  temp = data_merge.loc[data_merge['STATUSCODE_x'] != 'Complaint Closed']
  cond1 = temp['STATUSCODE_x'] == 'New Complaint'
  cond2 =  temp['STATUSCODE_x'] == 'Sent to other office'
  temp = temp[cond1 | cond2 ]
  temp = temp.filter(['BANK_NAME',
                      'COMPLAINT_REF_NO',
                      'COMPLAINTNAME',
                      'BANKNAME',
                      'STATUSCODE_x',
                      'ASSIGNEDTONAME',
                      'CREATEDON_x'])
  
  temp = temp.rename(columns = {'BANK_NAME':'Issuer Bank',
                        'COMPLAINT_REF_NO':'Complaint No.',
                        'COMPLAINTNAME':'Complainants Name',
                        'BANKNAME':'Acquirer Bank',
                        'STATUSCODE_x':'RBIO Status Code',
                        'ASSIGNEDTONAME':'RBIO DO Name',
                        'CREATEDON_x':'Origin Date'})
  td = _datetime.date.today()
  temp['Pending Weekdays from today'] =  temp['Origin Date'].apply(lambda x: np.busday_count(x, td))
  temp = temp.set_index('Issuer Bank')
  temp.drop_duplicates(subset='Complaint No.', keep = 'first', inplace = True)
  temp = temp.loc[temp['Pending Weekdays from today'] > periodNR]
  #temp2 is all maintainable complaints pending with DOs only along with pending age of NO Record as on today column added 
  temp2 = data_V
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'New Complaint']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to other office']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to Secretary BO']
  temp2 = temp2.loc[temp2['RBIO Status Code'] != 'Sent to BO']
  temp2 = temp2.set_index('Issuer Bank')
  td = _datetime.date.today()
  temp2['Pending With Bank from today'] =  temp2['NO Record Last changed on'].apply(lambda x: np.busday_count(x, td))
  #temp3 has information required cases past the given period 
  temp3 = temp2
  temp3 = temp3.loc[temp3['NO Record Status'] == 'Information Required']
  temp3 = temp3.loc[temp3['Pending With Bank from today'] > periodIR]
  #temp4 has 10-1 notice issued cases past the given period
  temp4 = temp2
  temp4 = temp4.loc[temp4['NO Record Status'] == '10-1 Notice Issued']
  temp4 = temp4.loc[temp4['Pending With Bank from today'] > period101]
  #temp5 has Issue Advisories issued cases past the given period
  temp5 = temp2
  temp5 = temp5.loc[temp5['NO Record Status'] == ' Advisory Issued']
  temp5 = temp5.loc[temp5['Pending With Bank from today'] > periodAI]
  #temp6 has all which are in sent to rbi, advisory complied stage
  temp6 = temp2
  cond1 = temp6['NO Record Status'] == 'Sent to RBI'
  cond2 =  temp6['NO Record Status'] == 'Advisory Complied'
  temp6 = temp6[cond1 | cond2]
  temp6 = temp6.sort_values("Complaint No.", ascending=True)
  #calling to_xl_VGS_DO with all dataframes and DO Name
  nos = []
  checked = {}
  for index in data_V['Issuer Bank']:
    if index not in checked:
      checked[index] = True
      nos.append(index)
  for no in nos:
    cnt=0
    templist = [temp,temp3,temp4,temp5]
    for dfs in templist:
      t = len(dfs.loc[dfs.index == no])
      cnt = cnt + t
    if cnt>0:
      to_xl_VGS_BW(temp,temp3,temp4,temp5,temp6, no)

def to_xl_VGS_BW(temp,temp3,temp4,temp5,temp6, name):
  fname = re.sub(r'[":\-();*!@#$%^&=`~+,.<>?/\n"]', "_",name)
  writer = pd.ExcelWriter("{}-MultiCriteria.xlsx".format(fname), engine = 'xlsxwriter')
  temp = temp[temp.index==name]
  temp3 = temp3[temp3.index==name]
  temp4 = temp4[temp4.index==name]
  temp5 = temp5[temp5.index==name]
  temp6 = temp6[temp6.index==name]
  temp.to_excel(writer, sheet_name="New Recpts Past Period")
  temp3.to_excel(writer, sheet_name="Info Reqd Past Period")
  temp4.to_excel(writer, sheet_name="10-1 Past Period")
  temp5.to_excel(writer, sheet_name="Issue Advisory Past Period")
  temp6.to_excel(writer, sheet_name="This Bank Has Replied")

  workbook = writer.book
  for sheets in writer.sheets:  
    worksheet = writer.sheets[sheets]
    
    format1 = workbook.add_format({'num_format':'0'})
    format2 = workbook.add_format({'bold': True, 'font_color': 'red'})
    format4 = workbook.add_format({'bold': True})
    
    worksheet.set_column('A:A', 35, None)
    worksheet.set_column('B:B', 35, format1)
    worksheet.set_column('C:C', 35, None)
    worksheet.set_column('D:D', 35, None)
    worksheet.set_column('E:E', 35, format2)
    worksheet.set_column('F:F', 35, None)
    worksheet.set_column('G:G', 35, format2)
    worksheet.set_column('H:H', 35, format4)
    worksheet.set_column('I:I', 35, None)
    worksheet.set_column('J:J', 35, None)
    worksheet.set_column('K:K', 35, None)
    worksheet.set_column('L:L', 35, format2)
  
  writer.save()
  writer.close()

def graph():
  global data
  created_on = data.loc[:,['CREATEDON']]
  created_on = created_on.sort_values(by=['CREATEDON'])
  created_on = created_on['CREATEDON'].groupby(created_on['CREATEDON'].apply(lambda x: x.strftime('%Y-%U'))).count().to_frame()
  created_on = created_on.rename(columns = {'CREATEDON':'Count'})
  N=len(created_on)
  created_on.plot(kind='area',use_index=True, y = 'Count')
  plt.title('Weekly Complaints Receipt Count Since June 2019')
  plt.xticks(range(len(created_on.index)), created_on.index, fontsize=5)
  plt.gca().margins(x=0)
  plt.gcf().canvas.draw()
  tl = plt.gca().get_xticklabels()
  maxsize = max([t.get_window_extent().width for t in tl])
  m = 0.2 # inch margin
  s = maxsize/plt.gcf().dpi*N+2*m
  margin = m/plt.gcf().get_size_inches()[0]
  plt.gcf().subplots_adjust(left=margin, right=1.-margin)
  plt.gcf().set_size_inches(s, plt.gcf().get_size_inches()[1])
  plt.xlabel('YYYY-WW (Year-Week) [WW From 1-52]')
  plt.ylabel('Complaint Count')
  plt.show()
  
def graph_D():
  global data
  closed_on = data.loc[:,['COMPLAINTCLOSEDON']]
  closed_on = closed_on.sort_values(by=['COMPLAINTCLOSEDON'])
  closed_on = closed_on[pd.notnull(closed_on['COMPLAINTCLOSEDON'])]
  closed_on = closed_on['COMPLAINTCLOSEDON'].groupby(closed_on['COMPLAINTCLOSEDON'].apply(lambda x: x.strftime('%Y-%U'))).count().to_frame()
  closed_on = closed_on.rename(columns = {'COMPLAINTCLOSEDON':'Count'})
  N=len(closed_on)
  closed_on.plot(kind='area',color = 'green',use_index=True, y = 'Count')
  plt.title('Weekly Complaints Disposal Count Since June 2019')
  plt.xticks(range(len(closed_on.index)), closed_on.index, fontsize=5)
  plt.gca().margins(x=0)
  plt.gcf().canvas.draw()
  tl = plt.gca().get_xticklabels()
  maxsize = max([t.get_window_extent().width for t in tl])
  m = 0.2 # inch margin
  s = maxsize/plt.gcf().dpi*N+2*m
  margin = m/plt.gcf().get_size_inches()[0]
  plt.gcf().subplots_adjust(left=margin, right=1.-margin)
  plt.gcf().set_size_inches(s, plt.gcf().get_size_inches()[1])
  plt.xlabel('YYYY-WW (Year-Week) [WW From 1-52]')
  plt.ylabel('Complaint Count')
  plt.show()

def clicked():
  buttonUD.configure(text='Done. Load Again?', bg='purple')

window = tk.Tk()
window.title('CMS Data Ship')
window.geometry("1350x768")
window.configure(bg='black')
buttonG = tk.Button(window, text = 'Show Receipt Graph', bg='yellow', fg='black',command = graph, borderwidth = 6,font = "Helvetica 10 bold")
buttonG.grid(row=5,column=5)
buttonG.place(x=150, y=150)
buttonGD = tk.Button(window, text = 'Show Disposal Graph', bg='yellow', fg='black',command = graph_D, borderwidth = 6,font = "Helvetica 10 bold")
buttonGD.grid(row=5,column=5)
buttonGD.place(x=150, y=220)
buttonUD = tk.Button(window, text = 'Load/Refresh Data', bg='white', fg='red',command = lambda:[csv_reader(),clicked()], borderwidth = 6,font = "Helvetica 10 bold")
buttonUD.grid(row=5,column=5)
buttonUD.place(x=1080, y=50) 
buttonPAO = tk.Button(window, text = 'Set Date-3 Blue Button Reports-->', bg='yellow', fg='black',command = lambda:[PAO(),counter()], borderwidth = 6,font = "Helvetica 10 bold")
buttonPAO.grid(row=5,column=5)
buttonPAO.place(x=400, y=150)
buttonDOW = tk.Button(window, text = 'DO-WISE-Start Exploding', bg='blue', fg='white',command = to_xl_DOW ,borderwidth = 6,font = "Helvetica 10 bold")
buttonDOW.grid(row=5,column=5)
buttonDOW.place(x=450, y=280)
buttonNOW = tk.Button(window, text = 'DETAILED ISSUER BANK-WISE-Start Exploding', bg='blue', fg='white',command = to_xl_NOW ,borderwidth = 6,font = "Helvetica 10 bold")
buttonNOW.grid(row=5,column=5)
buttonNOW.place(x=120, y=350)
buttonANOW = tk.Button(window, text = 'DETAILED ACQUIRER BANK-WISE-Start Exploding', bg='blue', fg='white',command = to_xl_ANOW ,borderwidth = 6,font = "Helvetica 10 bold")
buttonANOW.grid(row=5,column=5)
buttonANOW.place(x=570, y=350)
buttonSTR = tk.Button(window, text = 'Report of Bank Replies Recd. between', bg='green', fg='white',command = STR ,borderwidth = 6,font = "Helvetica 10 bold")
buttonSTR.grid(row=5,column=5)
buttonSTR.place(x=220, y=430)
v2 = StringVar(window, value='dd-mm-yyyy')
entryF = tk.Entry(window, width = 13,  textvariable = v2)
entryF.grid(row=5,column=5)
entryF.place(x = 550, y = 438)
text4 = tk.Text(window, height = 1, width = 5, bg = 'black', fg = 'yellow', font = "Courier 10 italic bold")
text4.place(x=653,y=438)
text4.insert(tk.END, 'AND')
text4.configure(state='disabled')
v3 = StringVar(window, value='dd-mm-yyyy')
entryTO = tk.Entry(window, width = 13,  textvariable = v3)
entryTO.grid(row=5,column=5)
entryTO.place(x = 695,y = 438)
v4 = StringVar(window, value='0')
entryIR = tk.Entry(window, width = 5,  textvariable = v4)
entryIR.grid(row=5,column=5)
entryIR.place(x = 559, y = 538)
text5 = tk.Text(window, height = 1, width = 9, bg = 'black', fg = 'yellow', font = "Courier 10 italic bold")
text5.place(x=639,y=538)
text5.insert(tk.END, 'PDD(10-1)')
text5.configure(state='disabled')
text6 = tk.Text(window, height = 1, width = 15, bg = 'black', fg = 'yellow', font = "Courier 10 italic bold")
text6.place(x=403,y=538)
text6.insert(tk.END, 'PDD(Info. Req.)')
text6.configure(state='disabled')
text7 = tk.Text(window, height = 1, width = 15, bg = 'black', fg = 'yellow', font = "Courier 10 italic bold")
text7.place(x=813,y=538)
text7.insert(tk.END, 'PDD(Adv. Issd.)')
text7.configure(state='disabled')
text8 = tk.Text(window, height = 1, width = 17, bg = 'black', fg = 'yellow', font = "Courier 10 italic bold")
text8.place(x=153,y=538)
text8.insert(tk.END, 'PDD(New Receipts)')
text8.configure(state='disabled')
v5 = StringVar(window, value='0')
entry101 = tk.Entry(window, width = 5,  textvariable = v5)
entry101.grid(row=5,column=5)
entry101.place(x = 735,y = 538)
v6 = StringVar(window, value='0')
entryAI = tk.Entry(window, width = 5,  textvariable = v6)
entryAI.grid(row=5,column=5)
entryAI.place(x = 970,y = 538)
v7 = StringVar(window, value='0')
entryNR = tk.Entry(window, width = 5,  textvariable = v7)
entryNR.grid(row=5,column=5)
entryNR.place(x = 330,y = 538)
buttonVGSBW = tk.Button(window, text = 'BANK-WISE Explode FFC PDD', bg='orange', fg='black',command = VGS_BW ,borderwidth = 6,font = "Helvetica 10 bold")
buttonVGSBW.grid(row=5,column=5)
buttonVGSBW.place(x=50, y=638)
buttonVGSDOWISE = tk.Button(window, text = 'DO-WISE Explode FFC PDD', bg='orange', fg='black',command = VGS_DO ,borderwidth = 6,font = "Helvetica 10 bold")
buttonVGSDOWISE.grid(row=5,column=5)
buttonVGSDOWISE.place(x=350, y=638)
buttonVGSALL = tk.Button(window, text = 'TOTAL FFC PDD', bg='orange', fg='black',command = VGS_DO_ALL ,borderwidth = 6,font = "Helvetica 10 bold")
buttonVGSALL.grid(row=5,column=5)
buttonVGSALL.place(x=260, y=589)
buttonVGSIOWISE = tk.Button(window, text = 'IO/RO-WISE Explode Drafts', bg='orange', fg='black',command = VGS_Drafts ,borderwidth = 6,font = "Helvetica 10 bold")
buttonVGSIOWISE.grid(row=5,column=5)
buttonVGSIOWISE.place(x=720, y=638)
buttonVGSDRALL = tk.Button(window, text = 'TOTAL Drafts Pending', bg='orange', fg='black',command = VGS_Drafts_ALL ,borderwidth = 6,font = "Helvetica 10 bold")
buttonVGSDRALL.grid(row=5,column=5)
buttonVGSDRALL.place(x=740, y=589)

v = StringVar(window, value='Enter date in dd-mm-yyyy format')
entryS = tk.Entry(window, width = 30,  textvariable = v)
entryS.grid(row=5,column=5)
entryS.place(x = 700,y = 155)
v1 = StringVar(window, value='BO New Delhi II')
entryT = tk.Entry(window, width = 14,  textvariable = v1)
entryT.grid(row=5,column=5)
entryT.place(x = 80,y = 30)
buttonT = tk.Button(window, text = 'Change Territory', bg='purple', fg='red',command = territory ,borderwidth = 6,font = "Helvetica 10 bold")
buttonT.grid(row=5,column=5)
buttonT.place(x=60, y=60)
text1 = tk.Text(window, height =3 , width=40, bg = 'light grey', fg = 'blue', font = "Arial 15 bold")
text1.pack()
text1.insert(tk.END, "       ....Dedicated to the finest boss in RBI...." +'\n' + '\n' + '                    (press a button below)')
text1.configure(state='disabled')
text2 = tk.Text(window, height =4 , width=122, bg = 'light grey', fg= 'blue',font = "Courier 10 italic")
text2.place(x=85,y=680)
text2.insert(tk.END, 'Note: Just put the BO ND 2 and the NOPNO csv files as recd from Nagpur Data Centre by re-naming them as data and NOPNO respectively. Set PDD values zero to get all pending in any or all categories. Dont forget to load the data before starting.\n PDD - Past Due Days(from today-excluding weekends),   FFC - Full Fledged Complaints (Non-drafts). These 4 PDD values lead to Logical OR operations, means if any one condition of PDD yields result, it will produce the excel file,else,no output.')
text2.configure(state='disabled')
count_all = tk.Button(window, text = '     ', bg='black', fg='yellow',command = to_xl_all, borderwidth = 6,font = "Helvetica 10 bold")
count_all.place(x=970, y=150)
text3 = tk.Text(window, height = 1, width = 5, bg = 'black', fg = 'yellow', font = "Courier 10 italic bold")
text3.place(x=1030,y=160)
text3.insert(tk.END, 'TOTAL')
text3.configure(state='disabled')
count_all_M = tk.Button(window, text = '     ', bg='black', fg='yellow',command = to_xl_M, borderwidth = 6,font = "Helvetica 10 bold")
count_all_M.place(x=970, y=190)
text4 = tk.Text(window, height = 1, width = 13, bg = 'black', fg = 'yellow', font = "Courier 10 italic bold")
text4.place(x=1030,y=200)
text4.insert(tk.END, 'Maintainable')
text4.configure(state='disabled')
text5 = tk.Text(window, height = 3, width = 40, bg = 'black', fg = 'yellow', font = "Courier 10 italic bold")
text5.place(x=780,y=235)
text5.insert(tk.END, 'Click on the no. to Export to Excel- it will automatically be saved to the same folder where this program was run from.')
text5.configure(state='disabled')
window.mainloop()
