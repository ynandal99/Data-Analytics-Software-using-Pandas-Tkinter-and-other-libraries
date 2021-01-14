# -*- coding: utf-8 -*-
"""
Created on Sat May  9 09:44:09 2020

@author: Nandal
"""

import pandas as pd
import _datetime

#MAKE SURE TO IMPORT CSV INTO EXCEL FIRST, DELIMITER TAB AND COMMAS- AS DMY IN ALL DATE FIELDS (USING DATA->FROM TEXT), THEN APPLY FORMATTING IN EACH DATE FIELD - CUSTOM- PRESSING CTRL +1 THEN SELECET *14/03/2020 OR *DD/MM/YYYY FORMAT 
attempt = 0
while True:
  selector = str(input('For BO - enter 1 , for NBFCO - enter 2, for ODT - enter 3 (type exit to exit) :'))

  if selector=='1' or selector =='2' or selector =='3':
    #Input For Quarter
    qbeg , qend= input('Enter Quarter Beginning Date in the format dd-mm-yyyy: '), input('Enter Quarter Ending Date in the format dd-mm-yyyy: ')
    
    qbeg = pd.to_datetime(qbeg, format='%d-%m-%Y').date()
    qend = pd.to_datetime(qend, format='%d-%m-%Y').date()
    
    print('\n\n')
    office = input('If your file has data for multiple offices, please enter one office name (such as \'BO Kanpur\' or \'BO New Delhi II\' or \'ODT New Delhi II\' etc., else type \'run\' to keep running if the file is for one office only (enter without quotes & words are case sensitive) : ')
    
    data = pd.read_excel('file.xlsx')
    # data = data.dropna(how='all')
    if office != 'run':
      data = data[data['TERRITORYNAME'] == office]
    data = data[data['ORIGIN'] != 'FRC Portal']
    data = data[data['ORIGIN'] != 'Sub Judice Portal']
    data.loc[data['STATUSCODE'] == 'Assign to other Regulatory bodies', 'COMPLAINTCLOSEDON'] = data['LASTMODIFIEDON']
    data.loc[data['STATUSCODE'] == 'Assign to other Regulatory bodies', 'STATUSCODE'] =  'Complaint Closed'
    data.loc[data['LAYOUT_NM'] == 'Draft Complaints', 'DATE_OF_REGISTRATION'] = data['CREATEDON']
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
    data['DATE_OF_REGISTRATION'].fillna(data['CREATEDON'], inplace=True)
    del data['CREATEDON']
    data = data.rename(columns = {'DATE_OF_REGISTRATION':'CREATEDON'})
    data['CREATEDON'] = pd.to_datetime(data['CREATEDON'], format='%d-%m-%Y').dt.date
    data['LASTMODIFIEDON'] = pd.to_datetime(data['LASTMODIFIEDON'], format='%d-%m-%Y').dt.date
    data['COMPLAINTCLOSEDON'] = pd.to_datetime(data['COMPLAINTCLOSEDON'], format='%d-%m-%Y').dt.date
    
    totrecd = data[(data['CREATEDON'] >= qbeg) & (data['CREATEDON'] <= qend)]
    
    data = data.filter(['COMPLAINT_REF_NO', 'STATUSCODE', 'LAYOUT_NM', 'CREATEDON', 'LASTMODIFIEDON', 'COMPLAINTCLOSEDON', 'ORIGIN', 'SUBCATEGORY'])
    totrecd = totrecd.filter(['COMPLAINT_REF_NO', 'BANK_NAME', 'STATUSCODE', 'LAYOUT_NM', 'CREATEDON', 'LASTMODIFIEDON', 'COMPLAINTCLOSEDON', 'ORIGIN', 'SUBCATEGORY'])
      
    drafts = data.loc[data['LAYOUT_NM'] == 'Draft Complaints']
    
    totrecd_drafts = totrecd.loc[totrecd['LAYOUT_NM'] == 'Draft Complaints']
    
    tables = {'Volume of Complaints for given quarter (including drafts): ':len(totrecd.axes[0]), 'Of which, Main Complaints: ':len(totrecd.axes[0]) - len(totrecd_drafts.axes[0]), 'Of which, Drafts: ':len(totrecd_drafts.axes[0])}
    
    keyss = tables.keys()
    
    with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
      f.write('\n'+'---------------Note:::: This data excludes FRC/Subjudice automatically closed Complaints--------------' + '\n\n\n')
      f.write('\n' + '=================================================================='+ '\n \n')
      f.write('FOR QUARTER : '+ qbeg.strftime('%d')+ ' '+ qbeg.strftime('%b')+ ' '+ qbeg.strftime('%Y') + ' --TO-- ' + qend.strftime('%d')+ ' '+ qend.strftime('%b')+' '+ qend.strftime('%Y') +'\n')
      f.write('\n' + '==================================================================' + '\n \n')
      for t in keyss:
        f.write((str(t) + str(tables[t])))
        f.write('\n \n')
    
    checker = {}
    modes = []
    modelist = totrecd['ORIGIN'].to_list()
    for i in modelist:
      if i not in checker:
        checker[i] = True
        modes.append(i)
    
    update_mode = {'Mode of which, : ' + str(i) + ': ': modelist.count(i) for i in modes}
    
    keyss = update_mode.keys()
    
    with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
      f.write('\n' + '------------------------------------------------------------------'+'\n\n')
      for t in keyss:
        f.write((str(t) + str(update_mode[t])))
        f.write('\n \n')
    
    tables.update(update_mode)
    
    if selector =='1':
      checker1 = {}
      grounds = []
      grounds_list = totrecd['SUBCATEGORY'].to_list()
      for j in grounds_list:
        if j not in checker1:
          checker1[j] = True
          grounds.append(j)
      lister = {j + ': ': grounds_list.count(j) for j in grounds}
      
      lister['Relating to Cards: '] = lister['ATM/DEBIT CARD: '] + lister['CREDIT CARD: ']
      del lister['ATM/DEBIT CARD: '] 
      del lister['CREDIT CARD: ']
      lister['Relating to Electronic / Mobile Banking: '] = lister['MOBILE BANKING / ELECTRONIC BANKING: ']
      del lister['MOBILE BANKING / ELECTRONIC BANKING: ']
      lister['Relating to non-adherence to FPC: '] = lister['FAIR PRACTICES: ']
      del lister['FAIR PRACTICES: ']
      lister['Relating to Charges & Loans: '] = lister['CHARGES WITHOUT PRIOR NOTICE: '] + lister['LOANS AND ADVANCE- HOUSING: '] + lister['LOANS AND ADVANCE- OTHERS: ']
      del lister['CHARGES WITHOUT PRIOR NOTICE: ']
      del lister['LOANS AND ADVANCE- HOUSING: '] 
      del lister['LOANS AND ADVANCE- OTHERS: ']
      lister['Relating to Pension Payments: '] = lister['PENSION: ']
      del lister['PENSION: ']
      lister['Others : '] = 0
      lister['Not Available (Draft Complaints): '] = lister[' : ']
      del lister[' : ']
      for k in list(lister):
        if (k != 'Relating to Cards: ') and (k != 'Relating to Pension Payments: ') and (k != 'Relating to Charges & Loans: ') and (k != 'Relating to non-adherence to FPC: ') and (k != 'Relating to Electronic / Mobile Banking: ') and (k != 'Not Available (Draft Complaints): ') and(k != 'Others : '):
            lister['Others : '] += lister [k]
            del lister[k]
      
      keyss = lister.keys()
      
      with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
        f.write('\n' + '------------------------------------------------------------------'+'\n\n')
        for t in keyss:
          f.write((str(t) + str(lister[t])))
          f.write('\n \n')
      
      tables.update(lister)
    
    elif selector == '2':
      checker1 = {}
      grounds = []
      grounds_list = totrecd['SUBCATEGORY'].to_list()
      for j in grounds_list:
        if j not in checker1:
          checker1[j] = True
          grounds.append(j)
      
      lister = {str(j) + ': ': int(grounds_list.count(j)) for j in grounds}
      
      lister['Non Adherence to FPC: '] = lister['Non-Adherence to Fair Practices Code: '] + lister['FAIR PRACTICES: ']
      del lister['FAIR PRACTICES: ']
      del lister['Non-Adherence to Fair Practices Code: ']
      
      lister['Lack of Transparency: '] = lister['No Transparency in Contract/Loan: ']
      del lister['No Transparency in Contract/Loan: ']
      
      lister['Levy of Charges without notice: '] = lister['CHARGES WITHOUT PRIOR NOTICE: '] + lister['Levying of Charges without Notice: ']
      del lister['CHARGES WITHOUT PRIOR NOTICE: '] 
      del lister['Levying of Charges without Notice: ']
      
      lister['Others : '] = 0
      lister['Not Available (Draft Complaints): '] = lister[' : ']
      del lister[' : ']
      for k in list(lister):
        if (k != 'Lack of Transparency: ') and (k != 'Non Adherence to FPC: ') and (k != 'Levy of Charges without notice: ') and (k != 'Not Available (Draft Complaints): ') and(k != 'Others : '):
            lister['Others : '] += lister [k]
            del lister[k]
      
      keyss = lister.keys()
      
      with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
        f.write('\n' + '------------------------------------------------------------------'+'\n\n')
        for t in keyss:
          f.write((str(t) + str(lister[t])))
          f.write('\n \n')
      
      tables.update(lister)
    
    else:
      checker1 = {}
      grounds = []
      grounds_list = totrecd['SUBCATEGORY'].to_list()
      for j in grounds_list:
        if j not in checker1:
          checker1[j] = True
          grounds.append(j)
      
      lister = {j + ': ': grounds_list.count(j) for j in grounds}
      
      lister['Others : '] = 0
      lister['Not Available (Draft Complaints): '] = lister[' : ']
      del lister[' : ']
      for k in list(lister):
        if (k != 'Fund Transfers/UPI/BBPS/Bharat QR Code: ') and (k != 'Mobile/Electronic Fund Transfers: ') and (k != 'Not Available (Draft Complaints): ') and(k != 'Others : '):
            lister['Others : '] += lister [k]
            del lister[k]
      
      keyss = lister.keys()
      
      with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
        f.write('\n' + '------------------------------------------------------------------'+'\n\n')
        for t in keyss:
          f.write((str(t) + str(lister[t])))
          f.write('\n \n')
  
      tables.update(lister)
    
    banksall = totrecd['BANK_NAME'].to_list()
    checker2 = {}
    banks_list = []
    
    for bank in banksall:
      if bank not in checker2:
        checker2[bank] = True
        banks_list.append(bank)
    try:
      for i in range(len(banks_list)):
        if str(banks_list[i]) == 'nan':
          del banks_list[i]
    except:
      None  
    banks_dict = {str(item) + ' (no. for given quarter) : ': banksall.count(item) for item in banks_list}
    try: 
      del banks_dict['Others (no. for given quarter) : ']
    except:
      try:
        del banks_dict['nan (no. for given quarter) : ']
      except:
        None
    sortedbanks = {k:v for k,v in sorted(banks_dict.items(), key = lambda items: items[1], reverse = True)}
    
    top5dict = dict(list(sortedbanks.items())[0:5]) 
    
    keyss = top5dict.keys()
    
    with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
      f.write('\n' + '------------------------------------------------------------------'+'\n\n')
      for t in keyss:
        f.write((str(t) + str(top5dict[t]) +' - %age to total volume in given quarter: '+str(round((top5dict[t]/tables['Volume of Complaints for given quarter (including drafts): '])*100,2))+' AND, %age to Main complaints in given quarter: '+str(round((top5dict[t]/tables['Of which, Main Complaints: '])*100,2))))
        f.write('\n \n')
    
    #for ALL DISPOSED DURING THE QUARTER
    
    cond_closed_1 = data['STATUSCODE'] == 'Complaint Closed'
    cond_closed_2 = data['STATUSCODE'] == 'Not a Complaint'
    
    all_closed_main = data[cond_closed_1]
    all_closed_drafts = data[cond_closed_2]
    
    cond_closed_3 = all_closed_main['COMPLAINTCLOSEDON'] >= qbeg
    cond_closed_4 = all_closed_main['COMPLAINTCLOSEDON'] <= qend
    
    cond_closed_5 = all_closed_drafts['LASTMODIFIEDON'] >= qbeg
    cond_closed_6 = all_closed_drafts['LASTMODIFIEDON'] <= qend
    
    closed_during_qtr_main = all_closed_main[cond_closed_3 & cond_closed_4]
    closed_during_qtr_drafts = all_closed_drafts[cond_closed_5 & cond_closed_6]
    
    closed_during_qtr_ALL = pd.concat([closed_during_qtr_main, closed_during_qtr_drafts])
    
    closed_list = {'Total Disposed during the quarter: ': len(closed_during_qtr_ALL.axes[0]), 'Of which, Drafts: ':len(closed_during_qtr_drafts.axes[0]), 'Of which, Main Complaints: ':len(closed_during_qtr_main.axes[0])}
    
    keyss = closed_list.keys()
    
    with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
      f.write('\n' + '------------------------------------------------------------------'+'\n\n')
      for t in keyss:
        f.write((str(t) + str(closed_list[t])))
        f.write('\n \n')
    
    
    tables.update(closed_list)
    
    # A - for pending at the END of the quarter
    
    main = data.loc[data['LAYOUT_NM'] != 'Draft Complaints']
    
    cond_pending_main1 = main['STATUSCODE'] != 'Complaint Closed'
    cond_pending_main2 = main['CREATEDON'] <= qend
    pending_main1 = main[cond_pending_main1]                              #all pending main
    
    #Main complaints pending in still open - END
    pending_main_in_end_open = main[cond_pending_main1 & cond_pending_main2]
    
    cond_closed_main1 = (all_closed_main['COMPLAINTCLOSEDON'] > qend)
    pending_main_in_closed1 = all_closed_main[cond_closed_main1]
    
    cond_closed_main2 = pending_main_in_closed1['CREATEDON'] >= qbeg
    cond_closed_main3 = pending_main_in_closed1['CREATEDON'] <= qend
    
    #Main complaints pending in closed - END
    pending_main_in_end_closed = pending_main_in_closed1[cond_closed_main2 & cond_closed_main3]
    
    cond_pending_drafts1 = drafts['STATUSCODE'] != 'Not a Complaint'
    cond_pending_drafts2 = drafts['CREATEDON'] <= qend
    pending_drafts1 = drafts[cond_pending_drafts1]                        #all pending drafts
    
    #Drafts pending in still open - END
    pending_drafts_in_end_open = drafts[cond_pending_drafts1 & cond_pending_drafts2]
    
    cond_closed_drafts1 = (all_closed_drafts['LASTMODIFIEDON'] > qend)
    pending_drafts_in_closed1 = all_closed_drafts[cond_closed_drafts1]
    
    cond_closed_drafts2 = pending_drafts_in_closed1['CREATEDON'] >= qbeg
    cond_closed_drafts3 = pending_drafts_in_closed1['CREATEDON'] <= qend
    
    #Drafts pending in closed - END
    pending_drafts_in_end_closed = pending_drafts_in_closed1[cond_closed_drafts2 & cond_closed_drafts3]
    
    # B - for pending at the BEGINNING of the quarter
    
    cond_closed_drafts4 = all_closed_drafts['LASTMODIFIEDON'] >= qbeg
    cond_closed_drafts5 = all_closed_drafts['CREATEDON'] < qbeg
    
    pending_drafts_at_beg_closed1 = all_closed_drafts[cond_closed_drafts4 & cond_closed_drafts5]
    cond_closed_drafts6 = pending_drafts_at_beg_closed1['LASTMODIFIEDON'] <= qend
    
    #Draft pending in closed - BEGINNING
    pending_drafts_at_beg_closed = pending_drafts_at_beg_closed1[cond_closed_drafts6]
    
    cond_pending_drafts3 = pending_drafts1['CREATEDON'] < qbeg
    
    #Draft pending in still open - BEGINNING
    pending_drafts_at_beg_open = pending_drafts1[cond_pending_drafts3]
    
    cond_closed_main4 = all_closed_main['COMPLAINTCLOSEDON'] >= qbeg
    cond_closed_main5 = all_closed_main['CREATEDON'] < qbeg
    
    pending_main_at_beg_closed1 = all_closed_main[cond_closed_main4 & cond_closed_main5]
    cond_closed_main6 = pending_main_at_beg_closed1['COMPLAINTCLOSEDON'] <= qend
    
    #Main complaints pending in closed - BEGINNING
    pending_main_at_beg_closed = pending_main_at_beg_closed1[cond_closed_main6]
    
    cond_pending_main3 = pending_main1['CREATEDON'] < qbeg
    
    #Main complaints pending in still open - BEGINNING
    pending_main_at_beg_open = pending_main1[cond_pending_main3]
    
    #FINAL PENDING DICT
    pending_beg_list = {'Total Pending at the beginning of the quarter: ':len(pending_drafts_at_beg_closed.axes[0]) + len(pending_drafts_at_beg_open.axes[0])+ len(pending_main_at_beg_closed.axes[0]) + len(pending_main_at_beg_open.axes[0]), 'Of which, Drafts: ':len(pending_drafts_at_beg_closed.axes[0]) + len(pending_drafts_at_beg_open.axes[0]), 'Of which, Main Complaints: ':len(pending_main_at_beg_closed.axes[0]) + len(pending_main_at_beg_open.axes[0])}
    
    pending_end_list = {'Total Pending at the end of the quarter: ':len(pending_main_in_end_open.axes[0]) + len(pending_main_in_end_closed.axes[0])+ len(pending_drafts_in_end_open.axes[0])+ len(pending_drafts_in_end_closed.axes[0]), 'Of Which Drafts: ':len(pending_drafts_in_end_open.axes[0])+ len(pending_drafts_in_end_closed.axes[0]), 'Of Which, Main Complaints: ':len(pending_main_in_end_open.axes[0]) + len(pending_main_in_end_closed.axes[0])}
    
    keyss = pending_beg_list.keys()
    
    with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
      f.write('\n' + '------------------------------------------------------------------'+'\n\n')
      for t in keyss:
        f.write((str(t) + str(pending_beg_list[t])))
        f.write('\n \n')
        
    keyss = pending_end_list.keys()
    
    with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
      f.write('\n' + '------------------------------------------------------------------'+'\n\n')
      for t in keyss:
        f.write((str(t) + str(pending_end_list[t])))
        f.write('\n \n')
    
    
    #UPDATING IN TABLES
    tables.update(pending_beg_list)
    tables.update(pending_end_list)
    
    #CALCULATING AGE-WISE (GREATER THAN 1 MONTH 2 MONTH ETC)
    
    agedata = pd.concat([pending_main_in_end_open, pending_main_in_end_closed, pending_drafts_in_end_open, pending_drafts_in_end_closed])
    
    agedata['Pending in months at end'] = agedata['CREATEDON'].map(lambda x: (qend.year - x.year)*12 + (qend.month - x.month))
    
    age_list = agedata['Pending in months at end'].tolist()
    
    pendency_dict = { 'Months '+ str(a) : age_list.count(a) for a in age_list}
    
    pendency = {}
    pendency['Less than 1 month: '] =  0
    pendency['Between 1-2 months: '] =  0
    pendency['Between 2-3 months: '] =  0
    pendency['Greater than 3 months: '] =  0
    
    
    for key in pendency_dict:
      if int(key[6:]) == 0:
        pendency['Less than 1 month: '] += pendency_dict[key]
      elif int(key[6:]) >= 1 and int(key[6:]) < 2:
        pendency['Between 1-2 months: '] += pendency_dict[key]
      elif int(key[6:]) >= 2 and int(key[6:]) < 3:
        pendency['Between 2-3 months: '] += pendency_dict[key]
      elif int(key[6:]) >= 3:
        pendency['Greater than 3 months: '] += pendency_dict[key]
      else:
        break
    
    keyss = pendency.keys()
    
    with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
      f.write('\n' + '------------------------------------------------------------------'+'\n\n')
      for t in keyss:
        f.write( (str(t) + str(pendency[t])))
        f.write('\n \n')
    
    perc_pendency = {'%age of pending complaints at the end of the quarter: ': round((pending_end_list['Total Pending at the end of the quarter: '] /(pending_beg_list['Total Pending at the beginning of the quarter: '] + tables['Volume of Complaints for given quarter (including drafts): '])) * 100 , 2) }
    
    keyss = perc_pendency.keys()
    
    with open('QUARTERLY DATA {}.txt'.format(qbeg.strftime('%b')+qbeg.strftime('%y') + ' to ' + qend.strftime('%b')+ qend.strftime('%y')),'a+') as f:
      f.write('\n' + '------------------------------------------------------------------'+'\n\n')
      for t in keyss:
        f.write( (str(t) + str(perc_pendency[t])))
        f.write('\n \n')
    print('\n\n\n'+'Success.........!!!!   Check file in same directory as this program')
    break
  elif selector == 'exit':
    print('Exiting on user Request !! ')
    attempt +=1
    print('Attempt: '+str(attempt))
    break
  else:
    print('Try Again with proper input !!!')
    attempt +=1
    print('Attempt: '+str(attempt))
    continue