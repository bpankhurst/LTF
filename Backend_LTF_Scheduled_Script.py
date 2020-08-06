# -*- coding: utf-8 -*-
"""
Created on Mon Mar  9 11:38:53 2020

@author: BPankhurst
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Mar  6 17:42:39 2020

@author: BPankhurst
"""
                                                   
import pandas as pd
import numpy as np
import os
import datetime
import glob


#Set up path in which to save email attatchments to
path = os.path.expanduser(r'C:\Users\bpankhurst\Box\LTF_Project\FINAL UPLOAD FILES')

daily_path = r'C:\Users\bpankhurst\Alvarez and Marsal\Project Diamond in the Rough - LTF and Backlog Dashboard\Daily Data uploads'
manual_path = r'C:\Users\bpankhurst\Alvarez and Marsal\Project Diamond in the Rough - LTF and Backlog Dashboard\ECPR_LTF_BackLog Manual Input files'
pillar_path = r'C:\Users\bpankhurst\Alvarez and Marsal\Project Diamond in the Rough - General\Pillar_list'
agress_path = r'C:\Users\bpankhurst\Box\LTF_Project\FINAL UPLOAD FILES'
save_path = r'C:\Users\bpankhurst\Box\LTF_Project\Upload_trial'
email = 'bpankhurst@alvarezandmarsal.com'

#Set todays date so that we can restrict only to atttatchments recieved today
today = datetime.date.today()
    
 # Conect to outlook   
from win32com.client import Dispatch
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

#Now select folders USED COMMENTTED OUT FOR LOOPS (28-31 & 38-40) IF UNSURE OF EMAIL FOLDER STRUCTURE

for i in range(1,5):
    try:
        print(outlook.Folders.Item(i))
    except Exception:
            print('No more')

#Select root folder
#root_folder = outlook.Folders.Item(4)
root_folder = outlook.Folders.Item(email)



#Select correct rootfolder: 
for folder in root_folder.Folders:
    print (folder.Name)
  
inbox_folder = root_folder.Folders['Inbox']

#iF YOU HAVE NOT CREATED A SEPERATE FOLDER FOR ATTATCHMENTS THEN SKIP TO 46 AND USE messages = inbox_folder.Items 

#Now select subfolder (if necessary)
for folder in inbox_folder.Folders:
    print(folder.Name)
    break

BIRST_subfolder = inbox_folder.Folders['BIRST']

#Select all emails within this subfolder as messages
messages = BIRST_subfolder.Items

def saveattachemnts(subject):
    for message in messages:
        if message.Subject[:4]== subject and message.Senton.date() == today :
            # body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            print('File saved: '+ str(attachment))
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                if message.Subject == subject and message.Unread:
                    message.Unread = False
                break
      
            
saveattachemnts('TEST') 
saveattachemnts('Test') 



os.chdir(manual_path)


#Read in SBU lookup
SBU = pd.read_excel('SBU_Country_Lookup.xlsx')


#Read in HRIS emial address list
email = pd.read_excel('MD Email Addresses_HRIS.xlsx')
email.info()

email.columns = email.columns.str.replace(" ", "_")

email = email[['Employee_ID','Full_Name', 'First_Name','Last_Name','Email_Address']]
email.info()


#Read in Targets
target = pd.read_excel('MD Targets 2020_24_04_WORKING FILE2.xlsx')
target.info()

target.rename(columns={'2020 SBU': 'SBU'}, inplace=True)

target.columns = target.columns.str.replace(" ", "_")
target.columns = target.columns.str.replace(".", "_")

#target['Target_FY20'] = target['Target_FY20'].str.replace(" ","")
target['Target_FY20'] = pd.to_numeric(target['Target_FY20'], downcast='float')


target['LTF_Target'] = pd.to_numeric(target['LTF_Target'], downcast='float')


target['FY19_Actual'] = pd.to_numeric(target['FY19_Actual'], downcast='float')



target1 = target[['Resource_Type','Capability_Pillar','2020_Division','2020_Resource_Type_Ranked','Emp_ID','Emp_Name','SBU','Country_','Status','Start_Date','Target_FY20','LTF_Target','FY19_Actual']]



House = pd.read_csv('LWC Assignments by House_Updated_19032020.csv') 

House.columns = House.columns.str.replace(" ", "_")

House.info()

House['New_LWC'] = pd.to_numeric(House['New_LWC'], downcast='float')

House['New_RC'] = pd.to_numeric(House['New_RC'], downcast='float')



os.chdir(pillar_path)

#Read in Pillar lookup
pillar = pd.read_excel('Pillars1.xlsx', 'Pillars1')
pillar.info()

pillar = pillar[['Employee_ID','full_name', 'Pillar', 'Employee_SBU', 'Pillar_clean', 'Division', 'ECPR_Division']]


#os.chdir(r'C:\Users\bpankhurst\Box\LTF_Project\FINAL UPLOAD FILES\OLD')
#MD_Tree = pd.read_excel('Project Trees run 5th November 2019 PL Edit.xlsx')
os.chdir(daily_path)


for file in glob.glob('GLOBAL_COLLECTIONS*.csv'):
    Col = pd.read_csv(file)

#Col = pd.read_csv('GLOBAL COLLECTIONS 2020.csv')
#Col = pd.read_csv('GLOBAL_COLLECTIONS_2020_v2.csv')


Col.columns = Col.columns.str.replace(" ", "_")
Col.columns = Col.columns.str.replace("(", "_")
Col.columns = Col.columns.str.replace(")", "")

Col['MD_Summary_EUR'] = Col['MD_Summary_EUR'].str.replace(",","").astype(float)

Col['MD_Summary_EUR'] = pd.to_numeric(Col['MD_Summary_EUR'])
Col['Period'] = pd.to_numeric(Col['Period'])

Col.columns = Col.columns.str.replace(" ", "_")

Col.info()

for file in glob.glob('Tree Browser by MD*.xlsx'):
    MD_Tree = pd.read_excel(file,sheet_name='Tree Browser by MD')

#MD_Tree = pd.read_excel('Tree Browser by MD.xlsx', sheet_name='Tree Browser by MD')

MD_Tree.info()

#MD_Tree.rename(columns={'dim1': 'Project', 'xdim1': 'Project (T)','dim2' : '}, inplace=True)

MD_Tree['Percentage'] = pd.to_numeric(MD_Tree['Percentage'], downcast='float')
#.round(2)
MD_Tree['Period_to'] = pd.to_numeric(MD_Tree['Period_to'], downcast='integer')


MD_Tree_Mar = MD_Tree[[ 'Employee (T)', 'Project', 'Project (T)', 'Project SBU',
       'Project SBU (T)', 'Tt', 'Period_to','Percentage', 'Project Status']]

MD_Tree_Mar_1 = MD_Tree_Mar.dropna(how = 'any', subset = ['Employee (T)'])
MD_Tree_Mar = MD_Tree_Mar_1

MD_Tree.columns
MD_Tree_Mar.columns = MD_Tree_Mar.columns.str.replace(" ", "_")
MD_Tree_Mar.columns = MD_Tree_Mar.columns.str.replace("(", "_")
MD_Tree_Mar.columns = MD_Tree_Mar.columns.str.replace(")", "")

MD_Tree_Mar.columns

#Read in AR file for invoice tab

for file in glob.glob('AR Open Items ECPR*.xlsx'):
    invoice = pd.read_excel(file,sheet_name='AR Open Items ECPR')

invoice.columns = invoice.columns.str.replace(" ", "_")
invoice.columns = invoice.columns.str.replace("(", "")
invoice.columns = invoice.columns.str.replace(")", "")
invoice.columns = invoice.columns.str.replace(".", "")
invoice.info()

for file in glob.glob('Global Client list*.xlsx'):
    Client = pd.read_excel(file)
    
#Client = pd.read_excel('Global Client List as at 29th June 2020.xlsx')

Client['CustomerID'] = pd.to_numeric(Client['CustomerID'], downcast='integer')

#Client['CustomerID'] = Client['CustomerID'].fillna(0).astype(int).astype(str)
#Client['CustomerID']  = Client['CustomerID'].replace({'0':np.nan}) 

Client.columns = Client.columns.str.replace(" ", "_")
Client.columns = Client.columns.str.replace("(", "_")
Client.columns = Client.columns.str.replace(")", "")

#MD_Tree.rename(columns={'dim1':
Client.rename(columns = {'Pif_Dropdown_List_-_Industry_Type_T' : 'Industry'}, inplace = True)

Client = Client[['Md', 'Md_T', 'Project_SBU', 'Project', 'Project_T',
       'CustomerID', 'CustomerID_T', 'Industry','TC', 'Start_from', 'Completed']]

Client = Client[Client['CustomerID']!=825550]

Client.info()


os.chdir(agress_path)


import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

sheet = client.open('New Project Entry Form (Responses)').sheet1

prime_r = sheet.get_all_records()

Prime = pd.DataFrame(prime_r)

Prime.columns = Prime.columns.str.replace(" ", "_")

Prime = Prime[Prime['reason_for_removal']=='']

Prime['Lead_MD']= Prime['Employee']

Prime = Prime[['Timestamp','Project_Clean', 'Customer','email_address','Employee','LWC_1','RC_1',
                  'Lead_MD','Employee_2','LWC_2','RC_2', 'Pipeline_Amount', 'FY21_Pipeline_Amount',
                  'Conversion_Probability', 'Sector', 'Type_of_work','Project_start_date']]

#'Hard_Backlog','Soft_Backlog'

Prime['Project'] = Prime['Project_Clean']

Prime.info()

Prime['Soft_Backlog'] = np.where(Prime.Conversion_Probability==1,Prime['Pipeline_Amount'],0)
Prime['Pipeline_value'] = np.where(Prime.Conversion_Probability<1,Prime['Pipeline_Amount'],0)
Prime['FY21_Soft_Backlog'] = np.where(Prime.Conversion_Probability==1,Prime['FY21_Pipeline_Amount'],0)
Prime['FY21_Pipeline_value'] = np.where(Prime.Conversion_Probability<1,Prime['FY21_Pipeline_Amount'],0)




Prime = Prime[['Timestamp','Project', 'Customer','email_address','Employee','LWC_1','RC_1',
                  'Lead_MD','Employee_2','LWC_2','RC_2', 'Pipeline_Amount', 'FY21_Pipeline_Amount',
                  'Pipeline_value', 'FY21_Pipeline_value',
                  'Conversion_Probability', 'Soft_Backlog','FY21_Soft_Backlog', 'Sector', 'Type_of_work', 'Project_start_date']]


Prime['LWC_2'] = pd.to_numeric(Prime['LWC_2'], downcast='float')
Prime['RC_2'] = pd.to_numeric(Prime['RC_2'], downcast='float')


#Prime['Pipeline_value'] = pd.to_numeric(Prime['Pipeline_value'], downcast='float')

Prime.info()


sheet = client.open('LTF/Backlog Responses').sheet1

HBL_1 = sheet.get_all_records()

HBL = pd.DataFrame(HBL_1)

HBL.info()

HBL = HBL[['Project','New_Hard_Backlog','New_Soft_Backlog', '% WIP to keep','% AR to keep']]

HBL_1 = HBL 

HBL_1.columns = ['Project','HBL','SBL','WIP_keep','AR_keep']

HBL_1['HBL'] = pd.to_numeric(HBL_1['HBL'], downcast = 'integer')

HBL_1['SBL'] = pd.to_numeric(HBL_1['SBL'], downcast = 'integer')

HBL_1['WIP_keep'] = pd.to_numeric(HBL_1['WIP_keep'], downcast = 'float')

HBL_1['AR_keep'] = pd.to_numeric(HBL_1['AR_keep'], downcast = 'float')


HBL_1.info()


HRIS = pd.read_excel('HRIS_Europe_Headcount_By_Date (36).xlsx', skiprows=[0])

HRIS.info()



HRIS.columns = HRIS.columns.str.replace(" ", "_")
HRIS.columns = HRIS.columns.str.replace("(", "_")
HRIS.columns = HRIS.columns.str.replace(")", "")

Wip = pd.read_csv('Birst Wip.csv')

Wip['WIP_Amount_EUR'] = Wip['WIP_Amount_EUR'].str.replace(",","").astype(float)

Wip['WIP_Amount_EUR'] = pd.to_numeric(Wip['WIP_Amount_EUR'])

Wip.info()


AR = pd.read_csv('Global AR.csv')

AR['AR_Amount_EUR'] = AR['AR_Amount_EUR'].str.replace(",","").astype(float)

AR['AR_Amount_EUR'] = pd.to_numeric(AR['AR_Amount_EUR'])

AR.info()

MD_Birst = pd.read_csv('MDTREEEURO_Update.csv',skiprows=[0,1])


MD_Birst.columns = ['Project','Project (T)','delete', 'Employee (T)','Project SBU', 'Period_Start', 'LWC', 'Referral']

MD_Birst.info()

MD_Birst1 = MD_Birst[['Employee (T)', 'Project', 'Project (T)', 'Project SBU','Period_Start', 'LWC', 'Referral']]


MD_Tree_Mar = MD_Tree[[ 'Employee (T)', 'Project', 'Project (T)', 'Project SBU',
       'Project SBU (T)', 'Tt', 'Period_to','Percentage', 'Project Status']]

MD_Tree_Mar_1 = MD_Tree_Mar.dropna(how = 'any', subset = ['Employee (T)'])
MD_Tree_Mar = MD_Tree_Mar_1

MD_Tree.columns
MD_Tree_Mar.columns = MD_Tree_Mar.columns.str.replace(" ", "_")
MD_Tree_Mar.columns = MD_Tree_Mar.columns.str.replace("(", "_")
MD_Tree_Mar.columns = MD_Tree_Mar.columns.str.replace(")", "")

MD_Tree_Mar.columns



#Read in historical Collections data
#col_2018 = pd.read_csv('GLOBAL_COLLECTIONS_2018.csv')
#col_2019 = pd.read_csv('GLOBAL_COLLECTIONS_2019.csv')

#col_2018['MD_Summary_EUR'] = col_2018['MD_Summary_EUR'].str.replace(",","").astype(float)
#col_2018['MD_Summary_EUR'] = pd.to_numeric(col_2018['MD_Summary_EUR'])
#col_2018['Period'] = pd.to_numeric(col_2018['Period'])

#col_2019['MD_Summary_EUR'] = col_2019['MD_Summary_EUR'].str.replace(",","").astype(float)
#col_2019['MD_Summary_EUR'] = pd.to_numeric(col_2019['MD_Summary_EUR'])
#col_2019['Period'] = pd.to_numeric(col_2019['Period'])


os.chdir(save_path)

HBL_1.to_csv('HBL.txt', index = False, sep = '~')

email.to_csv('email.txt', index=False, sep='~')
invoice.to_csv('invoice.txt', index=False, sep='~')

SBU.to_csv('SBU.txt', index = False, sep='~')

Client.to_csv('Client.txt', index = False, sep = '~')

pillar.to_csv('Pillar.txt', index = False, sep = '~')

target1.to_csv('Targets.txt',index=False,sep='~')

Prime.to_csv('Prime.txt', index=False, sep='~') 

House.to_csv('House.txt', index=False, sep='~')   
           
HRIS.to_csv('HRIS.txt', index=False, sep='~')   

Wip.to_csv('WIP_trial.txt', index=False, sep='~')    

AR.to_csv('AR_trial.txt', index=False, sep='~')

Col.to_csv('Col_trial.txt', index=False, sep='~')
          
           
MD_Tree_Mar.to_csv('MD_Tree_Mar.txt', index=False, sep='~')         

MD_Birst1.to_csv('MD_Tree_Birst.txt', index = False, sep = '~')


print('finished')
