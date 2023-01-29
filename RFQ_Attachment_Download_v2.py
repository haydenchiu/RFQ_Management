import win32com.client
import pandas as pd
from exchangelib import DELEGATE, Account, Credentials, Configuration
from bs4 import BeautifulSoup, NavigableString, Tag
import os
import os.path
from os import path
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
import dateutil.parser
import time
import numpy as np
from pathlib import Path
import re
import sys
import shutil
from win32com import client
import getpass
import glob
import re


user_name = getpass.getuser()

#access forwarder email mapping
os.chdir(fr'C:\Users\{user_name}\Groupe SEB\Supply Chain Data Automation - Documents\Programs\08 RFQ Attachment Download\01 Forwarder List')
ff_mail_df = pd.read_excel('RFQ Forwarder email list.xlsx')

EMAIL_ACCOUNT = 'SEB Asia Logistics'  # e.g. 'good.employee@importantcompany.com'
ITER_FOLDER = 'pending'  # e.g. 'IterationFolder'
MOVE_TO_FOLDER = 'processed'  # e.g 'ProcessedFolder'
SAVE_AS_PATH = r'I:/Logistic Dept/Forwarder Performance Report/ST Team/Air Freight/2. RFQ/RFQ Result/Automate_test/'  # e.g.r'C:\DownloadedCSV'
ERROR_LOG_PATH = r'I:/Logistic Dept/Forwarder Performance Report/ST Team/Air Freight/2. RFQ/RFQ Result/Automate_test/01 Error Log/'
#EMAIL_SUBJ_SEARCH_STRING = 'Duplicate'  # e.g. 'Email to download'
report_date = dt.strftime(dt.date(dt.now()),format="%m-%d-%Y")
#report_start_date = (dt.date(dt.now()) - relativedelta(months=3)).replace(day=1)

def process_RFQ_email():

    out_app = client.gencache.EnsureDispatch('Outlook.Application')
    out_namespace = out_app.GetNamespace("MAPI")

    out_iter_folder = out_namespace.Folders[EMAIL_ACCOUNT].Folders['Inbox'].Folders[ITER_FOLDER]

    out_move_to_folder = out_namespace.Folders[EMAIL_ACCOUNT].Folders['Inbox'].Folders[MOVE_TO_FOLDER]
    
    emails = out_iter_folder.Items
    emails.Sort('[SentOn]', True)

    print(f'Number of email in this batch: {out_iter_folder.Items.Count}')
    
    error_counter = 0

    for i, mail in enumerate(emails):
        try:
            print(i)
            if mail.Class==43 and mail.SenderEmailType=='EX':
                sender_email = mail.Sender.GetExchangeUser().PrimarySmtpAddress
                print(sender_email)
            else:
                sender_email = mail.SenderEmailAddress
                print(sender_email)

            if mail.Attachments.Count > 0:
                print(mail.Subject)
                print(mail.SentOn)
                print(mail.ReceivedTime)

                attachments = mail.Attachments
                for file in attachments:
                    if ('SA' in file.FileName) and (('.xls'in file.FileName) or ('.xlsx'in file.FileName) or ('.xlsm'in file.FileName)):
                        print(file.FileName)
                        #matchObj = re.search(r'SA\d{7}',file.FileName)
                        matchObj = re.search(r'SA[\d]{4}-[\d]{4}',file.FileName)
                        ff_name = ff_mail_df[ff_mail_df['Email']==sender_email]['Forwarder'].unique()[0]
                        time_str = f'{str(mail.SentOn)[:10]}_{str(mail.SentOn)[11:13]}_{str(mail.SentOn)[14:16]}_{str(mail.SentOn)[17:19]}'

                        if not os.path.exists(SAVE_AS_PATH + matchObj.group()):#Check if the RFQ directory already exist, if no create dir
                            os.makedirs(SAVE_AS_PATH + matchObj.group())
                            print(SAVE_AS_PATH + matchObj.group())
                            print(f'RFQ Code: {matchObj.group()}')
                            print(ff_mail_df[ff_mail_df['Email']==sender_email]['Forwarder'].unique()[0])
                            print(f'{str(mail.SentOn)[:10]}')
                            print(f'{str(mail.SentOn)[:10]}_{str(mail.SentOn)[11:13]}_{str(mail.SentOn)[14:16]}_{str(mail.SentOn)[17:19]}')
                            
                            
                        else:
                            print('Saved to existing RFQ folder')
                            print(SAVE_AS_PATH + matchObj.group())
                            print(str(f'{file.FileName}'))
                        
                        j = 0
                        while os.path.exists(SAVE_AS_PATH + matchObj.group() + '/' + ff_name + ' - ' + f'{time_str}' + '-' + f'({j})' + '-' + str(f'{file.FileName}')):#Check if the attachment already exist in our folder, if yes increase the index in file name by 1
                            j += 1
                        
                        print(SAVE_AS_PATH + matchObj.group() + '/' + ff_name + ' - ' + f'{time_str}' + '-' + f'({j})' + '-' + str(f'{file.FileName}'))
                        file.SaveAsFile(SAVE_AS_PATH + matchObj.group() + '/' + ff_name + ' - ' + f'{time_str}' + '-' + f'({j})' + '-' + str(f'{file.FileName}'))

            print('\n')
            

        except Exception as e:
            print(i)
            print(e)
            error_counter += 1
            file = open(ERROR_LOG_PATH + f'error_{report_date}.log', 'w')
            file.write(f'Problem on iteration: {i}')
            file.close()
            print('\n')
            continue

    for mail in list(emails):
        mail.Move(out_move_to_folder)

    print(f'Number of error in this batch: {error_counter}')


def RFQ_Summary():
    
    input_summary_dir = r'I:\Logistic Dept\Forwarder Performance Report\ST Team\Air Freight\2. RFQ\RFQ Summary'
    input_SA_dir = r'I:\Logistic Dept\Forwarder Performance Report\ST Team\Air Freight\2. RFQ\RFQ Result\Automate_test'
    output_dir = r'I:\Logistic Dept\Forwarder Performance Report\ST Team\Air Freight\2. RFQ\RFQ Summary'
    
    
    #Read Send out Summary
    os.chdir(input_summary_dir)
    sum_df = pd.read_excel('AIR RFQ.xlsm',sheet_name='RFQ summary')
    
    sum_df['RFQ no.'] = sum_df['RFQ no.'].str.upper()
    
    #display(sum_df)
    
    #Read Received RFQ
    os.chdir(input_SA_dir)
    
    xlsm_files = glob.glob(f'{input_SA_dir}\**\*.xlsm', recursive=True)
    
    xlsx_files = glob.glob(f'{input_SA_dir}\**\*.xlsx', recursive=True)
    
    xls_files = glob.glob(f'{input_SA_dir}\**\*.xls', recursive=True)
    
    dir_list = [xlsm_files, xlsx_files, xls_files]
    
    frames = []
    
    col_dic = {0:'Forwarder Identity',1:'Total Amount',2:'Service',3:'Transit Time in working days'}
    
    for dir_ in dir_list:
    
        for file in dir_:
            try:
                #print(os.path.basename(file))
                df = pd.read_excel(file,sheet_name='RFQ FORM',header=73,nrows=4,usecols=[6])

                rfq_nbr = re.split(r'\\',os.path.dirname(file))[-1]
                
                matchObj = re.search(r'\d{4}-\d{2}-\d{2}_\d{2}_\d{2}_\d{2}',os.path.basename(file)) #Search Date string in file name

                email_receive_date = matchObj.group()

                matchObj2 = re.search(r'\d{4}-\d{2}-\d{2}_\d{2}_\d{2}_\d{2}-[(]\d+[)]',os.path.basename(file)) #Search Date string in file name

                version_nbr = matchObj2.group().split('(', 1)[1].split(')')[0] #select only the integer inside parentheses

                #print(version_nbr)

                df = df.T

                df.rename(columns=col_dic,inplace=True)

                df = df.reset_index(drop=True)

                df['RFQ no.'] = rfq_nbr
                
                df['Quotation Receive Date'] = pd.to_datetime(email_receive_date,format='%Y-%m-%d_%H_%M_%S')

                df['Version Number'] = version_nbr

                frames.append(df)

            except Exception as e:
                print(e)
                print(f'Cannot access RFQ infomation in: {os.path.basename(file)}')
                pass
        
        sa_df = pd.concat(frames)
        
    sa_df = sa_df.append(sa_df)
    
    sa_df = sa_df.drop_duplicates()
    
    sa_df['RFQ no.'] = sa_df['RFQ no.'].str.upper()

    sa_df.sort_values(by=['RFQ no.','Forwarder Identity','Quotation Receive Date','Version Number'],inplace=True)

    sa_df['First Quotation Receive Date'] = sa_df.groupby(['RFQ no.','Forwarder Identity'],sort=False)['Quotation Receive Date'].transform('min')

    sa_df = sa_df.drop_duplicates(subset=['RFQ no.','Forwarder Identity'],keep='last')
    
    print(sa_df.columns)
    #display(sa_df)
    
    #Left Join Send-out Summary & sa_df
    
    r_df = pd.merge(sum_df,sa_df,how='left',on=['RFQ no.'])
    
    #display(r_df)
    
    os.chdir(output_dir)
    r_df.to_excel('AIR RFQ Matched_test.xlsx',index=False)
    
    return(r_df)

def RFQ_Summary_Tender_Rate():
    
    input_summary_dir = r'I:\Logistic Dept\Forwarder Performance Report\ST Team\Air Freight\2. RFQ\RFQ Summary'
    input_SA_dir = r'I:\Logistic Dept\Forwarder Performance Report\ST Team\Air Freight\2. RFQ\RFQ Original\2023\Tender Rate request'
    output_dir = r'I:\Logistic Dept\Forwarder Performance Report\ST Team\Air Freight\2. RFQ\RFQ Summary'
    
    
    #Read Send out Summary
    os.chdir(input_summary_dir)
    sum_df = pd.read_excel('AIR RFQ.xlsm',sheet_name='RFQ summary Tender Rate')
    
    sum_df['RFQ no.'] = sum_df['RFQ no.'].str.upper()
    
    #display(sum_df)
    
    #Read Tender Rate RFQ
    os.chdir(input_SA_dir)
    
    xlsm_files = glob.glob(f'{input_SA_dir}\**\*.xlsm', recursive=True)
    
    xlsx_files = glob.glob(f'{input_SA_dir}\**\*.xlsx', recursive=True)
    
    xls_files = glob.glob(f'{input_SA_dir}\**\*.xls', recursive=True)
    
    dir_list = [xlsm_files, xlsx_files, xls_files]
    
    frames = []
    
    col_dic = {0:'Forwarder Identity',1:'Total Amount',2:'Service',3:'Transit Time in working days'}
    
    for dir_ in dir_list:
    
        for file in dir_:
            try:
                #print(os.path.basename(file))
                df = pd.read_excel(file,sheet_name='RFQ FORM',header=73,nrows=4,usecols=[6])
                
                print(os.path.basename(file))
                
                rfq_nbr = re.split(r' ',os.path.basename(file))[0]
                
                print(rfq_nbr)
                
                #matchObj = re.search(r'\d{4}-\d{2}-\d{2}_\d{2}_\d{2}_\d{2}',os.path.basename(file)) #Search Date string in file name

                #email_receive_date = matchObj.group()
                
                modification_date = dt.fromtimestamp(os.path.getmtime(os.path.basename(file)))
                
                #print(modification_date)
                
                creation_date = dt.fromtimestamp(os.path.getctime(os.path.basename(file)))
                
                #print(creation_date)

                df = df.T

                df.rename(columns=col_dic,inplace=True)

                df = df.reset_index(drop=True)

                df['RFQ no.'] = rfq_nbr
                
                df['Quotation Receive Date'] = modification_date
                
                df['First Quotation Receive Date'] = creation_date
                
                
                frames.append(df)

            except Exception as e:
                print(e)
                print(f'Cannot access RFQ Tender Rate data infomation in: {os.path.basename(file)}')
                pass
        
        sa_df = pd.concat(frames)
        
    sa_df = sa_df.append(sa_df)
    
    sa_df = sa_df.drop_duplicates()
    
    sa_df['RFQ no.'] = sa_df['RFQ no.'].str.upper()

    sa_df.sort_values(by=['RFQ no.','Forwarder Identity'],inplace=True)

    sa_df = sa_df.drop_duplicates(subset=['RFQ no.','Forwarder Identity'],keep='last')
    
    #print(sa_df.columns)
    
    #Left Join Send-out Summary & sa_df
    
    r_df = pd.merge(sum_df,sa_df,how='left',on=['RFQ no.'])
    
    os.chdir(output_dir)
    r_df.to_excel('AIR RFQ Tendor Rate Matched_test.xlsx',index=False)
    
    return(r_df)

if __name__=='__main__':
    process_RFQ_email()
    RFQ_Summary()
    RFQ_Summary_Tender_Rate()
    print('Report complete.')
