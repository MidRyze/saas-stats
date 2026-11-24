# -*- coding: utf-8 -*-
import os
import io
from io import StringIO
import calendar
import datetime
from dateutil.relativedelta import relativedelta
import requests
import jwt
import csv
import gzip
import shutil
import time
import streamlit as st
import numpy as np
import pandas as pd
import openpyxl
import seaborn
import gspread
import gspread.utils
from google.cloud import storage
from gspread_formatting import CellFormat, format_cell_range, numberFormat
import matplotlib
import plotly.express as px
import plotly.graph_objects as go
from streamlit_gsheets import GSheetsConnection
##--##-----------FUNCTIONS---------##--##
##-----------!! TIME CALC !!-----------##
# Calculate time #
def check_last_day(month:int, year='', check=False)->int:
    """
    Returns Final day of the targeted month
    """
    if month==2 and year=='':
        _today=datetime.datetime.today()
        year=datetime.datetime.strftime(_today,'%Y')
        year=int(year)
    elif month==2:
        # year=datetime.datetime.strptime(year,'%Y').strftime('%Y')
        year=int(year)
    if month==2:
        if year%4==0 and year%100!=0:
            last_day_of_feb=29
        elif year%4==0 and year%100==0 and year%400==0:
            last_day_of_feb=29
        else:
            last_day_of_feb=28
        return last_day_of_feb
    last_day_of_month={
        1:31,
        3:31,
        4:30,
        5:31,
        6:30,
        7:31,
        8:31,
        9:30,
        10:31,
        11:30,
        12:31,
    }
    return_values=last_day_of_month[month]
    return return_values
# Calculate date and or modify format #
def modify_dates(date:str,date_format:str,change:dict|list|int='',output_format='',
                    check=False)->dict:
    """
    Returns updated date and or format
    [Header Field]:[Value]
    Args:
    date: Enter date, make sure the format follows this\n
    date_format: the order it's written (eg: YYYYMMDD)\n
    change: add values to date (eg:1 day 1 month = {d:1,m:1}, if there's no letter, it will follow the format order)\n
    output_format: Change the output format (Default: Returns as original format)\n
    """
    allowed_char=set('%#dmy')
    date_format=date_format.lower()
    if output_format.lower()==date_format or output_format=='':
        output_format_use=date_format
        status=f'convert_to_dates(check_update_date) - Done, no format change ✅'
    else:
        output_format_use=output_format
        status=f'convert_to_dates(check_update_date) - Done ✅'
    date_format=''.join(each for each in date_format if each in allowed_char)
    output_format_use=''.join(each for each in output_format_use if each in allowed_char)
    
    if check:print(f'~~~{['STOP']}');return ## ------------------ CHECK ------------------ ##

    if isinstance(date, int):
        date=str(date)
    if len(date)!=len(date_format) or isinstance(change,(dict, list)) and len(change)>3:
        return_values={
            'status':f'convert_to_dates(check_update_date) - LENGTH ERR: Length error',
            'response':2,
            'months_to_update':None,
            }
        return return_values
    day_splice_start=date_format.find('d')
    day_length=date_format.count('d')
    day_splice_end=day_length+day_splice_start
    month_splice_start=date_format.find('m')
    month_length=date_format.count('m')
    month_splice_end=month_length+month_splice_start
    year_splice_start=date_format.find('y')
    year_length=date_format.count('y')
    year_splice_end=year_length+year_splice_start

    start_day=int(date[day_splice_start:day_splice_end])
    start_month=int(date[month_splice_start:month_splice_end])
    start_year=int(date[year_splice_start:year_splice_end])

    # APPLY CHANGES TO DATE (IF ANY)
    start_date=datetime.datetime(start_year,start_month,start_day)
    add_day=0
    add_month=0
    add_year=0
    ordering={
        day_splice_start:[day_length, add_day, 'day'],
        month_splice_start:[month_length, add_month, 'month'],
        year_splice_start:[year_length, add_year, 'year'],
            }
    if isinstance(change, dict):
        keys=change.keys()
        if 'd' in keys:
            add_day=int(int(change['d']))
        if 'm' in keys:
            add_month=int(int(change['m']))
        if 'y' in keys:
            add_year=int(int(change['y']))
    elif isinstance(change, list):
        for i, each in enumerate(change):
            if i==0:
                ordering[i][1]=ordering[i][1]+each
                next_element=i+ordering[i][0]
            else:
                ordering[next_element][1]=ordering[next_element][1]+each
                next_element=next_element+ordering[next_element][0]
        if ordering[day_splice_start][1]!=0:
            add_day=ordering[day_splice_start][1]
        if ordering[month_splice_start][1]!=0:
            add_month=ordering[month_splice_start][1]
        if ordering[year_splice_start][1]!=0:
            add_year=ordering[year_splice_start][1]
    elif isinstance(change, (str,int)) and change != '':
        change=int(change)
        if ordering[0][2]=='day':
            add_day=change
        elif ordering[0][2]=='month':
            add_month=change
        else:
            add_year=change
    else:
        pass
    new_date=start_date+relativedelta(years=add_year,months=add_month,days=add_day)
    new_day=new_date.day
    new_month=new_date.month
    new_year=new_date.year

    output_day_length=output_format_use.count('d')
    output_month_length=output_format_use.count('m')
    output_year_length=output_format_use.count('y')
    output_day_start=output_format_use.find('d')
    output_month_start=output_format_use.find('m')
    output_year_start=output_format_use.find('y')
    output_listed_order=sorted([output_day_start,
                                output_month_start,
                                output_year_start])
    if output_day_length>1 or '%d' in output_format:
        output_day_format='%d'
    elif output_day_length==1 or '%#d' in output_format:
        output_day_format='%#d' # WIN based
        # output_day_format='%-d' # UNIX based
    else:
        output_day_format=''

    if output_month_length>1 or '%m' in output_format:
        output_month_format='%m'
    elif output_month_length==1 or '%m' in output_format:
        output_month_format='%#m' # WIN based
        # output_month_format='%-m' # UNIX based
    else:
        output_month_format=''

    if output_year_length>2 or '%Y' in output_format:
        output_year_format='%Y'
    elif output_year_length>0 or '%y' in output_format:
        output_year_format='%y'
    else:
        output_year_format=''

    formatted_date={
        output_day_start:output_day_format,
        output_month_start:output_month_format,
        output_year_start:output_year_format,
    }

    output_format=f'{formatted_date[output_listed_order[0]]}{formatted_date[output_listed_order[1]]}{formatted_date[output_listed_order[2]]}'
    output_date_str=datetime.datetime(new_year,new_month,new_day).strftime(output_format)
    output_date_dtime=datetime.datetime.strptime(output_date_str,output_format)

    return_values={
                'status':status,
                'response':0,
                'content':[output_date_str, output_date_dtime],
                }
    return return_values

# Calculate time #
def check_update_date(latest_date, delay=2, target_day=25, start_day=1, 
                        overwrite=[], retrieve_eval=False, check=False):
    """
    Returns Status if time is valid for updating GSheet
    [Header Field]:[Value]
    Args:
    latest_date:
    delay: How many month(s) to check for latest report\n
    target_day: The targeted day to check, 28 is the latest to account for February\n
    overwrite: Set custom date (DDMMYYYY)\n
    """
    def convert_to_dates(latest_date,current_date='',check=False): # FORMAT DDMMYYYY
        if current_date=='':
            current_date=0
        try:
            int(current_date)
        except Exception as e:
            return_values={
                        'status':f'convert_to_dates(check_update_date) - FORMAT ERR: {e}',
                        'response':2,
                        'months_to_update':None,
                        }
            return return_values
        late_month=int(latest_date[2:4])
        if check:print(f'{[latest_date]}');return ## ------------------ CHECK ------------------ ##
        year=int(latest_date[4:])
        target_month=late_month+1

        if target_month>12:
            target_month=1
            year+=1
        year=str(year)
        target_date=year+'{:02d}'.format(target_month)

        try:
            return_values={
                        'curr_day':int(current_date[:2]),
                        'curr_month':int(current_date[2:4]),
                        'curr_year':int(current_date[4:]),
                        'late_month':late_month,
                        'target_date':target_date,
                        'late_year':int(latest_date[4:]),
                        }
            return return_values
        except:
            return_values={
                        'late_month':late_month,
                        'target_date':target_date,
                        'late_year':int(latest_date[4:]),
                        }
            return return_values
        # return late_month, late_year, curr_day, curr_month, curr_year
    
    if retrieve_eval:
        retrieve_eval_date=convert_to_dates(latest_date,
                                            # check=True,
                                            )
        print(retrieve_eval_date)
        retrieve_eval={
            # delay,
            # target_day,
            'delay':delay,
            'target_day':target_day,
            'late_month':retrieve_eval_date['late_month'],
            'target_date':retrieve_eval_date['target_date'],
            'late_year':retrieve_eval_date['late_year'],
        }
        return retrieve_eval

    target_day=int(target_day)
    if target_day>28:
        target_day=28

    if len(overwrite)!=0:
        latest_date=str(overwrite[0])
        if len(latest_date)!=8:
            return_values={
                'status':f'check_update_date - LENGTH ERR: overwrite_latest_date-Length error',
                'response':1,
                'months_to_update':None,
                }
            return return_values
        try:
            current_date=str(overwrite[1])
            datetime.datetime.strptime(current_date,'%d%m%Y')
        except:
            current_date=datetime.date.today()
            current_date=current_date.strftime('%d%m%Y')
    else:
        current_date=datetime.date.today()
        current_date=current_date.strftime('%d%m%Y')
        latest_date=latest_date.replace(',','')
        latest_date=datetime.datetime.strptime(latest_date,'%b %d %Y')
        latest_date=latest_date.strftime('%d%m%Y')

    var_modify_dates_current=modify_dates(current_date,'ddmmyyyy')
    var_modify_dates_target=modify_dates(latest_date,'ddmmyyyy',{'m':1},'yyyymmdd')
    var_modify_dates_latest=modify_dates(latest_date,'ddmmyyyy')
    late_month=int(var_modify_dates_latest['content'][0][2:4])
    late_year=int(var_modify_dates_latest['content'][0][4:])
    target_date=var_modify_dates_target['content'][0]
    curr_day=int(var_modify_dates_current['content'][0][:2])
    curr_month=int(var_modify_dates_current['content'][0][2:4])
    curr_year=int(var_modify_dates_current['content'][0][4:])

    # var_convert_to_dates=convert_to_dates(latest_date,current_date)
    # late_month=var_convert_to_dates['late_month']
    # late_year=var_convert_to_dates['late_year']
    # target_date=var_convert_to_dates['target_date']+'{:02d}'.format(start_day)
    # curr_day=var_convert_to_dates['curr_day']
    # curr_month=var_convert_to_dates['curr_month']
    # curr_year=var_convert_to_dates['curr_year']
    # late_month, late_year, curr_day, curr_month, curr_year=convert_to_dates(latest_date,current_date)
    # PRETTY SURE THE FUNCTION CAN JUST BE ONE HUGE CODE HERE WITHOUT REPEATING
    if late_year>curr_year or (late_year==curr_year and late_month>curr_month):
        latest_date=current_date
        late_month=int(var_modify_dates_latest['content'][0][2:4])
        late_year=int(var_modify_dates_latest['content'][0][4:])
        # late_month=var_convert_to_dates['late_month']
        # late_year=var_convert_to_dates['late_year']
        # target_date=var_convert_to_dates['target_date']+'{:02d}'.format(start_day)
        # curr_day=var_convert_to_dates['curr_day']
        # curr_month=var_convert_to_dates['curr_month']
        # curr_year=var_convert_to_dates['curr_year']

    try:
        datetime.datetime.strptime(latest_date,'%d%m%Y')
    except Exception as e:
        return_values={
                        'status':f'FORMAT ERR: {e}',
                        'response':2,
                        'months_to_update':None,
                        }
        return return_values
    
    year_difference=curr_year-late_year
    curr_month+=year_difference*12
    month_diff=curr_month-late_month
    months_to_update=month_diff-delay

    if month_diff>delay or month_diff==delay and curr_day >= target_day:
        # update_gsheet(new_data,target_df)
        status=f'UPDATE: {month_diff} month(s) total difference. {months_to_update} month(s) to update.'
    else:
        return_values={
                        'status':f'NO UPDATE',
                        'response':3,
                        'months_to_update':None,
                        }
        return return_values

    if check:print(f'{[type(target_date_),type(target_date)]}');return ## ------------------ CHECK ------------------ ##
    return_values={
                    'status':status,
                    'response':0,
                    'months_to_update':months_to_update,
                    'latest_date':latest_date,
                    'target_date':target_date,
                    'current_date':current_date,
                    'late_month':late_month,
                    'late_year':late_year,
                    'curr_day':curr_day,
                    'curr_month':curr_month,
                    'curr_year':curr_year,
                    }
    return return_values

##-------------!! FETCH !!-------------##
#-- APPLE --#
# Apple KEYS #
def apple_dt(check=False, ipynb=False, **kwargs)->dict:
    """
    Returns Keys
    """
    if ipynb==False:
        api_values=st.secrets["apple_data"]
    else:
        kwargs=kwargs['kwargs']
        api_values={each:kwargs[each] for each in kwargs}

    API_APPLE={
        'ISSUER_ID':api_values['issuer_id'],
        'KEY_ID':api_values['key_id'],
        'PRIVATE_KEY':api_values['private_key'],
        'VENDOR_NUM':api_values['vendor_id'],
        'URL':api_values['url'],
    }
    if check:print(f'{[API_APPLE['URL']]}');return ## ------------------ CHECK ------------------ ##
    return API_APPLE
# Apple Listing #
def apple_listing(check=False)->dict:
    """
    Returns list of Apps, the Dirs & Breakdowns
    """
    var_app_list=[ # TEMP
                'MasjidProMY',      #00
                'USIM MasjidPro',   #01
                'MISPRO Sarawak',   #02
                'Islah Junior',     #03
                'MWtopup',          #04
                'DIDIKcomm',        #05
                ]
    var_app_listing={i:each for i,each in enumerate(var_app_list)}

    return_values={
        'apps':var_app_listing,
        'type_identifier':{     # DESCRIPTION:
            '1'    :'1',        # Free or paid app - iOS, iPadOS, visionOS, watchOS
            '1b'   :'1-B',      # App Bundle - iOS, iPadOS, visionOS app bundle
            'f1b'  :'F1-B',     # App Bundle - Mac app bundle
            '1e'   :'1E',       # Paid app - Custom iOS app
            '1ep'  :'1EP',      # Paid app - Custom iPadOS app
            '1eu'  :'1EU',      # Paid app - Custom universal app
            '1f'   :'1F',       # Free or paid app - Universal app, excluding tvOS
            '1t'   :'1T',       # Free or paid app - iPad apps
            '3'    :'3',        # Re-download - App update (iOS, tvOS, visionOS, watchOS)
            '3f'   :'3F',       # Re-download - Universal app, excluding tvOS
            '7'    :'7',        # Update - App update (iOS, tvOS, visionOS, watchOS)
            '7f'   :'7F',       # Update - Universal app, excluding tvOS
            '7t'   :'7T',       # Update - App update (iPadOS, visionOS)
            'f1'   :'F1',       # Free or paid app - Mac
            'f7'   :'F7',       # Update - App update (Mac)
            'fi1'  :'FI1',      # In-App Purchase - Mac
            'ia1'  :'IA1',      # In-App Purchase - In-App Purchase (iOS, iPadOS, visionOS)
            'ia1m' :'IA1-M',    # In-App Purchase - In-App Purchase (Mac)
            'ia3'  :'IA3',      # Restored In-App Purchase - Non consumable In-App Purchase
            'ia9'  :'IA9',      # In-App Purchase - Non-renewing subscription (iOS, iPadOS, visionOS)
            'ia9m' :'IA9-M',    # In-App Purchase - Subscription (Mac)
            'iay'  :'IAY',      # In-App Purchase - Auto-renewable subscription (iOS, iPadOS, visionOS)
            'iaym' :'IAY-M',    # In-App Purchase
            },
        'breakdown':{                                   # COLUMN BREAKDOWN:
            'title'        :'Title',                    # App name
            'units'        :'Units',                    # Number of ...
            'begin date'   :'Begin Date',               # Start date
            'end date'     :'End Date',                 # End date
            'product type' :'Product Type Identifier',  # 
            },
        }
    if check:print(f'{[]}');return ## ------------------ CHECK ------------------ ##
    return return_values
# Apple Request #
def apple_auth(ISSUER_ID='',KEY_ID='',PRIVATE_KEY='',VENDOR_NUM='',URL='',
               timer=1140,check=False)->dict:
    """
    Returns Keys & Response
    """        
    if type(timer)!=int:
        print(f"ERR: timer not int type")
        return 1
    if ISSUER_ID=='' or KEY_ID=='' or PRIVATE_KEY=='' or URL=='':
        var_API_APPLE= apple_dt()
        ISSUER_ID=var_API_APPLE['ISSUER_ID']
        KEY_ID=var_API_APPLE['KEY_ID']
        PRIVATE_KEY=var_API_APPLE['PRIVATE_KEY']
        VENDOR_NUM=var_API_APPLE['VENDOR_NUM']
        URL=var_API_APPLE['URL']
    header_token={
        "alg": "ES256",     # Encryption Algorithm -> All JWTs for App Store Connect API must be signed with ES256 encryption.
        "kid": KEY_ID,      # Key Identifier -> for example, 2X9R4HXF34
        "typ": "JWT",       # Token Type
    }
    payload={
        'iss':ISSUER_ID,
        'iat':int(time.time()),
        'exp':int(time.time())+timer,
        'aud':'appstoreconnect-v1',
    }
    token=jwt.encode(payload, PRIVATE_KEY, algorithm=header_token['alg'], headers=header_token) # CONNECTS APPLE SERVICES VIA API
    headers_response={'Authorization':f'Bearer {token}',}
    response=requests.get(URL,headers=headers_response)
    if check:print(f'{[VENDOR_NUM]}');return ## ------------------ CHECK ------------------ ##
    if response.status_code==200:
        apps_data=response.json().get('data',[])
        response=0
    else:
        response=response.json()['errors'][0]
        apps_data=''
        print(f"{response.status_code}, {response['code']}\n{response['title']} {response['detail']}")
        response=1

    return_values={
        'headers_response':headers_response,
        'response':response,
        'apps_data':apps_data,
        'ISSUER_ID':ISSUER_ID,
        'KEY_ID':KEY_ID,
        'PRIVATE_KEY':PRIVATE_KEY,
        'VENDOR_NUM':VENDOR_NUM,
        'URL':URL,
    }
    return return_values
# Retrive Apple Data #
def apple_df(start_date:str='',frequency_type:str='YEARLY',
             filter_req:dict={},failover=False,
             terminal_dump=False,ipynb_overwrite=False,check=False)->dict:
    """
    Returns Install Sum & The DF
    [Header Field]:[Value]
    Args:
    start_date: Date format 'YYYYMMDD'
    frequency_type: 'YEARLY'/'MONTHLY'/'~~WEEKLY~~'/'DAILY'
    filter_req: Example {
            'Title':'MasjidProMY',
            'Product Type Identifier':'1F',
            }
    """
    end_date_var='End Date'
    return_values={
            'status':None,
            'message':None,
            'sum':None,
            'entire_df':None,
            # 'start_date':start_date, #DEPRECIATED
            # 'end_date':end_date, #DEPRECIATED
            # 'frequency_failover':frequency_failover, #DEPRECIATED
            }
    def modify_df_columns(df,target_columns=[],mode='keep',check=False):
        '''
        mode = keep | drop
        '''
        target_columns=[str.lower(i) for i in target_columns]
        columns={str.lower(each_col): each_col for each_col in df.columns}
        if check:print(f'{[]}');return ## ------------------ CHECK ------------------ ##
        if mode=='drop':
            for each in target_columns:
                try:
                    df.drop(columns=columns[each],inplace=True)
                except:
                    print(f'Column {each} not found ⚠️')
                    return 1
            return df
        else:
            for each in target_columns:
                try:
                    df=df[[columns[each] for each in target_columns]]
                except:
                    print(f'Column {each} not found ⚠️')
                    return 1
            return df
    def terminal_print(csv_data,frequency_type):
        # with gzip.open(io.BytesIO(response_data.content),mode='rt',encoding='utf-8',errors='ignore') as f:
        #     csv_data=csv.DictReader(f,delimiter='\t')
        print(f"{'App Name':<30} | {'Country':<25} | {'Units':<10} | {'Type':<10} | {'Timeframe':<10}")
        print("-" * 115)
        for row in csv_data:
            print(row)
            # col1 = row.get('Title', 'Unknown')
            # col2 = row.get('Country Code', 'Unknown')
            # col3 = row.get('Units', '0')
            # col4 = row.get('Product Type Identifier', 'Unknown') #-1/1F DOWNLOADS, 3/3F REDOWNLOADS, 7/7F UPDATES
            # col5 = row.get('Begin Date', 'Unknown')
            # print(f"{col1:<30} | {col2:<25} | {col3:<10} | {col4:<10} | {col5+' '+str.capitalize(frequency_type):<10}")
                # if col1==filter_app: #and col4=='1F' or col4=='3F':
                #     print(f"{col1:<30} | {col2:<25} | {col3:<10} | {col4:<10} | {col5+' '+str.capitalize(frequency_type):<10}")
                # else:
                #     print(f"{col1:<30} | {col2:<25} | {col3:<10} | {col4:<10} | {col5+' '+str.capitalize(frequency_type):<10}")

    var_apple_listing=apple_listing()
    if not isinstance(ipynb_overwrite, dict):
        var_apple_auth_data=apple_auth()
    else:
        ISSUER_ID=ipynb_overwrite['ISSUER_ID']
        KEY_ID=ipynb_overwrite['KEY_ID']
        PRIVATE_KEY=ipynb_overwrite['PRIVATE_KEY']
        VENDOR_NUM=ipynb_overwrite['VENDOR_NUM']
        URL=ipynb_overwrite['URL']
        var_apple_auth_data=apple_auth(
            ISSUER_ID=ISSUER_ID,
            KEY_ID=KEY_ID,
            PRIVATE_KEY=PRIVATE_KEY,
            VENDOR_NUM=VENDOR_NUM,
            URL=URL,
            )
    apps_data=var_apple_auth_data['apps_data']
    headers_response=var_apple_auth_data['headers_response']
    VENDOR_NUM=var_apple_auth_data['VENDOR_NUM']

    if start_date=='':
        start_date=datetime.datetime.today().strftime('%Y-%m-%d')
        month=int(start_date[5:7])-1
        year=int(start_date[:4])
        if month==-1:
            # month=str(12)
            # year=str(int(start_date[:4])-1)
            month=12
            year=int(start_date[:4])-1
        month='{:02d}'.format(month)
        year='{:0004d}'.format(year)
        start_date=f'{year}-{month}-{start_date[8:]}'
    else:
        day=start_date[6:]
        month=start_date[4:6]
        year=start_date[:4]
        start_date=f'{year}-{month}-{day}'
    if frequency_type=='YEARLY':
        start_date_adjusted=start_date[:4]
    elif frequency_type=='MONTHLY':
        start_date_adjusted=start_date[:7]
    
    os.system('cls')
    print('START EXTRACTING...')
    apps_list={}
    for apps in apps_data:
        sku = apps['attributes']['sku']
        apps_list[sku] = {
            "app_id": apps['id'],
            "name": apps['attributes']['name'],
            "bundle_id": apps['attributes']['bundleId'],
            "sku": sku,
            "total_downloads": 0
        }
        # print(apps_list)
    report_url="https://api.appstoreconnect.apple.com/v1/salesReports"
    params={
        "filter[frequency]":frequency_type,
        "filter[reportDate]":start_date_adjusted, # YYYY-MM-DD, Not required, if frequency = daily
        "filter[reportSubType]":'SUMMARY',
        "filter[reportType]":'SALES',
        "filter[vendorNumber]":f'{VENDOR_NUM}',
        # "filter[version]":'1_0', # Not required
        }
    
    report_response=requests.get(report_url,headers=headers_response,params=params)
    
    if report_response.status_code==410 and failover:
        start_date_adjusted=start_date[:4]
        params["filter[frequency]"]='YEARLY'
        params["filter[reportDate]"]=start_date_adjusted
        report_response=requests.get(report_url,headers=headers_response,params=params)
        return_values['status']=1
        return_values['frequency_failover']=True # TO DEPRECIATE
        return_values['start_date']=start_date
    elif report_response.status_code!=200:
        print(report_response.text)
        return_values['status']=f'{report_response.text}, END ⚠️'
        return return_values
    
    with gzip.open(io.BytesIO(report_response.content),mode='rt',encoding='utf-8',errors='ignore') as f:
        target_ap_df=pd.read_csv(f, sep='\t')
        
    filter_ap_df=target_ap_df
    filter_ap_df=modify_df_columns(filter_ap_df,[
        'Title',
        'Units',
        'Begin Date',
        'End Date',
        'Product Type Identifier',
    ],mode='keep')
    for keys in filter_req:
        # print(target_ap_df[])
        filter_ap_df=filter_ap_df[filter_ap_df[keys]==filter_req[keys]]

    if terminal_dump:
        terminal_print(filter_ap_df,frequency_type)

    # Sum 'Units' column
    columns={i: each_col for i,each_col in enumerate(filter_ap_df.columns)}
    ios_sum_install=sum(filter_ap_df[columns[1]].tolist()) #🟢🔻 to implement
    # try:
    #     end_date=filter_ap_df[end_date_var].iloc[-1]
    #     end_date=end_date.replace('/','')
    #     end_date=end_date[4:]+end_date[:2]+end_date[2:4]
    # except:
    #     end_date=f'\'{end_date_var}\' unretrievable, check pulled table'
    return_values['sum']=ios_sum_install
    return_values['entire_df']=filter_ap_df
    if check:print(f'=={[]}');return ## ------------------ CHECK ------------------ ##

    if return_values['status'] is None:
        return_values['status']=0
        return_values['message']=f'Successful, END 🟢',
    else:
        end_date=filter_ap_df[end_date_var].iloc[-1]
        end_date=end_date.replace('/','')
        end_date=end_date[4:]+end_date[:2]+end_date[2:4]
        return_values['message']=f'Unretrievable, frequency_type to \'YEARLY\'',
        return_values['end_date']=end_date
        
    return return_values

#-- GOOGLE --#
# Google KEYS #
def google_dt(check=False, ipynb=False, **kwargs)->vars:
    """
    Returns Keys & Bucket
    """
    if ipynb==False:
        api_values=st.secrets["google_data"]
        bucket_value=st.secrets["google_bucket"]["bucket_name"]
    else:
        kwargs=kwargs['kwargs']
        api_values={each:kwargs[each] for each in kwargs}
        bucket_value=api_values['bucket_name']

    API_GOOGLE={
        'type':api_values['type'],
        'project_id':api_values['project_id'],
        'private_key_id':api_values['private_key_id'],
        'private_key':api_values['private_key'],
        'client_email':api_values['client_email'],
        'client_id':api_values['client_id'],
        'auth_uri':api_values['auth_uri'],
        'token_uri':api_values['token_uri'],
        'auth_provider_x509_cert_url':api_values['auth_provider_x509_cert_url'],
        'client_x509_cert_url':api_values['client_x509_cert_url'],
        'universe_domain':api_values['universe_domain'],
    }
    GOOGLE_BUCKET=bucket_value
    if check:print(f'{API_GOOGLE,GOOGLE_BUCKET}');return ## ------------------ CHECK ------------------ ##
    return API_GOOGLE,GOOGLE_BUCKET
# Goggle Listing #
def google_listing(check=False)->dict:
    """
    Returns list of Apps, the Dirs & Breakdowns
    """
    var_app_list=[ # TEMP
        'com.hpcs.school_bus_pro',                  #00
        'com.otoreport.app6stopup',                 #01
        'com.otoreport.denttopup',                  #02
        'com.otoreport.kitakedai',                  #03
        'com.otoreport.magicpay',                   #04
        'com.otoreport.mobil',                      #05
        'com.otoreport.mynzpay',                    #06
        'com.otoreport.onesolutioncentre',          #07
        'com.otoreport.reloadjean',                 #08
        'com.otoreport.smartpay',                   #09
        'com.otoreport.topupbsg',                   #10
        'com.otoreport.topupcreative',              #11
        'com.otoreport.topuphrz',                   #12
        'com.otoreport.topuporange',                #13
        'com.otoreport.topuppro',                   #14
        'com.otoreport.topuprz',                    #15
        'com.otoreport.topupsaiful',                #16
        'com.otoreport.xwormpay',                   #17
        'com.otoreport.zareload',                   #18
        'intesh.smartool',                          #19
        'my.fizah.smart_mengaji',                   #20
        'my.hpcs.didikcomm',                        #21
        'my.hpcs.ecommerce_shoes',                  #22
        'my.hpcs.islamic_tourism',                  #23
        'my.hpcs.masjidpro.masjidpro',              #24
        'my.hpcs.mftripapps',                       #25
        'my.hpcs.mispro',                           #26
        'my.hpcs.pibgpro',                          #27
        'my.hpcs.usim_masjidpro.usim_masjidpro',    #28
        'my.realreload.topup',                      #29
        ]
    var_app_listing={i:each for i,each in enumerate(var_app_list)}

    return_values={
        'apps':var_app_listing,
        'dir':{                                                          # DESCRIPTION:
            'buyers_7d'           :'acquisition/buyers_7d/',             # 7-day buyer data
            'retained_installers' :'acquisition/retained_installers/',   # 
            'reviews'             :'reviews/',                           # reviews
            'crashes'             :'stats/crashes/',                     # crashes
            'installs'            :'stats/installs/',                    # installs
            'ratings'             :'stats/ratings/',                     # ratings
            'ratings_v2'          :'stats/ratings_v2/',                  # 
            'store_performance'   :'stats/store_performance/',           # store performance
            },
        'breakdown':{                           # BREAKDOWN BY:
            'app_version' :'app_version.csv',   # App version
            'carrier'     :'carrier.csv',       # Carrier
            'country'     :'country.csv',       # Country/region
            'device'      :'device.csv',        # Device
            'language'    :'language.csv',      # Language
            'os_version'  :'os_version.csv',    # Android version
            'overview'    :'overview.csv',      # Overview
            },
        }
    if check:print(f'{[]}');return ## ------------------ CHECK ------------------ ##
    return return_values
# Goggle Request #
def google_auth(GOOGLE_KEYS:dict='', GOOGLE_BUCKET:str='', check=False)->dict:
    """
    Returns Bucket & Listings (Blobs)
    """  
    os.system('cls')
    if check:print(f'{['test']}');return ## ------------------ CHECK ------------------ ##
    print('ESTABLISHING CONNECTION...')
    # print(type(GOOGLE_KEYS))
    if isinstance(GOOGLE_KEYS, dict) and GOOGLE_BUCKET!='':
        var_API_GOOGLE=GOOGLE_KEYS
        bucket_name=GOOGLE_BUCKET
    else:
        var_API_GOOGLE,bucket_name=google_dt()
    
    storage_client = storage.Client.from_service_account_info(var_API_GOOGLE) # CONNECTION STARTS
    bucket = storage_client.bucket(bucket_name)
    var_listing=google_listing()
    install_dir=var_listing['dir']['installs']
    csv_listing = bucket.list_blobs(prefix=install_dir)

    return_values={
                    'bucket':bucket,
                    'csv_listing':csv_listing,
                    }
    return return_values
# Retrive Goggle Data #
def google_df(target_app:int, target_month:str='', ipynb_overwrite=False, check=False, stop_run=int(0))->dict:
    #google_df(bucket, target_app:int, target_month:int='', check=False, stop_run=int(0))->dict: OLD ARGU
    """
    Returns Install Sum & The DF
    [Header Field]:[Value]
    Args:
    ~~bucket: From google_auth()~~
    target_app: Following App Listing (int)
    target_month: YYYYMM (str)
    """
    summary=None
    status=None
    var_listing=google_listing()
    # var_google_auth=google_auth()
    if not isinstance(ipynb_overwrite, (dict, tuple)):
        var_google_auth=google_auth()
        bucket=var_google_auth['bucket']
    else:
        GOOGLE_KEYS=ipynb_overwrite[0]
        GOOGLE_BUCKET=ipynb_overwrite[1]
        var_google_auth=google_auth(
            GOOGLE_KEYS,
            GOOGLE_BUCKET,
            )
        bucket=var_google_auth['bucket']
    try:
        target_app=int(target_app)
    except:
        status=1
        return_values={
                    'status':status,
                    'message':'google_df - ERR, TYPE',
                    'sum':None,
                    'entire_df':None,
                    'summary_df':summary,
                    }
        return return_values
    os.system('cls')
    # print('EXECUTE RETRIEVAL')
    app=var_listing['apps'][target_app]
    breakdown=var_listing['breakdown']['overview']
    
    if target_month=='':
        latest_date=datetime.datetime.today()
        year=latest_date.year
        latest_month=latest_date.month
        # _day,var_delay_month=check_update_date(latest_date,retrieve_eval=True)
        # target_month=latest_month-var_delay_month
        var_delay_month=check_update_date(latest_date,retrieve_eval=True)
        var_delay_month=var_delay_month['delay']
        target_month=latest_month-var_delay_month
        if target_month<0:
            target_month=target_month+12
            year=year-1
        target_month=str(year)+'{:02d}'.format(target_month)
    else:
        target_month=target_month[:6]
            

    print('START EXTRACTING...')
    try:
        file=bucket.blob(f"stats/installs/installs_{app}_{target_month}_{breakdown}")
        data=StringIO(file.download_as_bytes().decode("utf-16"))
        go_df=pd.read_csv(data, sep=',')
        columns=go_df.columns.tolist()
        go_df=go_df.drop(columns=columns[3:])
        go_df=go_df.set_index(columns[0])
        android_sum_install=sum(go_df[columns[2]].tolist())
        summary=f'Sum: {android_sum_install}, DONE ✅'
        status=0
    except:
        go_df=pd.DataFrame()
        android_sum_install=0
        summary=f'Record not found, DONE ✅'
        status=2
    if check:print(f'--{[go_df]}');return ## ------------------ CHECK ------------------ ##
    # print('------>','return_values');return

    # os.system('cls')
    return_values={
                    'status':status,
                    'message':'google_df process complete',
                    'sum':android_sum_install,
                    'entire_df':go_df,
                    # 'summary_df':summary,
                    }
    return return_values
    ''' OLD CODE
    # #ORI STRUCTURE-> google_df(bucket,target_app,csv_listing='',target_month='',check=False,stop_run=int(0)):
    # df_list=[]
    # counter=1
    # for file in csv_listing:
    #     if 'overview' in file.name and target_app in file.name:
    #         data=StringIO(file.download_as_bytes().decode("utf-16"))
    #         df=pd.read_csv(data, sep=',')
    #         df_list.append(df)
    #         if stop_run and counter==stop_run:
    #             break
    #         counter+=1
    # if df_list:
    #     entire_df=pd.concat(df_list,ignore_index=True)
    #     columns=entire_df.columns.tolist()
    #     entire_df=entire_df.drop(columns=columns[3:])
    #     entire_df=entire_df.set_index(columns[0])
    #     os.system('cls')
    #     print(entire_df,'DONE ✅')
    #     print(df.iloc[-1])
    # else:
    #     entire_df=''
    #     print('No data available. DONE ✅')
    # return entire_df'''

#-- GSHEET --#
# GSheet KEYS #
def gsheet_connection(check=False)->str:
    """
    Returns Install Sum & The DF
    [Header Field]:[Value]
    Args:
    ~~bucket: From google_auth()~~
    target_app: Following App Listing (int)
    target_month: MM (int)
    """
    # st_conn=st.connection("gsheet_data", type=GSheetsConnection)
    gsheet_cred=gspread.service_account_from_dict(st.secrets["connections"]["gsheet_data"])
    url=str(st.secrets["connections"]["gsheet_data"]['spreadsheet'])
    prefix='/d/'
    suffix='/edit'
    target_workbook=url[url.find(prefix)+len(prefix):url.find(suffix)]

    if check:print(f'{[]}');return ## ------------------ CHECK ------------------ ##
    # Establish Connections #
    st_conn=gsheet_cred.open_by_key(target_workbook) # CONNECTS GOOGLE SERVICES VIA API
    return st_conn
# Retrive GSheet #
def gsheet_df(target_sheet, check=False)->dict:
    """
    Returns DF values from GSheet & the Connection to it
    [Header Field]:[Value]
    Args:
    target_sheet: The targeted tab on said GSheet
    """
    df_conn=gsheet_connection()
    gspread_df = df_conn.worksheet(target_sheet)
    df = gspread_df.get_all_records()
    df = pd.DataFrame(df)
    df_columns=df.columns.to_list()

    last_row=list(df.iloc[-1])
    df_last_date=last_row[0]
    df_last_val_android=last_row[2].item()
    df_last_val_apple=last_row[4].item()

    return_values={
                    'df_connection':df_conn,
                    'gspread_df':gspread_df,
                    'df':df,
                    'df_columns':df_columns,
                    'df_last_date':df_last_date,
                    'df_last_val_android':df_last_val_android,
                    'df_last_val_apple':df_last_val_apple,
                    }
    if check:print(f'{[return_values]}');return ## ------------------ CHECK ------------------ ##
    return return_values
# Update GSheet #
def update_gsheet(new_data, data_type=[], custom_formatting=[],
                  header=True,
                  refresh=False, 
                  retrieve=False,
                  check=False,
                  )->None:
        # https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/cells#CellFormat
        # https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/cells#NumberFormat
        # https://developers.google.com/workspace/sheets/api/guides/formats
    # def update_gsheet(new_data, target_df, target_sheet, latest_date, refresh=False, retrieve=False,
        """
        Returns status of operations
        [Header Field]:[Value]
        Args:
        NOT YET COMPLETED!
        """
        if data_type==[]:
            return_values={
                'status':0,
                'message':'update_gsheet - Done, no formatting ✅',
                'content':[new_data,target_df],
                }
            return return_values

        # column_len=len(data_type)
        # if column_len!=len(custom_formatting) and custom_formatting!=[]:
        #     return_values={
        #                 'status':1,
        #                 'message':'update_gsheet - ERR, Custom formatting ',
        #                 'content':[],
        #                 }
        #     return return_values
        
        format_pattern={   # IF CUSTOM FORMATTING, USE t
            't0':'TEXT',
            't1':'NUMBER',
            't2':'PERCENT',
            't3':'CURRENCY',
            't4':'DATE',
            't5':'TIME',
            't6':'DATE_TIME',
            't7':'SCIENTIFIC',
            'def0':['DATE','mmm d, yyyy'],
            'def1':['DATE','mmmm d, yyyy'],
            'def2':['DATE','d-M-yy'],
            'def3':['DATE','dd-MM-yyyy+'],
            'def4':['NUMBER','#,##0'],              # NUM
            'def5':['NUMBER','#,##0.00'],           # NUM WITH DECIMAL
            'def6':['CURRENCY','\"RM\"#,##0.00'],
        }
        max_row=1 # (*1/2/3/4)
        target_entry=1

        var_gsheet_df=gsheet_df(target_sheet)
        target_df=var_gsheet_df['gspread_df']
        target_df.append_row(new_data, value_input_option='USER_ENTERED') # APPENDS DATA TO GOOGLE SHEETS
        if check:print(f'~~~{['STOP']}');return ## ------------------ CHECK ------------------ ##
        table_len=len(var_gsheet_df['df'].index)+header # HEADER - BOOL converts to int 1/0
        
        custom_formatting_index=0
        curr_col_num=0
        for i, each in enumerate(data_type):
            curr_col_num+=1
            curr_col=curr_col_num-1
            print(new_data[curr_col],curr_col)
            # print(curr_col_num)
            if i==0 and isinstance(each, int):
                return_values={
                    'status':1,
                    'message':'update_gsheet - FORMAT ERR: First entry can\'t be number',
                    'content':None,
                    }
                print('err')
                return return_values
            # print(i,each)
            if isinstance(each, int):
                n=each
                while n>0:
                    if current_data_type!='':
                        format_cell=CellFormat(numberFormat=numberFormat(type=current_data_type,
                                                                        pattern=current_pattern)
                                                                        )
                        cell_coor=gspread.utils.rowcol_to_a1(row=(table_len),col=curr_col_num*max_row)
                        format_cell_range(target_df,cell_coor,format_cell)
                    print('---',new_data[curr_col],data_type[i-1],current_data_type,n)
                    n-=1
                    if n>0:
                        curr_col_num+=1
                        curr_col=curr_col_num-1
                continue
            elif each=='-':
                current_data_type=prev_data_type
                current_pattern=prev_pattern
            elif each=='':
                current_data_type=''
                current_pattern=''
                print('~~~',new_data[curr_col],None,current_data_type)
                continue

            try:
                current_data_type=format_pattern[data_type[i]]
                current_pattern=format_pattern[data_type[i]]
                # print(current_data_type)
                if 'def' in data_type[i]:
                    current_data_type=current_data_type[0]
                    current_pattern=current_pattern[1]
                else:
                    current_pattern=custom_formatting[custom_formatting_index]
                    custom_formatting_index+=1
            except:
                return_values={
                    'status':2,
                    'message':'update_gsheet - ERR, Invalid format',
                    'content':None,
                    }
                print('err')
                return return_values

            format_cell=CellFormat(numberFormat=numberFormat(type=current_data_type,
                                                            pattern=current_pattern)
                                                            )
            cell_coor=gspread.utils.rowcol_to_a1(row=(table_len),col=curr_col_num*max_row)
            format_cell_range(target_df,cell_coor,format_cell)
            prev_data_type=current_data_type
            prev_pattern=current_pattern
            print('>>>',new_data[curr_col],data_type[i],current_data_type)

        # if max_row not reached: row*column_len
        # e=new_data.to_dict()
        # new_data=['2026-08-21',9,135555,4,75000,'2026-10-01 11:48:27.453538']
        # result,status=check_update_date(latest_date); if result!=0:print(status);return
        # col=gspread.utils.rowcol_to_a1(row=1,col=target_entry*max_row)
        # cell_coor=f'{col}{str(len(var_gsheet_df['df'].index)+1)}'
        # formatted_date=CellFormat(numberFormat=numberFormat(type='DATE',pattern="mmm d, yyyy"))
        # format_cell_range(target_df,col,formatted_date)
        if refresh:
            target_df=gsheet_df(target_sheet)
        if retrieve:
            target_df=gsheet_df(target_sheet)
        # updated_df=pd.concat([target_df,new_data],ignore_index=True)
        # st.dataframe(updated_df)
        # gsheet_updt('testing_updt',new_data,run_connection=False,df_connection=mpm_df_connection)
        return_values={
            'status':0,
            'message':'update_gsheet - Done ✅',
            'content':[new_data,target_df],
            }
        return return_values
        st_conn=gsheet_connection()
        worksheet=st_conn.worksheet(target_df)
        print(worksheet)
        return


def _exec():
    os.system('cls')
    target_sheet='_final'
    target_sheet='testing_updt'
    color_1="#A2AAAD"
    color_2="#a4c639"
    color_3="#CE0E2D"
    color_1="#c3cd3ec2"
    color_2="#3e45cdc2"
    ##--##-----------------------------##--##
    ##-----------INIT ST LAYOUT------------##
    def layout():
        st.set_page_config(page_title="App Store Data Analysis", page_icon=":bar_chart:", layout="wide")
        st.set_page_config(initial_sidebar_state=500)

        st.sidebar.selectbox('Aplikasi',('MasjidPro','USIMMasjidPro'),label_visibility='visible')
        color_1="#A2AAAD"
        color_2="#a4c639"
        color_3="#CE0E2D"
        color_1="#c3cd3ec2"
        color_2="#3e45cdc2"

        st.html(f"""
        <style>
        section[data-testid="stSidebar"] {{
            width: 250px !important;
        }}
        .stHorizontalBlock{{
            gap: 0rem;
        }}
        .st-key-top_1_container{{
            background: {color_1};
            border-radius: 0.5rem 0rem 0rem 0.5rem;
        }}
        .st-key-top_2_container{{
            background: {color_2};
            border-radius: 0rem;
        }}
        button[aria-label="Fullscreen"] {{
            display: none;
        }}
        [data-testid="stHeaderActionElements"] {{
            display: none;
        }}
        [data-testid="stImage"] img {{
            border-radius: 0rem ;
        }}
        </style>
        """)
        #'''
        # a[aria-label="Link to heading"] {{
        #     display: none;
        #         }}
        # .st-emotion-cache-gi0tri{{
        #     display: none;
        #         }}
        # .st-emotion-cache-i18xme{{
        #     display: none;
        #         }}
        # '''
    layout()

    ##--##-----------------------------##--##
    ##-----------EXE BUTTON ACT------------##
    def execute(target_sheet:str, target_app:int, target_df, custom_date='', check=''):
    # execute(target_sheet, latest_date, target_df, check=''): OLD ARGU
        os.system('cls')

        # FOR TESTING
        if check!='':
            if check=='updt_test':
                def final_run():
                    current_date=datetime.datetime
                    var_gsheet_df=gsheet_df(
                                        target_sheet,
                                        # check=True,
                                        )
                    res_check_update_date=check_update_date(
                                    var_gsheet_df['df_last_date'],
                                    # overwrite='01042026',
                                    # check=True,
                                    )
                    if res_check_update_date['response']!=0:
                        print(res_check_update_date['status'])
                        return
                    # else:
                    #     print(res_check_update_date['status'])
                    var_target_date=res_check_update_date['target_date'] #YYYYMMDD
                    var_starting_day=var_target_date[6:]
                    var_target_month=int(var_target_date[:6])
                    cumulative_sum_google=var_gsheet_df['df_last_val_android']
                    cumulative_sum_apple=var_gsheet_df['df_last_val_apple']
                    apple_df_filter_req={
                                        'Title':'MasjidProMY',
                                        'Product Type Identifier':'1F',
                                        }
                    # print(res_check_update_date['status'],n);return
                    # if var_target_month>12:
                    #     year=str(var_target_month)[:4]
                    #     year=str(int(year)+1)
                    #     var_target_month=int(year+'01')
                    # if var_target_date>12:
                    #     year=str(var_target_date)[:4]
                    #     year=str(int(year)+1)
                    #     var_target_date=int(year+'01')
                    # new_data=[];new_df=[]
                    # n=2;var_target_month=202613;var_target_date='20261301'
                    n=res_check_update_date['months_to_update']
                    skip_apple_df=False
                    remarks=''
                    while n>0:
                        var_gdf=google_df(
                                target_app=target_app,
                                target_month=var_target_month,
                                #  check=True,
                                )
                        var_adf=apple_df(
                                frequency_type='MONTHLY',
                                start_date=var_target_date,
                                filter_req=apple_df_filter_req,
                                # check=True,
                                )
                        cumulative_sum_google+=var_gdf['sum']
                        frequency_failover=var_gdf['frequency_failover']

                        if frequency_failover:
                            remarks='Apple - Skipped to end of the year because logs expired'
                            end_date=var_gdf['end_date']
                            cumulative_sum_apple+=var_gdf['sum']
                            skip_apple_df=True
                        else:
                            cumulative_sum_apple+=var_gdf['sum']
                        date=datetime.datetime.strptime(str(var_target_month)+var_starting_day, "%Y%m%d")
                        formatted_date=f"{date.strftime('%b')} {date.day}, {date.year}"
                        append_data=[
                            formatted_date,
                            var_gdf['sum'],
                            cumulative_sum_google,
                            var_adf['sum'],
                            cumulative_sum_apple,
                            remarks,
                            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            ]
                        update_gsheet(append_data,data_type=[
                            'def0',
                            'def4',
                            3,
                            ],
                            # check=True,
                        )
                        #FORMATTING ISSUE - MIGHT BE RUNTIME ISSUE
                        # if check fail, run update_gsheet again, in update_gsheet?
                        var_gsheet_df=gsheet_df(
                                            target_sheet,
                                            # check=True,
                                            )
                        res_check_update_date=check_update_date(
                                        var_gsheet_df['df_last_date'],
                                        # overwrite='01042026',
                                        # check=True,
                                        )
                        var_target_date=res_check_update_date['target_date'] #YYYYMMDD
                        var_starting_day=var_target_date[6:]
                        var_target_month=int(var_target_date[:6])
                        # # new_data.append(var_gdf['sum'])
                        # # new_df.append(var_gdf['entire_df'])
                        # var_target_month+=1
                        # date_increment(var_target_date, 1)
                        # if var_target_month>12 or var_target_date>12:
                        #     check_update_date(var_target_date,retrieve_eval=True)
                        n-=1
                        if n==4:
                            os.system('cls');print('-!DONE!-');return
                    # print(new_data,new_df)
                    # st.write('--> '+str(last_row)) #-TESTING ONLY
                final_run()
                def test():
                    # var_target_date=str(datetime.datetime.now().strptime('20260301',"%Y%m%d").date())
                    # var_target_date=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    var_target_date='20250802'
                    var_target_date='20261001'
                    # print(var_target_date);return
                    apple_df_filter_req={
                                        'Title':'MasjidProMY',
                                        'Product Type Identifier':'1F',
                                        }
                    # google_df_filter_req=
                    var_adf=apple_df(
                            frequency_type='YEARLY',
                            start_date=var_target_date,
                            filter_req=apple_df_filter_req,
                            # check=True,
                            )
                    var_gdf=google_df(
                            target_app=26,
                            target_month=var_target_date,
                            # check=True,
                            )
                    print('var_adf')
                    if var_adf['status'] < 2:
                        st.dataframe(var_adf['entire_df'])
                    if var_adf['status'] == 0:
                        st.dataframe(var_gdf['entire_df'])
                    return
                # test()
            '''
            def testing_only(new_data='',test='A',target_sheet='',update=False):
                if test=='A':
                    file=r'C:\Work\06_PYTHON\_Google_pull_data\.env\mock_data\test.xlsx'
                    df=pd.read_excel(file)
                    new_val=df.iloc[-1]
                    new_val=new_val.to_dict()
                    new_val.pop('Package Name')
                    new_val['A\'Droid listing acquisitions']=new_val.pop('Daily Device Installs')
                    new_val.update({
                        'Android milestone':135555,
                        'Apple listing acquisitions':4,
                        'Apple milestone':75000,
                        'Timestamp':str(datetime.datetime.now()),
                        })
                    # retrieve_val_2=gsheet_df('testing_updt',df_connection=mpm_df_connection)
                    # mpm_android_df_test=retrieve_val_2['df']
                    # return new_val,mpm_android_df_test
                    # print(mpm_android_df_test)
                    print('EXECUTING A')
                    return new_val
                if test=='B-G':
                    google_dt()
                    var_google_auth=google_auth(
                                            #    check=True,
                                               )
                    var_google_df=google_df(
                                            var_google_auth['bucket'],
                                            target_app=24
                                            # var_target_app, OLD ARGU
                                            # var_google_auth['csv_listing'],
                                            # check=True,
                                            # stop_run=25,
                                            )
                    st.dataframe(var_google_df['df'])
                if test=='B-A':
                    # st.write('APPLE OFFLINE');return
                    apple_dt(
                            # check=True
                            )
                    apple_auth(
                            #   check=True
                              )
                    var_filter_req={
                            'Title':'MasjidProMY',
                            'Product Type Identifier':'1F',
                            }
                    var_apple_df= apple_df(
                                    start_date='2026-03-02',
                                    frequency_type='MONTHLY',
                                    filter_req=var_filter_req,
                                    # check=True,
                                    terminal_dump=True,
                                    )
                    st.dataframe(var_apple_df['df'])
                    return
                if test=='B':
                    try:
                        new_data=pd.DataFrame([new_data])
                    except Exception as e:
                        st.write('ABORT, ERROR!')
                        print(f'ERROR!:{e}')
                        return
                    if isinstance(target_df, pd.DataFrame):
                        st.write('current data')
                        st.dataframe(target_df)
                    else:
                        pass
                    st.write('new data')
                    st.dataframe(new_data)
                    print('B')
                    # return
                if update:
                    print('UPDT')
                    update_gsheet(new_data,
                                  target_df,
                                  target_sheet=target_sheet,
                                  latest_date=latest_date,
                                  refresh=False,
                                  retrieve=True,
                                #   check=True,
                                  )
                    st.write('updated data')
                    # st.dataframe(new_data)
            # st.dataframe([new_data])
            print('EXECUTING TEST')
            new_data=testing_only(test='A')
            if check=='date':
                retrieved_dates=check_update_date(latest_date,overwrite=['01042027','01042029'])
                gsheet_connection(check)
                return
            if check=='table':
                testing_only(new_data=new_data,test='B')
                # print(target_df)
            if 'updt' in check:
                testing_only(new_data=new_data,
                             test='B',
                             target_sheet=target_sheet,
                             update=True
                             )
                # update_gsheet(new_data,target_df,target_sheet) #can't update immediately multiple times because the retrieved data is not the latest (has grace period) 
            if check=='table_g':
                testing_only(new_data=new_data,test='B-G')
            if check=='table_a':
                testing_only(new_data=new_data,test='B-A')
            #'''
                
        return
        # print('mpm_android_df');
        # # Write back to Google Sheets
        # conn.update(worksheet="Sheet1", data=updated_df)
        lastest_val_android=android_fetch_val-df_last_val_android
        lastest_val_apple=apple_fetch_val-df_last_val_apple
        col_1='Date'
        col_2='A\'Droid listing acquisitions'
        col_3='Android milestone'
        col_4='Apple listing acquisitions'
        col_5='Apple milestone'
        new_row=pd.DataFrame({
            '{col_1}:{current_date}'
            '{col_2}:{lastest_val_android}'
            '{col_3}:{android_fetch_val}'
            '{col_4}:{lastest_val_apple}'
            '{col_5}:{apple_fetch_val}'
            })
        mpm_df_connection.update(
            worksheet='Sheet1',
            usecols=[0,2,4],
        )
        pass


    ##--##-----------------------------##--##
    ##----------EXE RETRIEVE DATA----------##
    _=r'''
    retrieve_val_1=gsheet_df(target_sheet)
    mpm_df_connection=retrieve_val_1['df_connection']
    mpm_latest_date=retrieve_val_1['df_last_date']

    # MasjidPro #
    mpm_df=retrieve_val_1['df']
    mpm_gspread_df=retrieve_val_1['gspread_df']
    mpm_android_latest_value=retrieve_val_1['df_last_val_android']
    mpm_apple_latest_value=retrieve_val_1['df_last_val_apple']
    # st.write(mpm_android_df)
    st.write(mpm_df)

    ##--##-----------------------------##--##
    ##------------INIT COLUMNS-------------##
    col1,col2,col3=st.columns([1,5,1])
    with col1.container(key='top_1_container',border=True,horizontal_alignment='center'):
        st.image(r'media\mpm.svg')
        st.subheader('MasjidProMy',text_alignment='center')
    with col2.container(key='top_2_container',border=True,horizontal_alignment='center'):
        st.image(r'media\ump.svg')
        st.subheader('USIMMasjidPro',text_alignment='center')
        val=[int(mpm_android_latest_value),int(mpm_apple_latest_value)] ; entity=['Android','Apple']
        fig = px.pie(
                    values=val,
                    names=entity,
                    hole=0.4,
                    color=entity,
                    color_discrete_map={
                        entity[0]: "#3261ce",
                        entity[1]: "#32ce5c",
                    },  # Custom color mapping
        )
        fig.update_layout(showlegend=False, 
                          paper_bgcolor="#ffffff00",
                          height=170,
                          margin=dict(t=0,b=0,l=0,r=0)
                          )
        fig.update_traces(textinfo="label+percent",
                          textposition="inside", 
                          textfont_size=14)
        fig_2 = go.Figure(go.Pie(
                    values=val,
                    labels=entity,
                    hole=0.4,
                    marker=dict(colors=[color_1,color_2,]), # Custom color mapping
        ))
        fig_2.update_layout(showlegend=False, 
                          paper_bgcolor="#ffffff00",
                          height=170,
                          margin=dict(t=0,b=0,l=0,r=0)
                          )
        fig_2.update_traces(textinfo="label+percent",
                          textposition="inside", 
                          textfont_size=10,
                          )
        # st.plotly_chart(fig, use_container_width=True)
        st.plotly_chart(fig_2, width=True)
        df=dict(
            _0=['Android',f'{"{:,}".format(val[0])}'],
            _1=['Apple',f'{"{:,}".format(val[1])}'],
            )
        df=pd.DataFrame.from_dict(df,orient='index',columns=['Device','Value'])
        st.dataframe(df, hide_index=True)
        col_1_1,col_1_2=st.columns([1,1])
        with col_1_1.container():
            st.write(f'Total: {"{:,}".format(val[0]+val[1])}')
        with col_1_2.container():
            content=f'Data as of: {mpm_latest_date}'
            col_1_2.markdown(f"<div style='text-align: right;'>{content}</div>", unsafe_allow_html=True)
            # st.write(f'Date: {mpm_latest_date}')
            pass

    
        # if "counter" not in st.session_state:
        #     st.session_state.counter = 0
        # # 2. Define your callback function
        # def increment_counter(amount):
        #     st.session_state.counter += amount
        # 3. Use on_click and pass arguments via args=
    col_1_3,col_1_4,col_1_5,col_1_6,col_1_7=st.columns([1,1,1,1,1])
    with col_1_3.container():
        if st.button("Test date"):
            mpm_latest_date=''
            mpm_gspread_df=''
            execute(target_sheet,mpm_latest_date,mpm_df,check='date')
            pass
    with col_1_4.container():
        if st.button("Test retrieve"):
            execute(target_sheet,mpm_latest_date,mpm_df,check='table')
            pass
    with col_1_5.container():
        if st.button("Test update"):
            execute(target_sheet,mpm_latest_date,mpm_gspread_df,check='table_updt')
            pass
    col_1_3,col_1_4,col_1_5,col_1_6,col_1_7=st.columns([1,1,1,1,1])
    with col_1_6.container():
        if st.button("Test retrieve google"):
            mpm_latest_date=''
            mpm_gspread_df=''
            execute(target_sheet,mpm_latest_date,mpm_gspread_df,check='table_g')
    with col_1_7.container():
        if st.button("Test retrieve apple"):
            mpm_latest_date=''
            mpm_gspread_df=''
            execute(target_sheet,mpm_latest_date,mpm_gspread_df,check='table_a')
    #'''

    col_2_1,col_2_2,col_2_3,col_2_4=st.columns([1,1,1,1])
    with col_2_1.container():
        var_app=st.text_input('The application (number)',placeholder='Eg: MasjidProMy')
        try:
            int(var_app)
        except:
            print('Invalid')
    with col_2_2.container():
        var_sheet=st.text_input('Target sheet',value='testing_updt') #testing_updt
        var_sheet=str.lower(var_sheet)
    with col_2_3.container():
        var_time=st.text_input('Start date (Month & Year) - Optional',placeholder='MMYYYY')
        var_time=str.lower(var_time)
    with col_2_4.container():
        if st.button("Apply Update"):
            if var_app=='' or var_sheet=='':
                print('incomplete')
            else:
                mpm_latest_date=''
                mpm_gspread_df=''
                # print(f'--{var_app,var_time,var_sheet}\n--')
                execute(target_sheet,var_app,mpm_gspread_df,check='updt_test')
            # execute(target_sheet,mpm_latest_date,mpm_gspread_df,check='table_a')
        
    # col1.selectbox(label='Aplikasi',label_visibility='visible',options=['MasjidPro','USIMMasjidPro'])
    
if __name__=='__main__':
    def packages():
        try:
            import os
            import datetime
            import requests
            import jwt
            import gzip
            import shutil
            import time
            import io
            import csv
            # from google.cloud import storage
            # from google_play_scraper import app
            import streamlit as st
            import numpy as np
            import pandas as pd
            import seaborn
            import matplotlib
            print('success')
        except Exception as e:
            print('---',e)
    # packages()
    _exec()
    def test():
        # 1. Sample Data (using public image URLs for logos)
        data = {
            "Product Name": ["Alpha App", "Beta Tool", "Gamma Suite"],
            "Logo": [
                "https://images.unsplash.com/photo-1618005182384-a83a8bd57fbe?w=100&h=100&fit=crop",
                "https://images.unsplash.com/photo-1579546929518-9e396f3cc809?w=100&h=100&fit=crop",
                "https://images.unsplash.com/photo-1557683316-973673baf926?w=100&h=100&fit=crop",
            ],
            "Downloads": [1524000, 843200, 3120000],
            "Android (%)": [65, 40, 55],
            "Apple (%)": [35, 60, 45],
            "Registered Users": [452000, 210000, 980000],
            "Entity": [["Android", "Apple"], ["Android", "Apple"], ["Android", "Apple"]],
            "Percentage": [[65, 35],[40, 60],[55, 45]],
        }

        fig = px.pie(
                values=[65, 35],
                names=["Android", "Apple"],
                hole=0.5,  # This turns the pie chart into a donut chart
                color=["Android", "Apple"],
                color_discrete_map={
                    "Android": "#3261ce",
                    "Apple": "#32ce5c",
                },  # Custom color mapping
            )
        # 3. Optional: Customize text layout to show both percent and label clearly
        fig.update_traces(textinfo="percent+label", textfont_size=14)

        df = pd.DataFrame(data)

        # 2. Configure column appearance and formatting
        column_configurations = {
            "Product Name": st.column_config.TextColumn(
                "Product", help="Name of the application", width="medium"
            ),
            "Logo": st.column_config.ImageColumn(
                "Logo", help="App brand logo", width="small"
            ),
            "Downloads": st.column_config.NumberColumn(
                "Total Downloads", format="%d 📥", help="Total lifetime downloads"
            ),
            "Android (%)": st.plotly_chart(fig, width=True),
            # "Android (%)": st.column_config.ProgressColumn(
            #     "Android Share", format="%d%%", min_value=0, max_value=100
            # ),
            "Apple (%)": st.column_config.ProgressColumn(
                "Apple Share", format="%d%%", min_value=0, max_value=100
            ),
            "Registered Users": st.column_config.NumberColumn(
                "Registered Users", format="%d 👤", help="Total registered accounts"
            ),
        }

        # 3. Render the interactive table
        st.dataframe(
            df,
            column_config=column_configurations,
            use_container_width=True,
            hide_index=True,
        )
        # import streamlit as st

        # 1. Prepare your data
        # data = {"Entity": ["Android", "Apple"], "Percentage": [34, 66]}

        # 2. Create a Donut chart using Plotly Express
        

        

        # # 4. Render the chart inside your Streamlit app
        # st.subheader("Mobile OS Market Share")
    # test()