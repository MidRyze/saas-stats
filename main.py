def exec():
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

    ##--##-----------FUNCTIONS---------##--##
    ##--------------!! API !!--------------##
    # Retrive Google Sheet #
    def gsheet_connection(check=False):
        # st_conn=st.connection("gsheet_data", type=GSheetsConnection)
        gsheet_cred=gspread.service_account_from_dict(st.secrets["connections"]["gsheet_data"])
        url=str(st.secrets["connections"]["gsheet_data"]['spreadsheet'])
        prefix='/d/'
        suffix='/edit'
        target_workbook=url[url.find(prefix)+len(prefix):url.find(suffix)]
        if check:
            print(f'{target_workbook}')
            return
        # Establish Connections #
        st_conn=gsheet_cred.open_by_key(target_workbook) # CONNECTS GOOGLE SERVICES VIA API
        return st_conn

    # Retrive Google Bucket #
    def google_dt(check=False):
        if check:
            print(f'{[]}')
            return
        print(gbucket_data_creds())
        pass

    # Retrive Apple Console #
    def apple_dt(check=False):
        if check:
            print(f'{[]}')
            return
        print(apple_data_creds())
        pass

    def apple_data(ISSUER_ID,KEY_ID,PRIVATE_KEY_PATH,generate_token=True):
        def gen_token():
            '''
            alg:ES256                                       # Encryption Algorithm -> All JWTs for App Store Connect API must be signed with ES256 encryption.
            kid:Your private key ID from App Store Connect  # Key Identifier -> for example, 2X9R4HXF34
            typ:JWT                                         # Token Type
            '''
            header={
                "alg": "ES256",
                "kid": KEY_ID,
                "typ": "JWT",
            }
            payload={
                'iss':ISSUER_ID,
                'iat':int(time.time()),
                'exp':int(time.time())+1140,
                'aud':'appstoreconnect-v1',
            }
            with open(PRIVATE_KEY_PATH, 'rb') as f:
                priv_key=f.read()

            token=jwt.encode(payload, priv_key, algorithm=header['alg'], headers=header)
            # print(token)
            return token
        
        def generate_apple_token(token):
            header={
                "Authorization": f"Bearer {token}",
            }
            url="https://api.appstoreconnect.apple.com/v1/apps"
            response=requests.get(url, headers=header)

            if response.status_code==200:
                data=response.json()
                # print('>>>',data)
            else:
                json_response=response.json()['errors'][0]
                print(f"{response.status_code}, {json_response['code']}\n{json_response['title']} {json_response['detail']}")

        def dwnld_apple_data(VENDOR_NUM,token,start_date:str,frequency_type:str='YEARLY'):
            """
            [Header Field]:[Value]
            Args:
            start_date (str): Date format 'YYYY-MM-DD'
            frequency_type (str): 'YEARLY'/'MONTHLY'/'~~WEEKLY~~'/'DAILY'
            """
            if start_date=='':
                start_date=datetime.datetime.today().strftime('%Y-%m-%d')
                month=int(start_date[5:7])-1
                year=int(start_date[:4])
                if month==-1:
                    month=str(12)
                    year=str(int(start_date[:4])-1)
                month='{:02d}'.format(month)
                year='{:0004d}'.format(year)
                start_date=f'{year}-{month}-{start_date[8:]}'
            if frequency_type=='YEARLY':
                start_date=start_date[:4]
            elif frequency_type=='MONTHLY':
                start_date=start_date[:7]
            # print(start_date,type(start_date));return
            # start_date='2026-08-23'
            vendor_num=VENDOR_NUM
            url="https://api.appstoreconnect.apple.com/v1/apps"
            headers={'Authorization':f'Bearer {token}',}
            response=requests.get(url,headers=headers)
            apps_data=response.json().get('data',[])
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
            report_url="https://api.appstoreconnect.apple.com/v1/salesReports"
            params={
                "filter[frequency]":frequency_type,
                "filter[reportDate]":start_date, # YYYY-MM-DD, Not required, if frequency = daily
                "filter[reportSubType]":'SUMMARY',
                "filter[reportType]":'SALES',
                "filter[vendorNumber]":f'{vendor_num}',
                # "filter[version]":'1_0', # Not required
                }
            report_response=requests.get(report_url,headers=headers,params=params)
            
            if report_response.status_code!=200:
                print(report_response.text);return
            
            with gzip.open(io.BytesIO(report_response.content),mode='rt',encoding='utf-8',errors='ignore') as f:
                _csv=csv.DictReader(f,delimiter='\t')

            if generate_token:
                new_token=gen_token()

    def gbucket_data_creds():
        return st.secrets["google_data"]
    def apple_data_creds():
        return st.secrets["apple_data"]

    # Calculate time #
    def check_update_date(latest_date, delay=2, target_day=25, overwrite=[]):
        """
        delay: How many month(s) to check for latest report\n
        target_day: The targeted day to check, 28 is the latest to account for February\n
        """
        def convert_to_dates(latest_date,current_date):
            late_month=int(latest_date[2:4])
            late_year=int(latest_date[4:])
            curr_day=int(current_date[:2])
            curr_month=int(current_date[2:4])
            curr_year=int(current_date[4:])
            return late_month, late_year, curr_day, curr_month, curr_year

        target_day=int(target_day)
        if target_day>28:
            target_day=28

        if len(overwrite)!=0:
            latest_date=str(overwrite[0])
            if len(latest_date)!=8:
                print('LENGTH ERR')
                return 2,'LENGTH ERR'
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

        late_month, late_year, curr_day, curr_month, curr_year=convert_to_dates(latest_date,current_date)
        try:
            datetime.datetime.strptime(latest_date,'%d%m%Y')
        except Exception as e:
            print(f'FORMAT ERR: {e}')
            return 3,'FORMAT ERR'

        if late_year>curr_year or (late_year==curr_year and late_month>curr_month):
            latest_date=current_date
            late_month, late_year, curr_day, curr_month, curr_year=convert_to_dates(latest_date,current_date)
        
        year_difference=curr_year-late_year
        curr_month+=year_difference*12

        if (curr_month-late_month)>delay or (curr_month-late_month)==delay and curr_day >= target_day:
            # update_gsheet(new_data,target_df)
            print(f'UPDATE: {curr_month-late_month} month(s) difference')
            return 0
        else:
            print(f'NO UPDATE')
            return 1,'NO UPDATE'

        data_date={
            'latest_date':latest_date,
            'current_date':current_date,
            'late_month':late_month,
            'late_year':late_year,
            'curr_day':curr_day,
            'curr_month':curr_month,
            'curr_year':curr_year,
        }
        return data_date

    #-- GSheet --#
    # Retrive GSheet #
    def gsheet_df(target_sheet):
        df_conn=gsheet_connection()
        # df=df_conn.read(
        #     worksheet=target_sheet,
        #     ttl='8m',
        #     # usecols=[0,2,4],
        # )
        gspread_df = df_conn.worksheet(target_sheet)
        df = gspread_df.get_all_records()
        df = pd.DataFrame(df)
        # print('>>>',df_conn)

        last_row=list(df.iloc[-1])
        df_last_date=last_row[0]
        df_last_val_android=last_row[2].item()
        df_last_val_apple=last_row[4].item()
        # st.write('--> '+str(last_row)) #-TESTING ONLY
        values={
            'df':df,
            'gspread_df':gspread_df,
            'df_connection':df_conn,
            'df_last_date':df_last_date,
            'df_last_val_android':df_last_val_android,
            'df_last_val_apple':df_last_val_apple,
                }
        
        return values
    
    # Update GSheet #
    def update_gsheet(new_data,target_df,target_sheet,latest_date,refresh=False,retrieve=False,check=False):
        e=new_data.to_dict()
        print(e)
        return
        result,status=check_update_date(latest_date)
        if result!=0:
            print(status)
            return
        if refresh:
            target_df=gsheet_df(target_sheet)
        # new_data=pd.DataFrame([new_data],columns=target_df.columns)
        if check:
            print(f'{new_data}')
            return
        # convert new_data to list
        new_data=['2026-08-21',9,135555,4,75000,'2026-10-01 11:48:27.453538']
        target_df.append_row(new_data)
        if retrieve:
            target_df=gsheet_df(target_sheet)
        # updated_df=pd.concat([target_df,new_data],ignore_index=True)
        # st.dataframe(updated_df)
        # gsheet_updt('testing_updt',new_data,run_connection=False,df_connection=mpm_df_connection)
        return new_data,target_df
        st_conn=gsheet_connection()
        worksheet=st_conn.worksheet(target_df)
        print(worksheet)
        return

    ##--##-----------------------------##--##
    ##-----------EXE BUTTON ACT------------##
    def execute(target_sheet, latest_date, target_df, check=''):
        os.system('cls')

        # FOR TESTING
        if check!='':
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
                    return new_val
                if test=='B-G':
                    google_dt()
                if test=='B-A':
                    apple_dt()
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
            print('TEST')
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

        # print('mpm_android_df');
        return
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
    retrieve_val_1=gsheet_df(target_sheet)
    mpm_df_connection=retrieve_val_1['df_connection']
    mpm_latest_date=retrieve_val_1['df_last_date']

    # MasjidPro #
    mpm_df=retrieve_val_1['df']
    mpm_gspread_df=retrieve_val_1['gspread_df']
    mpm_android_latest_value=retrieve_val_1['df_last_val_android']
    mpm_apple_latest_value=retrieve_val_1['df_last_val_apple']
    # st.write(mpm_android_df)

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

        # update_data=st.text_input('Test text update')

    
        # if "counter" not in st.session_state:
        #     st.session_state.counter = 0
        # # 2. Define your callback function
        # def increment_counter(amount):
        #     st.session_state.counter += amount
        # 3. Use on_click and pass arguments via args=
        col_1_3,col_1_4,col_1_5,col_1_6,col_1_7=st.columns([1,1,1,1,1])
        with col_1_3.container():
            if st.button("Test date"):
                execute(target_sheet,mpm_latest_date,mpm_df,check='date')
        with col_1_4.container():
            if st.button("Test retrieve"):
                execute(target_sheet,mpm_latest_date,mpm_df,check='table')
        with col_1_5.container():
            if st.button("Test update"):
                execute(target_sheet,mpm_latest_date,mpm_gspread_df,check='table_updt')
        with col_1_6.container():
            if st.button("Test retrieve google"):
                execute(target_sheet,mpm_latest_date,mpm_gspread_df,check='table_g')
        with col_1_7.container():
            if st.button("Test retrieve apple"):
                execute(target_sheet,mpm_latest_date,mpm_gspread_df,check='table_a')
        
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
    import os
    import datetime
    import requests
    import jwt
    import gzip
    import shutil
    import time
    import streamlit as st
    import numpy as np
    import pandas as pd
    import openpyxl
    import seaborn
    import gspread
    import matplotlib
    import plotly.express as px
    import plotly.graph_objects as go
    from streamlit_gsheets import GSheetsConnection
    exec()
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