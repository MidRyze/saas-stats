def packages():
    try:
        import os
        import datetime
        import requests
        import jwt
        import gzip
        import shutil
        import time
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

def exec():
    # st.title('Aplication')
    from streamlit_gsheets import GSheetsConnection
    st.set_page_config(page_title="App Store Data Analysis", page_icon=":bar_chart:", layout="wide")
    st.set_page_config(initial_sidebar_state=500)
    # .css-15zrgzn{{display:none}}

    df_conn=st.connection("gsheets", type=GSheetsConnection)
    df=df_conn.read()

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
    }}
    .st-key-top_2_container{{
        background: {color_2};
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
    # a[aria-label="Link to heading"] {{
    #     display: none;
    #         }}
    # .st-emotion-cache-gi0tri{{
    #     display: none;
    #         }}
    # .st-emotion-cache-i18xme{{
    #     display: none;
    #         }}

    st.sidebar.selectbox('Aplikasi',('MasjidPro','USIMMasjidPro'),label_visibility='visible')
    col1,col2,col3,col4=st.columns([1,1,1,1])
    with col1.container(key='top_1_container',border=True,horizontal_alignment='center'):
        st.image(r'media\mpm.svg')
        st.subheader('MasjidProMy',text_alignment='center')
    with col2.container(key='top_2_container',border=True,horizontal_alignment='center'):
        st.image(r'media\ump.svg')
        st.subheader('USIMMasjidPro',text_alignment='center')
        val=[10,8] ; entity=['Android','Apple']
        st.write(f'Android: {val[0]}')
        st.write(f'Apple: {val[1]}')
        fig = px.pie(
                    values=val,
                    names=entity,
                    hole=0.5,
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
        st.plotly_chart(fig, use_container_width=True)
        st.write(f'Total: {val[0]+val[1]}')
        
    # col1.selectbox(label='Aplikasi',label_visibility='visible',options=['MasjidPro','USIMMasjidPro'])

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
        "Android (%)": st.plotly_chart(fig, use_container_width=True),
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
    

if __name__=='__main__':
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
    import seaborn
    import matplotlib
    import plotly.express as px
    exec()
    # test()