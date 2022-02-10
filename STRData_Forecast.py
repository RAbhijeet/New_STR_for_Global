"""
CREATED BY : Chakradhar T.

statement: Automate forecast report using STR DATA AND BUDGET FILE.


c_year = current year

"""


import pandas as pd
import calendar
import datetime
import numpy as np
import ForeCast_TY as fc

c_month = datetime.datetime.today().month
c_year = datetime.datetime.today().year

eDate = datetime.date.today() - pd.tseries.offsets.MonthEnd(1)

def str_forecast(str_file, map_file,bget_path, str_fcfname,std_path):

    map_df = pd.read_excel(map_file)
    map_df['keycol'] = map_df.Year.astype(str) + map_df.Month.astype(str)
    print('Reading STR_DATA......')
    df = pd.read_excel(str_file)
    df = df[df['Date'] <= eDate]
    # ---------------------------Rooms and Rev Columns Calculation-------------------

    df['Hotel Rms TY'] = (df['Hotel TY Occ'] * df['Hotel Cap'] / 100).round(0)

    df['Comp Rms TY'] = (df['Comp TY Occ'] * df['Comp Cap'] / 100).round(0)

    df['Hotel Rms LY'] = (df['Hotel LY Occ'] * df['Hotel Cap'] / 100).round(0)

    df['Comp Rms LY'] = (df['Comp LY Occ'] * df['Comp Cap'] / 100).round(0)

    # -------------------------------------------------------------------------------

    df['Hotel Rev TY'] = df['Hotel Rms TY'] * df['Hotel TY ADR']

    df['Comp Rev TY'] = df['Comp Rms TY'] * df['Comp TY ADR']

    df['Hotel Rev LY'] = df['Hotel Rms LY'] * df['Hotel LY ADR']

    df['Comp Rev LY'] = df['Comp Rms LY'] * df['Comp LY ADR']

    # df2 = pd.DataFrame(df, columns=['Date', 'DOW', 'Month', 'Year', 'Hotel Rms TY', 'Comp Rms TY', 'Hotel Rms LY',
    #                                 'Comp Rms LY', 'Hotel Rev TY', 'Comp Rev TY', 'Hotel Rev LY', 'Comp Rev LY'])

    df['mth'] = pd.to_datetime(df.Month, format='%b').dt.month
    fc_df = pd.DataFrame(df[['Year', 'mth', 'Hotel Cap', 'Comp Cap', 'Hotel Rms TY', 'Comp Rms TY', 'Hotel Rms LY',
                             'Comp Rms LY', 'Hotel TY ADR', 'Comp TY ADR', 'Hotel LY ADR', 'Comp LY ADR',
                             'Hotel Rev TY',
                             'Comp Rev TY', 'Hotel Rev LY', 'Comp Rev LY', 'Hotel TY Occ', 'Comp TY Occ',
                             'Hotel LY Occ',
                             'Comp LY Occ']])

    fc_df1 = fc_df.groupby(['Year', 'mth']).sum().reset_index()

    # CREATE KEYCOL TO MAP FUTURE FORECAST WITH FORECAST FILE/DF (fc_df1)
    fc_df1['keycol'] = fc_df1['Year'].astype(str) + fc_df1['mth'].astype(str)
    # COLUMNS CREATION WITH FORMULA
    fc_df1['Hotel TY ADR'] = fc_df1['Hotel Rev TY'] / fc_df1['Hotel Rms TY']

    fc_df1['Comp TY ADR'] = fc_df1['Comp Rev TY'] / fc_df1['Comp Rms TY']

    fc_df1['Hotel LY ADR'] = fc_df1['Hotel Rev LY'] / fc_df1['Hotel Rms LY']

    fc_df1['Comp LY ADR'] = fc_df1['Comp Rev LY'] / fc_df1['Comp Rms LY']

    fc_df1['Hotel TY Occ'] = fc_df1['Hotel Rms TY'] / fc_df1['Hotel Cap'] * 100

    fc_df1['Hotel LY Occ'] = fc_df1['Hotel Rms LY'] / fc_df1['Hotel Cap'] * 100

    fc_df1['Comp TY Occ'] = fc_df1['Comp Rms TY'] / fc_df1['Comp Cap'] * 100

    fc_df1['Comp LY Occ'] = fc_df1['Comp Rms LY'] / fc_df1['Comp Cap'] * 100

    fc_df1['Hotel TY RevPAR'] = fc_df1['Hotel TY ADR'] * fc_df1['Hotel TY Occ'] / 100

    fc_df1['Comp TY RevPAR'] = fc_df1['Comp TY ADR'] * fc_df1['Comp TY Occ'] / 100

    fc_df1['Hotel LY RevPAR'] = fc_df1['Hotel LY ADR'] * fc_df1['Hotel LY Occ'] / 100

    fc_df1['Comp LY RevPAR'] = fc_df1['Comp LY ADR'] * fc_df1['Comp LY Occ'] / 100

    fc_df1['Hotel Occ Forecast'] = 1

    fc_df1['Comp Occ Forecast'] = 2

    fc_df1['Hotel Rate Forecast'] = 3

    fc_df1['Comp Rate Forecast'] = 4

    fc_df1['MPI'] = (fc_df1['Hotel TY Occ'] / fc_df1['Comp TY Occ']).round(4)
    fc_df1 = fc_df1.replace(np.inf, np.nan)
    fc_df1['ARI'] = (fc_df1['Hotel TY ADR'] / fc_df1['Comp TY ADR']).round(4)

    fc_df1['RGI'] = (fc_df1['Hotel TY RevPAR'] / fc_df1['Comp TY RevPAR']).round(4)

    fc_df1['MPI Forecast'] = 5

    fc_df1['ARI Forecast'] = 6

    fc_df1['RGI Forecast'] = 7

    fc_key_df = map_df.merge(fc_df1, on='keycol', how='left')
    # FORECAST CALCULATION FOR ROOMS
    print('Calculating competitor forecast.....')
    fc_rms_df = fc.strForecast(df, 'Comp Rms TY')               # CALL ForeCast_LY from room columns
    fc_rms_df = fc_rms_df.rename(columns={'Forecast': 'Rsm_Forecast'})

    fc_rms_df['Year'] = c_year                                  # filter data for current year
    fc_rms_df = fc_rms_df[fc_rms_df.mth >= c_month]             # filter data for current month
    fc_rms_df = pd.DataFrame(fc_rms_df)

    fc_rms_df.mth = fc_rms_df.mth.astype(int)
    # CREATE KEYCOL TO MAP CALCULATED COMPETITOR FORECAST TO THE CURRENT MONTH ONWARD
    fc_rms_df['keycol'] = fc_rms_df['Year'].astype(str) + fc_rms_df['mth'].astype(str)
    fc_rms_df = fc_rms_df[['keycol', 'Rsm_Forecast']]
    #  MERGING ROOMS FORECAST DF WITH FORECAST DF
    fc_key_df = fc_key_df.merge(fc_rms_df, on='keycol', how='left')
    # FORECAST CALCULATION FOR REVENUE
    fc_rev_df = fc.strForecast(df, 'Comp Rev TY')               # CALL ForeCast_LY from revenue columns
    fc_rev_df = fc_rev_df.rename(columns={'Forecast': 'Rev_Forecast'})
    fc_rev_df['Year'] = c_year                                  # filter data for current year
    fc_rev_df = fc_rev_df[fc_rev_df.mth >= c_month]             # filter data for current month
    fc_rev_df = pd.DataFrame(fc_rev_df)
    fc_rev_df.mth = fc_rev_df.mth.astype(int)
    # CREATE KEYCOL TO MAP CALCULATED COMPETITOR FORECAST TO THE CURRENT MONTH ONWARD
    fc_rev_df['keycol'] = fc_rev_df['Year'].astype(str) + fc_rev_df['mth'].astype(str)
    fc_rev_df = fc_rev_df[['keycol', 'Rev_Forecast']]
    # MERGING ROOMS FORECAST DF WITH FORECAST DF
    fc_key_df = fc_key_df.merge(fc_rev_df, on='keycol', how='left')

    fc_key_df['Hotel Cap'] = fc_df['Hotel Cap'].values[0]
    fc_key_df['Comp Cap'] = fc_df['Comp Cap'].values[0]
    # FORECAST CALCULATION FOR COMPETITOR
    fc_key_df['Comp Occ Forecast'] = fc_key_df['Rsm_Forecast'] / (fc_key_df['Comp Cap'] * fc_key_df['Days']) * 100
    fc_key_df['Comp Rate Forecast'] = fc_key_df['Rev_Forecast'] / fc_key_df['Rsm_Forecast']

    # READ Budget_Forecast FILE
    bd_df = pd.read_excel(bget_path)
    #
    bd_df = bd_df.groupby(['Year', 'Month'], sort=False).sum().reset_index()
    bd_df['Month'] = pd.to_datetime(bd_df.Month, format='%b').dt.month
    # CREATE KEYCOL TO MAP CALCULATED FORECAST TO CURRENT MONTH ONWARD
    bd_df['keycol'] = bd_df['Year'].astype(str) + bd_df['Month'].astype(str)
    # FILTERING DATA FOR CURRENT YEAR AND MONTH
    bd_df = bd_df[bd_df.Year >= c_year]
    bd_df = bd_df[bd_df.Month >= c_month]
    # COLLECTING REQUIRED COLUMNS FOR CALCULATING HOTEL FORECAST
    bd_df = pd.DataFrame(bd_df[['keycol', 'Fcst_Rooms', 'Fcst_Revenue']])
    # MERGING BUDGET DF WITH FORECAST DF
    fc_key_df = fc_key_df.merge(bd_df, on='keycol', how='left')
    # FORECAST CALCULATION FOR HOTEL
    print('Calculating Hotel forecast.....')
    fc_key_df['Hotel Occ Forecast'] = fc_key_df['Fcst_Rooms'] / (fc_key_df['Hotel Cap'] * fc_key_df['Days']) * 100
    fc_key_df['Hotel Rate Forecast'] = fc_key_df['Fcst_Revenue'] / fc_key_df['Fcst_Rooms']

    # ===== FORECAST CALCULATION FOR MPI ARI And RGI ======#
    print('Calculating MPI ARI And RGI forecast.....')
    fc_key_df['MPI Forecast'] = (fc_key_df['Hotel Occ Forecast'] / fc_key_df['Comp Occ Forecast']).round(4)

    fc_key_df['ARI Forecast'] = (fc_key_df['Hotel Rate Forecast'] / fc_key_df['Comp Rate Forecast']).round(4)

    fc_key_df['RGI Forecast'] = (((fc_key_df['Hotel Occ Forecast']) * (fc_key_df['Hotel Rate Forecast'])) / ((fc_key_df['Comp Occ Forecast']) * (fc_key_df['Comp Rate Forecast']))).round(4)

    fc_key_df = fc_key_df[['Year_x', 'Month', 'Hotel Cap', 'Comp Cap', 'Hotel Rms TY', 'Comp Rms TY', 'Hotel Rms LY',
                           'Comp Rms LY', 'Hotel TY ADR', 'Comp TY ADR', 'Hotel LY ADR', 'Comp LY ADR', 'Hotel Rev TY',
                           'Comp Rev TY', 'Hotel Rev LY', 'Comp Rev LY', 'Hotel TY Occ', 'Comp TY Occ', 'Hotel LY Occ',
                           'Comp LY Occ', 'Hotel TY RevPAR', 'Comp TY RevPAR', 'Hotel LY RevPAR', 'Comp LY RevPAR',
                           'Hotel Occ Forecast', 'Comp Occ Forecast', 'Hotel Rate Forecast', 'Comp Rate Forecast',
                           'MPI', 'ARI', 'RGI', 'MPI Forecast', 'ARI Forecast', 'RGI Forecast']]
    # RENAME COLUMNS NAME TO STANDARD COLUMN NAME
    fc_key_df = fc_key_df.rename(columns={'Year_x': 'Year', 'Hotel Rms TY': 'Hotel TY Sold', 'Comp Rms TY': 'Comp TY Sold',
                 'Hotel Rms LY': 'Hotel LY Sold', 'Comp Rms LY': 'Comp LY Sold', 'Hotel Rev TY':'Hotel Revenue TY',
                'Comp Rev TY': 'Comp Revenue TY', 'Hotel Rev LY':'Hotel Revenue LY', 'Comp Rev LY': 'Comp Revenue LY'})
    # CONVERT MONTH NUMBER INTO MONTH NAME 'SORT'
    fc_key_df['Month'] = fc_key_df['Month'].apply(lambda x: calendar.month_abbr[x])
    # print('Creating Forecast report for {} hotel'.format(htl_code))
    fc_key_df.to_excel(std_path +'/'+'{}'.format(str_fcfname),'Sheet1')

# str_fcfname ='D:\Swapnil All Program\Leela_str_project\STR_Data/STR_Data_Forecast.xlsx'
# str_file='D:\Swapnil All Program\Leela_str_project\STR_Data/STR_Data.xlsx'
# map_file = 'D:\Swapnil All Program\Leela_str_project/yearmonthMAP.xlsx'
# htl_code ='UDP'
# bget_path ='D:\Swapnil All Program\Leela_str_project/Budget_Forecast.xlsx'
# str_forecast(str_file, map_file, htl_code,bget_path, str_fcfname)