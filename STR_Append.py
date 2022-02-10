"""

CREATED BY: Chakrdhar T.

MODULE NAME: STR COPY AND PAST

Statement: 'Copy the STR provided by hotel and paste it to the STR Data and STR Data forecast'

Copy required columns data from STR file and paste into STR DATA AND STR DATA FORECAST file with satisfied all condition

CONDITION: append all data from input template file of hotel for current year.

"""
import pandas as pd


def str_copy_paste(htl_file, str_file, htl_code):
    # READ INPUT FILE FROM HOTEL
    df = pd.read_excel(htl_file,  header=[7])
    # COLLECTING REQUIRED COLUMNS
    df = pd.DataFrame(df, columns=['Unnamed: 1', 'Unnamed: 2', 'My Prop', 'Comp Set', 'My Prop.1', 'Comp Set.1',
                                   'Unnamed: 12', 'Unnamed: 13','My Prop.3', 'Comp Set.3', 'My Prop.4', 'Comp Set.4',
                                   'Unnamed: 23', 'Unnamed: 24', 'Unnamed: 34', 'Unnamed: 35'])
    # RENAME COLUMNS IN TO STANDARD NAME FOR STR FILE
    df = df.rename(columns={'Unnamed: 1': 'Date', 'Unnamed: 2': 'DOW', 'My Prop': 'Hotel TY Occ', 'Comp Set': 'Comp TY Occ',
                            'My Prop.1': 'Hotel LY Occ', 'Comp Set.1': 'Comp LY Occ', 'Unnamed: 12': 'Rank Occ',
                            'Unnamed: 13': 'Rk Chg Occ', 'My Prop.3': 'Hotel TY ADR', 'Comp Set.3': 'Comp TY ADR',
                            'My Prop.4': 'Hotel LY ADR', 'Comp Set.4': 'Comp LY ADR', 'Unnamed: 23': 'Rank ADR',
                            'Unnamed: 24': 'Rk Chg ADR', 'Unnamed: 34': 'Rank RevPAR', 'Unnamed: 35': 'Rk Chg RevPAR'})

    # DROP NULL ROWS OF 'DOW' COLUMN NULL ROWS OF 'DOW' COLUMN DOES NOT CONTAIN ANY RECODE FOR DATE
    df = df[df.DOW.notnull()]

    # CONVERT DATE INTO STANDARD FORMAT
    df['Date'] = pd.to_datetime(df.Date, format='%m/%d/%Y')
    df['Date'] = pd.to_datetime(df.Date, format='%Y-%m-%d')

    # EXTRACTING SHORT FROM DATE COLUMN
    df['Month'] = df['Date'].apply(lambda x: x.strftime('%b'))
    # EXTRACTING YEAR FROM DATE COLUMN
    df['Year'] = df['Date'].dt.year
    # COLLECT REQUIRED COLUMNS FORM DF AND ADDING NEW COLUMNS ['Hotel Cap', 'Comp Cap'] WHICH NOT PRESENT IN INPUT FILE
    df = pd.DataFrame(df, columns=['Date', 'DOW', 'Month', 'Year', 'Hotel Cap', 'Comp Cap', 'Hotel TY Occ', 'Comp TY Occ',
                                   'Hotel LY Occ', 'Comp LY Occ', 'Rank Occ', 'Rk Chg Occ', 'Hotel TY ADR', 'Comp TY ADR',
                                   'Hotel LY ADR', 'Comp LY ADR', 'Rank ADR', 'Rk Chg ADR', 'Rank RevPAR', 'Rk Chg RevPAR'])

    # READ STR FILE
    str_df = pd.read_excel(str_file)
    # CONVERT DATE INTO STANDARD FORMAT
    str_df['Date'] = pd.to_datetime(str_df.Date)
    # COLLECT FIRST ROW VALUE OF FOLLOWING COLUMNS FROM STR FILE TO ASSIGNED CONSTANT VALUE TO THE HOTEL DF COLUMNS
    # hotel = str_df['Hotel Cap'].values[0]
    # comp_c = str_df['Comp Cap'].values[0]
    hotel = 115
    comp_c = 969
    # ASSIGNED CONSTANT VALUE TO THE HOTEL DF COLUMNS ['Hotel Cap', 'Comp Cap']
    df['Hotel Cap'] = hotel
    df['Comp Cap'] = comp_c
    # COLLECTING APPENDED STR DF IN NEW_DF
    new_df = str_df.append(df, sort=False)
    new_df['Date'] = pd.to_datetime(new_df.Date, format='%Y-%m-%d')      # CONVERT DATE INTO STANDARD FORMAT
    # REMOVING OLD DATA FROM STR FILE AFRTER ADDING NEW STR DATA FROM HOTEL FILE
    new_df.drop_duplicates(subset="Date", keep='last', inplace=True)
    new_df['Date'] = new_df['Date'].dt.date  # EXTACTING DATA FROM DATETIME FORMAT
    # DUMP STR WITH UPDATED DATA

    new_df['Hotel TY Sold'] = ((new_df['Hotel Cap'] * new_df['Hotel TY Occ']) / 100).round(2)
    new_df['Comp TY Sold'] = ((new_df['Comp Cap'] * new_df['Comp TY Occ']) / 100).round(2)
    new_df['Hotel LY Sold'] = ((new_df['Hotel Cap'] * new_df['Hotel LY Occ']) / 100).round(2)
    new_df['Comp LY Sold'] = ((new_df['Comp Cap'] * new_df['Comp LY Occ']) / 100).round(2)
    new_df['Hotel Revenue TY'] = ((new_df['Hotel Cap'] * new_df['Hotel TY Occ']) / 100) * new_df['Hotel TY ADR']
    new_df['Comp Revenue TY'] = ((new_df['Comp Cap'] * new_df['Comp TY Occ']) / 100) * new_df['Comp TY ADR']
    new_df['Hotel Revenue LY'] = ((new_df['Hotel Cap'] * new_df['Hotel LY Occ']) / 100) * new_df['Hotel LY ADR']
    new_df['Comp Revenue LY'] = ((new_df['Comp Cap'] * new_df['Comp LY Occ']) / 100) * new_df['Comp LY ADR']

    new_df = pd.DataFrame(new_df, columns=['Date', 'DOW', 'Month', 'Year', 'Hotel Cap', 'Comp Cap', 'Hotel TY Occ','Comp TY Occ',
                                           'Hotel LY Occ', 'Comp LY Occ', 'Rank Occ', 'Rk Chg Occ', 'Hotel TY ADR','Comp TY ADR',
                                           'Hotel LY ADR', 'Comp LY ADR', 'Rank ADR', 'Rk Chg ADR', 'Rank RevPAR','Rk Chg RevPAR',
                                           'Hotel TY Sold','Comp TY Sold', 'Hotel LY Sold', 'Comp LY Sold','Hotel Revenue TY',
                                           'Comp Revenue TY', 'Hotel Revenue LY', 'Comp Revenue LY'])

    print('Saving updated STR_DATA for {} hotel'.format(htl_code))
    new_df.to_excel(str_file, index=False)

# path='D:\Swapnil All Program\Leela_str_project\STR_Data'
# htl_file='D:\Swapnil All Program\Leela_str_project/Cmp_Daily_11292019_23644795.xls'
# str_file='D:\Swapnil All Program\Leela_str_project\STR_Data/STR_Data.xlsx'
# htl_code ='UDP'
# str_copy_paste(htl_file, str_file,htl_code)
