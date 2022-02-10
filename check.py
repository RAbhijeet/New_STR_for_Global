import pandas as pd
import os
from datetime import datetime,date

def readStr(htl, strpath, sheet,str_file):

    for file in os.listdir(strpath):
        if file.__contains__(htl) == True:
            print(file)
            newname = file.rsplit("_", 1)[1]
            datename, ext = newname.split('.')
            file_date = pd.to_datetime(datename, format="%d%b%Y").date()
            fdate_date = file_date.__format__('%d.%m.%Y')
            date_col = file_date.__format__('%d-%b-%y')
            file_year = file_date.__format__('%Y')
            file_month = file_date.__format__('%m')
            file_day = int(file_date.__format__('%d'))

            #REading The STR File

            df = pd.read_excel(strpath + file, sheet_name=sheet, skiprows=22, header=None)
            df.drop(columns=0, inplace=True)

            df2 = df.loc[2:22, 2:file_day+1]
            df_Trans = df2.T
            df_trans1 = df_Trans.drop(columns=[5, 6, 7, 10, 11, 12, 13, 14, 17, 18, 19, 22])
            df_trans1 = df_trans1.rename(columns={2:"Day",3:"Hotel TY Occ",4:"Comp TY Occ",8:"Hotel Occ % Change",
                                                9:"Comp Occ % Change", 15:"Hotel TY ADR", 16:"Comp TY ADR",
                                                20:"Hotel ADR % Change", 21:"Comp ADR % Change"})
            df_trans1["Month"] = int(file_month)
            df_trans1["Year"] = int(file_year)
            df_trans1["Date"] = pd.to_datetime(df_trans1[["Year","Month","Day"]])
            # df_trans1["Month"] = pd.DatetimeIndex(df_trans1["Date"]).month
            df_trans1["Month"] = df_trans1["Date"].dt.strftime("%b")
            df_trans1["DOW"] = df_trans1["Date"].dt.weekday_name.str[:3]
            # df_trans1["Date"] = pd.DatetimeIndex(df_trans1["Date"]).date
            # df_trans1.to_excel(strpath+"\STR_Data_new.xlsx",index= False)

            ###### Calculation of Last Year Occ and ADR

            df_trans1["Hotel LY Occ"] = df_trans1["Hotel TY Occ"]/(1+(df_trans1["Hotel Occ % Change"]/100))
            df_trans1["Hotel LY ADR"] = df_trans1["Hotel TY ADR"] / (1 + (df_trans1["Hotel ADR % Change"] / 100))
            df_trans1["Comp LY Occ"] = df_trans1["Comp TY Occ"] / (1 + (df_trans1["Comp Occ % Change"] / 100))
            df_trans1["Comp LY ADR"] = df_trans1["Comp TY ADR"] / (1 + (df_trans1["Comp ADR % Change"] / 100))

            # Adding blank columns for rank
            df_trans1["Rank Occ"] = ""
            df_trans1["Rk Chg Occ"] = ""
            df_trans1["Rank ADR"] = ""
            df_trans1["Rk Chg ADR"] = ""
            df_trans1["Rank RevPar"] = ""
            df_trans1["Rk Chg RevPar"] = ""

            df_new = pd.DataFrame(df_trans1, columns=['Date', 'DOW', 'Month', 'Year', 'Hotel Cap', 'Comp Cap',
                                                         'Hotel TY Occ', 'Comp TY Occ', 'Hotel LY Occ',
                                                         'Comp LY Occ', 'Rank Occ', 'Rk Chg Occ', 'Hotel TY ADR',
                                                         'Comp TY ADR', 'Hotel LY ADR', 'Comp LY ADR',
                                                         'Rank ADR', 'Rk Chg ADR', 'Rank RevPAR', 'Rk Chg RevPAR'])

            df_new.to_excel(strpath + "\STR_Data_new.xlsx", index=False)

            # READ STR FILE
            str_df = pd.read_excel(str_file)
            # CONVERT DATE INTO STANDARD FORMAT
            str_df['Date'] = pd.to_datetime(str_df.Date)
            # COLLECT FIRST ROW VALUE OF FOLLOWING COLUMNS FROM STR FILE TO ASSIGNED CONSTANT VALUE TO THE HOTEL DF COLUMNS
            # hotel = str_df['Hotel Cap'].values[0]
            # comp_c = str_df['Comp Cap'].values[0]
            hotel = 115
            comp_c = 630
            # ASSIGNED CONSTANT VALUE TO THE HOTEL DF COLUMNS ['Hotel Cap', 'Comp Cap']
            df_new['Hotel Cap'] = hotel
            df_new['Comp Cap'] = comp_c
            # COLLECTING APPENDED STR DF IN NEW_DF
            df_fin = str_df.append(df_new, sort=False)
            df_fin['Date'] = pd.to_datetime(df_fin.Date, format='%Y-%m-%d')  # CONVERT DATE INTO STANDARD FORMAT
            # REMOVING OLD DATA FROM STR FILE AFRTER ADDING NEW STR DATA FROM HOTEL FILE
            df_fin.drop_duplicates(subset="Date", keep='last', inplace=True)
            df_fin['Date'] = df_fin['Date'].dt.date  # EXTACTING DATA FROM DATETIME FORMAT
            # DUMP STR WITH UPDATED DATA

            # df_fin['Hotel TY Sold'] = ((df_fin['Hotel Cap'] * df_fin['Hotel TY Occ']) / 100).round(2)
            # df_fin['Comp TY Sold'] = ((df_fin['Comp Cap'] * df_fin['Comp TY Occ']) / 100).round(2)
            # df_fin['Hotel LY Sold'] = ((df_fin['Hotel Cap'] * df_fin['Hotel LY Occ']) / 100).round(2)
            # df_fin['Comp LY Sold'] = ((df_fin['Comp Cap'] * df_fin['Comp LY Occ']) / 100).round(2)
            # df_fin['Hotel Revenue TY'] = ((df_fin['Hotel Cap'] * df_fin['Hotel TY Occ']) / 100) * df_fin['Hotel TY ADR']
            # df_fin['Comp Revenue TY'] = ((df_fin['Comp Cap'] * df_fin['Comp TY Occ']) / 100) * df_fin['Comp TY ADR']
            # df_fin['Hotel Revenue LY'] = ((df_fin['Hotel Cap'] * df_fin['Hotel LY Occ']) / 100) * df_fin['Hotel LY ADR']
            # df_fin['Comp Revenue LY'] = ((df_fin['Comp Cap'] * df_fin['Comp LY Occ']) / 100) * df_fin['Comp LY ADR']

            df_fin['Hotel TY Sold'] = ((df_fin['Hotel Cap'] * df_fin['Hotel TY Occ']) / 100)
            df_fin['Comp TY Sold'] = ((df_fin['Comp Cap'] * df_fin['Comp TY Occ']) / 100)
            df_fin['Hotel LY Sold'] = ((df_fin['Hotel Cap'] * df_fin['Hotel LY Occ']) / 100)
            df_fin['Comp LY Sold'] = ((df_fin['Comp Cap'] * df_fin['Comp LY Occ']) / 100)
            df_fin['Hotel Revenue TY'] = ((df_fin['Hotel Cap'] * df_fin['Hotel TY Occ']) / 100) * df_fin['Hotel TY ADR']
            df_fin['Comp Revenue TY'] = ((df_fin['Comp Cap'] * df_fin['Comp TY Occ']) / 100) * df_fin['Comp TY ADR']
            df_fin['Hotel Revenue LY'] = ((df_fin['Hotel Cap'] * df_fin['Hotel LY Occ']) / 100) * df_fin['Hotel LY ADR']
            df_fin['Comp Revenue LY'] = ((df_fin['Comp Cap'] * df_fin['Comp LY Occ']) / 100) * df_fin['Comp LY ADR']

            df_fin = pd.DataFrame(df_fin,
                                  columns=['Date', 'DOW', 'Month', 'Year', 'Hotel Cap', 'Comp Cap', 'Hotel TY Occ',
                                           'Comp TY Occ',
                                           'Hotel LY Occ', 'Comp LY Occ', 'Rank Occ', 'Rk Chg Occ', 'Hotel TY ADR',
                                           'Comp TY ADR',
                                           'Hotel LY ADR', 'Comp LY ADR', 'Rank ADR', 'Rk Chg ADR', 'Rank RevPAR',
                                           'Rk Chg RevPAR',
                                           'Hotel TY Sold', 'Comp TY Sold', 'Hotel LY Sold', 'Comp LY Sold',
                                           'Hotel Revenue TY',
                                           'Comp Revenue TY', 'Hotel Revenue LY', 'Comp Revenue LY'])

            print('Saving updated STR_DATA for {} hotel'.format(htl))
            df_fin.to_excel(str_file, index=False)


if __name__ =='__main__':
    # file = 'P&L_Apr21.xlsx'
    htl = "sfonv"
    sheet = "Daily by Month"
    str_path = r'C:\ftp\sfonv\STR_Data/'

    str_file = r'C:\ftp\sfonv\STR_Data/STR_Data.xlsx'

    readStr(htl, str_path, sheet,str_file)