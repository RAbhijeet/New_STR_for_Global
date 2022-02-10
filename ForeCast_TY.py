"""
CREATED BY : Chakradhar T.


c_year = current year

"""


import pandas as pd
import numpy as np
import datetime
import matplotlib.pyplot as ptl


def polypfit(dataframe):
    ptl.plot(dataframe)
    ptl.legend(dataframe)
    return ptl


year = datetime.datetime.today()
c_year = year.year - 1
# c_year = 2021


def strForecast(file_df, column):

    # COLLECTING REQUIRED COLUMNS FROM INPUT DF
    df = file_df[['Date', 'Comp Rms TY', 'Comp Rev TY']]
    df = df.groupby('Date').sum().reset_index()
    df['Date'] = pd.to_datetime(df['Date'], format="%Y-%m-%d")
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%b-%y')
    df['Month'] = pd.DatetimeIndex(df['Date']).month
    df['Year'] = pd.DatetimeIndex(df['Date']).year
    raw_df = pd.DataFrame(df, columns=['Month', 'Year', 'Comp Rms TY', 'Comp Rev TY'])
    df = df[df.Year <= c_year]
    df = pd.DataFrame(df, columns=['Month', 'Year', 'Comp Rms TY', 'Comp Rev TY'])
    df = df.groupby(['Month', 'Year']).sum().reset_index()
    df = pd.DataFrame(df, columns=['Month', 'Year', 'Comp Rms TY', 'Comp Rev TY'])

    # FORECAST CALCULATION FROM PROVIDED COLUMNS
    avg_df = df[column].mean()                      # average of column

    m_df = df[['Month', column]]
    m_df = m_df.groupby(['Month']).mean().reset_index()
    m_df = m_df.rename(columns={'Month': 'mth', column: 'Base'})
    df['mth'] = m_df['mth']
    df['Base'] = m_df['Base']

    y1 = df['Year'].max()-1

    df_1 = df[['Month', 'Year', column]]
    df_1 = df_1[df_1['Year'] == y1]
    df_1 = df_1.groupby(['Month', 'Year']).mean().reset_index()
    df_1 = df_1.rename(columns={column: y1})

    y2 = df['Year'].max()
    df_2 = df[['Month', 'Year', column]]
    df_2 = df_2[df_2['Year'] == y2]
    df_2 = df_2.groupby(['Month', 'Year']).mean().reset_index()
    df_2 = df_2.rename(columns={column: y2})

    df[y1] = df_1[y1]
    df[y2] = df_2[y2]

    df['Index'] = df['Base']/avg_df
    df['YOY'] = df[y2]/df[y1]
    df=pd.DataFrame(df)

    df['YOY Index'] = df['Index'] * df['YOY']

    z = np.polyfit(df['Month'], df[column], 2)

    x = c_year - min(df.Year) + 1

    poly_fit = (z[0]*x**2) + (z[1]*x) + z[2]

    m_df['Poly Fit'] = poly_fit
    df['Poly Fit'] = m_df['Poly Fit']
    df['Pre'] = df['Index'] * df['Poly Fit']
    df['Forecast'] = df['YOY Index'] * df['Poly Fit']

    df1 = pd.DataFrame(df[['mth', 'Forecast']])
    df1=df1.replace(np.inf, np.nan)
    df1 = df1.dropna()
    return df1

