import pandas as pd
import sys
sys.path.insert(0,r'C:\programs\revnomix\bin\src\Leela str forecast\src')
import STR_Append as strapp
import STRData_Forecast as strfc
# ======================================= INPUT FILES ==================================================
list_dir ={ 
            # 'tlm':'Leela Mumbai',
            # 'tlb':'Leela Bangalore',
            # 'tlu':'Leela Udaipur',
           # 'tlacd':'Leela Convention Delhi',
            'sfonv':'sfonv'
           #  'tlc':'Leela Chennai',
           #  'tld':'Leela Delhi',
           #  'tlg':'Leela Goa'
            }
for ls_code in list_dir:
    htl_name=list_dir[ls_code]
    htl_code = ls_code
    hotel_file_name = htl_name+'.xls'
    # ======================================================================================================
    str_fname = 'STR_Data.xlsx'
    bdgfc_fname = 'Budget_Forecast.xlsx'
    str_fcname = 'STR_Data_Forecast.xlsx'
    # =======================================================================================================
    mapping_fname = 'yearmonthMAP.xlsx'
    # ==================================== Leela FOLDER PATH ==================================================
    path =r'C:\ftp'
    std_path = path+'/{}/STR_Data'.format(ls_code)
    # =======================================================================================================
    # str_path = std_path + '\\' + htl_code + r''
    htl_file = std_path + '\\' + hotel_file_name
    str_file = std_path + '\\' + str_fname
    bdgfc_file = path + '/{}/Mapping_Files/'.format(ls_code) + bdgfc_fname
    mapping_file = r'C:/programs/revnomix/bin/src/Leela str forecast'+ '//'+ mapping_fname
    # =======================================================================================================
    strapp.str_copy_paste(htl_file, str_file, htl_code)
    print('STR_DATA UPDATED')

    strfc.str_forecast(str_file, mapping_file, bdgfc_file, str_fcname,std_path)
    print('FORECAST GENERATED')
