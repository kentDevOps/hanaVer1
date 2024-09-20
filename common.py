import glob
from log import *
import pandas as pd
import numpy as np
from openpyxl import load_workbook

def getRelativeFile(folder_name,file_name):
    strAbsPath = os.path.abspath(sys.argv[0])
    strCrrPath = os.path.dirname(strAbsPath)
    strFilePath = os.path.join(strCrrPath,folder_name) + file_name
    file_path = glob.glob(strFilePath)
    return file_path
def BOMprocess():
    file_path = getRelativeFile('BOM','\*BOM*.xlsx')
    df = pd.read_excel(file_path[0])
    result = df.groupby(['Mã sản phẩm', 'Mã NPL'], as_index=False).agg({'Lượng NL, VT thực tế sử dụng để sản xuất một sản phẩm ': 'sum'})
    # Lọc mã NPL Có  KD (Đóng Thuế hoặc là TC)
    df_Kd = result[result['Mã NPL'].str.contains('KD', na=False)]
    print(df_Kd)
    # Lọc mã NPL không KD (Miễn Thuế)
    df_None_Kd = result[~result['Mã NPL'].str.contains('KD', na=False)]
    print('df_None_Kd:')  
    print(df_None_Kd)   
    # Đọc file miễn thuế để lấy ra cột tên nguyên liệu
    df_loc_mienThue = mienThueProcess()
    # Đọc file Đóng thuế và TC để lấy ra cột tên nguyên liệu
    df_loc_dongThue = dongThue_Tc_Process()
    #lấy cột mã NPL từ df_None_kd
    ma_npl_df_kd = df_None_Kd.iloc[:,1].rename('npl')
    print('df_loc_mienThue:') 
    print(df_loc_mienThue)
    #lọc lấy tên hàng theo ma_npl_df_kd từ df_loc_mienThue
    '''filtered_df2 = df_loc_mienThue[df_loc_mienThue['Mã NPL'].isin(df_None_Kd['Mã NPL'])]
    print('filtered_df2:') 
    print(filtered_df2)
    df_mienThue_tenHang = pd.merge(df_None_Kd,filtered_df2[['Mã NPL', 'tenHang']],on='Mã NPL',how='left')'''
    df_loc_mienThue = df_loc_mienThue.drop_duplicates(subset=['Mã NPL'])
    df_mienThue_tenHang = pd.merge(df_None_Kd,df_loc_mienThue,on='Mã NPL',how='left')
    print('Lấy Ra Được frame chứa Tên Hàng Theo Các Mã Được Miễn Thuế:')
    print(df_mienThue_tenHang)
    #lọc lấy tên hàng theo ma_npl_df_kd từ df_loc_dongThue
    #lấy cột mã NPL từ df_kd
    ma_npl_df_kd_CoThue = df_Kd.iloc[:,1].rename('npl')
    df_loc_dongThue = df_loc_dongThue.drop_duplicates(subset=['Mã NPL'])
    df_dongThue_tenHang = pd.merge(df_Kd,df_loc_dongThue,on='Mã NPL',how='left')    
    print('Lọc Tên Hàng Theo Các Mã Phải Đóng Thuế:') 
    print(df_dongThue_tenHang)
    df_BOM = pd.concat([df_dongThue_tenHang, df_mienThue_tenHang], axis=0)
    print(df_BOM)
    colNpl =result.iloc[:,1]
    colSoLuong =result.iloc[:,2]
    colTenHang =result.iloc[:,3]
    dfReport = pd.read_excel('temp.xlsx',sheet_name='rp')
    lastRowRp = len(dfReport) + 1
    dfReport.iloc[:,:] = np.nan    
    print(result)
    with pd.ExcelWriter('temp.xlsx' , engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
        dfReport.to_excel(writer, sheet_name='rp', index=False, startcol= 0,startrow=22, header=False) 
        colNpl.to_excel(writer, sheet_name='rp', index=False, startcol= 1,startrow=22, header=False)  
        colSoLuong.to_excel(writer, sheet_name='rp', index=False, startcol= 5,startrow=22, header=False)    
def mienThueProcess():
    file_path = getRelativeFile('mienThue','\*mienThue*.xlsx')
    df = pd.read_excel(file_path[0]) 
    #df_loc = df[['Mã NPL/SP','Tên hàng']]
    df_loc = df.iloc[2:,[39,41]]
    df_loc.columns = ['Mã NPL','tenHang']
    return df_loc
def dongThue_Tc_Process():
    file_path_dongThue = getRelativeFile('dongThue','\*dongThue*.xlsx')
    file_path_Tc = getRelativeFile('tc','\*TC*.xlsx')
    df_dongThue = pd.read_excel(file_path_dongThue[0]).iloc[2:,[3,45]]
    df_dongThue.columns = ['Mã NPL','tenHang']
    print("dongThue File :")
    print(df_dongThue)
    df_tc = pd.read_excel(file_path_Tc[0],sheet_name='TC').iloc[:,[8,10]]
    df_tc.columns = ['Mã NPL','tenHang']
    print("Tc File :")
    print(df_tc)    
    df_merged = pd.concat([df_dongThue, df_tc], axis=0)
    print("Tong Hop File :")
    print(df_merged)    
    return df_merged
def tcTest():
    file_path_Tc = getRelativeFile('tc','\*TC*.xlsx')
    df_tc = pd.read_excel(file_path_Tc[0],sheet_name="TC")
    
    #df_tc = pd.read_excel(file_path_Tc[0])#.iloc[2:,[8,10]]
   # df_tc = df_tc.iloc[:,3]
    print(df_tc.shape) 
    return df_tc