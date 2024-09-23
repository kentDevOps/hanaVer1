import glob
from log import *
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import openpyxl

def getRelativeFile(folder_name,file_name):
    strAbsPath = os.path.abspath(sys.argv[0])
    strCrrPath = os.path.dirname(strAbsPath)
    strFilePath = os.path.join(strCrrPath,folder_name) + file_name
    file_path = glob.glob(strFilePath)
    return file_path
def BOMprocess():
    file_path = getRelativeFile('BOM','\*BOM*.xlsx')
    df = pd.read_excel(file_path[0])
    print('In Ra BOM nguyên bản...')
    print(df)
    result = df.groupby(['Mã sản phẩm', 'Mã NPL'], as_index=False).agg({'Lượng NL, VT thực tế sử dụng để sản xuất một sản phẩm ': 'sum'})
    print('In Ra BOM Sau Khi Gộp Theo Mã SP, Mã NL , Gộp Lượng Định Mức...')
    print(result)   
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
    '''colNpl =df_BOM.iloc[:,1]
    colSoLuong =df_BOM.iloc[:,2]
    colTenHang =df_BOM.iloc[:,3]
    colHs =df_BOM.iloc[:,4]
    colDv =df_BOM.iloc[:,5]
    dfReport = pd.read_excel('temp.xlsx',sheet_name='rp')
    lastRowRp = len(dfReport) + 1
    dfReport.iloc[:,:] = np.nan    
    print('Bắt Đầu Ghi Mã NPL, Tên Hàng , Định Mức Sản Phẩm...')
    with pd.ExcelWriter('temp.xlsx' , engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
        dfReport.to_excel(writer, sheet_name='rp', index=False, startcol= 0,startrow=22, header=False) 
        colNpl.to_excel(writer, sheet_name='rp', index=False, startcol= 1,startrow=22, header=False)  
        colTenHang.to_excel(writer, sheet_name='rp', index=False, startcol= 2,startrow=22, header=False)    
        colSoLuong.to_excel(writer, sheet_name='rp', index=False, startcol= 5,startrow=22, header=False) 
        colHs.to_excel(writer, sheet_name='rp', index=False, startcol= 3,startrow=22, header=False) 
        colDv.to_excel(writer, sheet_name='rp', index=False, startcol= 4,startrow=22, header=False)
    print('Quá Trình Ghi Mã NPL, Tên Hàng , Định Mức Sản Phẩm Kết Thúc!!!') '''  
    return df_BOM
def exportToReport(maSp):
    df = pd.DataFrame({'maSP': [maSp]})
    with pd.ExcelWriter('temp.xlsx' , engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
        df['maSP'].to_excel(writer, sheet_name='rp', index=False, startcol= 11,startrow=3, header=False) 
def mienThueProcess():
    file_path = getRelativeFile('mienThue','\*mienThue*.xlsx')
    df = pd.read_excel(file_path[0]) 
    #df_loc = df[['Mã NPL/SP','Tên hàng']]
    df_loc = df.iloc[2:,[39,41,40,46]]
    df_loc.columns = ['Mã NPL','tenHang','hs','dv']
    return df_loc
def dongThue_Tc_Process():
    file_path_dongThue = getRelativeFile('dongThue','\*dongThue*.xlsx')
    file_path_Tc = getRelativeFile('tc','\*TC*.xlsx')
    df_dongThue = pd.read_excel(file_path_dongThue[0]).iloc[2:,[3,45,44,50]]
    df_dongThue.columns = ['Mã NPL','tenHang','hs','dv']
    print("dongThue File :")
    print(df_dongThue)
    df_tc = pd.read_excel(file_path_Tc[0],sheet_name='TC').iloc[:,[8,10,9,11]]
    df_tc.columns = ['Mã NPL','tenHang','hs','dv']
    print("Tc File :")
    print(df_tc)    
    df_merged = pd.concat([df_dongThue, df_tc], axis=0)
    print("Tong Hop File :")
    print(df_merged)    
    return df_merged
def tcTest():
    file_path_Tc = getRelativeFile('BOM','\*BOM*.xlsx')
    df_tc = pd.read_excel(file_path_Tc[0])
    #unique_values = df_tc['Mã sản phẩm'].unique()
    #df_tc = pd.read_excel(file_path_Tc[0])#.iloc[2:,[8,10]]
   # df_tc = df_tc.iloc[:,3]
    df_filtered = df_tc.drop_duplicates(subset=['Mã sản phẩm', 'Số lượng sản phẩm'], keep='first')
    print(df_filtered) 
    #print(unique_values) 
    #print(str(len(unique_values)))
    return df_tc
def exportSlsp(maSP):
    file_path_Tc = getRelativeFile('BOM','\*BOM*.xlsx')
    df_tc = pd.read_excel(file_path_Tc[0]) 
    df_filtered = df_tc.drop_duplicates(subset=['Mã sản phẩm', 'Số lượng sản phẩm'], keep='first')  
    filtered_df = df_filtered[df_filtered['Mã sản phẩm'] == maSP]
    print(filtered_df)
    so_luong_san_pham = filtered_df[['Số lượng sản phẩm']]
    print('slsp : ')
    print(so_luong_san_pham)
    so_luong_san_pham.columns = ['slsp']
    with pd.ExcelWriter('temp.xlsx' , engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
        so_luong_san_pham['slsp'].to_excel(writer, sheet_name='rp', index=False, startcol= 12,startrow=6, header=False)
    return so_luong_san_pham.iloc[0,0] 
def exportBasicInfor(maSP,df_BOM):
    colNpl =df_BOM.iloc[:,1]
    colSoLuong =df_BOM.iloc[:,2]
    colTenHang =df_BOM.iloc[:,3]
    colHs =df_BOM.iloc[:,4]
    colDv =df_BOM.iloc[:,5]
    dfReport = pd.read_excel('temp.xlsx',sheet_name='rp')
    lastRowRp = len(dfReport) + 1
    dfReport.iloc[:,:] = np.nan    
    print('Bắt Đầu Ghi Mã NPL, Tên Hàng , Định Mức Sản Phẩm...')
    with pd.ExcelWriter('temp.xlsx' , engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
        dfReport.to_excel(writer, sheet_name='rp', index=False, startcol= 0,startrow=22, header=False) 
        colNpl.to_excel(writer, sheet_name='rp', index=False, startcol= 1,startrow=22, header=False)  
        colTenHang.to_excel(writer, sheet_name='rp', index=False, startcol= 2,startrow=22, header=False)    
        colSoLuong.to_excel(writer, sheet_name='rp', index=False, startcol= 5,startrow=22, header=False) 
        colHs.to_excel(writer, sheet_name='rp', index=False, startcol= 3,startrow=22, header=False) 
        colDv.to_excel(writer, sheet_name='rp', index=False, startcol= 4,startrow=22, header=False)    