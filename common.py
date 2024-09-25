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
def cifProcess(row):
    if row['cif'] == 'EXW':
        return row['donGia_max'] * 1.02
    elif row['cif'] == 'FCA':
        return row['donGia_max'] * 1.005
    elif row['cif'] == 'FOB':
        return row['donGia_max'] * 1.005
    elif row['cif'] == 'DAP':
        return row['donGia_max'] * 0.99   
    else:
        return row['donGia_max'] 

def BOMprocess():
    file_path = getRelativeFile('BOM','\*BOM*.xlsx')
    df = pd.read_excel(file_path[0])
    print('---------------------------------------------------------------------')
    print('In Ra BOM nguyên bản...')
    print(df)
    result1 = df.groupby(['Mã sản phẩm', 'Mã NPL'], as_index=False).agg({'Lượng NL, VT thực tế sử dụng để sản xuất một sản phẩm ': 'sum'})
    print('---------------------------------------------------------------------')
    print('In Ra BOM Sau Khi Gộp Theo Mã SP, Mã NL , Gộp Lượng Định Mức...')
    print(result1)     
    # Gộp thêm cột "NGÀY TỜ KHAI XUẤT KHẨU" vào result
    result = result1.merge(df[['Mã sản phẩm', 'Mã NPL', 'NGÀY TỜ KHAI XUẤT KHẨU','Tỷ giá']].drop_duplicates(),
                        on=['Mã sản phẩm', 'Mã NPL'],
                        how='left')
    print('---------------------------------------------------------------------')
    print('In Ra BOM Sau Khi Thêm Ngày Xuất NPL...')
    print(result)   
    # Lọc mã NPL Có  KD (Đóng Thuế hoặc là TC)
    df_Kd = result[result['Mã NPL'].str.contains('KD', na=False)]
    df_Kd.rename(columns={df_Kd.columns[1]: 'npl'}, inplace=True)
    print('---------------------------------------------------------------------')
    print('df_Kd:') 
    print(df_Kd)
    # Lọc mã NPL không KD (Miễn Thuế)
    df_None_Kd = result[~result['Mã NPL'].str.contains('KD', na=False)]
    df_None_Kd.rename(columns={df_None_Kd.columns[1]: 'npl'}, inplace=True)
    print('---------------------------------------------------------------------')
    print('df_None_Kd:')  
    print(df_None_Kd)   
    # Đọc file miễn thuế để lấy ra cột tên nguyên liệu
    df_loc_mienThue = mienThueProcess()
    print('---------------------------------------------------------------------')
    print('df_loc_mienThue:') 
    print(df_loc_mienThue)    
    # Đọc file Đóng thuế và TC để lấy ra cột tên nguyên liệu
    df_loc_dongThue = dongThue_Tc_Process()
    #lấy cột mã NPL từ df_None_kd
    ma_npl_df_kd = df_None_Kd.iloc[:,1].rename('npl')
    print('---------------------------------------------------------------------')
    print('df_loc_Dong Thue:') 
    print(df_loc_dongThue)
    #lọc lấy tên hàng theo ma_npl_df_kd từ df_loc_mienThue
    #df_loc_mienThue = df_loc_mienThue.drop_duplicates(subset=['npl']) 
    df_mienThue_tenHang = pd.merge(df_None_Kd,df_loc_mienThue,on='npl',how='left')
    df_mienThue_tenHang = df_mienThue_tenHang.drop_duplicates(subset=['Mã sản phẩm', 'npl'])
    print('---------------------------------------------------------------------')
    print('Lấy Ra Được frame chứa Tên Hàng Theo Các Mã Được Miễn Thuế:')
    print(df_mienThue_tenHang)
    #lọc lấy tên hàng theo ma_npl_df_kd từ df_loc_dongThue
    #lấy cột mã NPL từ df_kd
    ma_npl_df_kd_CoThue = df_Kd.iloc[:,1].rename('npl')
    df_loc_dongThue = df_loc_dongThue.drop_duplicates(subset=['npl'])
    df_dongThue_tenHang = pd.merge(df_Kd,df_loc_dongThue,on='npl',how='left')  
    df_dongThue_tenHang = df_dongThue_tenHang.drop_duplicates(subset=['Mã sản phẩm', 'npl'])  
    print('---------------------------------------------------------------------')
    print('Lọc Tên Hàng Theo Các Mã Phải Đóng Thuế:') 
    print(df_dongThue_tenHang)
    df_BOM = pd.concat([df_dongThue_tenHang, df_mienThue_tenHang], axis=0)
    print('---------------------------------------------------------------------')
    print(df_BOM)
    #df_STK = pd.concat([df_loc_mienThue, df_loc_dongThue], axis=0)
    return df_BOM
def exportToReport(maSp):
    df = pd.DataFrame({'maSP': [maSp]})
    with pd.ExcelWriter('temp.xlsx' , engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
        df['maSP'].to_excel(writer, sheet_name='rp', index=False, startcol= 11,startrow=3, header=False) 
def mienThueProcess():
    file_path = getRelativeFile('mienThue','\*mienThue*.xlsx')
    df = pd.read_excel(file_path[0]) 
    #df_loc = df[['Mã NPL/SP','Tên hàng']]
    df_loc = df.iloc[2:,[39,41,40,46,42,36,43,20,21,47]]
    df_loc.columns = ['npl','tenHang','hs','dv','xuatXu','cif','donGia','stk','ntk','tongSl']
    '''df_loc['donGia'] = pd.to_numeric(df_loc['donGia'])
    max_value = df_loc['donGia'].max()
    df_loc['donGia'] = max_value'''
    #df_loc['donGia'] = df_loc['donGia'].astype(float)  # Chuyển đổi kiểu dữ liệu nếu cần
    #df_loc['donGia'] = df_loc.apply(cifProcess, axis=1)
    print('---------------------------------------------------------------------')
    print(df_loc)
    df_loc['donGia_max'] = df_loc.groupby(['npl','cif'])['donGia'].transform('max')
    '''df_grouped = df_loc.groupby('Mã NPL').agg({
        'tenHang': 'first',  # Giữ lại giá trị đầu tiên của cột 'Mã NPL' hoặc có thể thay đổi tùy theo nhu cầu
        'hs': 'first',  # Tổng số lượng theo mã sản phẩm
        'dv': 'first',  # Giữ lại giá trị đầu tiên (hoặc sử dụng hàm phù hợp khác)
        'xuatXu': 'first',  # Giữ lại giá trị đầu tiên của đơn vị
        'cif': 'first',
        'donGia': 'max'  # Giữ lại giá trị đầu tiên của cột cif (nếu cần)
    }).reset_index()'''  
    return df_loc
def dongThue_Tc_Process():
    file_path_dongThue = getRelativeFile('dongThue','\*dongThue*.xlsx')
    file_path_Tc = getRelativeFile('tc','\*TC*.xlsx')
    df_dongThue = pd.read_excel(file_path_dongThue[0]).iloc[2:,[3,45,44,50,46,40,47,24,25,49]]
    df_dongThue.columns = ['npl','tenHang','hs','dv','xuatXu','cif','donGia','stk','ntk','tongSl']
    #chuyển đổi giá trị thành số và định dạng số thập phân
    df_dongThue['donGia'] = pd.to_numeric(df_dongThue['donGia'], errors='coerce')
    df_dongThue['donGia'] = df_dongThue['donGia'].round(6)  # Giữ tối đa 6 chữ số sau dấu phẩy    
    df_dongThue['donGia_max'] = df_dongThue.groupby(['npl','cif'])['donGia'].transform('max')
    '''df_dongThue['donGia'] = pd.to_numeric(df_dongThue['donGia'])
    max_value = df_dongThue['donGia'].max()
    df_dongThue['donGia'] = max_value''' 
    print('---------------------------------------------------------------------')
    print("dongThue File :")
    print(df_dongThue)
    df_tc = pd.read_excel(file_path_Tc[0],sheet_name='TC').iloc[:,[8,10,9,11,21,21,13,3,4,12]]
    df_tc.columns = ['npl','tenHang','hs','dv','xuatXu','cif','donGia','stk','ntk','tongSl']
    #chuyển đổi giá trị thành số và định dạng số thập phân
    df_tc['donGia'] = pd.to_numeric(df_tc['donGia'], errors='coerce')
    df_tc['donGia'] = df_tc['donGia'].round(6)  # Giữ tối đa 6 chữ số sau dấu phẩy

    df_tc['donGia_max'] = df_tc.groupby(['npl','cif'])['donGia'].transform('max')
    '''df_tc['donGia'] = pd.to_numeric(df_tc['donGia'])
    max_value = df_tc['donGia'].max()
    df_tc['donGia'] = max_value '''
    print('---------------------------------------------------------------------')    
    print("Tc File :")
    print(df_tc)    
    df_merged = pd.concat([df_dongThue, df_tc], axis=0)
    print('---------------------------------------------------------------------')
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
    print('---------------------------------------------------------------------')
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
def exportData(df_BOM):
    #dfReport = pd.read_excel('data.xlsx',sheet_name='temp')
    #lastRowRp = len(dfReport) + 1
    #dfReport.iloc[:,:] = np.nan    
    print('Bắt Đầu Ghi Mã NPL, Tên Hàng , Định Mức Sản Phẩm...')
    #with pd.ExcelWriter('data.xlsx' , engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
    #    df_BOM.to_excel(writer, sheet_name='temp', index=False, startcol= 0,startrow=1, header=False) 
    df_BOM.to_excel('data.xlsx', sheet_name='rp', index=False)