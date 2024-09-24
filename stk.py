import glob
from log import *
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import openpyxl

def xuliStk(dfBom,dfStk):
# Tạo danh sách lưu kết quả
    results = []

    # Duyệt qua từng dòng trong DFBOM_Final
    for idx, row in dfBom.iterrows():
        npl_code = row['npl']  # Mã npl từ DFBOM_Final
        sl_nhu_cau = row['slNhuCau']  # Số lượng nhu cầu
        stk_list = []  # Danh sách các stk đã chọn
        tong_sl = 0  # Tổng số lượng lấy được
        
        # Tìm kiếm trong df_loc_mienThue các dòng tương ứng với mã npl
        df_filtered = dfStk[dfStk['npl'] == npl_code]
    # Duyệt qua từng dòng của df_loc_mienThue tương ứng với npl
        for id, r in df_filtered.iterrows():
            stk = r['stk']  # Số lượng trong kho
            stk_list.append(str(stk))  # Nối tên stk
            tong_sl += r['tongSl']  # Tăng tổng số lượng
            
            # Kiểm tra nếu đã đủ số lượng
            if tong_sl >= sl_nhu_cau:
                break    
              # Nối các stk lại thành chuỗi
        stk_combined = ' -> '.join(stk_list)      
    # Lưu kết quả vào danh sách
    results.append({
        'npl': npl_code,
        'slNhuCau': sl_nhu_cau,
        'stk_combined': stk_combined,
        'tongSl': tong_sl
    })

    # Tạo DataFrame kết quả
    df_result = pd.DataFrame(results)

    # Hiển thị kết quả
    print(df_result)        
def locTrung():
    # Đường dẫn đến file Excel
    file_path = 'temp.xlsx'
    sheet_name = 'temp'  # Thay bằng tên sheet của bạn

    # Tải workbook và sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]

    # Lấy tất cả giá trị trong cột A
    values = []
    for cell in sheet['A']:
        values.append(cell.value)

    # Lọc các giá trị duy nhất
    unique_values = list(set(values))
    print(unique_values)