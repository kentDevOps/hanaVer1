from log import *
from common import *
import time
def mainPro():
    try:
        start_time = time.time()
        #420012393
        df_BOM = BOMprocess()
        '''df_BOM['donGia'] = df_BOM['donGia'].astype(float)  # Chuyển đổi kiểu dữ liệu nếu cần
        df_BOM['donGia'] = df_BOM.apply(cifProcess, axis=1)
        print('DFBOM after :')
        print(df_BOM)'''
        '''slSp = exportSlsp('450005951-03')
        df_BOM['slNhuCau'] = df_BOM['Lượng NL, VT thực tế sử dụng để sản xuất một sản phẩm '] * slSp
        print(df_BOM)'''
        #exportBasicInfor('450005951-03',df_BOM)        
    # Kiểm Tra đã có các File Cần chưa
        '''count_Bom = countFileInFolder('BOM')
        count_Tc = countFileInFolder('tc')
        count_mienThue = countFileInFolder('mienThue')
        count_dongThue = countFileInFolder('dongThue')
        if count_Bom == 0:
            print(f'File BOM Không Tồn Tại , Hãy Copy Vào Folder BOM')
            return
        elif count_Tc == 0:
            print(f'File TC Không Tồn Tại , Hãy Copy Vào Folder tc')
            return
        elif count_mienThue == 0:
            print(f'File MIỄN THUẾ Không Tồn Tại , Hãy Copy Vào Folder mienThue')
            return
        elif count_dongThue == 0:
            print(f'File ĐÓNG THUẾ Không Tồn Tại , Hãy Copy Vào Folder dongThue')
            return
    # xử lí file BOM , LẤY DỮ LIỆU'''
        #df_BOM = BOMprocess()
        '''for value in df_BOM['Mã sản phẩm']:
            exportToReport(value)
            exportSlsp(value)
            exportBasicInfor(value,df_BOM)
            '''
        end_time = time.time()
        print(f'Thời gian thực thi: {end_time - start_time} giây')
    except Exception as ex:
        logExp(str(ex))   

#check xem có phải Hàm main không và show form
if __name__ == "__main__":
    mainPro()