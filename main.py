from log import *
from common import *
from stk import *
import time
def mainPro():
    try:
        start_time = time.time()
        #420012393
        #locTrung()
        print('------------------------------------------------------------------------------------------------------------------------------------------')
        print('Bắt Đầu Xử Lí Data Từ BOM,tc,Miễn Thuế,Đóng THuế...')
        print('Hãy Kiên Nhẫn Chờ Đợi, Data của bạn khá nặng...')
        print('Quá Trình Xử Lí có thể mất 2 - 4 phút  ...')
        df_BOM = BOMprocess()
        #df_filtered = df_BOM[df_BOM['npl'] == '140400350']
        print('DFBOM Origin :')
        print(df_BOM)
        exportData(df_BOM)
        print('Kết Thúc Copy Data Đã Xử Lí!!!')
        df_BOM['donGia_max'] = df_BOM['donGia_max'].astype(float)  # Chuyển đổi kiểu dữ liệu nếu cần
        df_BOM['donGia_max'] = df_BOM.apply(cifProcess, axis=1)
        print('DFBOM after :')
        print(df_BOM)
        # Danh sách các giá trị không được phép
        df_BOM['donGia_moi'] = df_BOM.apply(
            lambda row: row['donGia'] / row['Tỷ giá'] if row['cif'] not in ['EXW', 'CIF','FCA','FOB','DAP'] else row['donGia'], 
            axis=1
        )  
        df_BOM['slNhuCau'] = df_BOM['Lượng NL, VT thực tế sử dụng để sản xuất một sản phẩm '] * 1434  
        print('DFBOM Final :')
        print(df_BOM)  
        '''with pd.ExcelWriter('temp.xlsx',mode='a', engine='openpyxl',if_sheet_exists='overlay') as writer:
            df_BOM.to_excel(writer, sheet_name='temp', startrow=1, startcol=0, index=False)'''
          #df_BOM[df_BOM['cif']=='MUA TRONG NƯỚC']       
        '''slSp = exportSlsp('420012393')
        
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
            copyReport(value)'''
        end_time = time.time()
        print('------------------------------------------------------------------------------------------------------------------------------------------')        
        print(f'Thời gian thực thi: {end_time - start_time} giây')
    except Exception as ex:
       logExp(str(ex))   

#check xem có phải Hàm main không và show form
if __name__ == "__main__":
    mainPro()