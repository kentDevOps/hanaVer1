from log import *
from common import *
import time
def mainPro():
    try:
        start_time = time.time()

    # Kiểm Tra đã có các File Cần chưa
        count_Bom = countFileInFolder('BOM')
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
    # xử lí file BOM , LẤY DỮ LIỆU
        BOMprocess()
        end_time = time.time()
        print(f'Thời gian thực thi: {end_time - start_time} giây')
    except Exception as ex:
        logExp(str(ex))   

#check xem có phải Hàm main không và show form
if __name__ == "__main__":
    mainPro()