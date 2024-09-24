from datetime import datetime
import os,sys
import shutil

def logExp(ex):
    strLogPath = getRelativePath()
    strTime =  datetime.now().strftime("%Y%m%d")
    strFilePath = strLogPath + r"\log_" + strTime + ".txt"
    strContents = "[{}] {}".format( datetime.now().strftime("%Y%m%d %H:%M:%S"),ex)
    if not os.path.exists(strFilePath):
        with open(strFilePath,"x") as logFile:
            logFile.writelines("\n")
            logFile.writelines(strContents)
    else:
        with open(strFilePath,"a") as logFile:
            logFile.writelines("\n")
            logFile.writelines(strContents)   
def getRelativePath():
    strAbsPath = os.path.abspath(sys.argv[0])
    strCrrPath = os.path.dirname(strAbsPath)
    strFilePath = os.path.join(strCrrPath,"log")
    if not os.path.exists(strFilePath):
        os.makedirs(strFilePath)
        return strFilePath
    else:      
        return strFilePath
def getRelativePath1(folder):
    strAbsPath = os.path.abspath(sys.argv[0])
    strCrrPath = os.path.dirname(strAbsPath)
    strFilePath = os.path.join(strCrrPath,folder)
    if not os.path.exists(strFilePath):
        os.makedirs(strFilePath)
        return strFilePath
    else:      
        return strFilePath    
def countFileInFolder(folder_name):
    sys_Path = os.path.abspath(sys.argv[0])
    base_Path = os.path.dirname(sys_Path)
    fol_path = os.path.join(base_Path,folder_name)
    all_files = os.listdir(fol_path)
    print(len(all_files))
    return all_files
def copyReport(mSp):
    strFilePath = getRelativePath1('report')
# Đường dẫn đến file Excel gốc
    source_path = "temp.xlsx"
    destination_path = os.path.join(strFilePath,str(mSp) + ".xlsx")

    # Đường dẫn đến file Excel mới sau khi sao chép và đổi tên
    

    # Sao chép file và đổi tên
    shutil.copy(source_path, destination_path)

    print(f"File đã được sao chép và đổi tên thành {destination_path}")    