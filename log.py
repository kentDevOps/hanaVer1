from datetime import datetime
import os,sys

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
def countFileInFolder(folder_name):
    sys_Path = os.path.abspath(sys.argv[0])
    base_Path = os.path.dirname(sys_Path)
    fol_path = os.path.join(base_Path,folder_name)
    all_files = os.listdir(fol_path)
    print(len(all_files))
    return all_files