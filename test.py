#导入pyExcel模块
import os
from  pyWord   import *

def handleWord(fileAddr):
    setWordVisible(True,False)
    openWord(fileAddr)
    closeWord()

def listFile(dir):    
    for root,dirs,files in os.walk(dir):
        for file in files:
            fileAddr = os.path.join(root,file)
            fileType = file.split(".")[-1]
            if("~" not in file):
                if(fileType == "doc" or fileType == "docx" or fileType == "wps"):
                    handleWord(fileAddr)


#word目录
wordDir = "C:\Documents and Settings\Administrator\桌面\汇总";
listFile(wordDir)        
#放到最后
quitWord()



