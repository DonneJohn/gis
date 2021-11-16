# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import os
import sys

# 获取Lij 的列表值
def getLijList(df):
    df_li = df.values.tolist()
    result = []
    for s_li in df_li:
        result.append(s_li[4])

    print("LijList:\n{0}".format(result))  # 格式化输出
    return result

#获取LMax的值
def getLMax(LijList):
    Lmax = max(LijList)
    print("Lmax:", Lmax)
    return Lmax

#获取ij 下标列表
def getIJList(df):
    df_li = df.values.tolist()
    result = []
    for s_li in df_li:
        result.append([int(s_li[2]), int(s_li[3])])
    print("i j two lists:\n{0}".format(result))
    return result

# 获取P表G列的所有值
def getPtableGlist(pf):
    pf_li = pf.values.tolist()
    result = []
    for s_li in pf_li:
        result.append(int(s_li[6]))
    print("PtableGlist:\n{0}".format(result))
    return result
    

# 获取p表index行的数据
def getPtableRowData(pf, index):
    pf_li = pf.values.tolist()
    result = pf_li[index]
    print("Ptable index item:\n{0}".format(result))
    return result
    
# 计算Gij
def calcuGij(Pi, Si, Pj, Sj, Lmax, Ci, Cj, Lij):
    result = Pi * Si * Pj * Sj * (Lmax**2) / (Ci * Cj * (Lij**2))
    print("Gij:\n{0}".format(result))
    return result

def saveResultToExcel(result, df, excelFilePath):
    ds = pd.DataFrame(result, columns=['Gij'])
    print(ds)
    #df = df.append(ds, ignore_index=False, sort=False, verify_integrity=True)
    df = df.append(ds, ignore_index=False, sort=False)
    df.to_excel(excelFilePath, index=False, header=True)
    #df = pd.DataFrame(result, columns=['Gij'])
    #df.to_excel(excelFilePath, index=False)



if __name__ == '__main__':
    print("hello")
    dirpath = os.getcwd()
    df = pd.read_excel(dirpath + "/" + "c2.xls")
    LijList = getLijList(df)
    Lmax = getLMax(LijList)
    ijList = getIJList(df)
    #print("ijList len", len(ijList))
    #print(type(ijList))
    
    pf = pd.read_excel(dirpath + "/" + "p.xls")
    pTableGlist = getPtableGlist(pf)
    
    GijList = []
    for index in range(len(ijList)):
        print("i:", ijList[index][0], "> j:", ijList[index][1], " > Lij:", LijList[index], " >ptableI index:", pTableGlist.index(ijList[index][0]),
        " >ptableJ index:", pTableGlist.index(ijList[index][1]), " >")
        i = ijList[index][0]
        j = ijList[index][1]
        Lij = LijList[index]
        pTableIIndex = pTableGlist.index(i)
        pTableJIndex = pTableGlist.index(j)
        pTableIIndexData = getPtableRowData(pf, pTableIIndex)
        Pi = pTableIIndexData[2]
        print("Pi:", Pi)
        Si = pTableIIndexData[4]/1000000
        print("Si:", Si)
        Ci = pTableIIndexData[7]
        print("Ci:", Ci)
        pTableJIndexData = getPtableRowData(pf, pTableJIndex)
        Pj = pTableJIndexData[2]
        print("Pj:", Pj)
        Sj = pTableJIndexData[4]/1000000
        print("Sj:", Sj)
        Cj = pTableJIndexData[7]
        print("Cj:", Cj)
        Gij = calcuGij(Pi, Si, Pj, Sj, Lmax, Ci, Cj, Lij)
        GijList.append(Gij)
        
    systemCodeType = sys.getfilesystemencoding()
    print("结果:\n{0}".format(GijList).decode('UTF-8').encode(systemCodeType)) # 格式化输出
    saveResultToExcel(GijList, df, dirpath + "/" + "c2.xls")