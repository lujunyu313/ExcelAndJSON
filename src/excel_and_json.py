 # encoding: utf-8
__author__ = 'goldlion'
__qq__ = 233424570
__email__ = 'gdgoldlion@gmail.com'

import xlrd
import sys
import getopt
import json
import time
import os
import re

import SheetManager

def mkdir(path):
    if os.path.exists(path):
        return
    os.makedirs(path) 

def usage():
    print("------------------------------------------------------------")
    print("-o: output file path like ./")
    print("-i: input file path like web.xlsx")
    print("-t:  output file type like js/json/lua default = json")
    print("-f:  output file format default = False")
    print("------------------------------------------------------------")

#单表模式
def singlebook():
    opts, args = getopt.getopt(sys.argv[2:], "hfi:o:t:")

    file_type = "json"
    file_format = False
    for op, value in opts:
        if op == "-i":
            file_path = value
        elif op == "-o":
            output_path = value
        elif op == "-t":
            file_type = value
        elif op == "-f":
            file_format = True
        elif op == "-h":
            #TODO 写说明文字
            usage()
            sys.exit()

    if not "file_path" in locals().keys():
        usage()
        sys.exit()
    elif not "output_path" in locals().keys():
        usage()
        sys.exit()

    print '读取并转换文件...'
    SheetManager.addWorkBook(file_path)
    sheetNameList = SheetManager.getSheetNameList()

    mkdir(output_path)
    print '导出文件到：', output_path

    for sheet_name in sheetNameList:
        #单表模式下，被引用的表不会输出
        if SheetManager.isReferencedSheet(sheet_name):
            continue

        sheetJSON = SheetManager.exportJSON(sheet_name)
        
        outputJSON = ''

        if file_type == 'js':
            outputJSON = 'var ' + sheet_name + ' = ' + sheetJSON
    
        oFile = sheet_name + '.' + file_type
        print '正在导出 ' , oFile

        f = file(output_path+oFile, 'w')
        f.write(outputJSON.encode('UTF-8'))
        f.close()

    print '导出文件结束'

#主表模式
def mainbook():
    opts, args = getopt.getopt(sys.argv[2:], "hfi:o:t:")
    file_type = "json"
    file_format = False
    for op, value in opts:
        if op == "-i":
            file_path = value
        elif op == "-o":
            output_path = value
        elif op == "-t":
            file_type = value
        elif op == "-f":
            file_format = True
        elif op == "-h":
            #TODO 写说明文字
            usage()
            sys.exit()

    if not "file_path" in locals().keys():
        usage()
        sys.exit()
    elif not "output_path" in locals().keys():
        usage()
        sys.exit()

    #获取主表各种参数#
    wb = xlrd.open_workbook(file_path)
    sh = wb.sheet_by_index(0)

    workbookPathList = []
    sheetList = []
    for row in range(sh.nrows):
        type = sh.cell(row,0).value

        if type == '__workbook__':
            pass
        else:
            sheetList.append([])
            sheet = sheetList[-1]
            sheet.append(type)

        for col in range(1,sh.ncols):
            value = sh.cell(row,col).value

            if type == '__workbook__' and value != '':
                workbookPathList.append(value)
            elif value != '':
                sheet.append(value)

 #检验是否有重复表名#
    print '表名重复检验...'
    sheetFlag = {}
    for sheet in sheetList:
        if '->' in sheet[0]:
            sheet_output_name = sheet[0].split('->')[1]
        else:
            sheet_output_name = sheet[0]

        sheetFlag[sheet_output_name] = 0    
        
    for sheet in sheetList:
        if '->' in sheet[0]:
            sheet_output_name = sheet[0].split('->')[1]
        else:
            sheet_output_name = sheet[0]

        sheetFlag[sheet_output_name] += 1
        if sheetFlag[sheet_output_name] > 1:
            print sheet_output_name, '表被重复使用'
            return   

    SheetManager.setSheetList(sheetList)

    #加载所有xlsx文件#
    print '读取并转换文件...'
    for workbookPath in workbookPathList:
        #读取所有sheet
        if os.path.isfile(workbookPath+".xlsx"):
            SheetManager.addWorkBook(workbookPath+".xlsx")
        elif os.path.isfile(workbookPath+".xlsm"):
            SheetManager.addWorkBook(workbookPath+".xlsm")
        elif os.path.isfile(workbookPath+".xls"):
            SheetManager.addWorkBook(workbookPath+".xls")
        else:
            print workbookPath, '不存在'

    mkdir(output_path)
    print '导出文件到：', output_path

    #保存所有的输出对象#
    tableList = {}

    #输出所有表#
    for sheet in sheetList:

        #表改名处理
        if '->' in sheet[0]:
            sheet_name = sheet[0].split('->')[0]
            sheet_output_name = sheet[0].split('->')[1]
        else:
            sheet_output_name = sheet_name = sheet[0]

        sheet_output_field = sheet[1:]
        output = ''

        if file_type == 'lua':
            output = "return " + "'" + SheetManager.exportLua(sheet_name,sheet_output_field) + "'"
        else:
            sheetJSON = SheetManager.exportJSON(sheet_name,sheet_output_field,file_format)
            
            if file_type == 'js':
                output = 'var ' + sheet_output_name + ' = '

            output += sheetJSON

            if file_type == 'js':
                output += ';\n'

            
        sheet_output_name = re.sub(r'_',"",sheet_output_name)
        oFile = sheet_output_name + '.' + file_type
        print '正在导出 ' , oFile

        f = file(output_path+oFile, 'w')
        f.write(output.encode('UTF-8'))
        f.close()
    
    print '导出文件结束'


if __name__ == '__main__':
    modelType =  sys.argv[1]
    t1 = time.time()
    if modelType == "singlebook":
        singlebook()
    elif modelType == "mainbook":
        mainbook()
    else:
        # usage()
        sys.exit()
    t2 = time.time()
    print 'use time: ' , (t2 - t1)