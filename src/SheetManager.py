# encoding: utf-8
__author__ = 'goldlion'
__qq__ = 233424570
__email__ = 'gdgoldlion@gmail.com'

import xlrd
import json
import Sheet

sheetDict = {}
sheetNameList = []
sheetList = []

def setSheetList(sheets):
    for sheet in sheets:
        sheetList.append(sheet)

def isNeedExport(name):
    for sheet in sheetList:
        if '->' in sheet[0]:
            sheet_name = sheet[0].split('->')[0]
        else:
            sheet_name = sheet[0]
        if sheet_name == name:
            return True
    return False

def addWorkBook(filepath):
    wb = xlrd.open_workbook(filepath)
    for sheet_index in range(wb.nsheets):
        sh = wb.sheet_by_index(sheet_index)
        if isNeedExport(sh.name):
            sheet = Sheet.openSheet(sh)
            addSheet(sheet)

def addSheet(sheet):
    sheetDict[sheet.name] = sheet
    sheetNameList.append(sheet.name)

def getSheet(name):
    return sheetDict[name]

def getSheetNameList():
    return sheetNameList

def exportJSON(name,sheet_output_field = [],format = False):
    return sheetDict[name].toJSON(sheet_output_field,format)

def exportLua(name,sheet_output_field = []):
    return sheetDict[name].toLua(sheet_output_field)

def isReferencedSheet(name):
    for sheetName in sheetDict:
        if name in sheetDict[sheetName].referenceSheets:
            return  True

    return False