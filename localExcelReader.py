# Local Excel File Pasing Module

from const import CONSTS
import utils
import openpyxl as xl


def readExcel(localXlsFile):
    READ_START_HEADER_ROW_IDX = None
    READ_START_CONTENTS_ROW_IDX = None
    READ_START_HEADER = CONSTS.readStartHeader
    READ_HEADER_LIST = CONSTS.readHeaderList
    # 중복된 열 또는 제외할 열
    EXCEPT_COL_LIST = CONSTS.exceptColList
    EXCEPT_WORDS = CONSTS.exceptWords
    
    # parsing 결과 dictionary { sheet명 : contents }
    # sheet가 여러개일 때를 고려.
    readResult = {}
    
    # 엑셀 Pasing
    wb = xl.load_workbook(localXlsFile)

    for sheet_nm in wb.sheetnames:
        sheet = wb[sheet_nm]
        
        # 시작 Header 행 찾는 부분
        READ_START_HEADER_ROW_IDX = utils.getHeaderRowIdx(sheet, READ_START_HEADER)
        if READ_START_HEADER_ROW_IDX == None: READ_START_HEADER_ROW_IDX = 1

        # 시작 Contents 행 찾는 부분
        READ_START_CONTENTS_ROW_IDX = utils.getContentsRowIdx(sheet, READ_START_HEADER_ROW_IDX)
        if READ_START_CONTENTS_ROW_IDX == None: raise Exception('Excel Contents Empty.')

        
        # Contents 수집하는 부분
        # 중복된 열 또는 제외할 열을 제거한 후, 시작한다.
        utils.deleteDuplicatedColumn(sheet, EXCEPT_COL_LIST)
        resultDic = utils.getSheetContents(sheet, READ_START_HEADER_ROW_IDX, READ_START_CONTENTS_ROW_IDX, READ_HEADER_LIST, EXCEPT_WORDS)
        
        # readResult[sheet_nm] = resultDic
        
    wb.close()
    
    return resultDic