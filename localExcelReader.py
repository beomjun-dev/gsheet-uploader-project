# Local Excel File Pasing Module

import utils
import openpyxl as xl


def readExcel(localXlsFile):
    READ_START_HEADER_ROW_IDX = None
    READ_START_CONTENTS_ROW_IDX = None
    READ_START_HEADER = '이용일'
    READ_HEADER_LIST = [ READ_START_HEADER, '이용가맹점', '원금', '적립예정' ]
    # 중복된 열 또는 제외할 열
    EXCEPT_COL_LIST = ['I']
    EXCEPT_WORDS = [ 
                    '버스', '지하철', '동부생명보험', '맞춤서비스수수료', 
                    '메리츠화재해상보험주식회사', '한화손해보험', '롯데손해보험주식회사', '삼성화재해상보험' 
                   ]
    
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