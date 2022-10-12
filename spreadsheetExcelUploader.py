# Google Spreadsheet Excel Uploader

import gspread
from gspread import Spreadsheet
from gspread import Worksheet
from gspread.utils import ValueInputOption
from datetime import datetime
import os
import re
import utils
import localExcelReader
from const import CONSTS
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

def connectSpreadsheet(json_keyfile_name, spreadsheet_url):
    """
    Google Spreadsheet 에 접속하고, doc 객체를 반환한다.

    Args:
        json_keyfile_name (string): Google Dev Spreadsheet API Key 파일명
        spreadsheet_url (string): 접근하려는 spreadsheet 주소

    Returns:
        Spreadsheet: 접근 성공한 Spreadsheet 객체
    """
    
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_name, CONSTS.JSON_KEY_FILE_SCOPE)
    gc = gspread.authorize(credentials)
    # 스프레스시트 문서 가져오기 
    spreadsheetDoc = gc.open_by_url(spreadsheet_url)
    
    return spreadsheetDoc


def uploadLocalXlsToSpreadsheet(spreadsheetDoc: Spreadsheet, localXlsFile):
    sheetName = CONSTS.SHEET_NAME
    # 시트 선택하기
    try:
        worksheet = spreadsheetDoc.worksheet(sheetName)
    except gspread.exceptions.WorksheetNotFound as err:
        worksheet = spreadsheetDoc.add_worksheet(sheetName, 1, 1, 0)
    
    # 이전에 설정된 'MergedCell' 설정을 해지
    expectedCellAddr, endAddr = getMergedCellAddr(worksheet, CONSTS.STR_EXCEPT_CONTENTS_LIST)
    worksheet.unmerge_cells(expectedCellAddr + ':' + endAddr)
    
    # 기존 내용을 삭제
    requests = {"requests": [{"updateCells": {"range": {"sheetId": worksheet._properties['sheetId']}, "fields": "*"}}]}
    spreadsheetDoc.batch_update(requests)
    
    # 업로드 전 필요한 수정작업을 수행한다.
    # Key 는 row num, value 는 사용 내역 건
    editedXlsContents = editXlsContents(list(localExcelReader.readExcel(localXlsFile).values()))
    
    # value_input_option=ValueInputOption.raw 시, cell value 앞에 ' 이 붙는다. ex) A1='2020. 5. 5
    worksheet.append_rows(editedXlsContents, value_input_option=ValueInputOption.user_entered)

def getMergedCellAddr(worksheet: Worksheet, mergeCellWord: str):
    # 'mergeCellWord' 셀 시작 주소
    expectedCellAddr = getHighlightAddress(worksheet, mergeCellWord)
    if expectedCellAddr:
        # 'mergeCellWord' 셀 끝 주소
        endAddr = chr(ord(expectedCellAddr[0:1]) + 4) + expectedCellAddr[1:]
        
        return expectedCellAddr, endAddr
    
    return '', ''

def editXlsContents(xlsContentList: list):
    """
    spreadsheet에 올라갈 contents를 편집하는 메서드.
    필요한 메서드를 추가하며 사용하면 되는데
    CONSTS.STR_EXCEPT_CONTENTS_LIST 을 만나면 편집을 중단하도록 작성하는 것을 권장.

    Args:
        xlsContentList (list): xls Contents List

    Returns:
        list: 편집된 xls Contents List
    """    
    
    addColumn(xlsContentList, 2, '카드(J)')
    addColumn(xlsContentList, 2, '')
    remainOnlyNum(xlsContentList, 5)
    calculateSubColumn(xlsContentList, 4, 5, 4)
    removeColumn(xlsContentList, 5)
    convertDateFormatTypes(xlsContentList, 0)
    
    return xlsContentList


def addColumn(xlsContentList: list, colIdx: int, cellText: str):
    """
    xlsContentList 의 colIdx에 해당하는 열에 일괄적으로 cellText를 추가한다.

    Args:
        xlsContentList (list): excel content list
        colIdx (int): 0 ~ n 범위의 숫자
        cellText (str): 삽입하려는 문자열
        
    Raises:
        Exception: colIdx 가 0 미만일 때 발생.
    """
    
    if colIdx < 0:
        raise Exception('Wrong enter column index range.')
    
    for content in xlsContentList:
        if CONSTS.STR_EXCEPT_CONTENTS_LIST in content:
            break
        
        content.insert(colIdx, cellText)


def remainOnlyNum(xlsContentList: list, colIdx: int):
    """
    xlsContentList 의 colIdx에 해당하는 열의 문자열에 일괄적으로 숫자를 제외한 문자를 제외한다.
    ex) 1,200P -> 1200 

    Args:
        xlsContentList (list): excel content list
        colIdx (int): 0 ~ n 범위의 숫자
        
    Raises:
        Exception: colIdx 가 0 미만일 때 발생.
    """
    
    if colIdx < 0:
        raise Exception('Wrong enter column index range.')
    
    for content in xlsContentList:
        if CONSTS.STR_EXCEPT_CONTENTS_LIST in content:
            break
        
        if len(content) > colIdx:
            # re module은 정규표현식을 지원한다. 
            content[colIdx] = re.sub(r"[^0-9]", "", content[colIdx])


def calculateSubColumn(xlsContentList: list, xColIdx: int, yColIdx: int, destColIdx: int):
    """
    xlsContentList 의 xColIdx 와 yColIdx 에 해당하는 열의 값을 subtraction 한 값을 destColIdx 에 해당하는 모든 열에 저장한다.

    Args:
        xlsContentList (list): excel content list
        xColIdx (int): 뺄셈을 할 숫자가 적힌 열을 의미하는 0 ~ n 범위의 숫자
        yColIdx (int): 뺄셈을 할 숫자가 적힌 열을 의미하는 0 ~ n 범위의 숫자
        destColIdx (int): 연산의 결과가 저장될 열을 의미하는 0 ~ n 범위의 숫자
    
    Raises:
        Exception: xColIdx, yColIdx, destColIdx 가 0 미만일 때 발생.
    """    
    
    if xColIdx < 0 or yColIdx < 0 or destColIdx < 0:
        raise Exception('Wrong enter column index range.')
    
    for content in xlsContentList:
        if CONSTS.STR_EXCEPT_CONTENTS_LIST in content:
            break
        
        maxLen = len(content)
        if maxLen > xColIdx and maxLen > yColIdx and maxLen > destColIdx:
            # re module은 정규표현식을 지원한다. 
            content[destColIdx] = int(content[xColIdx]) - int(content[yColIdx])


def removeColumn(xlsContentList: list, colIdx: int):
    """
    xlsContentList 의 colIdx에 해당하는 열을 일괄적으로 제거한다.

    Args:
        xlsContentList (list): excel content list
        colIdx (int): 0 ~ n 범위의 숫자
        
    Raises:
        Exception: colIdx 가 0 미만일 때 발생.
    """
    
    if colIdx < 0:
        raise Exception('Wrong enter column index range.')
    
    for content in xlsContentList:
        if CONSTS.STR_EXCEPT_CONTENTS_LIST in content:
            break
        
        if len(content) > colIdx:
            del content[colIdx]


def convertDateFormatTypes(xlsContentList: list, colIdx: int):
    """
    colIdx에 해당하는 열의 값을 달력 위젯이 표시되는 Date format 'yyyy. m. d' 의 형식으로 일괄적으로 변경한다.
    선택하여 붙여넣기를 통해(Ctrl + Shift + V) 날짜 서식이 적용된 Cell에 붙여넣어서 사용한다.

    Args:
        xlsContentList (list): excel content list
        colIdx (int): 0 ~ n 범위의 숫자
        
    Raises:
        Exception: colIdx 가 0 미만일 때 발생.
    """
    
    if colIdx < 0:
        raise Exception('Wrong enter column index range.')
    
    # https://docs.python.org/ko/3/library/datetime.html#strftime-strptime-behavior
    for content in xlsContentList:
        if CONSTS.STR_EXCEPT_CONTENTS_LIST in content:
            break
        
        if len(content) > colIdx:
            try:
            # 2022.05.12 -> 2022. 5. 12
                date_time_obj = datetime.strptime(content[colIdx], '%Y.%m.%d')
                convertDate = date_time_obj.strftime("%Y. %m. %d")
                # python datetime 모듈에서는 0으로 채운값을 리턴하기에 0을 제거하는 과정.
                content[colIdx] = convertDate.replace(' 0', ' ')
            except ValueError as err:
                # continue
                content.insert(colIdx, 'Failed Convert')


def updateSpreadsheet(spreadsheetDoc: Spreadsheet):
    """
    배경색을 입히거나, 속성을 적용하려 하는 용도로 사용한다.

    Args:
        spreadsheetDoc (Spreadsheet): Spreadsheet 객체
    """
    
    # sheetName = 'autoGeneratedSheet'
    worksheet = spreadsheetDoc.worksheet(CONSTS.SHEET_NAME)
    
    # 고정지출 셀 색상 설정
    worksheet.format(getHighlightAddressList(worksheet, CONSTS.FIXED_EXPENSE_WORDS), {
    "backgroundColor": {
      "red": 0.0,
      "green": 1.0,
      "blue": 0.0
    }})
    
    # '제외된 내역 리스트' 셀 색상 설정 및 가운데 정렬
    expectedCellAddr, endAddr = getMergedCellAddr(worksheet, CONSTS.STR_EXCEPT_CONTENTS_LIST)
    worksheet.format(expectedCellAddr, {
    "backgroundColor": {
      # https://www.tug.org/pracjourn/2007-4/walden/color.pdf
      "red": 0.8,
      "green": 1.0,
      "blue": 1.0
    },
    "horizontalAlignment": "CENTER",
    })
    
    # '제외된 내역 리스트' ROW 병합
    worksheet.merge_cells(expectedCellAddr + ':' + endAddr)


def getHighlightAddress(worksheet: Worksheet, word: str):
    """
    인자로 받은 문자열이 포함된 셀의 주소를 반환한다. 
    """
    
    patternedWord = re.compile('.*' + word + '.*')
    return worksheet.find(patternedWord).address if worksheet.find(patternedWord) else ''


def getHighlightAddressList(worksheet: Worksheet, wordList: list):
    """
    인자로 받은 문자열이 포함된 셀의 주소가 담긴 리스트를 반환한다. 
    """    
    resultList = []
    
    try :
        for word in wordList:
            # 패턴매칭을 통해 포함된 cell의 주소를 얻는다.
            patternedWord = re.compile('.*' + word + '.*')
            resultList.append(worksheet.find(patternedWord).address)
    except Exception as e:
        # FIXED_EXPENSE_WORDS 중, 엑셀에 비연속적으로 나오는 경우 고려.
        pass

    
    return resultList

def essentialFileCheck():
    targetExcelFolder = os.path.join(os.getcwd(), CONSTS.EXCEL_FILE_FOLDER)
    credentialFile = os.path.join(os.getcwd(), CONSTS.JSON_KEY_FILE_NAME)
    
    result = True
    
    if not os.path.isdir(targetExcelFolder):
        os.mkdir(targetExcelFolder)
        print('[ERROR] "' + targetExcelFolder + '" 에 파싱 대상인 File 이 없습니다.')
        result = False
    
    if not os.path.exists(credentialFile):
        print('[ERROR] keyfile 인 "' + credentialFile + '" 이 없습니다.')
        result = False
    
    return result
        
    
def main():
    
    if essentialFileCheck():
        
        localXlsFile = None
        
        # 현재 위치로부터 excel 폴더 내 가장 최근 수정한 xls 파일을 찾는다.
        root = os.listdir(os.getcwd())
        for file in root:
            if os.path.isdir(file) and file == CONSTS.EXCEL_FILE_FOLDER:
                if len(os.listdir(file)) != 0:
                    localXlsFile = utils.getMostRecentFile(file)
                    break

        if localXlsFile != None:
            # 파싱하려는 xls 파일의 확장자가 'xls'인 경우 'xlsx'로 변환한다.
            if localXlsFile.endswith("xls"):
                utils.convertXlsToXlsx(localXlsFile)
                # 가장 최근에 생성된 파일을 기준으로 한다.
                localXlsFile = utils.getMostRecentFile(file)
        
            spreadsheetDoc = connectSpreadsheet(CONSTS.JSON_KEY_FILE_NAME, CONSTS.SPREADSHEET_URL)
            uploadLocalXlsToSpreadsheet(spreadsheetDoc, localXlsFile)
            updateSpreadsheet(spreadsheetDoc)


if __name__ == '__main__':
    main()
