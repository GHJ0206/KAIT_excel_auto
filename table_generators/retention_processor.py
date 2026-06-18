"""
표명 : 기록물철별 보존기간 책정 현황

기능 :
기록물분류기준표 Excel 파일을 분석하여
보존기간별 기록물철 개수와 비율을 자동 산출한다.
"""



def build_input_path(institution, year):
    pass


def get_sheet_names(file_path):
    pass


def select_sheet(workbook):
    pass


def find_retention_column(sheet):
    pass


def classify_retention_values(sheet, retention_col, header_row):
    pass


def create_headers(counts):
    pass


def create_output_excel(counts, headers):
    pass


def process_excel(file_path, sheet_name):
    pass


def run_local():

    print("retention_processor 정상 실행됨")