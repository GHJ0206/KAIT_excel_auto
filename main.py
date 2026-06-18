from table_generators import t001_기록물철_보존기간_0000
from table_generators import t002_기록물철_보존기간_0000_복합포함
from table_generators import t003_기록물건_보존기간_3년
from table_generators import t004_예시


TABLES = {
    "1": ("기록물철별 보존기간 책정 현황(0000)", t001_기록물철_보존기간_0000.run),
    "2": ("기록물철별 보존기간 책정 현황(0000)_복합포함", t002_기록물철_보존기간_0000_복합포함.run),
    "3": ("기록물건별 보존기간 책정 현황(최근 3년간)", t003_기록물건_보존기간_3년.run),
    "4": ("예시 파일 입니다", t004_예시.run)
}


print("===== 표 생성 프로그램 =====")


for key, value in TABLES.items():
    print(f"{key}. {value[0]}")


choice = input("번호 선택: ")


if choice in TABLES:

    institution = input(
        "분석할 기관명을 띄어쓰기 없이 정확하게 입력해주세요: "
    )

    year = input(
        "기준년도를 한글 없이 숫자만 입력해주세요: "
    )

    TABLES[choice][1](institution, year)


else:
    print("잘못된 번호입니다.")