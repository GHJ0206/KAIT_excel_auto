from table_generators import retention_processor


TABLES = {
    "1": (
        "기록물철별 보존기간 책정 현황",
        retention_processor.run_local
    ),

    "2": (
        "추후 추가 예정",
        None
    )
}


print("===== 표 생성 프로그램 =====")


for key, value in TABLES.items():
    print(f"{key}. {value[0]}")


choice = input("번호 선택: ")


if choice in TABLES:

    if TABLES[choice][1] is None:
        print("아직 구현되지 않은 기능입니다.")

    else:
        TABLES[choice][1]()

else:
    print("잘못된 번호입니다.")