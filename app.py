import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os

st.set_page_config(page_title="보험 계약 환산기", layout="wide")
st.title("📊 보험 계약 실적 환산기 (컨벤션 & 썸머 기준)")

uploaded_file = st.file_uploader("📂 계약 목록 Excel 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file:
    base_filename = os.path.splitext(uploaded_file.name)[0]
    download_filename = f"{base_filename}_환산결과.xlsx"
    df = pd.read_excel(uploaded_file)

    st.subheader("✅ 업로드된 데이터")
    st.dataframe(df)

    # 환산율 계산 함수
    def classify(row):
        보험사 = str(row["보험사"])
        납기 = int(row["납입기간"])
        상품명 = str(row.get("상품명", ""))

        is_생보 = "생명" in 보험사
        is_한화생명 = 보험사 == "한화생명"
        is_손보_250 = 보험사 in ["한화손해보험", "삼성화재", "흥국화재", "KB손해보험"]
        is_손보_200 = 보험사 in ["롯데손해보험", "메리츠화재", "현대해상", "DB손해보험", "MG손해보험", "하나손해보험", "AIG손해보험"]
        is_저축_제외 = any(x in 상품명 for x in ["저축", "연금", "일시납", "적립금", "태아보험일시납"])

        # 컨벤션 기준
        if is_한화생명:
            conv_rate = 150
        elif is_손보_250:
            conv_rate = 250
        elif is_손보_200:
            conv_rate = 200
        elif is_생보:
            conv_rate = 100 if 납기 >= 10 else 50
        else:
            conv_rate = 0

        # 썸머 기준
        if is_저축_제외:
            summ_rate = 0
        elif is_한화생명:
            summ_rate = 150 if 납기 >= 10 else 100
        elif is_생보:
            summ_rate = 100 if 납기 >= 10 else 30
        elif is_손보_250:
            summ_rate = 200 if 납기 >= 10 else 100
        else:
            summ_rate = 100 if 납기 >= 10 else 50

        return pd.Series([conv_rate, summ_rate])

    df[["컨벤션율", "썸머율"]] = df.apply(classify, axis=1)
    df["컨벤션환산금액"] = df["보험료"] * df["컨벤션율"] / 100
    df["썸머환산금액"] = df["보험료"] * df["썸머율"] / 100

    # 스타일링용 복사본
    styled_df = df.copy()
    styled_df["계약일자"] = pd.to_datetime(styled_df["계약일자"].astype(str), format="%Y%m%d").dt.strftime("%Y년%m월%d일")
    styled_df["납입기간"] = styled_df["납입기간"].astype(str) + "년"
    styled_df["보험료"] = styled_df["보험료"].map("{:,.0f} 원".format)
    styled_df["컨벤션율"] = styled_df["컨벤션율"].astype(str) + "%"
    styled_df["썸머율"] = styled_df["썸머율"].astype(str) + "%"
    styled_df["컨벤션환산금액"] = styled_df["컨벤션환산금액"].map("{:,.0f} 원".format)
    styled_df["썸머환산금액"] = styled_df["썸머환산금액"].map("{:,.0f} 원".format)

    # 합계 계산
    convention_sum = df["컨벤션환산금액"].sum()
    summer_sum = df["썸머환산금액"].sum()

    # 엑셀 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "환산결과"
    for r_idx, row in enumerate(dataframe_to_rows(styled_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # 표 삽입
    end_col_letter = ws.cell(row=1, column=styled_df.shape[1]).column_letter
    end_row = ws.max_row
    table_ref = f"A1:{end_col_letter}{end_row}"
    table = Table(displayName="환산결과표", ref=table_ref)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = style
    ws.add_table(table)

    # 열 너비 자동 조정
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 10

    # 총합 행 추가
    sum_row = ws.max_row + 2
    ws.cell(row=sum_row, column=7, value="총 합계").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=sum_row, column=8, value="{:,.0f} 원".format(convention_sum)).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=sum_row, column=9, value="{:,.0f} 원".format(summer_sum)).alignment = Alignment(horizontal="center", vertical="center")
    for col in [7, 8, 9]:
        ws.cell(row=sum_row, column=col).font = Font(bold=True)

    # 저장
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    # Streamlit 출력
    st.subheader("📄 환산 결과 요약")
    st.dataframe(styled_df)

    st.subheader("📈 총합")
    st.write(f"▶ 컨벤션 기준 합계: **{convention_sum:,.0f} 원**")
    st.write(f"▶ 썸머 기준 합계: **{summer_sum:,.0f} 원**")

    st.download_button(
        label="📥 환산 결과 엑셀 다운로드",
        data=excel_output,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("📤 계약 목록 Excel 파일(.xlsx)을 업로드해주세요.")
