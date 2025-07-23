import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import os

st.set_page_config(page_title="보험 계약 환산기", layout="wide")
st.title("📊 보험 계약 실적 환산기 (컨벤션 & 썸머 기준)")

uploaded_file = st.file_uploader("📂 계약 목록 Excel 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file:
    # 업로드 파일명 추출
    base_filename = os.path.splitext(uploaded_file.name)[0]
    download_filename = f"{base_filename}_환산결과.xlsx"

    # 데이터 읽기
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error("❗ Excel 파일을 읽는 중 오류가 발생했습니다.")
        st.stop()

    st.subheader("✅ 업로드된 데이터")
    st.dataframe(df)

    # 환산 기준표 로딩
    try:
        rate_df = pd.read_csv("conversion_rates.csv")
    except FileNotFoundError:
        st.error("⚠️ conversion_rates.csv 파일이 없습니다.")
        st.stop()

    def classify(row):
        생보사 = ["한화생명"]
        손보_250 = ["한화손해보험", "삼성화재", "흥국화재", "KB손해보험"]
        손보_200 = ["롯데손해보험", "메리츠화재", "현대해상", "DB손해보험", "MG손해보험", "하나손해보험", "AIG손해보험"]
        보험사 = row["보험사"]
        납기 = int(row["납입기간"])

        if 보험사 in 생보사:
            유형 = "생명보험"
            세부 = 보험사
        elif 보험사 in 손보_250 + 손보_200:
            유형 = "손해보험"
            세부 = 보험사
        else:
            유형 = "생명보험" if "생명" in 보험사 else "손해보험"
            세부 = "기타생보" if 유형 == "생명보험" else "기타손보"

        기간조건 = "10년 이상" if 납기 >= 10 else "10년 미만"
        match = rate_df[
            (rate_df["보험사"] == 세부) &
            (rate_df["유형"] == 유형) &
            (rate_df["납입기간조건"] == 기간조건)
        ]
        if match.empty:
            return pd.Series([0, 0])
        else:
            return pd.Series([match["컨벤션율"].values[0], match["썸머율"].values[0]])

    # 환산율 계산
    df[["컨벤션율", "썸머율"]] = df.apply(classify, axis=1)
    df["컨벤션환산금액"] = df["보험료"] * df["컨벤션율"] / 100
    df["썸머환산금액"] = df["보험료"] * df["썸머율"] / 100

    # 스타일 적용용 복사본
    styled_df = df.copy()
    styled_df["계약일자"] = pd.to_datetime(styled_df["계약일자"].astype(str), format="%Y%m%d").dt.strftime("%Y년%m월%d일")
    styled_df["납입기간"] = styled_df["납입기간"].astype(str) + "년"
    styled_df["보험료"] = styled_df["보험료"].map("{:,.0f} 원".format)
    styled_df["컨벤션율"] = styled_df["컨벤션율"].astype(str) + "%"
    styled_df["썸머율"] = styled_df["썸머율"].astype(str) + "%"
    styled_df["컨벤션환산금액"] = styled_df["컨벤션환산금액"].map("{:,.0f} 원".format)
    styled_df["썸머환산금액"] = styled_df["썸머환산금액"].map("{:,.0f} 원".format)

    # 합계
    convention_sum = df["컨벤션환산금액"].sum()
    summer_sum = df["썸머환산금액"].sum()

    # 엑셀로 저장 (openpyxl 스타일 포함)
    wb = Workbook()
    ws = wb.active
    ws.title = "환산결과"
    for r_idx, row in enumerate(dataframe_to_rows(styled_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
    for column_cells in ws.columns:
    max_length = 0
    column = column_cells[0].column_letter
    for cell in column_cells:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    ws.column_dimensions[column].width = max_length + 2

    sum_row = ws.max_row + 2
    ws.cell(row=sum_row, column=7, value="총 합계").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=sum_row, column=8, value="{:,.0f} 원".format(convention_sum)).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=sum_row, column=9, value="{:,.0f} 원".format(summer_sum)).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=sum_row, column=7).font = Font(bold=True)
    ws.cell(row=sum_row, column=8).font = Font(bold=True)
    ws.cell(row=sum_row, column=9).font = Font(bold=True)

    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    st.subheader("📄 환산 결과 요약")
    st.dataframe(styled_df)

    st.subheader("📈 총합")
    st.write(f"▶ 컨벤션 기준 합계: **{df['컨벤션환산금액'].sum():,.0f} 원**")
    st.write(f"▶ 썸머 기준 합계: **{df['썸머환산금액'].sum():,.0f} 원**")

    st.download_button(
        label="📥 환산 결과 엑셀 다운로드",
        data=excel_output,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("📤 계약 목록 Excel 파일(.xlsx)을 업로드해주세요.")
