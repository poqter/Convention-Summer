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
    # 1. 필요한 컬럼만 로드
    columns_needed = ["계약일", "보험사", "상품명", "납입기간", "계속보험료", "쉐어율"]
    df = pd.read_excel(uploaded_file, usecols=columns_needed)

    # 2. 컬럼명 정규화 (내부에서 쓸 이름으로 바꿈)
    df.rename(columns={
        "계약일": "계약일자",
        "계속보험료": "보험료"
    }, inplace=True)

    # 3. 필수 항목 체크
    required_columns = {"계약일자", "보험사", "상품명", "납입기간", "보험료", "쉐어율"}
    if not required_columns.issubset(df.columns):
        st.error("❌ 업로드된 파일에 다음 항목이 모두 포함되어 있어야 합니다:\n" + ", ".join(required_columns))
        st.stop()

    # 필수 컬럼 체크
    required_columns = {"계약일자", "보험사", "상품명", "납입기간", "보험료", "쉐어율"}
    if not required_columns.issubset(df.columns):
        st.error("❌ 업로드된 파일에 다음 항목이 모두 포함되어 있어야 합니다:\n" + ", ".join(required_columns))
        st.stop()

    # 쉐어율 누락 확인
    if df["쉐어율"].isnull().any():
        st.error("❌ '쉐어율'에 빈 값이 포함되어 있습니다. 모든 행에 값을 입력해주세요.")
        st.stop()

    def classify(row):
        보험사원본 = str(row["보험사"])
        납기 = int(row["납입기간"])
        상품명 = str(row.get("상품명", ""))

        # 보험사 분류
        if 보험사원본 == "한화생명":
            보험사 = "한화생명"
        elif "생명" in 보험사원본:
            보험사 = "기타생보"
        elif 보험사원본 in ["한화손보", "삼성화재", "흥국화재", "KB손보보"]:
            보험사 = 보험사원본
        elif any(x in 보험사원본 for x in ["손해", "화재", "손보"]):
            보험사 = "기타손보"
        else:
            보험사 = 보험사원본  # 분류되지 않은 보험사는 그대로 사용

        # 조건 플래그
        is_한화생명 = 보험사 == "한화생명"
        is_기타생보 = 보험사 == "기타생보"
        is_손보_250 = 보험사 in ["한화손보", "삼성화재", "흥국화재", "KB손보"]
        is_기타손보 = 보험사 == "기타손보"

        # 컨벤션 기준
        if is_한화생명:
            conv_rate = 150
        elif is_손보_250:
            conv_rate = 250
        elif is_기타손보:
            conv_rate = 200
        elif is_기타생보:
            conv_rate = 100 if 납기 >= 10 else 50
        else:
            conv_rate = 0

        # 썸머 기준
        if is_한화생명:
            summ_rate = 150 if 납기 >= 10 else 100
        elif is_기타생보:
            summ_rate = 100 if 납기 >= 10 else 30
        elif is_손보_250:
            summ_rate = 200 if 납기 >= 10 else 100
        elif is_기타손보:
            summ_rate = 100 if 납기 >= 10 else 50
        else:
            summ_rate = 0

        return pd.Series([conv_rate, summ_rate])

    # 환산율 적용
    df[["컨벤션율", "썸머율"]] = df.apply(classify, axis=1)

    # 쉐어율 강제 변환 (퍼센트 서식/소수/문자 모두 대응)
    df["쉐어율"] = df["쉐어율"].apply(lambda x: float(str(x).replace('%','')) if pd.notnull(x) else x)

    # 실적 보험료 계산 (쉐어율 적용)
    df["실적보험료"] = df["보험료"] * df["쉐어율"] / 100

    # 환산금액 계산
    df["컨벤션환산금액"] = df["실적보험료"] * df["컨벤션율"] / 100
    df["썸머환산금액"] = df["실적보험료"] * df["썸머율"] / 100

    # 합계
    performance_sum = df["실적보험료"].sum()
    convention_sum = df["컨벤션환산금액"].sum()
    summer_sum = df["썸머환산금액"].sum()

    # 스타일링용 복사본
    styled_df = df.copy()
    # ✅ 계약일자 날짜 형식 변환 (오류 발생 방지 + 사용자 경고 메시지 추가)
    styled_df["계약일자"] = pd.to_datetime(styled_df["계약일자"], errors="coerce")

    # ⛔ 변환 실패한 항목이 있는 경우 경고 표시 (Streamlit 환경)
    invalid_dates = styled_df[styled_df["계약일자"].isna()]
    if not invalid_dates.empty:
        st.warning(f"⚠️ {len(invalid_dates)}건의 계약일자가 날짜로 인식되지 않았습니다. 엑셀에서 '2025-07-23'처럼 정확한 형식으로 입력해주세요.")

    # ✅ 날짜를 "YYYY-MM-DD" 문자열로 변환
    styled_df["계약일자"] = styled_df["계약일자"].dt.strftime("%Y-%m-%d")
    styled_df["납입기간"] = styled_df["납입기간"].astype(str) + "년"
    styled_df["보험료"] = styled_df["보험료"].map("{:,.0f} 원".format)
    styled_df["쉐어율"] = styled_df["쉐어율"].astype(str) + " %"
    styled_df["실적보험료"] = styled_df["실적보험료"].map("{:,.0f} 원".format)
    styled_df["컨벤션율"] = styled_df["컨벤션율"].astype(str) + " %"
    styled_df["썸머율"] = styled_df["썸머율"].astype(str) + " %"
    styled_df["컨벤션환산금액"] = styled_df["컨벤션환산금액"].map("{:,.0f} 원".format)
    styled_df["썸머환산금액"] = styled_df["썸머환산금액"].map("{:,.0f} 원".format)

    # 엑셀 출력
    wb = Workbook()
    ws = wb.active
    ws.title = "환산결과"

    for r_idx, row in enumerate(dataframe_to_rows(styled_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # 표 적용
    end_col_letter = ws.cell(row=1, column=styled_df.shape[1]).column_letter
    end_row = ws.max_row
    table = Table(displayName="환산결과표", ref=f"A1:{end_col_letter}{end_row}")
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    table.tableStyleInfo = style
    ws.add_table(table)

    # 열 너비
    for column_cells in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = max_len + 10

    # 총합 행
    sum_row = ws.max_row + 2
    ws.cell(row=sum_row, column=8, value="총 합계").alignment = Alignment(horizontal="center")
    ws.cell(row=sum_row, column=9, value="{:,.0f} 원".format(performance_sum)).alignment = Alignment(horizontal="center")
    ws.cell(row=sum_row, column=10, value="{:,.0f} 원".format(convention_sum)).alignment = Alignment(horizontal="center")
    ws.cell(row=sum_row, column=11, value="{:,.0f} 원".format(summer_sum)).alignment = Alignment(horizontal="center")
    for col in [8, 9, 10, 11]:
        ws.cell(row=sum_row, column=col).font = Font(bold=True)

    # 다운로드
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    st.subheader("📄 환산 결과 요약")
    st.dataframe(styled_df)

    st.subheader("📈 총합")
    st.write(f"▶ 실적보험료 합계: **{performance_sum:,.0f} 원**")
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
