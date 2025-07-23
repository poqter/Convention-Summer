import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="보험 계약 환산기", layout="wide")
st.title("📊 보험 계약 실적 환산기 (컨벤션 & 썸머 기준)")

# 업로드
uploaded_file = st.file_uploader("계약 목록 Excel 파일 업로드", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.write("✅ 업로드된 데이터 미리보기")
    st.dataframe(df)

    # 환산 기준 로드
    rate_df = pd.read_csv("conversion_rates.csv")

    def classify(row):
        # 보험 유형 분류
        생보사 = ["한화생명"]
        손보_250 = ["한화손해보험", "삼성화재", "흥국화재", "KB손해보험"]
        손보_200 = ["롯데손해보험", "메리츠화재", "현대해상", "DB손해보험", "MG손해보험", "하나손해보험", "AIG손해보험"]

        보험사 = row["보험사"]
        납기 = int(row["납입기간"])

        # 유형 분류
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

        # 환산율 찾기
        match = rate_df[
            (rate_df["보험사"] == 세부) &
            (rate_df["유형"] == 유형) &
            (rate_df["납입기간조건"] == 기간조건)
        ]
        if match.empty:
            return pd.Series([0, 0])
        else:
            return pd.Series([match["컨벤션율"].values[0], match["썸머율"].values[0]])

    # 환산율 적용
    df[["컨벤션율", "썸머율"]] = df.apply(classify, axis=1)
    df["컨벤션환산금액"] = df["보험료"] * df["컨벤션율"] / 100
    df["썸머환산금액"] = df["보험료"] * df["썸머율"] / 100

    # 출력
    st.subheader("📌 계약별 환산 결과")
    st.dataframe(df.style.format({"보험료": "{:.0f}", "컨벤션환산금액": "{:.0f}", "썸머환산금액": "{:.0f}"}))

    st.subheader("📈 총합")
    st.write(f"▶ 컨벤션 기준 합계: **{df['컨벤션환산금액'].sum():,.0f} 원**")
    st.write(f"▶ 썸머 기준 합계: **{df['썸머환산금액'].sum():,.0f} 원**")

    # 다운로드
    from openpyxl import Workbook
    from openpyxl.styles import Alignment
    from openpyxl.utils.dataframe import dataframe_to_rows
    from io import BytesIO
    import pandas as pd
    import os
    
    # 스타일 적용용 복사본
    styled_df = df.copy()
    styled_df["계약일자"] = pd.to_datetime(styled_df["계약일자"]).dt.strftime("%Y년%m월%d일")
    styled_df["납입기간"] = styled_df["납입기간"].astype(str) + "년"
    styled_df["보험료"] = styled_df["보험료"].map("{:,.0f} 원".format)
    styled_df["컨벤션율"] = styled_df["컨벤션율"].astype(str) + "배"
    styled_df["썸머율"] = styled_df["썸머율"].astype(str) + "배"
    styled_df["컨벤션환산금액"] = styled_df["컨벤션환산금액"].map("{:,.0f} 원".format)
    styled_df["썸머환산금액"] = styled_df["썸머환산금액"].map("{:,.0f} 원".format)

    # 엑셀 워크북 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "환산결과"
    for r_idx, row in enumerate(dataframe_to_rows(styled_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # 바이트 객체로 저장
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    # 업로드된 파일명에서 기본 이름 추출
    base_filename = os.path.splitext(uploaded_file.name)[0]
    final_filename = f"{base_filename}_환산결과.xlsx"

    # 다운로드 버튼
    st.download_button(
        label="📥 환산 결과 Excel 다운로드",
        data=excel_output,
        file_name=final_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("먼저 계약 목록 Excel 파일을 업로드해주세요.")

