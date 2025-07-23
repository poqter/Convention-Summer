import streamlit as st
import pandas as pd
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="보험 계약 환산기", layout="wide")
st.title("📊 보험 계약 실적 환산기 (컨벤션 & 썸머 기준)")

# 파일 업로드
uploaded_file = st.file_uploader("📂 계약 목록 Excel 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error("❗ Excel 파일을 읽는 중 문제가 발생했습니다.")
        st.stop()

    st.subheader("✅ 업로드된 데이터 미리보기")
    st.dataframe(df)

    # 환산 기준표 불러오기
    try:
        rate_df = pd.read_csv("conversion_rates.csv")
    except FileNotFoundError:
        st.error("⚠️ 환산 기준 파일 (conversion_rates.csv)을 찾을 수 없습니다.")
        st.stop()

    # 환산율 계산 함수
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

    # 환산율 적용
