import streamlit as st
import pandas as pd

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
    output = df.copy()
    output["컨벤션환산금액"] = output["컨벤션환산금액"].round(0)
    output["썸머환산금액"] = output["썸머환산금액"].round(0)
    st.download_button("📥 결과 다운로드", output.to_csv(index=False).encode("utf-8-sig"), "환산결과.csv", "text/csv")

else:
    st.info("먼저 계약 목록 Excel 파일을 업로드해주세요.")

