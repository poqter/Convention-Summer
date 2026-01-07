import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os, re

TABLE_SEQ = 0


# ─────────────────────────────────────────────────────────────
# 유틸
def unique_sheet_name(wb, base, limit=31):
    name = str(base)[:limit] if base else "Sheet"
    if name not in wb.sheetnames:
        return name
    i = 2
    while True:
        suffix = f"_{i}"
        cand = f"{name[:limit-len(suffix)]}{suffix}"
        if cand not in wb.sheetnames:
            return cand
        i += 1


def safe_table_name(base):
    name = re.sub(r"[^A-Za-z0-9_]", "_", base)
    if not re.match(r"^[A-Za-z_]", name):
        name = f"tbl_{name}"
    return name[:254]


def autosize_columns(ws, padding=5):
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + padding


def to_int_safe(x, default=0):
    try:
        return int(float(x))
    except:
        return default


# ─────────────────────────────────────────────────────────────
def run():
    st.set_page_config(page_title="보험 계약 환산기", layout="wide")
    st.title("📊 보험 계약 실적 환산기")

    # ── 사이드바 ─────────────────────────────
    with st.sidebar:
        st.header("🧭 사용 방법")
        st.markdown(
            """
            **한화라이프랩 전산**
            - 계약관리 → 보유계약 장기
            - 기간 설정 후 엑셀 다운로드
            """
        )
        SHOW_SUMMER = st.toggle("🌴 썸머 환산 계산 포함", value=False)

    uploaded_file = st.file_uploader("📂 계약 목록 Excel 업로드", type=["xlsx"])
    if not uploaded_file:
        return

    base_filename = os.path.splitext(uploaded_file.name)[0]
    download_filename = f"{base_filename}_환산결과.xlsx"

    # ── 데이터 로드 ──────────────────────────
    cols = [
        "수금자명", "계약일", "보험사", "상품명", "납입기간",
        "초회보험료", "쉐어율", "납입방법", "상품군2", "계약상태"
    ]
    df = pd.read_excel(uploaded_file, usecols=cols)

    # ── 제외 조건 ────────────────────────────
    df["납입방법"] = df["납입방법"].astype(str)
    df["상품군2"] = df["상품군2"].astype(str)
    df["계약상태"] = df["계약상태"].astype(str)

    excluded = df[
        df["납입방법"].str.contains("일시납") |
        df["상품군2"].str.contains("연금|저축") |
        df["계약상태"].str.contains("철회|해약")
    ]
    df = df.drop(excluded.index)

    # ── 컬럼 정리 ────────────────────────────
    df.rename(columns={
        "계약일": "계약일자",
        "초회보험료": "보험료"
    }, inplace=True)

    # ── 환산율 계산 로직 (요청 반영) ──────────
    def classify(row):
        보험사원본 = str(row["보험사"])
        납기 = to_int_safe(row["납입기간"])

        if 보험사원본 == "한화생명":
            보험사 = "한화생명"
        elif "생명" in 보험사원본:
            보험사 = "기타생보"
        elif 보험사원본 in ["KB손보", "한화손보", "흥국화재", "DB손보"]:
            보험사 = 보험사원본
        elif any(x in 보험사원본 for x in ["손해", "손보", "화재"]):
            보험사 = "기타손보"
        else:
            보험사 = 보험사원본

        # 컨벤션 환산율
        if 보험사 == "한화생명":
            conv = 150
        elif 보험사 == "기타생보":
            conv = 100 if 납기 >= 10 else 50
        elif 보험사 in ["KB손보", "한화손보"]:
            conv = 250
        elif 보험사 in ["흥국화재", "DB손보"]:
            conv = 300
        elif 보험사 == "기타손보":
            conv = 200
        else:
            conv = 0

        # 썸머 환산율 (기존 유지)
        if 보험사 == "한화생명":
            summ = 150
        elif 보험사 == "기타생보":
            summ = 100 if 납기 >= 10 else 30
        elif 보험사 in ["KB손보", "한화손보", "흥국화재", "DB손보"]:
            summ = 200
        elif 보험사 == "기타손보":
            summ = 100
        else:
            summ = 0

        return pd.Series([conv, summ])

    df[["컨벤션율", "썸머율"]] = df.apply(classify, axis=1)

    # ── 계산 ─────────────────────────────────
    df["보험료"] = df["보험료"].astype(float)
    df["실적보험료"] = df["보험료"]
    df["컨벤션환산금액"] = df["실적보험료"] * df["컨벤션율"] / 100
    df["썸머환산금액"] = df["실적보험료"] * df["썸머율"] / 100

    # ── 목표 ─────────────────────────────────
    CONV_TARGET = 1_800_000
    SUMM_TARGET = 3_000_000

    # ── 화면 표시 ────────────────────────────
    st.subheader("📄 전체 계약 환산 결과")
    st.dataframe(df, use_container_width=True)

    conv_sum = df["컨벤션환산금액"].sum()
    summ_sum = df["썸머환산금액"].sum()

    st.subheader("📈 총합")
    st.markdown(f"- 컨벤션 합계: **{conv_sum:,.0f} 원**")
    if SHOW_SUMMER:
        st.markdown(f"- 썸머 합계: **{summ_sum:,.0f} 원**")

    st.markdown(
        f"### 🎯 컨벤션 목표 대비: {conv_sum - CONV_TARGET:,.0f} 원"
    )

    # ── 엑셀 출력 ────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = "전체"

    for r, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)

    autosize_columns(ws)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "📥 환산 결과 엑셀 다운로드",
        data=output,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    run()
