import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os, re  # re: 표 이름 정규화용

# ✅ 테이블 이름 고유화 시퀀스(충돌 방지)
TABLE_SEQ = 0

# ✅ 기준값
CONV_TARGET = 1_800_000  # 환산 목표 180만원
SUMM_TARGET = 3_000_000  # (유지) 썸머 목표 필요 시 사용

MIN_COUNT = 5
HANWHA_MIN_PREMIUM = 20_000  # 한화생명 가동 2만원 이상 1건 필수


# ───────────────────────────────────────────────────────────────────
# 유틸: 시트 이름 유니크 보장(31자 제한 고려)
def unique_sheet_name(wb, base, limit=31):
    name = str(base)[:limit] if base else "Sheet"
    if name not in wb.sheetnames:
        return name
    i = 2
    while True:
        suffix = f"_{i}"
        trunc = limit - len(suffix)
        cand = f"{name[:trunc]}{suffix}"
        if cand not in wb.sheetnames:
            return cand
        i += 1


# 유틸: 헤더 안전 인덱스 조회(없으면 default로)
def header_idx(ws, name, default=None):
    for i in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=i).value == name:
            return i
    return default


# ✅ 표 이름을 안전하게(영문/숫자/밑줄만, 첫 글자 규칙 준수)
def safe_table_name(base: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_]", "_", base)
    if not re.match(r"^[A-Za-z_]", name):
        name = f"tbl_{name}"
    return name[:254]  # 엑셀 제한(여유)


# ✅ 유틸: 열 너비 자동화(전체 열 스캔 + 패딩)
def autosize_columns_full(ws, padding=10):
    for column_cells in ws.columns:
        max_len = max(
            len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells
        )
        ws.column_dimensions[column_cells[0].column_letter].width = max_len + padding


# ✅ 필수 조건 판정
def check_requirements(dfin: pd.DataFrame):
    count_ok = len(dfin) >= MIN_COUNT
    hanwha_ok = (
        (dfin["보험사"].astype(str).str.strip() == "한화생명")
        & (pd.to_numeric(dfin["보험료"], errors="coerce").fillna(0) >= HANWHA_MIN_PREMIUM)
    ).any()
    return count_ok, hanwha_ok


def run():
    st.set_page_config(page_title="보험 계약 환산기", layout="wide")

    # 👉 사이드바: 사용법 + 썸머 토글
    with st.sidebar:
        st.header("🧭 사용 방법")
        st.markdown(
            """
            **🖥️ 한화라이프랩 전산**  
            **- 📂 계약관리**  
            **- 📑 보유계약 장기**  
            **- ⏱️ 기간 설정**  
            **- 💾 엑셀 다운로드 후 파일 첨부하면 됩니다.**
            """
        )
        SHOW_SUMMER = st.toggle("🌞 썸머 기준 포함", value=False)

    st.title("📊 보험 계약 실적 환산기 (컨벤션{} 기준)".format(" & 썸머" if SHOW_SUMMER else ""))

    uploaded_file = st.file_uploader("📂 계약 목록 Excel 파일 업로드 (.xlsx)", type=["xlsx"])
    if not uploaded_file:
        st.info("📤 계약 목록 Excel 파일(.xlsx)을 업로드해주세요.")
        return

    base_filename = os.path.splitext(uploaded_file.name)[0]
    download_filename = f"{base_filename}_환산결과.xlsx"

    # 1) 필요한 컬럼 로드 (✅ 수금자명 포함)
    columns_needed = [
        "수금자명",
        "계약일",
        "보험사",
        "상품명",
        "납입기간",
        "초회보험료",
        "쉐어율",
        "납입방법",
        "상품군2",
        "계약상태",
    ]
    df = pd.read_excel(uploaded_file, usecols=columns_needed)

    # 제외 데이터 초기화
    excluded_df = pd.DataFrame()

    # 2) 제외 조건 처리: 일시납 / 연금성·저축성 / 철회|해약
    if {"납입방법", "상품군2", "계약상태"}.issubset(df.columns):
        before_count = len(df)

        df["납입방법"] = df["납입방법"].astype(str).str.strip()
        df["상품군2"] = df["상품군2"].astype(str).str.strip()
        df["계약상태"] = df["계약상태"].astype(str).str.strip()

        is_lumpsum = df["납입방법"].str.contains("일시납")
        is_savings = df["상품군2"].str.contains("연금성|저축성")
        is_cancelled = df["계약상태"].str.contains("철회|해약")

        is_excluded = is_lumpsum | is_savings | is_cancelled
        excluded_df = df[is_excluded].copy()
        df = df[~is_excluded].copy()

        excluded_count = before_count - len(df)
        if excluded_count > 0:
            st.warning(
                f"⚠️ 제외된 계약 {excluded_count}건 "
                f"(일시납 / 연금성·저축성 / 철회|해약 계약)이 계산에서 제외되었습니다."
            )

            st.subheader("🚫 제외된 계약 목록")
            excluded_display = excluded_df[
                ["수금자명", "계약일", "보험사", "상품명", "납입기간", "초회보험료", "납입방법"]
            ].copy()
            excluded_display.columns = ["수금자명", "계약일", "보험사", "상품명", "납입기간", "보험료", "납입방법"]
            st.dataframe(excluded_display, use_container_width=True)

            st.markdown("📝 **제외 계약별 사유:**")
            for _, row in excluded_df.iterrows():
                상품명 = row.get("상품명", "")
                사유들 = []
                if isinstance(row.get("납입방법", ""), str) and "일시납" in row["납입방법"]:
                    사유들.append("일시납")
                if isinstance(row.get("상품군2", ""), str) and (
                    "연금성" in row["상품군2"] or "저축성" in row["상품군2"]
                ):
                    사유들.append("연금/저축성")
                if isinstance(row.get("계약상태", ""), str) and "철회" in row["계약상태"]:
                    사유들.append("철회")
                if isinstance(row.get("계약상태", ""), str) and "해약" in row["계약상태"]:
                    사유들.append("해약")
                사유_텍스트 = " / ".join(사유들) if 사유들 else "제외 조건 미상"
                st.markdown(f"- ({상품명}) → 제외사유: {사유_텍스트}")

    # 3) 내부 컬럼명 정규화
    df.rename(columns={"계약일": "계약일자", "초회보험료": "보험료"}, inplace=True)

    # 4) 필수 항목 체크
    required_columns = {"수금자명", "계약일자", "보험사", "상품명", "납입기간", "보험료", "쉐어율"}
    if not required_columns.issubset(df.columns):
        st.error(
            "❌ 업로드된 파일에 다음 항목이 모두 포함되어 있어야 합니다:\n"
            + ", ".join(sorted(required_columns))
        )
        st.stop()

    if df["쉐어율"].isnull().any():
        st.error("❌ '쉐어율'에 빈 값이 포함되어 있습니다. 모든 행에 값을 입력해주세요.")
        st.stop()

    # 5) 환산율 분류 (컨벤션 & 썸머)
    def to_int_safe(x, default=0):
        try:
            return int(float(x))
        except Exception:
            return default

    # ✅ 기준 변경 반영:
    # - 생명보험(10년납 미만) 50%
    # - 생명보험(10년납 이상) 100%
    # - 한화생명(납기 무관) 150%
    # - 손해보험 200%
    # - KB손해, 한화손해 250%
    # - 흥국화재, DB손해 300%
    def classify(row):
        보험사원본 = str(row["보험사"]).strip()
        납기 = to_int_safe(row["납입기간"], 0)

        # 표기 흔들림 방어
        is_hanhwa_life = 보험사원본 == "한화생명"

        is_kb_nonlife = ("KB" in 보험사원본) and ("손" in 보험사원본)  # KB손보/KB손해
        is_hanhwa_nonlife = ("한화" in 보험사원본) and ("손" in 보험사원본) and (보험사원본 != "한화생명")
        is_db_nonlife = ("DB" in 보험사원본) and ("손" in 보험사원본)  # DB손보/DB손해
        is_heungkuk = ("흥국" in 보험사원본) and ("화재" in 보험사원본)  # 흥국화재

        is_nonlife_generic = any(x in 보험사원본 for x in ["손해", "손보", "화재", "해상"])
        is_life_other = (not is_hanhwa_life) and (("생명" in 보험사원본) or (보험사원본 in ["신한라이프"]))

        # 컨벤션율
        if is_hanhwa_life:
            conv_rate = 150
        elif is_db_nonlife or is_heungkuk:
            conv_rate = 300
        elif is_kb_nonlife or is_hanhwa_nonlife:
            conv_rate = 250
        elif is_nonlife_generic:
            conv_rate = 200
        elif is_life_other:
            conv_rate = 100 if 납기 >= 10 else 50
        else:
            conv_rate = 0

        # 썸머율: 기능은 유지(토글), 기준표가 따로 없으므로 동일하게 적용
        summ_rate = conv_rate

        return pd.Series([conv_rate, summ_rate])

    df[["컨벤션율", "썸머율"]] = df.apply(classify, axis=1)

    # 6) 수치 계산
    # ✅ 쉐어율은 "참고 컬럼"으로만 사용 (보험료가 이미 반영된 값이라는 전제 유지)
    df["쉐어율"] = df["쉐어율"].apply(lambda x: float(str(x).replace("%", "")) if pd.notnull(x) else x)

    df["실적보험료"] = df["보험료"]
    df["컨벤션환산금액"] = df["실적보험료"] * df["컨벤션율"] / 100
    df["썸머환산금액"] = df["실적보험료"] * df["썸머율"] / 100

    # 날짜 유효성 체크
    df["계약일자_raw"] = pd.to_datetime(df["계약일자"], errors="coerce")
    invalid_dates = df[df["계약일자_raw"].isna()]
    if not invalid_dates.empty:
        st.warning(
            f"⚠️ {len(invalid_dates)}건의 계약일자가 날짜로 인식되지 않았습니다. "
            f"엑셀에서 '2025-07-23'처럼 정확한 형식으로 입력해주세요."
        )

    # 7) 수금자명 선택 필터
    collectors = ["전체"] + sorted(df["수금자명"].astype(str).unique().tolist())
    selected_collector = st.selectbox("👤 수금자명 선택", collectors, index=0)
    show_df = df if selected_collector == "전체" else df[df["수금자명"].astype(str) == selected_collector].copy()

    # 화면 표시용 포맷팅
    def to_styled(dfin: pd.DataFrame) -> pd.DataFrame:
        _styled = dfin.copy()
        _styled["계약일자"] = pd.to_datetime(_styled["계약일자"], errors="coerce").dt.strftime("%Y-%m-%d")
        _styled["납입기간"] = _styled["납입기간"].astype(str) + "년"
        _styled["보험료"] = _styled["보험료"].map("{:,.0f} 원".format)
        _styled["쉐어율"] = _styled["쉐어율"].astype(str) + " %"
        _styled["실적보험료"] = _styled["실적보험료"].map("{:,.0f} 원".format)
        _styled["컨벤션율"] = _styled["컨벤션율"].astype(str) + " %"
        if SHOW_SUMMER:
            _styled["썸머율"] = _styled["썸머율"].astype(str) + " %"
        _styled["컨벤션환산금액"] = _styled["컨벤션환산금액"].map("{:,.0f} 원".format)
        if SHOW_SUMMER:
            _styled["썸머환산금액"] = _styled["썸머환산금액"].map("{:,.0f} 원".format)

        base_cols = ["수금자명", "계약일자", "보험사", "상품명", "납입기간", "보험료", "컨벤션율"]
        if SHOW_SUMMER:
            base_cols += ["썸머율"]
        base_cols += ["실적보험료", "컨벤션환산금액"]
        if SHOW_SUMMER:
            base_cols += ["썸머환산금액"]
        return _styled[base_cols]

    # 합계 함수
    def sums(dfin: pd.DataFrame):
        perf = dfin["실적보험료"].sum()
        conv = dfin["컨벤션환산금액"].sum()
        summ = dfin["썸머환산금액"].sum() if SHOW_SUMMER else 0
        return perf, conv, summ

    # ── 화면 표시(선택된 수금자 기준) ───────────────────────────
    st.subheader(f"📄 {'전체' if selected_collector=='전체' else selected_collector} 환산 결과")
    st.dataframe(to_styled(show_df), use_container_width=True)

    perf_sum, conv_sum, summ_sum = sums(show_df)
    conv_gap = conv_sum - CONV_TARGET
    summ_gap = (summ_sum - SUMM_TARGET) if SHOW_SUMMER else 0

    st.subheader("📈 총합")
    st.markdown(
        f"""
        <div style='border: 2px solid #1f77b4; border-radius: 10px; padding: 20px; background-color: #f7faff; margin-bottom: 20px;'>
            <h4 style='color:#1f77b4;'>📈 총합 요약</h4>
            <p><strong>▶ 실적보험료 합계:</strong> {perf_sum:,.0f} 원</p>
            <p><strong>▶ 컨벤션 기준 합계:</strong> {conv_sum:,.0f} 원</p>
            {f"<p><strong>▶ 썸머 기준 합계:</strong> {summ_sum:,.0f} 원</p>" if SHOW_SUMMER else ""}
        </div>
        """,
        unsafe_allow_html=True,
    )

    def gap_box(title, amount):
        if amount > 0:
            color = "#e6f4ea"
            txt = "#0c6b2c"
            sym = f"+{amount:,.0f} 원 초과"
        elif amount < 0:
            color = "#fdecea"
            txt = "#b80000"
            sym = f"{amount:,.0f} 원 부족"
        else:
            color = "#f3f3f3"
            txt = "#000000"
            sym = "기준 달성"
        return f"""
        <div style='border: 1px solid {txt}; border-radius: 8px; background-color: {color}; padding: 12px; margin: 10px 0;'>
            <strong style='color:{txt};'>{title}: {sym}</strong>
        </div>
        """

    st.markdown(gap_box("컨벤션 목표 대비", conv_gap), unsafe_allow_html=True)
    if SHOW_SUMMER:
        st.markdown(gap_box("썸머 목표 대비", summ_gap), unsafe_allow_html=True)

    # ✅ 화면: 필수 조건 체크(선택된 수금자 기준)
    st.subheader("✅ 필수 조건 체크")
    count_ok, hanwha_ok = check_requirements(show_df)

    def req_box(title, ok):
        color = "#e6f4ea" if ok else "#fdecea"
        txt = "#0c6b2c" if ok else "#b80000"
        mark = "✅ 충족" if ok else "❌ 미충족"
        return f"""
        <div style='border: 1px solid {txt}; border-radius: 8px; background-color: {color}; padding: 12px; margin: 10px 0;'>
            <strong style='color:{txt};'>{title}: {mark}</strong>
        </div>
        """

    st.markdown(req_box(f"필수 건수 {MIN_COUNT}건 이상", count_ok), unsafe_allow_html=True)
    st.markdown(
        req_box(f"한화생명 가동 {HANWHA_MIN_PREMIUM:,.0f}원 이상 1건", hanwha_ok),
        unsafe_allow_html=True,
    )

    # ── 수금자명별 요약(전체 집계) ──────────────────────────────
    st.subheader("🧮 수금자명별 요약")
    group = df.groupby("수금자명", dropna=False).agg(
        실적보험료합계=("실적보험료", "sum"),
        컨벤션합계=("컨벤션환산금액", "sum"),
        썸머합계=("썸머환산금액", "sum") if SHOW_SUMMER else ("실적보험료", "sum"),
    ).reset_index()

    if not SHOW_SUMMER:
        group.drop(columns=["썸머합계"], inplace=True)

    group["컨벤션_갭"] = group["컨벤션합계"] - CONV_TARGET
    if SHOW_SUMMER:
        group["썸머_갭"] = group["썸머합계"] - SUMM_TARGET

    disp_group = group.copy()
    for col in ["실적보험료합계", "컨벤션합계", "컨벤션_갭"]:
        disp_group[col] = disp_group[col].map("{:,.0f} 원".format)
    if SHOW_SUMMER:
        for col in ["썸머합계", "썸머_갭"]:
            disp_group[col] = disp_group[col].map("{:,.0f} 원".format)

    st.dataframe(disp_group, use_container_width=True)

    # ── 엑셀 출력 보조 유틸 ─────────────────────────────────────
    def write_table(ws, df_for_sheet: pd.DataFrame, start_row: int = 1, name_suffix: str = "A"):
        """df_for_sheet: 헤더 포함(문자 포맷 완료)"""
        r_idx = start_row - 1
        for r_idx, row in enumerate(
            dataframe_to_rows(df_for_sheet, index=False, header=True), start_row
        ):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.alignment = Alignment(horizontal="center", vertical="center")

        end_col_letter = ws.cell(row=start_row, column=df_for_sheet.shape[1]).column_letter
        last_row = r_idx if df_for_sheet.shape[0] > 0 else start_row

        global TABLE_SEQ
        TABLE_SEQ += 1
        raw_name = f"tbl_{ws.title}_{name_suffix}_{TABLE_SEQ}"
        display_name = safe_table_name(raw_name)

        table = Table(displayName=display_name, ref=f"A{start_row}:{end_col_letter}{last_row}")
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        table.tableStyleInfo = style
        ws.add_table(table)

        autosize_columns_full(ws, padding=5)
        return last_row

    # ✅ 제외 목록 생성: 비어도 스키마 유지
    def build_excluded_with_reason(exdf: pd.DataFrame) -> pd.DataFrame:
        base_cols = ["수금자명", "계약일자", "보험사", "상품명", "납입기간", "보험료", "납입방법", "제외사유"]
        if exdf is None or exdf.empty:
            return pd.DataFrame(columns=base_cols)

        tmp = exdf.copy()
        reasons = []
        for _, row in tmp.iterrows():
            r = []
            if isinstance(row.get("납입방법", ""), str) and "일시납" in row["납입방법"]:
                r.append("일시납")
            if isinstance(row.get("상품군2", ""), str) and (
                "연금성" in row["상품군2"] or "저축성" in row["상품군2"]
            ):
                r.append("연금/저축성")
            if isinstance(row.get("계약상태", ""), str) and "철회" in row["계약상태"]:
                r.append("철회")
            if isinstance(row.get("계약상태", ""), str) and "해약" in row["계약상태"]:
                r.append("해약")
            reasons.append(" / ".join(r) if r else "제외 조건 미상")
        tmp["제외사유"] = reasons

        tmp_disp = tmp[["수금자명", "계약일", "보험사", "상품명", "납입기간", "초회보험료", "납입방법", "제외사유"]].copy()
        tmp_disp.rename(columns={"계약일": "계약일자", "초회보험료": "보험료"}, inplace=True)

        tmp_disp["계약일자"] = pd.to_datetime(tmp_disp["계약일자"], errors="coerce").dt.strftime("%Y-%m-%d")
        tmp_disp["납입기간"] = tmp_disp["납입기간"].apply(lambda x: f"{int(x)}년" if pd.notnull(x) else "")
        tmp_disp["보험료"] = tmp_disp["보험료"].map(lambda x: "{:,.0f} 원".format(x) if pd.notnull(x) else "")
        return tmp_disp[base_cols]

    excluded_disp_all = build_excluded_with_reason(excluded_df)

    # ✅ 요약 시트 필수조건 한 줄
    def write_requirements_line(ws, base_row: int, dfin: pd.DataFrame):
        c_ok, h_ok = check_requirements(dfin)

        def mark(ok):
            return "✅" if ok else "❌"

        line = (
            f"필수 조건: "
            f"최소 {MIN_COUNT}건 이상 {mark(c_ok)}  |  "
            f"한화생명 가동 {HANWHA_MIN_PREMIUM:,.0f}원 이상 1건 {mark(h_ok)}"
        )
        cell = ws.cell(row=base_row, column=1, value=line)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center")

    # ✅ 총합/갭 블록(시트별)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    sum_fill = PatternFill("solid", fgColor="F2F2F2")

    def sums_and_gaps_block(ws, dfin_numeric: pd.DataFrame, start_row: int):
        perf, conv, summ = sums(dfin_numeric)

        col_conv_rate = header_idx(ws, "컨벤션율", 1)
        col_perf = header_idx(ws, "실적보험료", 2)
        col_conv_amt = header_idx(ws, "컨벤션환산금액", 3)
        col_summ_amt = header_idx(ws, "썸머환산금액", None)

        # 총 합계: 테이블 아래 한 줄 띄우고
        sum_row = start_row + 2
        ws.cell(row=sum_row, column=col_conv_rate, value="총 합계").alignment = Alignment(horizontal="center")
        cell_perf = ws.cell(row=sum_row, column=col_perf, value=f"{perf:,.0f} 원")
        cell_conv = ws.cell(row=sum_row, column=col_conv_amt, value=f"{conv:,.0f} 원")
        cell_perf.alignment = Alignment(horizontal="center")
        cell_perf.font = Font(bold=True)
        cell_conv.alignment = Alignment(horizontal="center")
        cell_conv.font = Font(bold=True)

        if SHOW_SUMMER and col_summ_amt:
            cell_summ = ws.cell(row=sum_row, column=col_summ_amt, value=f"{summ:,.0f} 원")
            cell_summ.alignment = Alignment(horizontal="center")
            cell_summ.font = Font(bold=True)

        for c in [col_conv_rate, col_perf, col_conv_amt] + ([col_summ_amt] if SHOW_SUMMER and col_summ_amt else []):
            cell = ws.cell(row=sum_row, column=c)
            cell.fill = sum_fill
            cell.border = thin_border

        def style_gap(amount):
            if amount > 0:
                return f"+{amount:,.0f} 원 초과", "008000"
            if amount < 0:
                return f"{amount:,.0f} 원 부족", "FF0000"
            return "기준 달성", "000000"

        gap_row = sum_row + 2
        txt, col = style_gap(conv - CONV_TARGET)
        ws.cell(row=gap_row, column=col_conv_amt, value="컨벤션 기준 대비").alignment = Alignment(horizontal="center")
        gap_cell = ws.cell(row=gap_row, column=col_perf, value=txt)
        gap_cell.alignment = Alignment(horizontal="center")
        gap_cell.font = Font(bold=True, color=col)

        if SHOW_SUMMER and col_summ_amt:
            txt2, col2 = style_gap(summ - SUMM_TARGET)
            ws.cell(row=gap_row + 1, column=col_conv_amt, value="썸머 기준 대비").alignment = Alignment(horizontal="center")
            gap_cell2 = ws.cell(row=gap_row + 1, column=col_perf, value=txt2)
            gap_cell2.alignment = Alignment(horizontal="center")
            gap_cell2.font = Font(bold=True, color=col2)
            return gap_row + 2

        return gap_row + 1

    # ── 엑셀 워크북 작성: 요약 + 수금자별 ───────────────────────
    wb = Workbook()

    # 요약 시트
    ws_summary = wb.active
    ws_summary.title = "요약"

    summary_fmt = group.copy()
    money_cols = ["실적보험료합계", "컨벤션합계", "컨벤션_갭"]
    if SHOW_SUMMER:
        money_cols += ["썸머합계", "썸머_갭"]

    for c in money_cols:
        summary_fmt[c] = summary_fmt[c].map(lambda x: f"{x:,.0f} 원")

    next_row = write_table(ws_summary, summary_fmt, start_row=1, name_suffix="SUM")

    # ✅ 요약 시트: 필수 조건 한 줄 표기(전체 df 기준)
    write_requirements_line(ws_summary, base_row=next_row + 2, dfin=df)

    # 요약 시트에 제외 목록 전체(있을 때만)
    if not excluded_disp_all.empty:
        ws_summary.cell(row=next_row + 4, column=1, value="제외 계약 목록").font = Font(bold=True)
        next_row = write_table(ws_summary, excluded_disp_all, start_row=next_row + 5, name_suffix="EXC")

    # 수금자별 상세 시트
    for collector in sorted(df["수금자명"].astype(str).unique().tolist()):
        sub = df[df["수금자명"].astype(str) == collector].copy()
        styled_sub = to_styled(sub)

        sheet_title = unique_sheet_name(wb, collector)
        ws = wb.create_sheet(title=sheet_title)

        next_row = write_table(ws, styled_sub, start_row=1, name_suffix="NORM")

        # 주요 금액 컬럼 최소 열 너비 20
        for header in ["실적보험료", "컨벤션환산금액", "썸머환산금액"]:
            idx = header_idx(ws, header)
            if idx:
                col_letter = ws.cell(row=1, column=idx).column_letter
                cur = ws.column_dimensions[col_letter].width
                ws.column_dimensions[col_letter].width = 20 if (cur is None or cur < 20) else cur

        next_row = sums_and_gaps_block(ws, sub, start_row=next_row)

        # 해당 수금자의 제외건(있다면)
        ex_sub = excluded_disp_all[excluded_disp_all["수금자명"].astype(str) == collector]
        if not ex_sub.empty:
            ws.cell(row=next_row + 1, column=1, value="제외 계약").font = Font(bold=True)
            write_table(ws, ex_sub, start_row=next_row + 2, name_suffix="EXC")

    # 저장/다운로드
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    st.download_button(
        label="📥 환산 결과 엑셀 다운로드 (요약 + 수금자별 시트 + 제외사유 포함)",
        data=excel_output,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    run()
