import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import math

st.set_page_config(
    page_title="나이스 지필평가 성적 취합",
    page_icon="📊",
    layout="wide"
)
st.title("📊 나이스 지필평가 학급별 일람표 → 성적 취합")
st.markdown("---")

with st.sidebar:
    st.header("📌 사용 방법")
    st.markdown("""
    **Step 1.** 전교생 명단 업로드
    - 열: `반`, `번호`, `성명`
    
    **Step 2.** 나이스 일람표 업로드
    - 지필평가 학급별 일람표 원본 엑셀
    - 여러 반이 한 파일에 있어도 OK
    
    **Step 3.** 취합 실행 & 다운로드
    """)

GRADE_CUTOFF = [10, 34, 66, 90, 100]

def get_boundaries(total: int) -> list:
    boundaries = [round(total * pct / 100) for pct in GRADE_CUTOFF]
    boundaries[-1] = total
    return boundaries

def calc_grade(rank, total):
    if pd.isna(rank) or pd.isna(total) or total == 0:
        return np.nan
    boundaries = get_boundaries(int(total))
    for grade, boundary in enumerate(boundaries, start=1):
        if rank <= boundary:
            return grade
    return 5

def calc_percentile(rank, total):
    if pd.isna(rank) or pd.isna(total) or total == 0:
        return np.nan
    return round((1 - (rank - 1) / total) * 100, 1)

def calc_rank(scores: pd.Series) -> pd.Series:
    return scores.rank(method="min", ascending=False)

def calc_grade_table(df_stat: pd.DataFrame) -> list:
    total      = len(df_stat)
    boundaries = get_boundaries(total)
    응시생수   = total
    원점수평균 = round(float(df_stat["원점수"].mean()), 2)
    rows     = []
    prev_end = 0
    for g in range(1, 6):
        grade_df   = df_stat[df_stat["예상등급"] == g]
        rank_start = prev_end + 1
        rank_end   = boundaries[g - 1]
        rank_range = f"{rank_start}~{rank_end}"
        if len(grade_df) == 0:
            cutline = "-"
        elif g == 5:
            cutline = 0
        else:
            cutline = round(float(grade_df["원점수"].min()), 1)
        rows.append({
            "등급":     f"{g}등급",
            "커트라인": cutline,
            "석차범위": rank_range,
        })
        prev_end = rank_end
    return rows, 응시생수, 원점수평균

def parse_subject_unit(raw: str) -> int:
    if not isinstance(raw, str):
        return None
    match = re.search(r"\((\d+)\)\s*$", raw.strip())
    return int(match.group(1)) if match else None

def parse_subject_name(raw: str) -> str:
    if not isinstance(raw, str):
        return str(raw)
    raw = raw.strip()
    if ":" in raw:
        raw = raw.split(":")[-1].strip()
    raw = re.sub(r"\(\d+\)\s*$", "", raw).strip()
    return raw

def is_subject_cell(val_str: str) -> bool:
    return bool(re.search(r".+:.+\(\d+\)\s*$", val_str))

def parse_nice_excel(file) -> tuple:
    df_raw = pd.read_excel(file, header=None, dtype=str)
    df_raw = df_raw.fillna("")
    all_records = []
    unit_dict   = {}

    block_start_rows = []
    for row_idx in range(len(df_raw)):
        row_vals = " ".join(df_raw.iloc[row_idx].values)
        if re.search(r"\d+학년도.+\d+학년\s+\d+반", row_vals):
            block_start_rows.append(row_idx)

    if not block_start_rows:
        st.error("학년도/반 정보를 찾을 수 없습니다.")
        return pd.DataFrame(), {}

    block_start_rows.append(len(df_raw))
    st.info(f"📋 감지된 반 블록 수: **{len(block_start_rows)-1}개**")

    for b_idx in range(len(block_start_rows) - 1):
        b_start = block_start_rows[b_idx]
        b_end   = block_start_rows[b_idx + 1]
        block   = df_raw.iloc[b_start:b_end].reset_index(drop=True)

        header_text = " ".join(block.iloc[0].values)
        cleaned     = re.sub(r"\d+학년도", "", header_text)
        cls_match   = re.search(r"(\d+)반", cleaned)
        if not cls_match:
            continue
        cls = int(cls_match.group(1))

        subject_row_idx = None
        for r in range(min(10, len(block))):
            if any(is_subject_cell(str(v)) for v in block.iloc[r].values):
                subject_row_idx = r
                break
        if subject_row_idx is None:
            st.warning(f"⚠️ {cls}반: 과목명 행을 찾을 수 없습니다.")
            continue

        subject_row  = block.iloc[subject_row_idx]
        subject_cols = {}
        for col_idx, val in enumerate(subject_row):
            val_str = str(val).strip()
            if is_subject_cell(val_str):
                subj_name = parse_subject_name(val_str)
                subj_unit = parse_subject_unit(val_str)
                subject_cols[col_idx] = subj_name
                if subj_name not in unit_dict and subj_unit is not None:
                    unit_dict[subj_name] = subj_unit

        if not subject_cols:
            st.warning(f"⚠️ {cls}반: 과목 열을 찾을 수 없습니다.")
            continue

        data_start   = subject_row_idx + 2
        num_col_idx  = None
        name_col_idx = None

        for r in range(data_start, min(data_start + 5, len(block))):
            for c_idx, val in enumerate(block.iloc[r]):
                if re.match(r"^\d+$", str(val).strip()) and num_col_idx is None:
                    num_col_idx = c_idx
                    break
            if num_col_idx is not None:
                break

        if num_col_idx is not None:
            for r in range(data_start, min(data_start + 5, len(block))):
                for c_idx in range(num_col_idx + 1,
                                   min(num_col_idx + 5, len(block.iloc[r]))):
                    v = str(block.iloc[r].iloc[c_idx]).strip()
                    if re.match(r"^[가-힣]{2,5}$", v):
                        name_col_idx = c_idx
                        break
                if name_col_idx is not None:
                    break

        if num_col_idx is None or name_col_idx is None:
            st.warning(f"⚠️ {cls}반: 번호/성명 열을 찾을 수 없습니다.")
            continue

        stop_keywords = ["응시생수", "총", "평", "학과", "합 계"]

        for r in range(data_start, len(block)):
            row      = block.iloc[r]
            row_text = " ".join(str(v) for v in row.values)
            if any(kw in row_text for kw in stop_keywords):
                break

            num_val = str(row.iloc[num_col_idx]).strip()
            if not re.match(r"^\d+$", num_val):
                continue

            name_val = str(row.iloc[name_col_idx]).strip()
            if not re.match(r"^[가-힣]{2,5}$", name_val):
                continue

            for col_idx, subject_name in subject_cols.items():
                score_val = str(row.iloc[col_idx]).strip()
                if score_val == "" or score_val == "nan":
                    score = np.nan
                elif re.match(r"^[\d\.]+$", score_val.replace(",", "")):
                    try:
                        score = float(score_val.replace(",", ""))
                    except:
                        score = np.nan
                else:
                    score = score_val

                all_records.append({
                    "반":     cls,
                    "번호":   int(num_val),
                    "성명":   name_val,
                    "과목명": subject_name,
                    "점수":   score
                })

    if not all_records:
        return pd.DataFrame(), unit_dict
    return pd.DataFrame(all_records), unit_dict

def make_subject_stat(result_wide, subj_cols):
    stat_dict = {}
    for subj in subj_cols:
        df_subj = result_wide[["반", "번호", "성명", subj]].copy()
        df_subj = df_subj.rename(columns={subj: "원점수"})
        df_subj = df_subj[df_subj["원점수"].notna()].copy()
        if len(df_subj) == 0:
            continue
        total = len(df_subj)
        df_subj["석차"]     = calc_rank(df_subj["원점수"]).astype(int)
        df_subj["예상등급"] = df_subj["석차"].apply(
            lambda r: calc_grade(r, total)).astype(int)
        df_subj["백분위"]   = df_subj["석차"].apply(
            lambda r: calc_percentile(r, total))
        df_subj = df_subj.sort_values(
            ["반", "번호"],
            key=lambda x: pd.to_numeric(x, errors="coerce")
        ).reset_index(drop=True)
        stat_dict[subj] = df_subj
    return stat_dict

def make_grade_wide(students_df, stat_dict, subj_cols):
    grade_wide = students_df.copy()
    for subj in subj_cols:
        if subj not in stat_dict:
            grade_wide[subj] = np.nan
            continue
        df_stat = stat_dict[subj][["반", "번호", "성명", "예상등급"]].copy()
        df_stat = df_stat.rename(columns={"예상등급": subj})
        grade_wide = pd.merge(grade_wide, df_stat,
                              on=["반", "번호", "성명"], how="left")
    grade_wide = grade_wide.sort_values(
        ["반", "번호"],
        key=lambda x: pd.to_numeric(x, errors="coerce")
    ).reset_index(drop=True)
    return grade_wide

def calc_weighted_grade(row, subj_cols, unit_dict):
    total_weight = 0
    total_score  = 0
    for subj in subj_cols:
        grade = row.get(subj, np.nan)
        unit  = unit_dict.get(subj, None)
        if pd.isna(grade) or unit is None:
            continue
        total_weight += unit
        total_score  += grade * unit
    if total_weight == 0:
        return np.nan
    return round(total_score / total_weight, 2)

# ── UI ────────────────────────────────────────────────────────
st.header("Step 1️⃣  전교생 명단 업로드")

col_a, col_b = st.columns([3, 2])
with col_a:
    student_file = st.file_uploader(
        "전교생 명단 엑셀 (반/번호/성명)", type=["xlsx","xls"], key="stu")
with col_b:
    sample_stu = pd.DataFrame({
        "반":[1,1,1,2,2],"번호":[1,4,5,1,2],
        "성명":["홍길동","홍길동2","박길동","김철수","이영희"]})
    buf_s = io.BytesIO()
    sample_stu.to_excel(buf_s, index=False)
    st.download_button("📥 전교생명단 샘플", buf_s.getvalue(), "전교생명단_샘플.xlsx")

students_df = None
if student_file:
    xl_s    = pd.ExcelFile(student_file)
    sheet_s = st.selectbox("시트 선택", xl_s.sheet_names, key="ssheet")
    raw_s   = pd.read_excel(student_file, sheet_name=sheet_s, dtype=str)
    st.dataframe(raw_s.head(), use_container_width=True)
    cols_s  = raw_s.columns.tolist()

    def detect(keywords, cols):
        for kw in keywords:
            for c in cols:
                if kw in str(c): return c
        return cols[0]

    mc1, mc2, mc3 = st.columns(3)
    c_col = mc1.selectbox("반열",   cols_s, index=cols_s.index(detect(["반","class"], cols_s)))
    n_col = mc2.selectbox("번호열", cols_s, index=cols_s.index(detect(["번호","num"],  cols_s)))
    m_col = mc3.selectbox("성명열", cols_s, index=cols_s.index(detect(["성명","이름","name"], cols_s)))

    students_df = raw_s[[c_col, n_col, m_col]].copy()
    students_df.columns = ["반","번호","성명"]
    students_df = students_df.dropna(subset=["성명"])
    students_df["반"]   = students_df["반"].str.strip()
    students_df["번호"] = students_df["번호"].str.strip()
    students_df["성명"] = students_df["성명"].str.strip()
    st.success(f"✅ 전교생 {len(students_df)}명 로드 완료")

st.markdown("---")
st.header("Step 2️⃣  나이스 지필평가 일람표 업로드")

nice_file = st.file_uploader(
    "나이스 지필평가 학급별 일람표 엑셀", type=["xlsx","xls"], key="nice")

nice_long_df = None
unit_dict    = {}
if nice_file:
    with st.spinner("파일 파싱 중..."):
        nice_long_df, unit_dict = parse_nice_excel(nice_file)

    if nice_long_df is not None and len(nice_long_df) > 0:
        subjects_found = sorted(nice_long_df["과목명"].unique().tolist())
        col_info1, col_info2 = st.columns(2)
        col_info1.metric("감지된 과목 수",   f"{len(subjects_found)}개")
        col_info2.metric("성적 데이터 행수", f"{len(nice_long_df)}행")

        with st.expander("📐 과목별 단위수 확인"):
            unit_df = pd.DataFrame(
                unit_dict.items(), columns=["과목명","단위수"]
            ).sort_values("과목명").reset_index(drop=True)
            st.dataframe(unit_df, use_container_width=True)

        with st.expander("📚 감지된 과목 목록"):
            for i, subj in enumerate(subjects_found, 1):
                st.write(f"{i}. {subj}")
        with st.expander("🔍 파싱 결과 미리보기"):
            st.dataframe(nice_long_df.head(30), use_container_width=True)
        st.success(f"✅ 파싱 완료: **{len(nice_long_df)}**개 레코드")

st.markdown("---")
st.header("Step 3️⃣  성적 취합 및 다운로드")

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    sort_subject = st.checkbox("과목명 가나다 정렬", value=True)
with col_opt2:
    add_stat = st.checkbox("과목별 통계 시트 추가 (석차/등급/백분위)", value=True)

if st.button("🚀 성적 취합 실행", type="primary", use_container_width=True):
    if students_df is None:
        st.error("❌ 전교생 명단을 업로드하세요.")
    elif nice_long_df is None or len(nice_long_df) == 0:
        st.error("❌ 나이스 성적 파일을 업로드하세요.")
    else:
        with st.spinner("성적 취합 중..."):
            nice_df = nice_long_df.copy()
            nice_df["반"]   = nice_df["반"].astype(str)
            nice_df["번호"] = nice_df["번호"].astype(str)
            nice_df["성명"] = nice_df["성명"].astype(str)

            nice_numeric = nice_df.copy()
            nice_numeric["점수_숫자"] = pd.to_numeric(
                nice_numeric["점수"], errors="coerce")

            pivot = nice_numeric.pivot_table(
                index=["반","번호","성명"],
                columns="과목명", values="점수_숫자",
                aggfunc="first"
            ).reset_index()
            pivot.columns.name = None

            subj_cols = [c for c in pivot.columns if c not in ["반","번호","성명"]]
            if sort_subject:
                subj_cols = sorted(subj_cols)
            pivot = pivot[["반","번호","성명"] + subj_cols]

            stu = students_df.copy()
            stu["반"]   = stu["반"].astype(str)
            stu["번호"] = stu["번호"].astype(str)

            result_wide = pd.merge(stu, pivot, on=["반","번호","성명"], how="left")
            result_wide["평균"] = result_wide[subj_cols].mean(axis=1).round(1)
            result_wide = result_wide.sort_values(
                ["반","번호"], key=lambda x: pd.to_numeric(x, errors="coerce")
            ).reset_index(drop=True)

            special = nice_df[
                pd.to_numeric(nice_df["점수"], errors="coerce").isna()
                & (nice_df["점수"] != "") & (nice_df["점수"].notna())
            ].copy()
            if len(special) > 0:
                special_pivot = special.pivot_table(
                    index=["반","번호","성명"], columns="과목명",
                    values="점수", aggfunc="first"
                ).reset_index()
                special_pivot.columns.name = None
                sc = [c for c in special_pivot.columns if c not in ["반","번호","성명"]]
                special_pivot = special_pivot.rename(columns={c: f"{c}_비고" for c in sc})
                result_wide = pd.merge(result_wide, special_pivot,
                                       on=["반","번호","성명"], how="left")

            stat_dict = {}
            if add_stat:
                stat_dict = make_subject_stat(result_wide, subj_cols)

            grade_wide = make_grade_wide(
                students_df.copy().assign(
                    반=students_df["반"].astype(str),
                    번호=students_df["번호"].astype(str)
                ),
                stat_dict, subj_cols
            )

            grade_wide["평균등급(단위수반영)"] = grade_wide.apply(
                lambda row: calc_weighted_grade(row, subj_cols, unit_dict),
                axis=1
            )

            st.session_state["result_wide"] = result_wide
            st.session_state["grade_wide"]  = grade_wide
            st.session_state["nice_long"]   = nice_df
            st.session_state["subj_cols"]   = subj_cols
            st.session_state["stat_dict"]   = stat_dict
            st.session_state["unit_dict"]   = unit_dict

        st.success(f"✅ 취합 완료! **{len(result_wide)}명** / **{len(subj_cols)}개 과목**")

# ── 결과 표시 및 다운로드 ─────────────────────────────────────
if "result_wide" in st.session_state:

    result_wide = st.session_state["result_wide"]
    grade_wide  = st.session_state["grade_wide"]
    nice_long   = st.session_state["nice_long"]
    subj_cols   = st.session_state["subj_cols"]
    stat_dict   = st.session_state["stat_dict"]
    unit_dict   = st.session_state["unit_dict"]

    c1, c2, c3 = st.columns(3)
    c1.metric("전체 학생", f"{len(result_wide)}명")
    c2.metric("과목 수",   f"{len(subj_cols)}개")
    응시건수 = result_wide[subj_cols].notna().sum().sum()
    전체건수 = len(result_wide) * len(subj_cols)
    c3.metric("응시 데이터", f"{응시건수}건 / {전체건수}건")

    st.subheader("📋 성적 취합 결과")
    st.dataframe(result_wide, use_container_width=True, height=400)

    search = st.text_input("🔍 학생 이름 검색")
    if search:
        found = result_wide[result_wide["성명"].str.contains(search)]
        st.dataframe(found, use_container_width=True) if len(found) \
            else st.warning("해당 학생 없음")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        wb = writer.book

        def mfmt(d):
            return wb.add_format(d)

        fmt_info_header    = mfmt({"font_size":11,"bg_color":"#C8B4E0","align":"center","valign":"vcenter","border":1,"text_wrap":True})
        fmt_subj_header    = mfmt({"font_size":11,"bg_color":"#C8E6C8","align":"center","valign":"vcenter","border":1,"text_wrap":True})
        fmt_avg_header     = mfmt({"font_size":11,"bg_color":"#FFF0B0","align":"center","valign":"vcenter","border":1,"text_wrap":True})
        fmt_wavg_header    = mfmt({"font_size":11,"bg_color":"#FFD9B0","align":"center","valign":"vcenter","border":1,"text_wrap":True})
        fmt_stat_score_hdr = mfmt({"font_size":11,"bg_color":"#C8E6C8","align":"center","valign":"vcenter","border":1,"text_wrap":True})
        fmt_stat_rank_hdr  = mfmt({"font_size":11,"bg_color":"#FFD9B0","align":"center","valign":"vcenter","border":1,"text_wrap":True})
        fmt_stat_grade_hdr = mfmt({"font_size":11,"bg_color":"#B0D4FF","align":"center","valign":"vcenter","border":1,"text_wrap":True})
        fmt_stat_pct_hdr   = mfmt({"font_size":11,"bg_color":"#FFF0B0","align":"center","valign":"vcenter","border":1,"text_wrap":True})
        fmt_info           = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1})
        fmt_score          = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0.00"})
        fmt_avg_score      = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0.0"})
        fmt_wavg_val       = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0.00"})
        fmt_empty          = mfmt({"bg_color":"#FFFFFF","border":1})
        fmt_special        = mfmt({"font_size":11,"align":"center","valign":"vcenter","border":1,"font_color":"#C00000","bg_color":"#FFE0E0"})
        fmt_cell           = mfmt({"font_size":11,"align":"center","valign":"vcenter","border":1})
        fmt_int            = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0"})
        fmt_pct            = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0.0"})
        fmt_gt_title       = mfmt({"font_size":11,"bg_color":"#C8B4E0","align":"center","valign":"vcenter","border":1})
        fmt_gt_label       = mfmt({"font_size":11,"bg_color":"#E8E8E8","align":"center","valign":"vcenter","border":1})
        fmt_gt_cut         = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0.0"})
        fmt_gt_range       = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1})
        fmt_gt_cnt         = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0"})
        fmt_gt_avg         = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0.00"})
        fmt_grade_val      = mfmt({"font_size":11,"bg_color":"#FFFFFF","align":"center","valign":"vcenter","border":1,"num_format":"0"})

        info_cols  = ["반", "번호", "성명"]
        COL_GAP    = 7
        COL_GTITLE = 8
        COL_GCUT   = 9
        COL_GRANGE = 10
        COL_GCNT   = COL_GTITLE
        COL_GAVG   = COL_GCUT

        def get_header_fmt(col_name):
            if col_name in info_cols:                return fmt_info_header
            elif col_name == "평균":                 return fmt_avg_header
            elif col_name == "평균등급(단위수반영)":  return fmt_wavg_header
            else:                                    return fmt_subj_header

        def write_data_cell(ws, row, col, val, col_name):
            if col_name in info_cols:
                ws.write(row, col, val, fmt_info)
            elif col_name == "평균":
                ws.write(row, col,
                         "" if pd.isna(val) else val,
                         fmt_empty if pd.isna(val) else fmt_avg_score)
            elif col_name == "평균등급(단위수반영)":
                ws.write(row, col,
                         "" if pd.isna(val) else val,
                         fmt_empty if pd.isna(val) else fmt_wavg_val)
            elif pd.isna(val):
                ws.write(row, col, "", fmt_empty)
            elif isinstance(val, str):
                ws.write(row, col, val, fmt_special)
            else:
                ws.write(row, col, val, fmt_score)

        # ── 시트1: 전체결과 ───────────────────────────────────
        all_cols = info_cols + subj_cols + ["평균"]
        ws_all = wb.add_worksheet("전체결과")
        writer.sheets["전체결과"] = ws_all
        ws_all.freeze_panes(1, 3)
        ws_all.set_row(0, 25)
        for c_idx, col_name in enumerate(all_cols):
            ws_all.write(0, c_idx, col_name, get_header_fmt(col_name))
            ws_all.set_column(c_idx, c_idx, 10)
        for r_idx, row in result_wide[all_cols].iterrows():
            for c_idx, col_name in enumerate(all_cols):
                write_data_cell(ws_all, r_idx+1, c_idx, row[col_name], col_name)

        # ── 시트2: 예상등급 ───────────────────────────────────
        ws_grd = wb.add_worksheet("예상등급")
        writer.sheets["예상등급"] = ws_grd
        ws_grd.freeze_panes(1, 3)
        ws_grd.set_row(0, 25)

        grade_cols = info_cols + subj_cols + ["평균등급(단위수반영)"]

        for c_idx, col_name in enumerate(grade_cols):
            # ✅ 과목명에만 (단위수) 붙이기
            if col_name in info_cols or col_name == "평균등급(단위수반영)":
                display_name = col_name
            else:
                unit = unit_dict.get(col_name, None)
                display_name = f"{col_name}({unit})" if unit is not None else col_name

            ws_grd.write(0, c_idx, display_name, get_header_fmt(col_name))
            ws_grd.set_column(c_idx, c_idx, 10)

        # ✅ 학생 데이터 (바로 1행부터, 단위수 별도 행 없음)
        for r_idx, row in grade_wide[grade_cols].iterrows():
            excel_row = r_idx + 1
            for c_idx, col_name in enumerate(grade_cols):
                val = row[col_name]
                if col_name in info_cols:
                    ws_grd.write(excel_row, c_idx, val, fmt_info)
                elif col_name == "평균등급(단위수반영)":
                    ws_grd.write(excel_row, c_idx,
                                 "" if pd.isna(val) else val,
                                 fmt_empty if pd.isna(val) else fmt_wavg_val)
                elif pd.isna(val):
                    ws_grd.write(excel_row, c_idx, "", fmt_empty)
                else:
                    ws_grd.write(excel_row, c_idx, int(val), fmt_grade_val)

        # ── 시트3~N: 반별 시트 ────────────────────────────────
        sorted_cls_list = sorted(result_wide["반"].unique(), key=lambda x: int(x))

        for cls in sorted_cls_list:
            grp      = result_wide[result_wide["반"] == cls]
            grp_subj = [c for c in subj_cols if grp[c].notna().any()]
            sheet_nm = f"{cls}반"
            ws_r = wb.add_worksheet(sheet_nm)
            writer.sheets[sheet_nm] = ws_r
            ws_r.freeze_panes(1, 3)
            ws_r.set_row(0, 25)
            cols_r = info_cols + grp_subj + ["평균"]
            for c_idx, col_name in enumerate(cols_r):
                ws_r.write(0, c_idx, col_name, get_header_fmt(col_name))
                ws_r.set_column(c_idx, c_idx, 10)
            for r_idx, (_, row) in enumerate(grp[cols_r].iterrows()):
                for c_idx, col_name in enumerate(cols_r):
                    write_data_cell(ws_r, r_idx+1, c_idx, row[col_name], col_name)

            last_row = len(grp) + 1
            ws_r.write(last_row+1, 0, "응시생수", fmt_info_header)
            ws_r.write(last_row+2, 0, "평  균",   fmt_info_header)
            ws_r.write(last_row+1, 1, "", fmt_empty)
            ws_r.write(last_row+1, 2, "", fmt_empty)
            ws_r.write(last_row+2, 1, "", fmt_empty)
            ws_r.write(last_row+2, 2, "", fmt_empty)

            for c_idx, col_name in enumerate(cols_r):
                if col_name in info_cols: continue
                col_avg = pd.to_numeric(grp[col_name], errors="coerce").mean()
                if col_name == "평균":
                    ws_r.write(last_row+1, c_idx, "", fmt_empty)
                    ws_r.write(last_row+2, c_idx,
                               round(col_avg,1) if not np.isnan(col_avg) else "",
                               fmt_avg_score)
                else:
                    cnt = grp[col_name].notna().sum()
                    ws_r.write(last_row+1, c_idx,
                               int(cnt) if cnt > 0 else "", fmt_cell)
                    ws_r.write(last_row+2, c_idx,
                               round(col_avg,1) if not np.isnan(col_avg) else "",
                               fmt_score)

        # ── 과목별 통계 시트 ──────────────────────────────────
        if stat_dict:
            stat_header_map = {
                "반":       fmt_info_header,
                "번호":     fmt_info_header,
                "성명":     fmt_info_header,
                "원점수":   fmt_stat_score_hdr,
                "석차":     fmt_stat_rank_hdr,
                "예상등급": fmt_stat_grade_hdr,
                "백분위":   fmt_stat_pct_hdr,
            }
            stat_width_map = {
                "반":6,"번호":6,"성명":10,
                "원점수":10,"석차":8,"예상등급":10,"백분위":8
            }
            stat_cols = ["반","번호","성명","원점수","석차","예상등급","백분위"]

            for subj, df_stat in stat_dict.items():
                sht_name = (subj[:28]+"_통계" if len(subj)>28 else subj+"_통계")
                ws_st = wb.add_worksheet(sht_name)
                writer.sheets[sht_name] = ws_st
                ws_st.freeze_panes(1, 3)
                ws_st.set_row(0, 25)

                for c_idx, col_name in enumerate(stat_cols):
                    ws_st.write(0, c_idx, col_name, stat_header_map[col_name])
                    ws_st.set_column(c_idx, c_idx, stat_width_map[col_name])

                ws_st.set_column(COL_GAP,    COL_GAP,    2)
                ws_st.set_column(COL_GTITLE, COL_GTITLE, 12)
                ws_st.set_column(COL_GCUT,   COL_GCUT,   10)
                ws_st.set_column(COL_GRANGE, COL_GRANGE, 12)

                grade_table, 응시생수, 원점수평균 = calc_grade_table(df_stat)

                for c in [COL_GTITLE, COL_GCUT, COL_GRANGE]:
                    ws_st.write(0, c, "", None)

                ws_st.write(1, COL_GTITLE, subj,        fmt_gt_title)
                ws_st.write(1, COL_GCUT,   "커트라인",  fmt_gt_title)
                ws_st.write(1, COL_GRANGE, "석차 범위", fmt_gt_title)

                for g_idx, g_row in enumerate(grade_table):
                    r = g_idx + 2
                    ws_st.write(r, COL_GTITLE, g_row["등급"], fmt_gt_label)
                    if g_row["커트라인"] == "-":
                        ws_st.write(r, COL_GCUT,   "-", fmt_gt_range)
                        ws_st.write(r, COL_GRANGE, "-", fmt_gt_range)
                    else:
                        ws_st.write(r, COL_GCUT,   g_row["커트라인"], fmt_gt_cut)
                        ws_st.write(r, COL_GRANGE, g_row["석차범위"],  fmt_gt_range)

                for c in [COL_GTITLE, COL_GCUT, COL_GRANGE]:
                    ws_st.write(7, c, "", None)

                ws_st.write(8, COL_GCNT, "응시생수",   fmt_gt_title)
                ws_st.write(8, COL_GAVG, "원점수평균", fmt_gt_title)
                ws_st.write(9, COL_GCNT, 응시생수,     fmt_gt_cnt)
                ws_st.write(9, COL_GAVG, 원점수평균,   fmt_gt_avg)

                for r_idx, row in df_stat[stat_cols].iterrows():
                    excel_row = r_idx + 1
                    for c_idx, col_name in enumerate(stat_cols):
                        val = row[col_name]
                        if pd.isna(val):
                            ws_st.write(excel_row, c_idx, "", fmt_empty)
                        elif col_name in ["석차","예상등급"]:
                            ws_st.write(excel_row, c_idx, int(val), fmt_int)
                        elif col_name == "백분위":
                            ws_st.write(excel_row, c_idx, val, fmt_pct)
                        elif col_name == "원점수":
                            ws_st.write(excel_row, c_idx, val, fmt_score)
                        else:
                            ws_st.write(excel_row, c_idx, val, fmt_info)

        nice_long.to_excel(writer, index=False, sheet_name="원본(Long형)")

    st.download_button(
        label="⬇️ 성적 취합 결과 엑셀 다운로드",
        data=output.getvalue(),
        file_name="지필평가_성적취합결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )