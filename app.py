# streamlit run app.py
# streamlit run app2.py
# git add .
# git commit -m "로고 추가 및 UI 수정"
# git push

import re
import streamlit as st
import pandas as pd
from datetime import date

FILE = "품질 교육 참석 현황.xlsx"
LOGO = "aaron_logo.jpg"  # 프로젝트 폴더에 로고 파일을 이 이름으로 넣어주세요.

META_COLS = ["그룹", "이름", "직위"]
ROWID_COL = "__rowid__"

# 날짜(참석) 컬럼 판단용 정규식: "YYYY-MM-DD"
DATE_COL_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# ======================
# ✅ 상단 로고 + 제목
# ======================
try:
    st.image(LOGO, width=250)
except Exception:
    st.warning("로고 파일(aaron_logo.jpg)을 찾지 못했습니다. 레포/폴더에 파일이 있는지 확인하세요.")

st.title("품질 교육 참석 관리 시스템")

# ======================
# ✅ 데이터 타입 안정화 유틸 (핵심 수정)
# ======================
def coerce_attendance_cols_to_string(df: pd.DataFrame) -> pd.DataFrame:
    """
    - 컬럼명을 문자열로 통일
    - META_COLS(그룹/이름/직위) 보정
    - 날짜 컬럼(YYYY-MM-DD)은 무조건 string dtype으로 강제
    - NaN은 빈칸으로 통일
    """
    df = df.copy()
    df.columns = df.columns.map(str)

    for col in META_COLS:
        if col not in df.columns:
            df[col] = ""

    # 메타 컬럼도 안전하게 string
    for col in META_COLS:
        df[col] = df[col].astype("string")

    # 참석(날짜) 컬럼만 string 강제 (여기서 dtype 충돌 원천 차단)
    for c in df.columns:
        if DATE_COL_RE.match(str(c)):
            df[c] = df[c].astype("string")

    # 결측치 -> 빈 문자열
    df = df.fillna("")
    return df


# ======================
# ✅ 시트명/데이터 로딩 (시트명 유지 저장을 위해 ExcelFile 사용)
# ======================
@st.cache_data(show_spinner=False)
def load_data(file_path: str):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheet1 = xls.sheet_names[0]
    sheet2 = xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0]

    df1 = pd.read_excel(xls, sheet_name=sheet1, engine="openpyxl")
    df2 = pd.read_excel(xls, sheet_name=sheet2, engine="openpyxl")

    # ✅ 핵심: 타입 안정화 적용
    df1 = coerce_attendance_cols_to_string(df1)
    df2 = coerce_attendance_cols_to_string(df2)

    return df1, df2, sheet1, sheet2


# 파일이 없으면 친절한 메시지
try:
    df1, df2, SHEET1_NAME, SHEET2_NAME = load_data(FILE)
except Exception as e:
    st.error(
        f"엑셀 파일을 불러오지 못했습니다: {FILE}\n\n"
        f"- 레포/폴더에 파일이 있는지 확인\n"
        f"- requirements.txt에 openpyxl 포함 여부 확인\n\n"
        f"에러: {e}"
    )
    st.stop()

# ======================
# ✅ 공통 유틸
# ======================
def make_display_list_and_index(df: pd.DataFrame):
    """
    표시용 리스트(그룹-이름(직위))와, 그 항목이 가리키는 실제 df index를 함께 리턴.
    """
    display = []
    idx_map = []

    # 필수 컬럼 보정
    for col in META_COLS:
        if col not in df.columns:
            df[col] = ""

    for ridx, row in df.iterrows():
        dept = str(row.get("그룹", "")).strip()
        name = str(row.get("이름", "")).strip()
        title = str(row.get("직위", "")).strip()
        display.append(f"{dept} - {name} ({title})")
        idx_map.append(ridx)

    return display, idx_map


def save_excel(file_path: str, df1_new: pd.DataFrame, df2_new: pd.DataFrame, sheet1: str, sheet2: str):
    """
    기존 파일의 시트명을 유지하여 덮어쓰기 저장.
    """
    # 저장 직전에도 한 번 더 타입 안정화 (선택이지만 매우 안전)
    df1_new = coerce_attendance_cols_to_string(df1_new)
    df2_new = coerce_attendance_cols_to_string(df2_new)

    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df1_new.to_excel(writer, sheet_name=sheet1, index=False)
        df2_new.to_excel(writer, sheet_name=sheet2, index=False)


def normalize_cell(v):
    # data_editor에서 비워진 값이 NaN으로 들어오는 경우 빈칸 처리
    if pd.isna(v):
        return ""
    return str(v)


def apply_editor_changes_to_df(original_df: pd.DataFrame, edited_view_df: pd.DataFrame):
    """
    edited_view_df: ROWID_COL 포함(원본 index를 담고 있음)
    원본(original_df)에 편집값을 반영(수정/삭제 포함)한다.
    """
    if ROWID_COL not in edited_view_df.columns:
        return original_df

    view = edited_view_df.copy()

    # rowid는 int로 고정
    view[ROWID_COL] = pd.to_numeric(view[ROWID_COL], errors="coerce").fillna(-1).astype(int)
    view = view.set_index(ROWID_COL)

    # 원본과 동일한 컬럼만 반영 (안전)
    cols_to_apply = [c for c in view.columns if c in original_df.columns]

    # 값 정규화(빈칸 삭제 반영)
    for c in cols_to_apply:
        view[c] = view[c].map(normalize_cell)

    # 해당 rowid들만 업데이트
    for ridx in view.index:
        if ridx in original_df.index:
            original_df.loc[ridx, cols_to_apply] = view.loc[ridx, cols_to_apply].values

    # ✅ 반영 후 타입 안정화
    original_df = coerce_attendance_cols_to_string(original_df)
    return original_df


def hidden_editable_table(df_to_show: pd.DataFrame, original_df_ref: pd.DataFrame, save_callback, editor_key: str):
    """
    ✅ 화면에는 기본적으로 st.dataframe만 보여주고,
    ✅ 아래에 '숨김(접힘)' 형태로 편집/삭제 모드를 제공.
    - 셀 삭제: 편집 모드에서 셀 값을 지우면(빈칸) 삭제 처리됨
    - 저장: 저장 버튼 1개로 반영
    """
    # 기본 화면(그대로): 읽기 전용 표
    st.dataframe(df_to_show, use_container_width=True)

    # 숨김 편집/삭제 모드
    with st.expander(" 편집/삭제", expanded=False):
        st.caption("※ 여기서만 표를 직접 수정/삭제할 수 있어요. 셀을 비우면 삭제(빈칸)로 저장됩니다.")

        # 원본 index 유지용 컬럼 추가(화면에는 숨기고, 반영에만 사용)
        view_df = df_to_show.copy()
        view_df.insert(0, ROWID_COL, view_df.index)

        edited_view = st.data_editor(
            view_df,
            use_container_width=True,
            hide_index=True,
            disabled=[ROWID_COL],
            key=editor_key
        )

        # 저장 버튼(최소 노출)
        if st.button("변경사항 저장", key=f"{editor_key}_save"):
            updated = apply_editor_changes_to_df(original_df_ref, edited_view)
            save_callback(updated)

# ======================
# ✅ 탭 구성
# ======================
tab1, tab2 = st.tabs(["부서별", "경영진-PM"])

# ======================
# ✅ 탭1: 부서별
# ======================
with tab1:
    st.header("부서별 교육 참석 입력")

    CATEGORY_RULES = {
        "구매/수출입": {
            "groups": {"전략구매팀", "수출입팀"},
            "include_names": set(),
            "exclude_names": set(),
        },
        "생산관리": {
            "groups": {"E&C제조기술부"},
            "include_names": set(),
            "exclude_names": set(),
        },
        "설계/개발": {
            "groups": {"E&C제조기술부", "E&C사업지원부"},
            "include_names": set(),
            "exclude_names": {"안세연"},
        },
        "영업": {
            "groups": {"해외영업"},
            "include_names": {"안세연"},
            "exclude_names": set(),
        },
        "인사": {
            "groups": {"인사총무팀"},
            "include_names": set(),
            "exclude_names": set(),
        },
    }

    # df1 보정
    df1 = coerce_attendance_cols_to_string(df1)

    category = st.selectbox("대분류 선택", list(CATEGORY_RULES.keys()), key="dept_category")
    rule = CATEGORY_RULES[category]

    groups = rule["groups"]
    include_names = rule["include_names"]
    exclude_names = rule["exclude_names"]

    # 마스크 생성
    group_mask = df1["그룹"].fillna("").astype(str).isin(groups)
    if include_names:
        name_include_mask = df1["이름"].fillna("").astype(str).isin(include_names)
    else:
        name_include_mask = pd.Series([False] * len(df1), index=df1.index)

    include_mask = group_mask | name_include_mask

    if exclude_names:
        exclude_mask = df1["이름"].fillna("").astype(str).isin(exclude_names)
        include_mask = include_mask & (~exclude_mask)

    filtered_df = df1.loc[include_mask].copy()

    st.caption(f"선택 대분류: **{category}** | 대상: **{len(filtered_df)}명**")

    if filtered_df.empty:
        st.warning("해당 대분류에 포함되는 인원이 없습니다. 분류 규칙/데이터를 확인해 주세요.")
        st.dataframe(df1, use_container_width=True)
    else:
        display_list, idx_map = make_display_list_and_index(filtered_df)
        selected = st.selectbox("대상 선택", display_list, key="dept_person")
        selected_idx = idx_map[display_list.index(selected)]

        selected_date = st.date_input("교육 날짜", value=date.today(), key="dept_date")
        col_name = selected_date.strftime("%Y-%m-%d")

        status = st.radio("참석 여부", ["O", "X", "사유 입력"], key="dept_status")

        reason = ""
        if status == "사유 입력":
            reason = st.text_input("사유", key="dept_reason", placeholder="예: 출장, 외근, 전주 참석 등")

        if st.button("저장", key="save_dept"):
            if col_name not in df1.columns:
                df1[col_name] = ""

            # ✅ 핵심: 저장 직전 컬럼 dtype을 string으로 강제
            df1[col_name] = df1[col_name].astype("string").fillna("")

            value = reason if status == "사유 입력" else status
            df1.loc[selected_idx, col_name] = str(value)

            save_excel(FILE, df1, df2, SHEET1_NAME, SHEET2_NAME)
            st.success("✅ 저장 완료! (즉시 반영됩니다)")

            st.cache_data.clear()
            st.rerun()

        st.subheader("📊 현재 현황 (선택한 대분류만)")

        # 최신 df1 기준으로 재필터 (rerun 시 최신 파일을 다시 로드함)
        refreshed_filtered = df1.loc[include_mask].copy()

        def save_filtered(updated_df1):
            save_excel(FILE, updated_df1, df2, SHEET1_NAME, SHEET2_NAME)
            st.success("✅ 표 변경사항 저장 완료! (수정/삭제 반영)")
            st.cache_data.clear()
            st.rerun()

        hidden_editable_table(
            df_to_show=refreshed_filtered,
            original_df_ref=df1,
            save_callback=save_filtered,
            editor_key=f"dept_table_editor_{category}"
        )

        with st.expander("전체 데이터 보기"):
            def save_all(updated_df1):
                save_excel(FILE, updated_df1, df2, SHEET1_NAME, SHEET2_NAME)
                st.success("✅ 전체 데이터 표 변경사항 저장 완료! (수정/삭제 반영)")
                st.cache_data.clear()
                st.rerun()

            hidden_editable_table(
                df_to_show=df1,
                original_df_ref=df1,
                save_callback=save_all,
                editor_key="dept_table_editor_all"
            )

# ======================
# ✅ 탭2: 경영진-PM
# ======================
with tab2:
    st.header("경영진-PM 교육 참석 입력")

    df2 = coerce_attendance_cols_to_string(df2)

    display_list2, idx_map2 = make_display_list_and_index(df2)
    selected2 = st.selectbox("대상 선택", display_list2, key="exec_person")
    idx2 = idx_map2[display_list2.index(selected2)]

    selected_date2 = st.date_input("교육 날짜", value=date.today(), key="exec_date")
    col_name2 = selected_date2.strftime("%Y-%m-%d")

    status2 = st.radio("참석 여부", ["O", "X", "사유 입력"], key="exec_status")

    reason2 = ""
    if status2 == "사유 입력":
        reason2 = st.text_input("사유", key="exec_reason", placeholder="예: 출장, 외근, 전주 참석 등")

    if st.button("저장", key="save_exec"):
        if col_name2 not in df2.columns:
            df2[col_name2] = ""

        # ✅ 핵심: 저장 직전 컬럼 dtype을 string으로 강제 (TypeError 방지)
        df2[col_name2] = df2[col_name2].astype("string").fillna("")

        value2 = reason2 if status2 == "사유 입력" else status2
        df2.loc[idx2, col_name2] = str(value2)

        save_excel(FILE, df1, df2, SHEET1_NAME, SHEET2_NAME)
        st.success("✅ 저장 완료! (즉시 반영됩니다)")

        st.cache_data.clear()
        st.rerun()

    st.subheader("📊 현재 현황")

    def save_df2(updated_df2):
        save_excel(FILE, df1, updated_df2, SHEET1_NAME, SHEET2_NAME)
        st.success("✅ 표 변경사항 저장 완료! (수정/삭제 반영)")
        st.cache_data.clear()
        st.rerun()

    hidden_editable_table(
        df_to_show=df2,
        original_df_ref=df2,
        save_callback=save_df2,
        editor_key="exec_table_editor"
    )
