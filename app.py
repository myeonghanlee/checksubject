from __future__ import annotations

import hashlib
import csv
import difflib
import io
import json
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
STANDARD_PATH = DATA_DIR / "reference_catalog.json"
APPROVAL_PATH = DATA_DIR / "20261108교육감승인과목목록_고등학교.csv"
APPROVAL_ORG_SEOUL = "서울특별시교육청"

st.set_page_config(
    page_title="고등학교 학점 배당표 검증",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      :root { --ink:#182b35; --muted:#657782; --line:#dce5e9; --teal:#087e82; --cream:#f6f8f7; }
      .stApp { background: #f6f8f7; color: var(--ink); }
      .block-container { max-width: 1500px; padding-top: 1.45rem; padding-bottom: 3rem; }
      [data-testid="stSidebar"] { background: #edf3f2; border-right: 1px solid #dce5e9; }
      .hero { background: linear-gradient(115deg,#123440 0%,#176c70 100%); color: white; border-radius: 18px; padding: 1.45rem 1.7rem; margin-bottom: 1rem; }
      .hero h1 { color: white; font-size: 1.8rem; margin: 0 0 .35rem 0; letter-spacing: -.03em; }
      .hero p { color: #d6e9e7; margin: 0; font-size: 1rem; }
      .eyebrow { color:#a8d8d3; text-transform:uppercase; letter-spacing:.13em; font-size:.75rem; font-weight:700; margin-bottom:.35rem; }
      .section-label { color:var(--muted); font-size:.84rem; font-weight:700; letter-spacing:.04em; text-transform:uppercase; margin:.65rem 0 .25rem; }
      .small-note { color:#657782; font-size:.86rem; line-height:1.55; }
      .rule-card { border:1px solid var(--line); background:white; border-radius:14px; padding:1rem 1.1rem; height:100%; }
      .rule-card strong { color:#173c49; }
      .stButton > button[kind="primary"] { background:#087e82; border-color:#087e82; border-radius:10px; font-weight:700; min-height:2.75rem; }
      div[data-testid="stMetric"] { background:white; border:1px solid var(--line); border-radius:14px; padding:.85rem 1rem; }
      .stDataFrame { border:1px solid var(--line); border-radius:10px; overflow:hidden; }
      .footer { color:#71818a; font-size:.78rem; border-top:1px solid var(--line); margin-top:2rem; padding-top:.8rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def match_key(value: Any) -> str:
    """Ignore trailing whitespace and terminal footnote asterisks only."""
    text = "" if value is None else str(value).rstrip()
    return re.sub(r"\*{1,2}$", "", text).rstrip()


def cell_text(value: Any) -> str:
    return "" if value is None else str(value)


@st.cache_data(show_spinner=False)
def load_standard_catalog() -> list[dict[str, str]]:
    if not STANDARD_PATH.exists():
        raise FileNotFoundError("기준 과목 카탈로그가 없습니다. scripts/build_reference_catalog.py를 실행하세요.")
    return json.loads(STANDARD_PATH.read_text(encoding="utf-8"))


def relation_indicates_professional(value: str) -> bool:
    compact = re.sub(r"\s+", "", value)
    return (
        "전문교과" in compact
        or "전문공통" in compact
        or "전공일반" in compact
        or "전공실무" in compact
    )


@st.cache_data(show_spinner=False)
def load_approval_catalog(professional_standard_names: tuple[str, ...]) -> tuple[dict[str, list[dict[str, str]]], dict[str, list[dict[str, str]]], dict[str, int]]:
    approved: dict[str, list[dict[str, str]]] = defaultdict(list)
    pending: dict[str, list[dict[str, str]]] = defaultdict(list)
    counts = {"approved_rows": 0, "pending_rows": 0, "excluded_rows": 0}
    professional_names = set(professional_standard_names)

    with APPROVAL_PATH.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        for raw in reader:
            name = match_key(raw.get("과목명", ""))
            if not name:
                counts["excluded_rows"] += 1
                continue
            school_level = raw.get("학교급", "")
            status = raw.get("상태", "")
            if school_level != "고등학교":
                counts["excluded_rows"] += 1
                continue
            record = {
                "name": name,
                "related": raw.get("관련교과(교과군)", ""),
                "agency": raw.get("소속(기관)", ""),
                "status": status,
                "date": raw.get("작성일", ""),
                "textbook": raw.get("교과서명", ""),
            }
            if status == "승인완료":
                counts["approved_rows"] += 1
                approved[name].append(record)
            elif status == "승인진행":
                counts["pending_rows"] += 1
                pending[name].append(record)
            else:
                counts["excluded_rows"] += 1

    return dict(approved), dict(pending), counts


def make_indexes(catalog: list[dict[str, str]]) -> tuple[dict[str, list[dict[str, str]]], set[str]]:
    by_name: dict[str, list[dict[str, str]]] = defaultdict(list)
    for item in catalog:
        by_name[match_key(item.get("name", ""))].append(item)
    professional_names = {
        name for name, entries in by_name.items()
        if any(entry.get("table") == "표 3" for entry in entries)
    }
    return dict(by_name), professional_names


def detect_header_row(ws: Any, max_scan: int = 150) -> int | None:
    for row in range(1, min(ws.max_row, max_scan) + 1):
        d_value = cell_text(ws.cell(row, 4).value).strip()
        m_value = cell_text(ws.cell(row, 13).value).strip()
        if d_value in {"과목", "과목명"} and m_value == "비고":
            return row
    return None


def note_labels_for(records: list[dict[str, str]], is_professional: bool) -> list[str]:
    labels: list[str] = []
    for record in records:
        agency = record.get("agency", "")
        if not agency:
            continue
        if agency == APPROVAL_ORG_SEOUL:
            label = "교육감 승인 과목"
        else:
            label = "타 시·도 교육감 승인 과목"
        if is_professional:
            label += "(전문)"
        if label not in labels:
            labels.append(label)
    return labels


def visual_name(value: str) -> str:
    if value.endswith(" "):
        return value.rstrip(" ") + "␠" * (len(value) - len(value.rstrip(" ")))
    if value.startswith(" "):
        return "␠" * (len(value) - len(value.lstrip(" "))) + value.lstrip(" ")
    return value


def close_candidates(key: str, known_names: list[str]) -> list[str]:
    if key.lstrip() != key:
        trimmed = key.lstrip()
        if trimmed in known_names:
            return [f"앞 공백 제거 시 후보: {trimmed}"]
    return difflib.get_close_matches(key, known_names, n=3, cutoff=0.78)


def review_row(
    row_number: int,
    raw_name: str,
    raw_note: str,
    standard_by_name: dict[str, list[dict[str, str]]],
    approvals: dict[str, list[dict[str, str]]],
    pending: dict[str, list[dict[str, str]]],
    known_names: list[str],
) -> dict[str, Any]:
    key = match_key(raw_name)
    standard_hits = standard_by_name.get(key, [])
    approval_hits = approvals.get(key, [])
    pending_hits = pending.get(key, [])

    is_professional = (
        any(item.get("table") == "표 3" for item in standard_hits)
        or any(relation_indicates_professional(row.get("related", "")) for row in approval_hits)
    )

    expected_labels: list[str] = []
    for item in standard_hits:
        if item.get("table") == "표 2" and "특목고 선택 과목" not in expected_labels:
            expected_labels.append("특목고 선택 과목")
        elif item.get("table") == "표 3" and "전문 교과" not in expected_labels:
            expected_labels.append("전문 교과")
    for label in note_labels_for(approval_hits, is_professional):
        if label not in expected_labels:
            expected_labels.append(label)

    evidence: list[str] = []
    for item in standard_hits:
        context = " · ".join(part for part in [item.get("table", ""), item.get("classification", ""), item.get("group", ""), item.get("category", ""), item.get("context", "")] if part)
        if context not in evidence:
            evidence.append(context)
    for row in approval_hits:
        agency_label = row.get("agency", "기관 미기재")
        related_label = row.get("related", "관련 교과 미기재")
        record_label = f"승인목록 · {agency_label} · {related_label} · 승인완료"
        if record_label not in evidence:
            evidence.append(record_label)

    expected_text = ", ".join(expected_labels) if expected_labels else "(추가 표기 없음)"
    actual_tags = [label for label in expected_labels if label in raw_note]
    missing_tags = [label for label in expected_labels if label not in raw_note]
    candidates = []
    status = "정상"
    reason = "기준 과목명과 일치하며 필요한 비고 표기가 확인됩니다." if expected_labels else "기준 과목명과 일치합니다. 이 과목에 요구된 추가 비고 표기는 없습니다."

    has_known_course = bool(standard_hits or approval_hits)
    if not has_known_course:
        status = "확인 필요"
        if pending_hits:
            agencies = ", ".join(sorted({row.get("agency", "") for row in pending_hits if row.get("agency")}))
            reason = f"승인목록에서 ‘승인진행’ 상태만 확인되어 승인 과목으로 확정할 수 없습니다. {agencies}".strip()
        else:
            reason = "표 1~3 및 승인완료 고등학교 과목 목록에서 정확히 일치하는 과목을 찾지 못했습니다."
        candidates = close_candidates(key, known_names)
        if re.search(r"\*{1,2}\s*$", raw_name):
            candidates.insert(0, "과목명 끝의 별표(*) 표시는 비교에서 제거했습니다.")
    elif missing_tags:
        status = "오류"
        reason = "M열 필수 표기가 없거나 문구가 일치하지 않습니다: " + ", ".join(missing_tags)
    else:
        if pending_hits:
            reason += " 다만 동일 과목명에 승인진행 기록도 있어 승인목록 최신 상태를 확인하세요."
            status = "확인 필요"

    if not standard_hits and not approval_hits and not pending_hits and raw_name[:1].isspace():
        reason += " 과목명 앞쪽 공백은 자동 제거하지 않았습니다."

    return {
        "엑셀 행": row_number,
        "D열 과목명": raw_name,
        "과목명(공백 확인용)": visual_name(raw_name),
        "판정": status,
        "M열 비고": raw_note,
        "필수 비고": expected_text,
        "확인된 필수 비고": ", ".join(actual_tags),
        "기준 근거": " | ".join(evidence) if evidence else "일치 근거 없음",
        "검토 내용": reason,
        "후보/안내": " | ".join(candidates),
    }


def read_workbook_sheets(file_bytes: bytes) -> tuple[list[str], dict[str, int | None]]:
    workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    names = workbook.sheetnames
    headers: dict[str, int | None] = {}
    for name in names:
        headers[name] = detect_header_row(workbook[name])
    workbook.close()
    return names, headers


def review_workbook(file_bytes: bytes, sheet_name: str, header_row: int, standard_by_name: dict[str, list[dict[str, str]]], approvals: dict[str, list[dict[str, str]]], pending: dict[str, list[dict[str, str]]], known_names: list[str]) -> tuple[pd.DataFrame, int]:
    workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = workbook[sheet_name]
    results: list[dict[str, Any]] = []
    skipped = 0
    for row_number in range(header_row + 1, ws.max_row + 1):
        raw_value = ws.cell(row_number, 4).value
        if raw_value is None or cell_text(raw_value) == "":
            skipped += 1
            continue
        raw_name = cell_text(raw_value)
        raw_note = cell_text(ws.cell(row_number, 13).value)
        results.append(review_row(row_number, raw_name, raw_note, standard_by_name, approvals, pending, known_names))
    workbook.close()
    result_columns = ["엑셀 행", "D열 과목명", "과목명(공백 확인용)", "판정", "M열 비고", "필수 비고", "확인된 필수 비고", "기준 근거", "검토 내용", "후보/안내"]
    return pd.DataFrame(results, columns=result_columns), skipped


def render_status_badges() -> None:
    st.markdown('<div class="hero"><div class="eyebrow">교육과정 편성 점검</div><h1>고등학교 학점 배당표 검증</h1><p>D열 과목명과 M열 비고를 기준 자료에 대조하고, 필요한 표기를 행별로 확인합니다.</p></div>', unsafe_allow_html=True)


try:
    catalog = load_standard_catalog()
    standard_by_name, professional_names = make_indexes(catalog)
    approval_by_name, pending_by_name, source_counts = load_approval_catalog(tuple(sorted(professional_names)))
except Exception as error:
    st.error(f"기준 자료를 불러오지 못했습니다: {error}")
    st.stop()

all_known_names = sorted(set(standard_by_name) | set(approval_by_name))

with st.sidebar:
    st.markdown("### 기준 자료")
    st.caption("과목명 끝의 빈칸과 말미의 *·**는 무시합니다. 앞쪽 공백·중간 띄어쓰기·기호는 자동 보정하지 않습니다.")
    st.markdown("**교육과정 표** · 표 1 보통교과 / 표 2 특목고 선택 / 표 3 전문교과")
    st.markdown("**승인 과목 목록** · 고등학교 · 승인완료만 판정 대상으로 사용")
    st.markdown("**서울특별시교육청** → 교육감 승인 과목")
    st.markdown("**그 밖의 시·도교육청** → 타 시·도 교육감 승인 과목")
    st.markdown("전문 과목이면 승인 분류에 **(전문)**을 덧붙입니다.")
    with st.expander("판정 표기 기준"):
        st.markdown("- 표 2 일치: `특목고 선택 과목`\n- 표 3 일치: `전문 교과`\n- 서울 승인완료: `교육감 승인 과목`\n- 타 시·도 승인완료: `타 시·도 교육감 승인 과목`\n- 승인 과목이 전문에 해당하면 두 승인 표기 모두 `(전문)`을 적용\n- 표 1만 일치하는 과목에는 이 검토 항목의 추가 비고를 요구하지 않음")
    st.markdown("---")
    st.markdown(f"**승인완료 행** · {source_counts['approved_rows']:,}\n\n**승인진행 행** · {source_counts['pending_rows']:,} (확정 판정 제외)")
    st.caption("기준 목록을 갱신할 때는 data 폴더의 원본을 교체하고 카탈로그 생성 스크립트를 실행하세요.")

render_status_badges()

st.markdown("#### 1. 학점 배당표 업로드")
st.markdown('<p class="small-note">Excel 파일(.xlsx 또는 .xlsm)을 올려 주세요. 원본은 수정하지 않습니다. 선택한 시트의 D열 과목명과 M열 비고를 읽습니다.</p>', unsafe_allow_html=True)
uploaded = st.file_uploader("학점 배당표 파일", type=["xlsx", "xlsm"], label_visibility="collapsed")

if uploaded is None:
    st.info("파일을 올리면 시트와 헤더를 감지합니다. 검토 결과는 화면에서 확인하고 CSV로 내려받을 수 있습니다.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="rule-card"><strong>정확한 과목명 대조</strong><br><span class="small-note">교육과정 표 1~3과 교육감 승인 과목 목록을 독립적으로 비교합니다.</span></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="rule-card"><strong>M열 비고 검토</strong><br><span class="small-note">과목 분류에 필요한 표기가 빠졌거나 다르면 오류로 표시합니다.</span></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="rule-card"><strong>판단 불가 행 별도 표시</strong><br><span class="small-note">정확히 일치하지 않거나 승인 진행 중인 항목은 확인 필요로 남깁니다.</span></div>', unsafe_allow_html=True)
    st.stop()

payload = uploaded.getvalue()
file_signature = hashlib.sha256(payload).hexdigest()
try:
    sheet_names, detected_headers = read_workbook_sheets(payload)
except Exception as error:
    st.error(f"엑셀을 읽을 수 없습니다. 파일 형식을 확인해 주세요. 상세: {error}")
    st.stop()

suggested_sheet = next((name for name in sheet_names if detected_headers[name] is not None), sheet_names[0])
left, right = st.columns([2, 1])
with left:
    sheet_name = st.selectbox("검토할 시트", sheet_names, index=sheet_names.index(suggested_sheet))
with right:
    detected = detected_headers.get(sheet_name)
    default_header = detected or 1
    header_row = st.number_input("헤더 행 (D=과목, M=비고)", min_value=1, max_value=10000, value=default_header, step=1)
if detected:
    st.caption(f"‘{sheet_name}’ 시트 {detected}행에서 D열 ‘과목’과 M열 ‘비고’ 헤더를 감지했습니다.")
else:
    st.warning("D열 과목 / M열 비고 헤더를 자동 감지하지 못했습니다. 헤더 행 번호를 직접 지정해 주세요.")

run_review = st.button("학점 배당표 검토하기", type="primary", use_container_width=True)
if run_review:
    with st.spinner("기준 자료와 대조하고 있습니다…"):
        result_df, skipped_rows = review_workbook(payload, sheet_name, int(header_row), standard_by_name, approval_by_name, pending_by_name, all_known_names)
    st.session_state["review_result"] = result_df
    st.session_state["review_file_name"] = uploaded.name
    st.session_state["review_sheet"] = sheet_name
    st.session_state["review_signature"] = (file_signature, sheet_name, int(header_row))
    st.session_state["review_skipped"] = skipped_rows

result_df = st.session_state.get("review_result")
if result_df is None:
    st.stop()

# Prevent accidental display of an earlier file's results after a new upload.
if st.session_state.get("review_file_name") != uploaded.name or st.session_state.get("review_signature") != (file_signature, sheet_name, int(header_row)):
    st.info("현재 선택한 파일 또는 시트에 대해 다시 검토 버튼을 눌러 주세요.")
    st.stop()

counts = result_df["판정"].value_counts().to_dict() if not result_df.empty else {}
ok_count = counts.get("정상", 0)
error_count = counts.get("오류", 0)
confirm_count = counts.get("확인 필요", 0)

st.markdown("---")
st.markdown(f"#### 2. 검토 결과 · {uploaded.name} / {sheet_name}")
metric_cols = st.columns(4)
metric_cols[0].metric("검토 과목", f"{len(result_df):,}")
metric_cols[1].metric("정상", f"{ok_count:,}")
metric_cols[2].metric("오류", f"{error_count:,}")
metric_cols[3].metric("확인 필요", f"{confirm_count:,}")
if st.session_state.get("review_skipped", 0):
    st.caption(f"D열 과목명이 비어 있어 검토에서 제외된 행: {st.session_state['review_skipped']:,}행")

if confirm_count:
    st.warning("‘확인 필요’ 항목은 과목명 또는 승인 상태를 사람이 확인해야 합니다. 유사 후보는 참고용이며 자동 일치로 처리하지 않았습니다.")

filters = st.multiselect("표시할 판정", ["정상", "오류", "확인 필요"], default=["오류", "확인 필요"] if error_count or confirm_count else ["정상"])
visible = result_df[result_df["판정"].isin(filters)] if filters else result_df.iloc[0:0]
columns = ["엑셀 행", "D열 과목명", "과목명(공백 확인용)", "판정", "M열 비고", "필수 비고", "기준 근거", "검토 내용", "후보/안내"]
st.dataframe(visible[columns], width="stretch", hide_index=True, height=500)

csv_bytes = result_df.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "검토 결과 CSV 다운로드",
    data=csv_bytes,
    file_name=f"{Path(uploaded.name).stem}_검토결과.csv",
    mime="text/csv",
    type="secondary",
)

with st.expander("판정 방식과 유의사항"):
    st.markdown(
        """
        - **정상**: 과목명이 기준 목록과 일치하고, 요구되는 비고 표기가 M열에서 확인됨.
        - **오류**: 과목명은 기준 목록과 일치하지만, 요구되는 비고 표기가 누락되었거나 문구가 다름.
        - **확인 필요**: 정확히 일치하는 기준 과목이 없거나, 승인완료가 아닌 승인진행 기록만 확인됨.
        - 과목명 비교에서는 끝의 빈칸과 말미의 `*`, `**`를 무시합니다. 앞쪽 공백·중간 띄어쓰기·기호·숫자·로마자 차이는 자동 보정하지 않습니다.
        - 기준 표 1의 보통교과만 해당하면 이 검토 규칙에서 M열 추가 표기를 요구하지 않습니다.
        - 승인 과목의 전문 여부는 승인 목록의 관련 교과 표기와 표 3의 전문교과 등재 정보를 근거로 판정하며, 근거가 표 안에 표시됩니다.
        - 업로드한 Excel은 앱의 검토 과정에서 읽기 전용으로 처리되며, 이 앱의 코드에서는 파일을 저장하거나 외부로 전송하지 않습니다.
        """
    )

st.markdown('<div class="footer">기준 목록은 제공된 문서를 바탕으로 재구성한 검토 보조 자료입니다. 최종 교육과정 편성·승인 여부는 원문 및 관할 교육청 자료와 대조해 확정해 주세요.</div>', unsafe_allow_html=True)
