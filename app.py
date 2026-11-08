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
NO_TEXTBOOK_GAZETTED_COURSES = {
    # 수학 · 융합 선택
    "수학과제 탐구",
    # 과학계열 · 과학 · 진로 선택
    "과학과제 연구",
    # 체육계열 · 진로 선택
    "기초 체육 전공 실기",
    "심화 체육 전공 실기",
    "고급 체육 전공 실기",
    "스포츠 경기 기술",
    # 예술계열 · 예술
    "음악 전공 실기",
    "합창⋅합주",
    "음악 공연 실습",
    "드로잉",
    "미술 전공 실기",
    "조형 탐구",
    "무용 기초 실기",
    "무용 전공 실기",
    "안무",
    "무용 제작 실습",
    "문장론",
    "시 창작",
    "소설 창작",
    "극 창작",
    "무대 미술과 기술",
    "연극 제작 실습",
    "촬영⋅조명",
    "편집⋅사운드",
    "영화 제작 실습",
    "사진 촬영",
    "사진 표현 기법",
    "영상 제작의 이해",
    "사진 감상과 비평",
}

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


def group_match_key(value: Any) -> str:
    """Normalize dot glyph variants only for curriculum subject-group labels."""
    text = match_key(value)
    return text.replace("∙", "⋅").replace("·", "⋅").replace("ㆍ", "⋅")


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


def make_indexes(catalog: list[dict[str, str]]) -> tuple[dict[str, list[dict[str, str]]], set[str], dict[str, list[dict[str, str]]]]:
    by_name: dict[str, list[dict[str, str]]] = defaultdict(list)
    groups: dict[str, list[dict[str, str]]] = defaultdict(list)
    group_aliases = {
        "사회(역사/도덕 포함)": ["사회"],
        "기술⋅가정/정보": ["기술⋅가정", "정보"],
        "제2외국어/한문": ["제2외국어", "한문"],
    }
    seen_groups: set[tuple[str, str, str]] = set()
    for item in catalog:
        by_name[match_key(item.get("name", ""))].append(item)
        if item.get("table") == "표 1" and item.get("group"):
            source_group = item["group"]
            for group_name in [source_group, *group_aliases.get(source_group, [])]:
                key = group_match_key(group_name)
                identity = (key, group_name, source_group)
                if identity in seen_groups:
                    continue
                seen_groups.add(identity)
                groups[key].append({
                    "name": group_name,
                    "table": "교과군",
                    "classification": "보통교과 교과군",
                    "group": source_group,
                    "category": "교과군 명칭",
                    "context": "",
                })
    professional_names = {
        name for name, entries in by_name.items()
        if any(entry.get("table") == "표 3" for entry in entries)
    }
    return dict(by_name), professional_names, dict(groups)


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


def approval_record_has_no_textbook(record: dict[str, str]) -> bool:
    """Only treat explicit no-textbook values as confirmed absence; blank is unknown."""
    value = re.sub(r"\s+", "", record.get("textbook", ""))
    return value == "없음" or any(
        phrase in value
        for phrase in ("교과서없음", "교과서가없음", "교과서가없는")
    )


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
    group_by_name: dict[str, list[dict[str, str]]],
    approvals: dict[str, list[dict[str, str]]],
    pending: dict[str, list[dict[str, str]]],
    known_names: list[str],
) -> dict[str, Any]:
    key = match_key(raw_name)
    standard_hits = standard_by_name.get(key, [])
    group_hits = group_by_name.get(group_match_key(key), [])
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
    for item in standard_hits + group_hits:
        context = " · ".join(part for part in [item.get("table", ""), item.get("classification", ""), item.get("group", ""), item.get("category", ""), item.get("context", "")] if part)
        if context not in evidence:
            evidence.append(context)
    for row in approval_hits:
        agency_label = row.get("agency", "기관 미기재")
        related_label = row.get("related", "관련 교과 미기재")
        record_label = f"승인목록 · {agency_label} · {related_label} · 승인완료"
        if record_label not in evidence:
            evidence.append(record_label)
    if key in NO_TEXTBOOK_GAZETTED_COURSES:
        evidence.append("고시 과목 목록 · 교과서 없음")

    expected_text = ", ".join(expected_labels) if expected_labels else "(추가 표기 없음)"
    actual_tags = [label for label in expected_labels if label in raw_note]
    missing_tags = [label for label in expected_labels if label not in raw_note]
    candidates = []
    status = "정상"
    if group_hits and not standard_hits and not approval_hits:
        reason = "표 1의 교과군 명칭으로 확인했습니다. 교과군 행에는 추가 비고 표기를 요구하지 않습니다."
    else:
        reason = "기준 과목명과 일치하며 필요한 비고 표기가 확인됩니다." if expected_labels else "기준 과목명과 일치합니다. 이 과목에 요구된 추가 비고 표기는 없습니다."

    has_known_course = bool(standard_hits or group_hits or approval_hits)
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

    if key in NO_TEXTBOOK_GAZETTED_COURSES or any(
        approval_record_has_no_textbook(record) for record in approval_hits
    ):
        status = "확인 필요"
        reason = "교과서가 없는 과목임"

    if not standard_hits and not group_hits and not approval_hits and not pending_hits and raw_name[:1].isspace():
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


def review_workbook(file_bytes: bytes, sheet_name: str, header_row: int, standard_by_name: dict[str, list[dict[str, str]]], group_by_name: dict[str, list[dict[str, str]]], approvals: dict[str, list[dict[str, str]]], pending: dict[str, list[dict[str, str]]], known_names: list[str]) -> tuple[pd.DataFrame, int]:
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
        results.append(review_row(row_number, raw_name, raw_note, standard_by_name, group_by_name, approvals, pending, known_names))
    workbook.close()
    result_columns = ["엑셀 행", "D열 과목명", "과목명(공백 확인용)", "판정", "M열 비고", "필수 비고", "확인된 필수 비고", "기준 근거", "검토 내용", "후보/안내"]
    return pd.DataFrame(results, columns=result_columns), skipped


def render_status_badges() -> None:
    st.markdown('<div class="hero"><div class="eyebrow">교육과정 편성 점검</div><h1>고등학교 학점 배당표 검증</h1><p>D열 과목명과 M열 비고를 기준 자료에 대조하고, 필요한 표기를 행별로 확인합니다.</p></div>', unsafe_allow_html=True)


try:
    catalog = load_standard_catalog()
    standard_by_name, professional_names, group_by_name = make_indexes(catalog)
    approval_by_name, pending_by_name, source_counts = load_approval_catalog(tuple(sorted(professional_names)))
except Exception as error:
    st.error(f"기준 자료를 불러오지 못했습니다: {error}")
    st.stop()

all_group_names = {item["name"] for entries in group_by_name.values() for item in entries}
all_known_names = sorted(set(standard_by_name) | set(approval_by_name) | all_group_names)

with st.sidebar:
    st.markdown("### 기준 자료")
    st.caption("과목명 끝의 빈칸과 말미의 *·**는 무시합니다. 교과군명은 별도 인식하며, 기술·가정의 점 기호 변형도 허용합니다.")
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
st.markdown('<p class="small-note">Excel 파일(.xlsx 또는 .xlsm)을 여러 개 한 번에 올릴 수 있습니다. 파일별 시트·헤더를 확인한 뒤 검토하세요. 원본은 수정하지 않습니다.</p>', unsafe_allow_html=True)
uploaded_files = st.file_uploader("학점 배당표 파일", type=["xlsx", "xlsm"], accept_multiple_files=True, label_visibility="collapsed")

if not uploaded_files:
    st.info("파일을 하나 이상 올리면 각 파일의 시트와 헤더를 감지합니다. 검토 결과는 파일별 요약과 함께 통합 CSV로 내려받을 수 있습니다.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="rule-card"><strong>정확한 과목명 대조</strong><br><span class="small-note">교육과정 표 1~3과 교육감 승인 과목 목록을 독립적으로 비교합니다.</span></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="rule-card"><strong>M열 비고 검토</strong><br><span class="small-note">과목 분류에 필요한 표기가 빠졌거나 다르면 오류로 표시합니다.</span></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="rule-card"><strong>판단 불가 행 별도 표시</strong><br><span class="small-note">정확히 일치하지 않거나 승인 진행 중인 항목은 확인 필요로 남깁니다.</span></div>', unsafe_allow_html=True)
    st.stop()

name_counts: dict[str, int] = {}
for uploaded in uploaded_files:
    name_counts[uploaded.name] = name_counts.get(uploaded.name, 0) + 1

file_configs: list[dict[str, Any]] = []
upload_signatures: list[tuple[str, str]] = []
configuration_errors: list[str] = []
for index, uploaded in enumerate(uploaded_files):
    payload = uploaded.getvalue()
    digest = hashlib.sha256(payload).hexdigest()
    upload_signatures.append((uploaded.name, digest))
    display_name = uploaded.name if name_counts[uploaded.name] == 1 else f"{uploaded.name} ({index + 1})"
    try:
        sheet_names, detected_headers = read_workbook_sheets(payload)
    except Exception as error:
        message = f"{display_name}: 엑셀을 읽지 못했습니다. {error}"
        configuration_errors.append(message)
        st.error(message)
        continue

    suggested_sheet = next((name for name in sheet_names if detected_headers[name] is not None), sheet_names[0])
    widget_id = f"{index}_{digest[:12]}"
    with st.expander(f"{display_name} · {uploaded.size / 1024:,.0f} KB", expanded=(index == 0)):
        left, right = st.columns([2, 1])
        with left:
            sheet_name = st.selectbox("검토할 시트", sheet_names, index=sheet_names.index(suggested_sheet), key=f"sheet_{widget_id}")
        with right:
            detected = detected_headers.get(sheet_name)
            default_header = detected or 1
            header_row = st.number_input("헤더 행 (D=과목, M=비고)", min_value=1, max_value=10000, value=default_header, step=1, key=f"header_{widget_id}_{sheet_names.index(sheet_name)}")
        if detected:
            st.caption(f"{detected}행에서 D열 과목과 M열 비고 헤더를 감지했습니다.")
        else:
            st.warning("헤더를 자동 감지하지 못했습니다. 과목명·비고 헤더가 있는 행 번호를 지정해 주세요.")
    file_configs.append({
        "filename": uploaded.name,
        "display_name": display_name,
        "payload": payload,
        "digest": digest,
        "sheet_name": sheet_name,
        "header_row": int(header_row),
    })

if not file_configs:
    st.warning("검토 가능한 Excel 파일이 없습니다. 올바른 .xlsx 또는 .xlsm 파일을 추가해 주세요.")
    st.stop()

batch_signature = hashlib.sha256(json.dumps({
    "uploaded": upload_signatures,
    "configs": [(item["digest"], item["sheet_name"], item["header_row"]) for item in file_configs],
    "errors": configuration_errors,
}, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()

run_review = st.button("업로드한 파일 모두 검토하기", type="primary", use_container_width=True)
if run_review:
    result_frames: list[pd.DataFrame] = []
    summaries: list[dict[str, Any]] = []
    review_errors: list[str] = []
    with st.spinner("기준 자료와 대조하고 있습니다…"):
        for item in file_configs:
            try:
                file_df, skipped_rows = review_workbook(item["payload"], item["sheet_name"], item["header_row"], standard_by_name, group_by_name, approval_by_name, pending_by_name, all_known_names)
                file_df.insert(0, "시트", item["sheet_name"])
                file_df.insert(0, "파일명", item["display_name"])
                result_frames.append(file_df)
                counts = file_df["판정"].value_counts().to_dict() if not file_df.empty else {}
                summaries.append({
                    "파일명": item["display_name"],
                    "시트": item["sheet_name"],
                    "검토 과목": len(file_df),
                    "정상": counts.get("정상", 0),
                    "오류": counts.get("오류", 0),
                    "확인 필요": counts.get("확인 필요", 0),
                    "빈 과목 행": skipped_rows,
                })
            except Exception as error:
                review_errors.append(f"{item['display_name']}: {error}")
    report_columns = ["파일명", "시트", "엑셀 행", "D열 과목명", "과목명(공백 확인용)", "판정", "M열 비고", "필수 비고", "확인된 필수 비고", "기준 근거", "검토 내용", "후보/안내"]
    result_df = pd.concat(result_frames, ignore_index=True) if result_frames else pd.DataFrame(columns=report_columns)
    st.session_state["review_result"] = result_df
    st.session_state["review_signature"] = batch_signature
    st.session_state["review_summaries"] = summaries
    st.session_state["review_errors"] = review_errors

result_df = st.session_state.get("review_result")
if result_df is None:
    st.stop()

# Prevent stale results from being shown after files or per-file settings change.
if st.session_state.get("review_signature") != batch_signature:
    st.info("현재 업로드 파일 또는 파일별 시트·헤더 설정이 바뀌었습니다. 다시 검토해 주세요.")
    st.stop()

counts = result_df["판정"].value_counts().to_dict() if not result_df.empty else {}
ok_count = counts.get("정상", 0)
error_count = counts.get("오류", 0)
confirm_count = counts.get("확인 필요", 0)

st.markdown("---")
st.markdown(f"#### 2. 검토 결과 · {len(st.session_state['review_summaries']):,}개 파일")
metric_cols = st.columns(4)
metric_cols[0].metric("검토 과목", f"{len(result_df):,}")
metric_cols[1].metric("정상", f"{ok_count:,}")
metric_cols[2].metric("오류", f"{error_count:,}")
metric_cols[3].metric("확인 필요", f"{confirm_count:,}")
summary_df = pd.DataFrame(st.session_state["review_summaries"])
st.dataframe(summary_df, width="stretch", hide_index=True)
if not summary_df.empty and summary_df["빈 과목 행"].sum():
    st.caption(f"D열 과목명이 비어 있어 제외된 행: 총 {int(summary_df['빈 과목 행'].sum()):,}행")
for message in st.session_state.get("review_errors", []):
    st.error(f"검토 실패 · {message}")

if confirm_count:
    st.warning("‘확인 필요’ 항목은 과목명 또는 승인 상태를 사람이 확인해야 합니다. 유사 후보는 참고용이며 자동 일치로 처리하지 않았습니다.")

filter_col1, filter_col2 = st.columns([1, 2])
with filter_col1:
    filters = st.multiselect("표시할 판정", ["정상", "오류", "확인 필요"], default=["오류", "확인 필요"] if error_count or confirm_count else ["정상"], key="status_filter")
with filter_col2:
    available_files = result_df["파일명"].drop_duplicates().tolist() if not result_df.empty else []
    selected_files = st.multiselect("표시할 파일", available_files, default=available_files, key="file_filter")
visible = result_df[result_df["판정"].isin(filters) & result_df["파일명"].isin(selected_files)] if filters and selected_files else result_df.iloc[0:0]
columns = ["파일명", "시트", "엑셀 행", "D열 과목명", "과목명(공백 확인용)", "판정", "M열 비고", "필수 비고", "기준 근거", "검토 내용", "후보/안내"]
st.dataframe(visible[columns], width="stretch", hide_index=True, height=500)

csv_bytes = result_df.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "전체 파일 통합 검토 결과 CSV 다운로드",
    data=csv_bytes,
    file_name="고등학교_학점배당표_통합검토결과.csv",
    mime="text/csv",
    type="secondary",
)

with st.expander("판정 방식과 유의사항"):
    st.markdown(
        """
        - **정상**: 과목명이 기준 목록과 일치하고, 요구되는 비고 표기가 M열에서 확인됨.
        - **오류**: 과목명은 기준 목록과 일치하지만, 요구되는 비고 표기가 누락되었거나 문구가 다름.
        - **확인 필요**: 정확히 일치하는 기준 과목이 없거나 승인완료가 아닌 승인진행 기록만 확인됨. 승인 목록에서 교과서가 없다고 명시된 과목과 별도 지정된 교과서 없음 고시 과목은 과목명·비고가 맞더라도 `교과서가 없는 과목임`으로 표시합니다.
        - 교과서 없음 고시 과목 목록은 코드의 `NO_TEXTBOOK_GAZETTED_COURSES`에 관리하며, 기준 근거에 `고시 과목 목록 · 교과서 없음`으로 표시합니다.
        - 과목명 비교에서는 끝의 빈칸과 말미의 `*`, `**`를 무시합니다. 앞쪽 공백·중간 띄어쓰기·기호·숫자·로마자 차이는 자동 보정하지 않습니다.
        - 표 1의 교과군명도 유효한 기준 항목으로 확인합니다. `기술⋅가정/정보`의 `기술·가정`, `정보`, `한문` 교과군을 포함하며, 기술·가정 교과군 이름에서만 `⋅`, `∙`, `·`, `ㆍ` 차이를 허용합니다.
        - 여러 파일을 올리면 파일별 선택 시트와 헤더 행을 따로 적용하고, 통합 결과에 파일명·시트명·엑셀 행을 기록합니다.
        - 기준 표 1의 보통교과만 해당하면 이 검토 규칙에서 M열 추가 표기를 요구하지 않습니다.
        - 승인 과목의 전문 여부는 승인 목록의 관련 교과 표기와 표 3의 전문교과 등재 정보를 근거로 판정하며, 근거가 표 안에 표시됩니다.
        - 업로드한 Excel은 앱의 검토 과정에서 읽기 전용으로 처리되며, 이 앱의 코드에서는 파일을 저장하거나 외부로 전송하지 않습니다.
        """
    )

st.markdown('<div class="footer">기준 목록은 제공된 문서를 바탕으로 재구성한 검토 보조 자료입니다. 최종 교육과정 편성·승인 여부는 원문 및 관할 교육청 자료와 대조해 확정해 주세요.</div>', unsafe_allow_html=True)
