# -*- coding: utf-8 -*-
"""청년상인 로우데이터 취합 Streamlit 앱.

담당자가 매달 원본 자료를 업로드하면 전월 최종본에 이번 달 블록을 이어붙인
로우데이터(취합) 엑셀을 만들어준다. 청년상인은 인스타그램(+블로그) 채널만
운영하므로 LX세미콘 자동화보다 단순화된 단일 흐름으로 구성했다.
취합(엑셀)과 PPT는 2단계로 분리되어 있고, PPT는 로우데이터가 먼저 만들어진
뒤에만(그리고 전월 PPT가 업로드된 경우에만) 실행할 수 있다.
"""
import io
import datetime as dt
import streamlit as st

import openpyxl

from core.parsers.publish_list import parse_publish_list
from core.parsers.content_raw import parse_content_raw
from core.parsers.ad_report import parse_ad_report
from core.parsers.partnership_report import parse_partnership_report
from core.raw_data import build_raw_data, detect_target_month
from core.ppt_report import build_ppt

st.set_page_config(page_title="청년상인 로우데이터 취합", layout="wide")

COLORS = {"blue": "#2563EB", "green": "#16A34A", "purple": "#7C3AED"}

st.markdown("""
<style>
div[data-testid="stVerticalBlockBorderWrapper"] > div > div[data-testid="stVerticalBlock"] {
    padding: 4px 2px 2px 2px;
}
.card-accent {
    height: 5px;
    border-radius: 6px;
    margin: -0.9rem -0.1rem 14px -0.1rem;
}
.card-title {
    font-weight: 700;
    font-size: 16px;
    margin-bottom: 2px;
}
.card-desc {
    color: #6b7280;
    font-size: 12.5px;
    margin-bottom: 14px;
}
.section-row {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-top: 6px;
}
.section-badge {
    width: 28px;
    height: 28px;
    min-width: 28px;
    border-radius: 50%;
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 14px;
}
.section-title {
    font-size: 22px;
    font-weight: 700;
}
.section-desc {
    color: #6b7280;
    font-size: 13px;
    margin: 4px 0 16px 38px;
}
</style>
""", unsafe_allow_html=True)


def section_header(num, title, desc, color_key):
    color = COLORS[color_key]
    st.markdown(f"""
    <div class="section-row">
      <div class="section-badge" style="background:{color};">{num}</div>
      <div class="section-title" style="color:{color};">{title}</div>
    </div>
    <div class="section-desc">{desc}</div>
    """, unsafe_allow_html=True)


def card_top(color_key, title, desc=None):
    color = COLORS[color_key]
    desc_html = f'<div class="card-desc">{desc}</div>' if desc else ""
    st.markdown(f"""
    <div class="card-accent" style="background:{color};"></div>
    <div class="card-title">{title}</div>
    {desc_html}
    """, unsafe_allow_html=True)


st.title("청년상인 로우데이터 취합")
st.caption("전월 최종본 + 이번 달 발행리스트 + 채널 원본 데이터를 합쳐 이번 달 로우데이터와 PPT를 완성합니다.")

if "xlsx_result" not in st.session_state:
    st.session_state.xlsx_result = None
if "ppt_result" not in st.session_state:
    st.session_state.ppt_result = None

# ---------------------------------------------------------------------------
# 1. 입력 파일
# ---------------------------------------------------------------------------
section_header(1, "입력 파일", "전월 파일과 이번 달 입력 파일을 업로드하세요. ★는 필수입니다.", "blue")

col1, col2 = st.columns(2)
with col1, st.container(border=True):
    card_top("blue", "전월 최종본", "이번 달 블록을 이어붙일 기준 파일입니다.")
    prev_raw = st.file_uploader("★ 전월 로우데이터(취합) 엑셀", type=["xlsx"], key="prev_raw")
    prev_ppt = st.file_uploader("전월 월간보고서 PPT (선택 — 업로드하면 이번 달 PPT도 만들 수 있습니다)",
                                 type=["pptx"], key="prev_ppt")
with col2, st.container(border=True):
    card_top("blue", "이번 달 입력 데이터")
    publish_list_file = st.file_uploader("★ 이번 달 콘텐츠 발행리스트", type=["xlsx", "xls"], key="publish_list")

st.divider()

# ---------------------------------------------------------------------------
# 2. 채널 원본 데이터
# ---------------------------------------------------------------------------
section_header(2, "채널 원본 데이터", "채널별 성과 및 광고 데이터를 업로드하세요. 없으면 해당 항목은 '확인 필요'로 비워둡니다.", "green")

col3, col4 = st.columns(2)
with col3, st.container(border=True):
    card_top("green", "인스타그램")
    content_raw_file = st.file_uploader(
        "★ 인스타그램 콘텐츠 원본 (플랫폼 CSV 또는 정리본 xlsx)",
        type=["csv", "xlsx"], key="content_raw")
with col4, st.container(border=True):
    card_top("green", "인스타그램 AD")
    ad_report_file = st.file_uploader("★ 인스타그램 AD 원본 (Meta Ads 내보내기)", type=["xlsx"], key="ad_report")
    partnership_file = st.file_uploader("제휴(인플루언서 체험단) 오픈 보고서 (선택)", type=["xlsx"], key="partnership")

st.divider()

# ---------------------------------------------------------------------------
# 3. 취합 실행
# ---------------------------------------------------------------------------
section_header(3, "취합 실행", "1) 로우데이터를 먼저 만들고, 2) 그 결과로 PPT를 만듭니다.", "purple")

xlsx_ready = bool(prev_raw and publish_list_file and content_raw_file and ad_report_file)
if not xlsx_ready:
    st.info("전월 최종본, 발행리스트, 인스타그램 콘텐츠 원본, AD 원본은 필수 입력입니다.")

col_step1, col_step2 = st.columns(2)

with col_step1, st.container(border=True):
    card_top("purple", "1) 로우데이터")
    if st.button("로우데이터 만들기", disabled=not xlsx_ready, type="primary", key="build_xlsx_btn"):
        try:
            year_hint = dt.date.today().year
            publish_lists = parse_publish_list(publish_list_file, year_hint)

            prev_bytes = io.BytesIO(prev_raw.getvalue())
            prev_wb_probe = openpyxl.load_workbook(prev_bytes, data_only=True)
            target_month = detect_target_month(prev_wb_probe)
            prev_bytes.seek(0)

            content_perf = parse_content_raw(io.BytesIO(content_raw_file.getvalue()), content_raw_file.name)
            ad_rows = parse_ad_report(io.BytesIO(ad_report_file.getvalue()))
            partnership = {"인스타": [], "블로그": []}
            if partnership_file is not None:
                partnership = parse_partnership_report(io.BytesIO(partnership_file.getvalue()))

            wb, wb_data, confirm = build_raw_data(prev_bytes, target_month, publish_lists,
                                                   content_perf, ad_rows, partnership)

            out = io.BytesIO()
            wb.save(out)
            xlsx_bytes = out.getvalue()

            st.session_state.xlsx_result = {
                "wb": wb, "wb_data": wb_data, "confirm_items": list(confirm.items),
                "target_month": target_month, "xlsx_bytes": xlsx_bytes,
            }
            st.session_state.ppt_result = None  # 로우데이터를 다시 만들면 이전 PPT 결과는 무효화
        except Exception as e:
            st.error(f"로우데이터 취합 중 오류가 발생했습니다: {e}")
            raise

    xres = st.session_state.xlsx_result
    if xres:
        st.success(f"{xres['target_month']} 블록을 이어붙인 로우데이터를 완성했습니다.")
        st.download_button(
            "로우데이터(취합) 엑셀 다운로드",
            data=xres["xlsx_bytes"],
            file_name=f"청년상인_{xres['target_month']}_월간리포트_로우데이터.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_xlsx",
        )

with col_step2, st.container(border=True):
    card_top("purple", "2) PPT")
    ppt_ready = bool(st.session_state.xlsx_result) and prev_ppt is not None
    if not st.session_state.xlsx_result:
        st.caption("먼저 왼쪽에서 로우데이터를 만들어주세요.")
    elif prev_ppt is None:
        st.caption("전월 월간보고서 PPT를 업로드하면 PPT도 만들 수 있습니다.")

    if st.button("PPT 만들기", disabled=not ppt_ready, type="primary", key="build_ppt_btn"):
        try:
            xres = st.session_state.xlsx_result
            from core.confirm import ConfirmLog
            confirm = ConfirmLog()
            confirm.items = list(xres["confirm_items"])

            ppt_bytes_in = io.BytesIO(prev_ppt.getvalue())
            prs = build_ppt(ppt_bytes_in, xres["wb"], xres["target_month"], confirm, wb_data=xres["wb_data"])
            ppt_out = io.BytesIO()
            prs.save(ppt_out)

            xres["confirm_items"] = list(confirm.items)  # PPT 단계에서 추가된 확인 필요 항목 반영
            st.session_state.ppt_result = {
                "ppt_bytes": ppt_out.getvalue(), "target_month": xres["target_month"],
            }
        except Exception as e:
            st.error(f"PPT 생성 중 오류가 발생했습니다: {e}")
            raise

    pres = st.session_state.ppt_result
    if pres:
        st.success(f"{pres['target_month']} 월간보고서 PPT를 완성했습니다.")
        st.download_button(
            "월간보고서 PPT 다운로드",
            data=pres["ppt_bytes"],
            file_name=f"청년상인_{pres['target_month']}_SNS_운영_월간보고서.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="dl_ppt",
        )

# ---------------------------------------------------------------------------
# 확인 필요 항목 (로우데이터/PPT 통합)
# ---------------------------------------------------------------------------
if st.session_state.xlsx_result:
    items = st.session_state.xlsx_result["confirm_items"]
    if items:
        st.warning(f"'확인 필요' 항목 {len(items)}건 — 원본에 없거나 자동 매칭에 실패해 수기 확인이 필요합니다.")
        st.dataframe(items, use_container_width=True)
    else:
        st.success("모든 항목이 자동으로 채워졌습니다.")
