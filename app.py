# -*- coding: utf-8 -*-
"""청년상인 로우데이터 취합 Streamlit 앱.

'로우데이터'와 '월간보고서 PPT'가 서로 독립된 두 흐름으로 나뉘어 있다.
- 로우데이터: 전월 최종본에 이번 달 블록을 이어붙인 엑셀을 만든다.
- 월간보고서 PPT: '최종 로우데이터'(로우데이터 흐름의 결과물이든, 사람이 팔로워/
  블로그 등 확인 필요 항목을 손으로 채운 뒤 다시 올린 파일이든)를 독립적으로
  업로드받아 전월 PPT에 채운다. 제휴 오픈보고서·팔로워 비중 CSV도 이 단계에서
  '최종 로우데이터' 엑셀에 반영한 뒤 PPT를 만든다(엑셀만 고치고, 그 값을 PPT가
  읽어가는 방식이라 나중에 엑셀만 수정해도 다시 PPT를 만들면 반영된다).
"""
import io
import datetime as dt
import streamlit as st

import openpyxl

from core.parsers.publish_list import parse_publish_list
from core.parsers.content_raw import parse_content_raw
from core.parsers.ad_report import parse_ad_report
from core.parsers.partnership_report import parse_partnership_report
from core.parsers.follower_target import parse_follower_target
from core.raw_data import (
    build_raw_data, detect_target_month, detect_latest_filled_month,
    append_partnership_block, write_follower_target,
)
from core.ppt_report import build_ppt
from core.confirm import ConfirmLog

st.set_page_config(page_title="청년상인 로우데이터 취합", layout="wide")

# 로우데이터=블루, 월간보고서 PPT=퍼플 두 계열로만 색을 쓴다(카드별 개별 색상 없음).
# 주의사항(수기 입력/확인 필요)에만 제한적으로 앰버를 쓴다.
COLORS = {"blue": "#2563EB", "purple": "#7C5CE8", "amber": "#F59E0B"}
PANEL_BG = {"blue": "#EFF6FF", "purple": "#F5F3FF"}
TEXT_PRIMARY = "#172033"
TEXT_SECONDARY = "#7B8494"
BORDER_COLOR = "#E5E7EB"
CARD_BG = "#FFFFFF"
AMBER_BG = "#FFF8E7"

st.markdown(f"""
<style>
.stApp {{ background: #FFFFFF; }}
h1, h2, h3, p, span, label, div {{ color: {TEXT_PRIMARY}; }}
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * {{ color: {TEXT_SECONDARY} !important; }}

/* 같은 행에 나란히 있는 카드(컬럼)들의 너비와 높이를 강제로 맞춘다.
   Streamlit의 stColumn은 기본이 display:block이고 그 안의 카드 컨테이너
   (stLayoutWrapper)는 flex-grow:0이라, 내용이 짧은 카드는 높이가 안 늘어나서
   같은 행 카드끼리 하단선이 들쭉날쭉해진다 - block을 flex-column으로 바꾸고
   flex-grow/height:100%를 체인 전체(컬럼 -> 카드 래퍼 -> 카드 내부 블록)에
   강제로 걸어 내용이 짧은 카드도 옆 카드 높이만큼 늘어나도록 만든다. */
div[data-testid="stHorizontalBlock"] {{
    align-items: stretch;
}}
div[data-testid="stColumn"] {{
    display: flex;
    flex-direction: column;
}}
div[data-testid="stColumn"] > div[data-testid="stVerticalBlock"] {{
    flex: 1 1 auto;
    height: 100%;
}}
div[data-testid="stColumn"] div[data-testid="stLayoutWrapper"] {{
    flex: 1 1 auto;
    height: 100%;
    display: flex;
    flex-direction: column;
}}

/* 카드(업로드 단위) 공통 스타일 - 흰 배경 + 연한 중립 테두리, 그림자/그라데이션 없음.
   모든 카드가 동일한 스타일을 쓰고, 영역 구분은 아래 패널 배경 틴트와 accent
   바/배지/버튼 색으로만 표현한다. */
div[data-testid="stColumn"] div[data-testid="stLayoutWrapper"] > div[data-testid="stVerticalBlock"] {{
    flex: 1 1 auto;
    height: 100%;
    background: {CARD_BG};
    border: 1px solid {BORDER_COLOR} !important;
    border-radius: 10px;
    box-shadow: none !important;
    padding: 14px 16px 12px 16px !important;
}}

/* 패널(로우데이터/월간보고서 PPT) 전체 배경 틴트 - 위 카드 규칙 다음에 와서
   같은 우선순위에서 소스 순서로 이긴다. key=panel_raw/panel_ppt로 지정한
   컨테이너에 Streamlit이 붙여주는 st-key-* 클래스를 사용. */
div.st-key-panel_raw > div[data-testid="stVerticalBlock"] {{
    background: {PANEL_BG["blue"]} !important;
    border: 1px solid {BORDER_COLOR} !important;
    border-radius: 14px !important;
}}
div.st-key-panel_ppt > div[data-testid="stVerticalBlock"] {{
    background: {PANEL_BG["purple"]} !important;
    border: 1px solid {BORDER_COLOR} !important;
    border-radius: 14px !important;
}}

.card-accent {{
    height: 4px;
    border-radius: 4px;
    margin: -0.9rem -0.1rem 8px -0.1rem;
}}
.card-title {{
    font-weight: 600;
    font-size: 15px;
    color: {TEXT_PRIMARY};
    margin-bottom: 2px;
}}
.card-desc {{
    color: {TEXT_SECONDARY};
    font-size: 12.5px;
    margin-bottom: 8px;
}}
.manual-note {{
    color: #8A5A0A;
    background: {AMBER_BG};
    border-left: 3px solid {COLORS["amber"]};
    border-radius: 4px;
    font-size: 11.5px;
    padding: 5px 9px;
    margin-top: 5px;
    line-height: 1.4;
}}
.section-row {{
    display: flex;
    align-items: center;
    gap: 10px;
    margin-top: 6px;
}}
.section-badge {{
    width: 26px;
    height: 26px;
    min-width: 26px;
    border-radius: 50%;
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 13px;
}}
.section-title {{
    font-size: 19px;
    font-weight: 700;
}}
.section-desc {{
    color: {TEXT_SECONDARY};
    font-size: 12.5px;
    margin: 4px 0 14px 36px;
}}
.sheet-title {{
    font-size: 22px;
    font-weight: 800;
    margin-bottom: 2px;
}}
div.stButton > button[kind="primary"] {{
    border-radius: 8px;
    box-shadow: none;
}}
div.st-key-build_xlsx_btn button[kind="primary"] {{
    background-color: {COLORS["blue"]};
    border-color: {COLORS["blue"]};
}}
div.st-key-build_ppt_btn button[kind="primary"] {{
    background-color: {COLORS["purple"]};
    border-color: {COLORS["purple"]};
}}
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


def manual_note(text):
    st.markdown(f'<div class="manual-note">✏️ 수기 입력/확인 필요: {text}</div>', unsafe_allow_html=True)


st.title("청년상인 로우데이터 취합")
st.caption("로우데이터와 월간보고서 PPT는 서로 독립적으로 만들 수 있습니다.")

for key in ("xlsx_result", "ppt_result"):
    if key not in st.session_state:
        st.session_state[key] = None

col_raw, col_ppt = st.columns(2, gap="large")

# =============================================================================
# 왼쪽: 로우데이터
# =============================================================================
with col_raw, st.container(border=True, key="panel_raw"):
    st.markdown(f'<div class="sheet-title" style="color:{COLORS["blue"]};">로우데이터</div>', unsafe_allow_html=True)

    section_header(1, "입력 파일", "전월 로우데이터와 이번 달 발행리스트를 업로드하세요. ★는 필수입니다.", "blue")
    c1, c2 = st.columns(2)
    with c1, st.container(border=True):
        card_top("blue", "전월 로우데이터", "이번 달 블록을 이어붙일 기준 파일입니다.")
        prev_raw = st.file_uploader("★ 전월 로우데이터(취합) 엑셀", type=["xlsx"], key="prev_raw")
    with c2, st.container(border=True):
        card_top("blue", "이번 달 발행리스트")
        publish_list_file = st.file_uploader("★ 이번 달 콘텐츠 발행리스트", type=["xlsx", "xls"], key="publish_list")
        manual_note("발행리스트에 있는 블로그 콘텐츠는 성과 원본이 없어 조회/공감/댓글이 "
                    "'확인 필요'로 남습니다 — 엑셀에서 직접 채워주세요.")

    section_header(2, "채널 원본 데이터", "채널별 성과 및 광고 데이터를 업로드하세요. 없으면 해당 항목은 '확인 필요'로 비워둡니다.", "blue")
    c3, c4 = st.columns(2)
    with c3, st.container(border=True):
        card_top("blue", "콘텐츠 데이터")
        content_raw_file = st.file_uploader(
            "★ 인스타그램 콘텐츠 원본 (플랫폼 CSV 또는 정리본 xlsx)",
            type=["csv", "xlsx"], key="content_raw")
    with c4, st.container(border=True):
        card_top("blue", "AD 데이터")
        ad_report_file = st.file_uploader("★ 인스타그램 AD 원본 (Meta Ads 내보내기)", type=["xlsx"], key="ad_report")
        manual_note("광고의 '타깃' 설정값은 원본에 없어 '확인 필요'로 남습니다 — 엑셀에서 직접 채워주세요. "
                    "제휴(인플루언서 체험단) 데이터는 오른쪽 'PPT 만들기'에서 입력합니다.")

    st.write("")
    xlsx_ready = bool(prev_raw and publish_list_file and content_raw_file and ad_report_file)
    if not xlsx_ready:
        st.info("전월 로우데이터, 발행리스트, 콘텐츠 데이터, AD 데이터는 필수 입력입니다.")

    if st.button("로우데이터 만들기", disabled=not xlsx_ready, type="primary", key="build_xlsx_btn",
                 use_container_width=True):
        try:
            year_hint = dt.date.today().year
            publish_lists = parse_publish_list(publish_list_file, year_hint)

            prev_bytes_raw = prev_raw.getvalue()
            prev_wb_probe = openpyxl.load_workbook(io.BytesIO(prev_bytes_raw), data_only=True)
            target_month = detect_target_month(prev_wb_probe)

            content_perf = parse_content_raw(io.BytesIO(content_raw_file.getvalue()), content_raw_file.name)
            ad_rows = parse_ad_report(io.BytesIO(ad_report_file.getvalue()))

            wb, block_cache, confirm = build_raw_data(io.BytesIO(prev_bytes_raw), target_month,
                                                       publish_lists, content_perf, ad_rows)

            out = io.BytesIO()
            wb.save(out)
            xlsx_bytes = out.getvalue()

            st.session_state.xlsx_result = {
                "confirm_items": list(confirm.items),
                "target_month": target_month, "xlsx_bytes": xlsx_bytes,
            }
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
            key="dl_xlsx", use_container_width=True,
        )
        items = xres["confirm_items"]
        if items:
            st.warning(f"'확인 필요' 항목 {len(items)}건")
            st.dataframe(items, use_container_width=True)
        else:
            st.success("모든 항목이 자동으로 채워졌습니다.")

# =============================================================================
# 오른쪽: 월간보고서 PPT
# =============================================================================
with col_ppt, st.container(border=True, key="panel_ppt"):
    st.markdown(f'<div class="sheet-title" style="color:{COLORS["purple"]};">월간보고서 PPT</div>', unsafe_allow_html=True)
    st.caption("로우데이터는 여기서 별도로 업로드합니다 — 로우데이터를 만든 뒤 팔로워/블로그 등 확인 필요 "
               "항목을 엑셀에서 직접 채우고 그 파일을 올려주세요.")

    section_header(1, "입력 파일", "전월 PPT 템플릿과 이번 달 최종 로우데이터를 업로드하세요. ★는 필수입니다.", "purple")
    cp1, cp2 = st.columns(2)
    with cp1, st.container(border=True):
        card_top("purple", "전월 월간보고서", "이번 달 PPT를 만들 때 템플릿으로 씁니다.")
        prev_ppt = st.file_uploader("전월 월간보고서 PPT (선택)", type=["pptx"], key="prev_ppt")
    with cp2, st.container(border=True):
        card_top("purple", "이번 달 로우데이터(최종)")
        final_raw_file = st.file_uploader("★ 최종 로우데이터(취합) 엑셀", type=["xlsx"], key="final_raw")
        manual_note("왼쪽에서 만든 뒤, 팔로워 수·블로그 성과처럼 '확인 필요'로 남은 값을 엑셀에서 "
                    "직접 채우고 나서 올려주세요. 수식은 그대로 두셔도 됩니다.")

    section_header(2, "기타 데이터(PPT용)", "필요할 때만 올리세요. 없으면 해당 항목은 자동으로 채워지지 않습니다.", "purple")
    c5, c6 = st.columns(2)
    with c5, st.container(border=True):
        card_top("purple", "인스타그램 팔로워")
        follower_csv_file = st.file_uploader(
            "인스타그램 팔로워 비중(CSV, 선택)", type=["csv"], key="follower_csv")
        manual_note("Meta 인스타그램 인사이트 '타깃' 내보내기 CSV입니다. 올리면 로우데이터에 "
                    "'인스타그램 팔로워 타깃' 시트가 추가되고 PPT의 성별/연령 차트도 갱신됩니다. "
                    "값은 그 시트에서 언제든 손으로 고칠 수 있습니다.")
    with c6, st.container(border=True):
        card_top("purple", "제휴 오픈보고서")
        partnership_file_ppt = st.file_uploader(
            "제휴(인플루언서 체험단) 오픈 보고서 (선택)", type=["xlsx"], key="partnership_ppt")
        manual_note("올리지 않으면 로우데이터에 '제휴' 시트가 추가되지 않습니다.")

    st.write("")
    ppt_ready = bool(final_raw_file) and bool(prev_ppt)
    if not final_raw_file:
        st.info("이번 달 로우데이터(최종) 엑셀은 필수 입력입니다.")
    elif not prev_ppt:
        st.info("위 '전월 월간보고서' 카드에 전월 PPT를 업로드해주세요(이번 달 PPT의 템플릿으로 사용됩니다).")

    if st.button("PPT 만들기", disabled=not ppt_ready, type="primary", key="build_ppt_btn",
                 use_container_width=True):
        try:
            prev_bytes_final = final_raw_file.getvalue()
            wb_final = openpyxl.load_workbook(io.BytesIO(prev_bytes_final))
            wb_final_data = openpyxl.load_workbook(io.BytesIO(prev_bytes_final), data_only=True)
            target_month_ppt = detect_latest_filled_month(wb_final)

            confirm = ConfirmLog()

            if partnership_file_ppt is not None:
                partnership = parse_partnership_report(io.BytesIO(partnership_file_ppt.getvalue()))
                append_partnership_block(wb_final, target_month_ppt,
                                          partnership["인스타"], partnership["블로그"])

            follower_target = None
            if follower_csv_file is not None:
                follower_target = parse_follower_target(io.BytesIO(follower_csv_file.getvalue()))
                write_follower_target(wb_final, target_month_ppt, follower_target)

            updated_xlsx_out = io.BytesIO()
            wb_final.save(updated_xlsx_out)
            updated_xlsx_bytes = updated_xlsx_out.getvalue()

            ppt_bytes_in = io.BytesIO(prev_ppt.getvalue())
            prs = build_ppt(ppt_bytes_in, wb_final, target_month_ppt, confirm,
                             block_cache=None, wb_data=wb_final_data,
                             follower_target=follower_target)
            ppt_out = io.BytesIO()
            prs.save(ppt_out)

            st.session_state.ppt_result = {
                "ppt_bytes": ppt_out.getvalue(),
                "updated_xlsx_bytes": updated_xlsx_bytes,
                "target_month": target_month_ppt,
                "confirm_items": list(confirm.items),
                "wrote_extra": bool(partnership_file_ppt) or bool(follower_csv_file),
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
            key="dl_ppt", use_container_width=True,
        )
        if pres["wrote_extra"]:
            st.download_button(
                "제휴/팔로워 타깃이 반영된 로우데이터 다운로드",
                data=pres["updated_xlsx_bytes"],
                file_name=f"청년상인_{pres['target_month']}_월간리포트_로우데이터_업데이트.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx_updated", use_container_width=True,
            )
        items = pres["confirm_items"]
        if items:
            st.warning(f"'확인 필요' 항목 {len(items)}건")
            st.dataframe(items, use_container_width=True)
        else:
            st.success("모든 항목이 자동으로 채워졌습니다.")
