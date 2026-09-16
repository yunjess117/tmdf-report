# -*- coding: utf-8 -*-
"""청년상인 로우데이터 취합 Streamlit 앱.

담당자가 매달 원본 자료를 업로드하면 전월 최종본에 이번 달 블록을 이어붙인
로우데이터(취합) 엑셀을 만들어준다. 청년상인은 인스타그램(+블로그) 채널만
운영하므로 LX세미콘 자동화보다 단순화된 단일 흐름으로 구성했다.
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

st.set_page_config(page_title="청년상인 로우데이터 취합", layout="wide")
st.title("청년상인 로우데이터 취합")
st.caption("전월 최종본 + 이번 달 발행리스트 + 채널 원본 데이터를 합쳐 이번 달 로우데이터를 완성합니다.")

st.header("1. 입력 파일")
col1, col2 = st.columns(2)
with col1:
    st.subheader("전월 최종본")
    prev_raw = st.file_uploader("전월 로우데이터 엑셀", type=["xlsx"], key="prev_raw")
    prev_ppt = st.file_uploader("전월 월간보고서 PPT (참고용, 현재 취합에는 사용하지 않음)",
                                 type=["pptx"], key="prev_ppt")
with col2:
    st.subheader("이번 달 입력 데이터")
    publish_list_file = st.file_uploader("이번 달 콘텐츠 발행리스트", type=["xlsx", "xls"], key="publish_list")

st.header("2. 채널 원본 데이터")
col3, col4 = st.columns(2)
with col3:
    content_raw_file = st.file_uploader(
        "인스타그램 콘텐츠 원본 (플랫폼 CSV 또는 정리본 xlsx)",
        type=["csv", "xlsx"], key="content_raw")
with col4:
    ad_report_file = st.file_uploader("인스타그램 AD 원본 (Meta Ads 내보내기)", type=["xlsx"], key="ad_report")
    partnership_file = st.file_uploader("제휴(인플루언서 체험단) 오픈 보고서", type=["xlsx"], key="partnership")

st.header("3. 취합 실행")

ready = prev_raw and publish_list_file and content_raw_file and ad_report_file
if not ready:
    st.info("전월 최종본, 발행리스트, 인스타그램 콘텐츠 원본, AD 원본은 필수 입력입니다.")

if st.button("취합 실행", disabled=not ready, type="primary"):
    try:
        year_hint = dt.date.today().year
        publish_lists = parse_publish_list(publish_list_file, year_hint)

        prev_bytes = io.BytesIO(prev_raw.read())
        prev_wb_probe = openpyxl.load_workbook(prev_bytes)
        target_month = detect_target_month(prev_wb_probe)
        prev_bytes.seek(0)

        content_perf = parse_content_raw(content_raw_file, content_raw_file.name)
        ad_rows = parse_ad_report(ad_report_file)
        partnership = {"인스타": [], "블로그": []}
        if partnership_file is not None:
            partnership = parse_partnership_report(partnership_file)

        wb, confirm = build_raw_data(prev_bytes, target_month, publish_lists,
                                      content_perf, ad_rows, partnership)

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        st.success(f"{target_month} 블록을 이어붙인 로우데이터를 완성했습니다.")
        st.download_button(
            "로우데이터(취합) 엑셀 다운로드",
            data=out,
            file_name=f"청년상인_{target_month}_월간리포트_로우데이터.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if confirm.items:
            st.warning(f"'확인 필요' 항목 {len(confirm.items)}건 — 원본에 없거나 자동 매칭에 실패해 수기 확인이 필요합니다.")
            st.dataframe(confirm.items, use_container_width=True)
        else:
            st.success("모든 항목이 자동으로 채워졌습니다.")
    except Exception as e:
        st.error(f"취합 중 오류가 발생했습니다: {e}")
        raise
