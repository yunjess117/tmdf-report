# -*- coding: utf-8 -*-
"""월간보고서 PPT 채움 로직.

전월 PPT(고정 슬라이드 구성)를 템플릿으로 열어, 표 헤더/첫 열 패턴으로 표 종류를
인식한 뒤 이번 달 데이터로 채운다. 슬라이드 순서가 아니라 표 내용으로 인식하므로
슬라이드가 재배치돼도 동작한다. 스크린샷 기반 수기 표(운영 캘린더, 팔로워
추이/성별연령 차트, 검색 노출 현황, 이벤트/제휴 소개 등)는 건드리지 않는다.

로우데이터(엑셀) 쪽은 합계/평균/누적 등을 전월과 같은 '살아있는 수식'으로 쓰기
때문에 openpyxl로는 그 계산값을 읽을 수 없다(엑셀이 열어야 계산됨). 그래서 PPT에
필요한 실제 숫자는 (1) 이번 달 몫은 build_raw_data가 계산 과정에서 이미 만들어
둔 파이썬 값(block_cache)을 그대로 쓰고, (2) 전월 이전 몫은 전월 최종본을
data_only=True로 한 번 더 읽은 스냅샷(wb_data, 엑셀이 마지막 저장 때 캐시해 둔
계산값)에서 가져온다.
"""
import datetime as dt
import re
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE

from core.ppt_table import (
    set_row_count, set_cell_text, set_cell_hyperlink, set_run_color,
    UP_COLOR, DOWN_COLOR, NEUTRAL_COLOR,
)
from core.excel_block import CONFIRM_NEEDED
from core.formatters import month_to_int


def _fmt_num(v):
    if v is None or v == CONFIRM_NEEDED:
        return CONFIRM_NEEDED
    if isinstance(v, float):
        v = round(v)
    return f"{v:,}"


def _fmt_pct(v):
    if v is None or v == CONFIRM_NEEDED:
        return CONFIRM_NEEDED
    return f"{v * 100:.1f}%"


def _fmt_date_md(d):
    if not isinstance(d, (dt.date, dt.datetime)):
        return CONFIRM_NEEDED
    weekday = "월화수목금토일"[d.weekday()]
    return f"{d.month}/{d.day}({weekday})"


def _change_text_and_color(prev_val, cur_val, cell):
    if not isinstance(prev_val, (int, float)) or not isinstance(cur_val, (int, float)) or prev_val == 0:
        set_cell_text(cell, CONFIRM_NEEDED)
        set_run_color(cell, NEUTRAL_COLOR)
        return
    pct = (cur_val - prev_val) / prev_val * 100
    if pct > 0:
        set_cell_text(cell, f"▲ {pct:.1f}%")
        set_run_color(cell, UP_COLOR)
    elif pct < 0:
        set_cell_text(cell, f"▼ {abs(pct):.1f}%")
        set_run_color(cell, DOWN_COLOR)
    else:
        set_cell_text(cell, "-")


def _numval(ws_live, ws_data, row, col):
    """전월 최종본에 남아있는 수식 셀(예: '=인스타그램!D6')은 openpyxl로 계산값을
    읽을 수 없으므로, data_only 스냅샷(ws_data)에 캐시된 계산값으로 대체한다."""
    v = ws_live.cell(row=row, column=col).value
    if isinstance(v, (int, float)):
        return v
    if ws_data is not None:
        v2 = ws_data.cell(row=row, column=col).value
        if isinstance(v2, (int, float)):
            return v2
    return None


def _anyval(ws_live, ws_data, row, col):
    """숫자뿐 아니라 문자열도 대상인 셀(제목/라벨 등)까지 포괄하는 버전.
    실행 중(수식이 아직 계산 안 된) 셀은 살아있는 값을, 재업로드된(엑셀에서 한 번
    열어 저장된) 파일의 수식 셀은 data_only 캐시값을 우선한다."""
    v = ws_live.cell(row=row, column=col).value
    if isinstance(v, str) and v.startswith("="):
        if ws_data is not None:
            v2 = ws_data.cell(row=row, column=col).value
            return v2
        return None
    return v


def _fmt_num_or_confirm(v):
    return CONFIRM_NEEDED if v is None else _fmt_num(v)


# ---------------------------------------------------------------------------
# '채널 팔로워' 시트 리더(사람이 수기로 채워 넣는 시트 - 수식이 아니라 일반 값이지만,
# 혹시 실제 엑셀로 열어보지 않아 캐시가 비어 있을 경우까지 대비해 _anyval/_numval로 읽는다)
# ---------------------------------------------------------------------------

_FOLLOWER_SHEET = "채널 팔로워"


def _read_follower_monthly(wb, wb_data):
    """월별 요약표(6~14행)에서 {'7월': {'insta': 2168, 'blog': 221}, ...}를 만든다."""
    if _FOLLOWER_SHEET not in wb.sheetnames:
        return {}
    ws = wb[_FOLLOWER_SHEET]
    ws_data = wb_data[_FOLLOWER_SHEET] if wb_data and _FOLLOWER_SHEET in wb_data.sheetnames else None
    out = {}
    for r in range(8, 15):
        label = _anyval(ws, ws_data, r, 3)
        if not label:
            continue
        out[label] = {"insta": _numval(ws, ws_data, r, 4), "blog": _numval(ws, ws_data, r, 7)}
    return out


def _read_follower_daily(wb, wb_data, month_label):
    """일별 표(17행~)에서 해당 월(구분열=month_label) 행만 뽑는다.
    반환: [{'date':.., 'insta':.., 'blog_views':.., 'blog_visits':..}, ...]"""
    if _FOLLOWER_SHEET not in wb.sheetnames:
        return []
    ws = wb[_FOLLOWER_SHEET]
    ws_data = wb_data[_FOLLOWER_SHEET] if wb_data and _FOLLOWER_SHEET in wb_data.sheetnames else None
    out = []
    for r in range(19, ws.max_row + 1):
        if _anyval(ws, ws_data, r, 2) != month_label:
            continue
        date_v = _anyval(ws, ws_data, r, 3)
        if not isinstance(date_v, (dt.date, dt.datetime)):
            continue
        out.append({
            "date": date_v,
            "insta": _numval(ws, ws_data, r, 4),
            "blog_views": _numval(ws, ws_data, r, 7),
            "blog_visits": _numval(ws, ws_data, r, 9),
        })
    return out


# ---------------------------------------------------------------------------
# '제휴' 시트 리더(가장 최근에 append_partnership_block()이 이어붙인 달의 몫만)
# ---------------------------------------------------------------------------

def _read_partnership_sections(wb, wb_data):
    if "제휴" not in wb.sheetnames:
        return {"insta": ([], None), "blog": ([], None)}
    ws = wb["제휴"]
    ws_data = wb_data["제휴"] if wb_data and "제휴" in wb_data.sheetnames else None

    insta_header_row = 5
    blog_marker_row = None
    for r in range(insta_header_row + 1, ws.max_row + 1):
        if _anyval(ws, ws_data, r, 2) == "블로그":
            blog_marker_row = r
            break
    insta_end = (blog_marker_row - 1) if blog_marker_row else ws.max_row
    blog_header_row = (blog_marker_row + 1) if blog_marker_row else None

    def last_block(header_row, end_row, n_cols):
        if header_row is None:
            return [], None
        totals = [r for r in range(header_row + 1, end_row + 1)
                  if _anyval(ws, ws_data, r, 3) == "합계"]
        if not totals:
            return [], None
        total_row = totals[-1]
        prev_total = totals[-2] if len(totals) >= 2 else header_row
        rows = [[_anyval(ws, ws_data, r, c) for c in range(2, 2 + n_cols)]
                for r in range(prev_total + 1, total_row)]
        return rows, total_row

    insta_rows, insta_total = last_block(insta_header_row, insta_end, 10)
    blog_rows, blog_total = last_block(blog_header_row, ws.max_row, 11) if blog_header_row else ([], None)
    return {"insta": (insta_rows, insta_total), "blog": (blog_rows, blog_total)}


def _detect_topic_col(ws, header_row, default=5, search_rows=40):
    for r in range(header_row + 1, header_row + 1 + search_rows):
        for c in range(2, 12):
            if ws.cell(row=r, column=c).value in ("평균", "합계"):
                return c
    return default


def _last_block_rows_fallback(ws, ws_data, max_col):
    """block_cache에 이번 달 몫이 없을 때(=세션이 다른 독립 업로드 파일에서 PPT를
    만들 때) 시트에서 직접 마지막 블록(데이터+평균+합계)을 읽어온다. 수식 셀은
    data_only 스냅샷 값으로 대체한다."""
    header_row = None
    for r in range(1, 20):
        if ws.cell(row=r, column=2).value == "NO.":
            header_row = r
            break
    if header_row is None:
        return [], None, None
    topic_col = _detect_topic_col(ws, header_row)
    totals = [r for r in range(header_row + 1, ws.max_row + 1)
              if ws.cell(row=r, column=topic_col).value == "합계"]
    if not totals:
        return [], None, None
    total_row = totals[-1]
    avg_row = total_row - 1
    prev_total = totals[-2] if len(totals) >= 2 else header_row
    data_start = prev_total + 1

    def read_row(r):
        return [_anyval(ws, ws_data, r, c) for c in range(2, max_col + 1)]

    rows = [read_row(r) for r in range(data_start, avg_row)]
    avgs = read_row(avg_row)
    tot = read_row(total_row)
    return rows, avgs, tot


class XlsxSource:
    """방금 build_raw_data로 만든 워크북 + block_cache(파이썬 계산값)에서
    PPT에 필요한 값을 뽑아주는 헬퍼."""

    def __init__(self, wb, target_month, block_cache, wb_data=None):
        self.wb = wb
        self.wb_data = wb_data
        self.block_cache = block_cache or {}
        self.summary = self.block_cache.get("_summary", {})
        self.target_month = target_month
        self.month_num = month_to_int(target_month)

    def _cumulative(self, sheet_for_history, row, col_base, first_month, current_value):
        """월별 '열'에 값이 쌓이는 표(운영요약의 KPI/광고비 표)에서 누적값을 만든다.

        current_value가 있으면(이번 세션에서 방금 계산한 값) 이전 달은 data_only
        스냅샷으로 채우고 이번 달은 그 값을 더한다. current_value가 없으면(=block_cache
        없이 독립적으로 업로드된 '최종 로우데이터'로 PPT를 만드는 경우) 이번 달 몫까지
        포함해 전부 data_only 스냅샷에서 읽는다(엑셀에서 한 번 열어 저장된 파일이라
        이번 달 수식도 계산값이 캐시돼 있다고 가정)."""
        ws = self.wb[sheet_for_history]
        ws_data = self.wb_data[sheet_for_history] if self.wb_data else None
        last_month = (self.month_num - 1) if isinstance(current_value, (int, float)) else self.month_num
        total = 0
        for m in range(first_month, last_month + 1):
            c = col_base + (m - first_month)
            v = _numval(ws, ws_data, row, c)
            if v:
                total += v
        if isinstance(current_value, (int, float)):
            total += current_value
        return total

    # -- 운영요약 KPI --------------------------------------------------
    def kpi_row(self, metric_row, current_value):
        ws = self.wb["운영요약"]
        ws_data = self.wb_data["운영요약"] if self.wb_data else None
        if current_value is None:
            month_col = 3 + (self.month_num - 7)
            current_value = _numval(ws, ws_data, metric_row, month_col)
        cum = self._cumulative("운영요약", metric_row, 3, 7, current_value)
        target = ws.cell(row=metric_row, column=10).value  # J열: 고정 목표치(항상 리터럴)
        rate = (cum / target) if isinstance(target, (int, float)) and target else None
        return {"cur": current_value, "cum": cum, "target": target, "rate": rate}

    # -- 채널(인스타/블로그) 월 요약 -----------------------------------
    def channel_summary(self, sheet, prev_month_num, cur_values):
        """인스타그램/블로그 시트의 월별 요약(행6~11)에서 전월 값은 data_only
        스냅샷으로 읽고, 당월 값은 이번 실행에서 계산한 파이썬 값을 우선 쓴다.
        cur_values에 없는(None) 항목은 당월 열도 data_only 스냅샷에서 읽는다
        (block_cache 없이 독립 업로드된 파일로 PPT를 만드는 경우)."""
        ws = self.wb[sheet]
        ws_data = self.wb_data[sheet] if self.wb_data else None
        cols = {"발행수": 4, "지표2": 5, "지표3": 6, "지표4": 7}
        prev_row = 6 + (prev_month_num - 7)
        cur_row = 6 + (self.month_num - 7)
        out = {}
        for name, c in cols.items():
            prev_v = _numval(ws, ws_data, prev_row, c)
            cur_v = cur_values.get(name)
            if cur_v is None:
                cur_v = _numval(ws, ws_data, cur_row, c)
            out[name] = (prev_v, cur_v)
        return out

    # -- 콘텐츠 발행 내역(인스타/블로그 공통), AD 데이터 참여/도달/동영상조회 공통 --
    def block_rows(self, sheet, max_col=17):
        """block_cache에 저장해 둔 이번 달 (데이터 행, 평균 행, 합계 행)을 준다.
        block_cache에 없으면(독립 업로드된 파일로 PPT를 만드는 경우) 시트에서 직접
        마지막 블록을 읽어온다(수식 셀은 data_only 스냅샷 값으로 대체)."""
        entry = self.block_cache.get(sheet)
        if entry:
            return entry["rows"], entry["avg"], entry["total"]
        ws = self.wb[sheet]
        ws_data = self.wb_data[sheet] if self.wb_data else None
        return _last_block_rows_fallback(ws, ws_data, max_col)

    # -- AD 지출 누적/예산 ------------------------------------------------
    def ad_cumulative_spend(self, ad_type, row):
        current = self.summary.get("ad_totals_by_type", {}).get(ad_type, {}).get("지출금액")
        return self._cumulative("운영요약", row, 3, 7, current)

    def budget_summary(self):
        ws = self.wb["운영요약"]
        ws_data = self.wb_data["운영요약"] if self.wb_data else None
        total_budget = _numval(ws, ws_data, 32, 5)  # E32: 총 가용예산(고정값)
        cum = sum(self.ad_cumulative_spend(t, r) for t, r in
                  (("참여", 25), ("도달", 26), ("동영상조회", 27)))
        remain = (total_budget - cum) if isinstance(total_budget, (int, float)) else None
        return {"cum": cum, "remain": remain, "total": total_budget}


# ---------------------------------------------------------------------------
# 표 인식/분기
# ---------------------------------------------------------------------------

def _row_texts(table, r):
    return [c.text.strip() for c in table.rows[r].cells]


def _classify_table(table):
    try:
        header = _row_texts(table, 0)
    except IndexError:
        return None

    if header and header[0] == "구분" and "누적" in header and "목표" in header:
        return "kpi"
    if len(header) >= 5 and "인스타그램" in header and "블로그" in header:
        # 표26 스타일: 인스타/블로그를 한 표 안에 나란히 비교(6월/7월/전월 대비 x2)
        return "monthly_compare_dual"
    if len(header) >= 3 and header[0] == "" and "증감" in header[-1]:
        row1 = _row_texts(table, 1) if len(table.rows) > 1 else []
        if row1 and "광고" in row1[0]:
            return "ad_overview"
        return "channel_compare"
    if header and header[0] == "월" and "누적 집행예산" in header:
        return "budget_summary"
    if header and header[0] in ("NO", "NO.") and "발행일" in header and "광고비" in header:
        return "content_instagram"
    if header and header[0] in ("NO", "NO.") and "발행일" in header and "공감" in header:
        return "content_blog"
    if header and header[0] in ("NO", "NO.") and "VTR" in header:
        # 동영상조회 표는 '노출'/'도달' 컬럼도 같이 가지고 있어 아래 도달 표 조건과도
        # 겹치므로 더 구체적인 VTR 조건을 먼저 검사해야 한다.
        return "ad_video"
    if header and header[0] in ("NO", "NO.") and "참여" in header and "타깃" in header:
        return "ad_participation"
    if header and header[0] in ("NO", "NO.") and "도달" in header and "타깃" in header:
        return "ad_reach"
    if header and header[0] == "주제":
        row1 = _row_texts(table, 1) if len(table.rows) > 1 else []
        row2 = _row_texts(table, 2) if len(table.rows) > 2 else []
        if row1 and row1[0] == "이미지":
            if row2 and row2[0] == "도달수":
                return "content_top_reach"
            if row2 and row2[0].startswith("총 반응수"):
                return "content_top_engagement"
    if header and header[0] in ("NO", "NO.") and "방문 시장 및 청년몰" in header and "인스타 아이디" in header:
        return "partnership_insta"
    if header and header[0] in ("NO", "NO.") and "방문 시장 및 청년몰" in header and "진행 블로거" in header:
        return "partnership_blog"
    return None


# ---------------------------------------------------------------------------
# 표별 열 매핑: (ppt 열 index, xlsx 행(B열부터) index, 서식 종류)
# ---------------------------------------------------------------------------

INSTAGRAM_CONTENT_SPEC = [
    (0, 0, "no"), (1, 2, "date"), (2, 3, "url_text"), (3, 5, "text"),
    (4, 6, "num"), (5, 7, "num"), (6, 8, "num"), (7, 9, "num"), (8, 10, "num"),
    (9, 11, "num"), (10, 12, "num"), (11, 13, "num"), (12, 14, "num"), (13, 15, "num"),
]
BLOG_CONTENT_SPEC = [
    (0, 0, "no"), (1, 2, "date"), (2, 3, "url_text"), (3, 5, "text"),
    (4, 6, "num"), (5, 7, "num"), (6, 8, "num"), (7, 9, "num"),
]
AD_PARTICIPATION_SPEC = [
    (0, 0, "no"), (1, 3, "text"), (2, 4, "text"), (3, 5, "text"),
    (4, 6, "date"), (5, 7, "date"), (6, 8, "num"), (7, 9, "num"),
    (8, 10, "num"), (9, 11, "num"), (10, 12, "num"), (11, 13, "num"),
]
AD_REACH_SPEC = [
    (0, 0, "no"), (1, 3, "text"), (2, 4, "text"), (3, 5, "text"),
    (4, 6, "date"), (5, 7, "date"), (6, 8, "num"), (7, 9, "num"),
    (8, 10, "num"), (9, 12, "num"), (10, 13, "num"),
]
AD_VIDEO_SPEC = [
    (0, 0, "text"), (1, 1, "text"), (2, 2, "date"), (3, 3, "text"), (4, 4, "text"),
    (5, 5, "text"), (6, 6, "date"), (7, 7, "date"), (8, 8, "num"), (9, 9, "num"),
    (10, 10, "pct"), (11, 11, "num"), (12, 12, "num"), (13, 13, "num"), (14, 14, "num"),
]

CONTENT_URL_IDX = 4


def _format_value(v, kind):
    if v is None:
        return ""
    if v == CONFIRM_NEEDED:
        return CONFIRM_NEEDED
    if isinstance(v, str):
        return v  # '평균'/'합계'/'-' 등은 그대로
    if kind in ("date",):
        return _fmt_date_md(v)
    if kind == "num":
        return _fmt_num(v)
    if kind == "pct":
        return _fmt_pct(v)
    if kind == "no":
        return str(int(v)) if isinstance(v, (int, float)) else str(v)
    return str(v)


_CONTENT_COUNT_RE = re.compile(r"^-\s*총\s*\d+건\s*\(영상\s*\d+건,\s*카드뉴스\s*\d+건\)$")


def fill_content_count_label(slide, rows):
    """'- 총 7건 (영상 1건, 카드뉴스 6건)' 텍스트를 이번 달 실제 건수로 갱신한다."""
    video_n = sum(1 for r in rows if r[5] == "영상")
    card_n = sum(1 for r in rows if r[5] == "카드뉴스")
    for shape in slide.shapes:
        if shape.has_text_frame and _CONTENT_COUNT_RE.match(shape.text_frame.text.strip()):
            set_cell_text(shape, f"- 총 {len(rows)}건 (영상 {video_n}건, 카드뉴스 {card_n}건)")


def fill_data_table(table, spec, rows, avgs, totals, url_idx=None, url_ppt_col=None):
    n = len(rows)
    trailer_idxs = (len(table.rows) - 2, len(table.rows) - 1)
    set_row_count(table, n, template_row_idx=1, trailer_row_idxs=trailer_idxs)

    all_rows = list(rows)
    if avgs is not None:
        all_rows.append(avgs)
    if totals is not None:
        all_rows.append(totals)

    title_ppt_col = next((pc for pc, _, kind in spec if kind == "url_text"), None)

    from core.ppt_table import fill_row
    for i, xlsx_row in enumerate(all_rows):
        ppt_row = table.rows[1 + i]
        values = [_format_value(xlsx_row[xi] if xi < len(xlsx_row) else None, kind)
                  for _, xi, kind in spec]
        # '평균'/'합계' 라벨은 시트마다 다른 열(예: 블로그는 URL 열)에 찍혀 있을 수
        # 있어, spec의 고정 열 대신 행 안에서 라벨 문자열을 직접 찾아 제목 칸에 넣는다.
        label = next((v for v in xlsx_row if v in ("평균", "합계")), None)
        if label is not None and title_ppt_col is not None:
            values[title_ppt_col] = label
        url = None
        if url_idx is not None and url_idx < len(xlsx_row):
            url = xlsx_row[url_idx]
            if not isinstance(url, str) or not url.startswith("http"):
                url = None
        fill_row(ppt_row, values, url_col=url_ppt_col, url=url)


# ---------------------------------------------------------------------------
# 표별 채움 함수
# ---------------------------------------------------------------------------

_KPI_METRIC_ROW = {1: 14, 2: 15, 3: 16, 4: 17}  # ppt 표 row -> 운영요약 시트 row
_KPI_CUR_KEY = {14: "발행수", 15: "도달수", 16: "영상조회수", 17: "참여수"}


def fill_kpi_table(table, src: XlsxSource, confirm):
    header = _row_texts(table, 0)
    if src.target_month not in header:
        confirm.add("PPT/KPI", "달성현황 표", f"헤더에 {src.target_month} 열이 없어 건너뜀(양식 확인 필요)")
        return {}
    col = header.index(src.target_month)
    cum_col, target_col, rate_col = header.index("누적"), header.index("목표"), header.index("달성률")
    instagram_totals = src.summary.get("instagram_totals", {})
    results = {}
    for r, metric_row in _KPI_METRIC_ROW.items():
        label = table.rows[r].cells[0].text.strip()
        cur_value = instagram_totals.get(_KPI_CUR_KEY[metric_row])
        d = src.kpi_row(metric_row, cur_value)
        set_cell_text(table.rows[r].cells[col], _fmt_num(d["cur"]))
        set_cell_text(table.rows[r].cells[cum_col], _fmt_num(d["cum"]))
        set_cell_text(table.rows[r].cells[target_col], _fmt_num(d["target"]))
        set_cell_text(table.rows[r].cells[rate_col], _fmt_pct(d["rate"]))
        results[label] = d
    return results


_PCT_TEXT_RE = re.compile(r"^\d+(\.\d+)?%$")
_CUM_TEXT_RE = re.compile(r"(누적\s*실적\s*)([\d,]+)(.*)", re.DOTALL)


def fill_kpi_donut_and_text(slide, results, confirm):
    """KPI 표 옆 도넛 차트(달성/잔여)·그 아래 퍼센트 텍스트·'누적 실적' 텍스트를 표
    값과 맞춘다. 도넛 차트 series명이 표의 행 라벨과 같아서(예: '콘텐츠 발행수')
    그걸로 어떤 지표의 카드인지 식별하고, 같은 카드 안의 텍스트 상자는 차트와
    비슷한 x좌표(같은 카드 폭 안)에 있다는 것으로 찾는다."""
    if not results:
        return
    charts = []
    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        chart = shape.chart
        try:
            cats = list(chart.plots[0].categories)
        except Exception:
            continue
        if cats == ["달성", "잔여"]:
            charts.append((shape, chart))
    if not charts:
        return

    text_shapes = [s for s in slide.shapes if s.has_text_frame and not s.has_chart]

    for shape, chart in charts:
        label = chart.plots[0].series[0].name
        d = results.get(label)
        if d is None or d.get("rate") is None:
            confirm.add("PPT/KPI", f"달성 도넛차트({label})", "달성률을 계산하지 못해 건너뜀")
            continue
        achieved = round(d["rate"] * 100, 1)
        remaining = max(0.0, round(100 - achieved, 1))
        cd = CategoryChartData()
        cd.categories = ["달성", "잔여"]
        cd.add_series(label, (achieved, remaining))
        chart.replace_data(cd)

        band_lo, band_hi = shape.left - 400000, shape.left + shape.width + 400000
        for ts in text_shapes:
            if not (band_lo <= ts.left <= band_hi):
                continue
            text = ts.text_frame.text.strip()
            if _PCT_TEXT_RE.match(text):
                set_cell_text(ts, f"{achieved:.1f}%")
            elif text.startswith("누적 실적"):
                m = _CUM_TEXT_RE.match(text)
                if m:
                    set_cell_text(ts, f"{m.group(1)}{_fmt_num(d['cum'])}{m.group(3)}")


def _parse_num(text):
    if text in (None, "", "-", CONFIRM_NEEDED):
        return None
    try:
        return float(str(text).replace(",", "").strip())
    except ValueError:
        return None


_CHANNEL_ROW_SOURCE = {
    "팔로워 수": None, "이웃수": None,
    "콘텐츠 발행 수": "발행수", "콘텐츠 수량": "발행수",
    "콘텐츠 도달수 (전체)": "지표2", "영상 조회수": "지표3", "콘텐츠 참여수": "지표4",
    " 방문자수": None, "방문횟수": None, "블로그 총  조회수": None, "콘텐츠 조회수": "지표2",
}
_CHANNEL_FOLLOWER_LABELS = {"팔로워 수": "insta", "이웃수": "blog"}


def fill_channel_compare_table(table, sheet, src: XlsxSource, confirm):
    header = _row_texts(table, 0)
    if len(header) < 3:
        return
    prev_month_num = src.month_num - 1
    prev_label = header[2]  # 이번 실행 전 '당월' 헤더 -> 새 '전월' 헤더

    if sheet == "인스타그램":
        t = src.summary.get("instagram_totals", {})
        cur_values = {"발행수": t.get("발행수"), "지표2": t.get("도달수"),
                      "지표3": t.get("영상조회수"), "지표4": t.get("참여수")}
    else:
        t = src.summary.get("blog_totals", {})
        cur_values = {"발행수": t.get("발행수")}
    summary = src.channel_summary(sheet, prev_month_num, cur_values)
    follower_monthly = _read_follower_monthly(src.wb, src.wb_data)
    cur_follower = follower_monthly.get(src.target_month, {})

    set_cell_text(table.rows[0].cells[1], prev_label)
    set_cell_text(table.rows[0].cells[2], src.target_month)

    for r in range(1, len(table.rows)):
        label = table.rows[r].cells[0].text.strip()
        # 기존 '당월' 열 값을 새 '전월' 열로 이동
        prev_display = table.rows[r].cells[2].text.strip()
        set_cell_text(table.rows[r].cells[1], prev_display)

        follower_key = _CHANNEL_FOLLOWER_LABELS.get(label)
        metric = _CHANNEL_ROW_SOURCE.get(label)
        if follower_key is not None:
            cur_v = cur_follower.get(follower_key)
        else:
            cur_v = summary.get(metric, (None, None))[1] if metric else None
        if follower_key is None and (metric is None or cur_v is None):
            set_cell_text(table.rows[r].cells[2], CONFIRM_NEEDED)
            confirm.add("PPT/채널요약", f"{sheet} - {label}", "자동 산출 소스가 없어 확인 필요")
        elif cur_v is None:
            set_cell_text(table.rows[r].cells[2], CONFIRM_NEEDED)
            confirm.add("PPT/채널요약", f"{sheet} - {label}",
                        "'채널 팔로워' 시트에서 이번 달 값을 찾지 못해 확인 필요")
        else:
            set_cell_text(table.rows[r].cells[2], _fmt_num(cur_v))

        prev_num = _parse_num(prev_display)
        cur_num = _parse_num(table.rows[r].cells[2].text)
        _change_text_and_color(prev_num, cur_num, table.rows[r].cells[3])


_AD_BUDGET_TOTAL_SOURCES = (
    ("AD 데이터 (참여)", 13, 15), ("AD 데이터 (도달)", 13, 15), ("AD 데이터 (동영상조회)", 14, 16),
)


def _ad_budget_total(src: XlsxSource):
    """참여/도달/동영상조회 세 시트의 '예산' 합계 행을 더한 이번 달 총 광고 예산.
    (운영요약 49행은 '지출' 합계라 '예산'과 값이 다를 수 있어 따로 계산한다.)"""
    total, found = 0, False
    for sheet, idx, max_col in _AD_BUDGET_TOTAL_SOURCES:
        _rows, _avg, totals = src.block_rows(sheet, max_col=max_col)
        if totals and idx < len(totals) and isinstance(totals[idx], (int, float)):
            total += totals[idx]
            found = True
    return total if found else None


def fill_ad_overview_table(table, src: XlsxSource, confirm):
    ws = src.wb["운영요약"]
    ws_data = src.wb_data["운영요약"] if src.wb_data else None
    set_cell_text(table.rows[0].cells[1], ws.cell(row=36, column=3).value)
    set_cell_text(table.rows[0].cells[2], ws.cell(row=36, column=4).value)

    for r in range(1, len(table.rows)):
        row_num = 36 + r
        # 이 표의 몇몇 행(특히 '총 광고 예산'의 전월 칸)은 다음 달 자동화 때 참조
        # 수식(예: '=C28')으로 바뀌어, 라이브 셀을 그대로 읽으면 수식 문자열이 그대로
        # 노출된다. data_only 스냅샷 값으로 대체하는 _numval을 거쳐서 읽는다.
        prev_v = _numval(ws, ws_data, row_num, 3)
        cur_v = _numval(ws, ws_data, row_num, 4)
        chg_v = _numval(ws, ws_data, row_num, 5)
        label = table.rows[r].cells[0].text.strip()
        if label == "총 광고 예산":
            # 운영요약 시트의 이 행은 '지출(집행 스펜드)' 합계라 '예산'과 다를 수 있어
            # 세 AD 시트의 예산(만원 올림) 합계 행을 따로 더해 채운다.
            budget_total = _ad_budget_total(src)
            if budget_total is None:
                confirm.add("PPT/광고개요", "총 광고 예산", "예산 합계를 계산하지 못해 확인 필요")
            else:
                cur_v = budget_total
        set_cell_text(table.rows[r].cells[1], _fmt_num(prev_v) if prev_v is not None else CONFIRM_NEEDED)
        set_cell_text(table.rows[r].cells[2], _fmt_num(cur_v) if cur_v is not None else CONFIRM_NEEDED)
        if chg_v is not None:
            set_cell_text(table.rows[r].cells[3], f"{chg_v:+.1f}%")
        else:
            set_cell_text(table.rows[r].cells[3], CONFIRM_NEEDED)


def fill_budget_summary_table(table, src: XlsxSource):
    d = src.budget_summary()
    row = table.rows[1]
    set_cell_text(row.cells[0], src.target_month)
    set_cell_text(row.cells[1], _fmt_num(d["cum"]))
    set_cell_text(row.cells[2], _fmt_num(d["remain"]))
    set_cell_text(row.cells[3], _fmt_num(d["total"]))


# ---------------------------------------------------------------------------
# 전월 데이터 비교 표(인스타그램+블로그 한 표) — 슬라이드5 '표26' 스타일
# ---------------------------------------------------------------------------

def fill_monthly_compare_dual_table(table, src: XlsxSource, confirm):
    header1 = _row_texts(table, 1) if len(table.rows) > 1 else []
    if len(header1) < 7:
        confirm.add("PPT/월간비교", "전월 데이터 비교 표", "표 구조가 예상과 달라 건너뜀(양식 확인 필요)")
        return
    prev_month_num = src.month_num - 1
    prev_label = f"{prev_month_num}월"
    for c in (1, 4):
        set_cell_text(table.rows[1].cells[c], prev_label)
    for c in (2, 5):
        set_cell_text(table.rows[1].cells[c], src.target_month)

    instagram_totals = src.summary.get("instagram_totals", {})
    blog_totals = src.summary.get("blog_totals", {})
    insta_summary = src.channel_summary("인스타그램", prev_month_num, {
        "발행수": instagram_totals.get("발행수"), "지표2": instagram_totals.get("도달수"),
        "지표3": instagram_totals.get("영상조회수"), "지표4": instagram_totals.get("참여수"),
    })
    blog_summary = src.channel_summary("블로그", prev_month_num, {"발행수": blog_totals.get("발행수")})
    follower_monthly = _read_follower_monthly(src.wb, src.wb_data)
    prev_f = follower_monthly.get(prev_label, {})
    cur_f = follower_monthly.get(src.target_month, {})
    prev_daily = _read_follower_daily(src.wb, src.wb_data, prev_label)
    cur_daily = _read_follower_daily(src.wb, src.wb_data, src.target_month)

    def blog_daily_sum(daily_rows, key):
        vals = [r[key] for r in daily_rows if r[key] is not None]
        return sum(vals) if vals else None

    for r in range(2, len(table.rows)):
        label = table.rows[r].cells[0].text.strip()
        if label == "콘텐츠 발행수":
            ig_prev, ig_cur = insta_summary["발행수"]
            bl_prev, bl_cur = blog_summary["발행수"]
        elif label == "팔로워수 (이웃수)":
            ig_prev, ig_cur = prev_f.get("insta"), cur_f.get("insta")
            bl_prev = blog_daily_sum(prev_daily, "blog_views")
            bl_cur = blog_daily_sum(cur_daily, "blog_views")
        elif label == "콘텐츠 총 도달수(조회수)":
            ig_prev, ig_cur = insta_summary["지표2"]
            bl_prev = blog_daily_sum(prev_daily, "blog_visits")
            bl_cur = blog_daily_sum(cur_daily, "blog_visits")
        elif label == "콘텐츠 총 참여수":
            ig_prev, ig_cur = insta_summary["지표4"]
            bl_prev, bl_cur = None, None
        else:
            confirm.add("PPT/월간비교", label, "알 수 없는 행이라 건너뜀")
            continue

        for prev_v, cur_v, col_prev, col_cur, col_pct in (
            (ig_prev, ig_cur, 1, 2, 3), (bl_prev, bl_cur, 4, 5, 6),
        ):
            set_cell_text(table.rows[r].cells[col_prev], _fmt_num_or_confirm(prev_v))
            set_cell_text(table.rows[r].cells[col_cur], _fmt_num_or_confirm(cur_v))
            _change_text_and_color(prev_v, cur_v, table.rows[r].cells[col_pct])
    confirm.add("PPT/월간비교", "블로그 - 팔로워수(이웃수)/도달수(조회수)",
                "블로그 쪽은 '이웃수'가 거의 안 바뀌어 '채널 팔로워' 시트의 일별 조회수/방문횟수"
                " 합계로 대체 산출함 - 의도한 지표가 맞는지 확인 필요")


# ---------------------------------------------------------------------------
# 우수 콘텐츠 TOP3(도달수/반응수) — 슬라이드9/10
# ---------------------------------------------------------------------------

_TOP_REACH_ROWS = [
    ("도달수", 8, "num"), ("조회수", 7, "num"), ("총 반응수*", 15, "num"),
    ("콘텐츠 유형", 5, "text"), ("광고비 (원)", 6, "num"),
]
_TOP_ENGAGEMENT_ROWS = [
    ("총 반응수*", 15, "num"), ("좋아요", 11, "num"), ("댓글", 12, "num"),
    ("공유", 13, "num"), ("저장", 14, "num"),
]
_TOP_TITLE_RE = re.compile(r"^\d+월\s*우수\s*콘텐츠$")


def _rename_top_content_subtitle(slide, target_month):
    for shape in slide.shapes:
        if shape.has_text_frame and _TOP_TITLE_RE.match(shape.text_frame.text.strip()):
            set_cell_text(shape, f"{target_month} 우수 콘텐츠")


def _delete_all_pictures(slide):
    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            shape._element.getparent().remove(shape._element)


def fill_content_top_table(table, src: XlsxSource, rank_by_idx, row_spec, confirm, label,
                            exclude_event=False):
    rows, _avg, _total = src.block_rows("인스타그램", max_col=17)
    candidates = rows
    if exclude_event:
        # '반응수 TOP3'는 슬라이드 자체에 '※ 이벤트 콘텐츠 제외' 안내가 있다. 이벤트
        # 콘텐츠는 제목에 '이벤트'가 들어가는 관례(예: '[이벤트] ...', '8월 팔로우 이벤트')를
        # 따르므로 제목에 그 글자가 있으면 순위 대상에서 뺀다.
        candidates = [row for row in rows if "이벤트" not in str(row[3] or "")]
    ranked = sorted(
        (row for row in candidates if isinstance(row[rank_by_idx], (int, float))),
        key=lambda row: row[rank_by_idx], reverse=True,
    )
    top = ranked[:3]
    if len(top) < 3:
        confirm.add("PPT/우수콘텐츠", label, f"순위를 매길 콘텐츠가 {len(top)}건뿐이라 3건을 못 채움")

    n_cols = len(table.columns)
    for col in range(1, n_cols):
        if col - 1 >= len(top):
            for r in range(len(table.rows)):
                set_cell_text(table.rows[r].cells[col], "-")
            continue
        content = top[col - 1]
        set_cell_text(table.rows[0].cells[col], content[3] if content[3] else CONFIRM_NEEDED)  # 주제
        set_cell_text(table.rows[1].cells[col], "")  # 이미지 자리(직접 넣을 스크린샷용, 비워둠)
        for r, (_row_label, idx, kind) in enumerate(row_spec, start=2):
            set_cell_text(table.rows[r].cells[col], _format_value(content[idx], kind))


# ---------------------------------------------------------------------------
# 제휴(인플루언서) 표 — 슬라이드22
# ---------------------------------------------------------------------------

def fill_partnership_table(table, kind, src: XlsxSource, confirm):
    sections = _read_partnership_sections(src.wb, src.wb_data)
    rows, _total_row = sections["insta" if kind == "partnership_insta" else "blog"]
    if not rows:
        confirm.add("PPT/제휴", kind, "'제휴' 시트에서 이번 달 몫을 찾지 못해 건너뜀")
        return

    trailer_idxs = (len(table.rows) - 1,)
    set_row_count(table, len(rows), template_row_idx=1, trailer_row_idxs=trailer_idxs)

    if kind == "partnership_insta":
        # 시트 열(B부터): NO,오픈일,아이디,내용요약,팔로워수,조회수,좋아요수,댓글수,비고(시장),URL
        col_map = [(1, 8, "text"), (2, 1, "date"), (3, 2, "text"), (4, 3, "text"),
                   (6, 4, "num"), (7, 5, "num"), (8, 6, "num"), (9, 7, "num")]
        url_idx, url_ppt_col = 9, 5
        total_cols = {6: "num", 7: "num", 8: "num", 9: "num"}
    else:
        # 시트 열(B부터): NO,오픈일,블로거,제목,일평균방문자,오픈일방문자,PV수치,댓글수,공감수,비고(시장),URL
        col_map = [(1, 9, "text"), (2, 1, "date"), (3, 2, "text"), (4, 3, "text"),
                   (6, 4, "num"), (7, 5, "num"), (8, 6, "num"), (9, 7, "num"), (10, 8, "num")]
        url_idx, url_ppt_col = 10, 5
        total_cols = {6: "num", 7: "num", 8: "num", 9: "num", 10: "num"}

    for i, row in enumerate(rows):
        ppt_row = table.rows[1 + i]
        set_cell_text(ppt_row.cells[0], str(i + 1))
        for ppt_col, xlsx_idx, kind_fmt in col_map:
            set_cell_text(ppt_row.cells[ppt_col], _format_value(row[xlsx_idx], kind_fmt))
        url = row[url_idx] if url_idx < len(row) else None
        if isinstance(url, str) and url.startswith("http"):
            set_cell_hyperlink(ppt_row.cells[url_ppt_col], url)

    total_ppt_row = table.rows[1 + len(rows)]
    set_cell_text(total_ppt_row.cells[2], "합 계")
    for ppt_col, kind_fmt in total_cols.items():
        xlsx_idx = next(xi for pc, xi, _ in col_map if pc == ppt_col)
        vals = [row[xlsx_idx] for row in rows if isinstance(row[xlsx_idx], (int, float))]
        set_cell_text(total_ppt_row.cells[ppt_col], _fmt_num(sum(vals)) if vals else "-")


def _rename_partner_count_label(slide, insta_n, blog_n):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text.startswith("- 총") and "인" in text:
            band = shape.top
            # 인스타그래머 라벨(위쪽)과 블로거 라벨(아래쪽) 중 더 가까운 쪽의 인원수를 쓴다.
            nearest_label = min(
                (s for s in slide.shapes if s.has_text_frame and
                 s.text_frame.text.strip() in ("인스타그래머", "블로거")),
                key=lambda s: abs(s.top - band), default=None,
            )
            if nearest_label is None:
                continue
            n = insta_n if nearest_label.text_frame.text.strip() == "인스타그래머" else blog_n
            set_cell_text(shape, f"- 총 {n}인")


# ---------------------------------------------------------------------------
# 팔로워 추이 차트(월간 3개월/일간) — 슬라이드8
# ---------------------------------------------------------------------------

_DATE_CAT_RE = re.compile(r"^\d+/\d+\(")
_YEAR_MONTH_CAT_RE = re.compile(r"^\d+년\s*\d+월$")


def _monthly_values(monthly, month_num, key):
    """month_num-2..month_num 3개월 값. 개별 월이 비어 있으면(None) 0으로 채워
    chart.replace_data()에 None이 들어가 차트가 깨지는 것을 막는다."""
    labels = [f"{m}월" for m in range(month_num - 2, month_num + 1)]
    values = [monthly.get(lbl, {}).get(key) for lbl in labels]
    return labels, [v if isinstance(v, (int, float)) else 0 for v in values]


def fill_follower_trend_charts(prs, wb, wb_data, target_month, confirm):
    """8P(인스타 월간/일간 팔로워 추이)와 16P(블로그 이웃수 추이, 카테고리가
    '26년 n월' 형식이라 8P와 다르게 구분) 네이티브 차트를 '채널 팔로워' 시트로 채운다.
    이번 달 값 자체가 없으면(수기 입력 전 파일) 차트를 건드리지 않고 건너뛴다 -
    None을 그대로 넣으면 PowerPoint에서 차트가 비거나 깨지는 문제가 있었다."""
    month_num = month_to_int(target_month)
    monthly = _read_follower_monthly(wb, wb_data)
    daily = _read_follower_daily(wb, wb_data, target_month)
    daily = [d for d in daily if d["insta"] is not None]
    daily.sort(key=lambda d: d["date"])
    cur_insta = monthly.get(target_month, {}).get("insta")
    cur_blog = monthly.get(target_month, {}).get("blog")

    found_monthly = found_blog_monthly = found_daily = False
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_chart:
                continue
            chart = shape.chart
            try:
                cats = list(chart.plots[0].categories)
            except Exception:
                continue
            if not cats:
                continue
            series_names = [s.name for s in chart.plots[0].series]
            if not any("팔로워" in n for n in series_names):
                continue

            if len(cats) == 3 and any(_YEAR_MONTH_CAT_RE.match(str(c)) for c in cats):
                # 블로그 '이웃 수 추이' 차트: 카테고리가 '26년 8월'처럼 연도가 붙어있다.
                found_blog_monthly = True
                if cur_blog is None:
                    confirm.add("PPT/블로그", "이웃 수 추이", "'채널 팔로워' 시트에 이번 달 이웃수가 없어 건너뜀")
                    continue
                m = re.match(r"^(\d+년)", str(cats[-1]))
                year_prefix = m.group(1) if m else ""
                month_labels, values = _monthly_values(monthly, month_num, "blog")
                labels = [f"{year_prefix} {lbl}".strip() for lbl in month_labels]
                cd = CategoryChartData()
                cd.categories = labels
                cd.add_series(series_names[0], values)
                chart.replace_data(cd)
            elif len(cats) == 3 and not any(_DATE_CAT_RE.match(str(c)) for c in cats):
                # 인스타그램 '월간 팔로워 수 추이' 차트: 카테고리가 'n월' 형식(연도 없음).
                found_monthly = True
                if cur_insta is None:
                    confirm.add("PPT/팔로워", "월간 팔로워 수 추이", "'채널 팔로워' 시트에 이번 달 팔로워 수가 없어 건너뜀")
                    continue
                labels, values = _monthly_values(monthly, month_num, "insta")
                cd = CategoryChartData()
                cd.categories = labels
                cd.add_series("팔로워", values)
                chart.replace_data(cd)
            elif _DATE_CAT_RE.match(str(cats[0])):
                found_daily = True
                if not daily:
                    confirm.add("PPT/팔로워", "일간 팔로워 수", "'채널 팔로워' 시트에 이번 달 수기 입력 값이 없어 건너뜀")
                    continue
                labels = [_fmt_date_md(d["date"]) for d in daily]
                values = [d["insta"] for d in daily]
                cd = CategoryChartData()
                cd.categories = labels
                cd.add_series(series_names[0], values)
                chart.replace_data(cd)
    if not found_monthly:
        confirm.add("PPT/팔로워", "월간 팔로워 수 추이", "전월 PPT에서 월간 팔로워 추이 차트를 찾지 못함(템플릿 구성 확인 필요)")
    if not found_blog_monthly:
        confirm.add("PPT/블로그", "이웃 수 추이", "전월 PPT에서 블로그 이웃수 추이 차트를 찾지 못함(템플릿 구성 확인 필요)")
    if not found_daily:
        confirm.add("PPT/팔로워", "일간 팔로워 수", "전월 PPT에서 일간 팔로워 차트를 찾지 못함(템플릿 구성 확인 필요)")


# ---------------------------------------------------------------------------
# 인스타그램 팔로워 성별/연령 비중 (네이티브 차트) 채움
# ---------------------------------------------------------------------------

def _force_percent_format(chart):
    """차트 값은 0~1 사이 소수(0.492)로 저장하지만, replace_data() 후에는 데이터
    레이블/축 서식이 초기화(일반 숫자 표시)될 수 있다. 퍼센트로 보이도록 서식을
    명시적으로 다시 지정한다(예: 0.492 -> '49.2%')."""
    try:
        dls = chart.plots[0].data_labels
        dls.number_format = "0.0%"
        dls.number_format_is_linked = False
    except Exception:
        pass
    try:
        axis = chart.value_axis
        axis.tick_labels.number_format = "0.0%"
        axis.tick_labels.number_format_is_linked = False
    except Exception:
        pass


def fill_gender_age_charts(prs, target, confirm):
    """슬라이드의 파이차트(카테고리=['남성','여성'])와 막대차트(연령대별 남/녀)를
    찾아 target(FollowerTarget)의 값으로 갱신한다. chart.replace_data()를 써서
    캐시된 값뿐 아니라 차트에 내장된 데이터 시트도 같이 갱신되므로, PowerPoint에서
    '데이터 편집'으로 열었을 때도 정상적으로 수기 수정이 가능하다."""
    if target is None or not target.age_labels:
        return
    filled = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_chart:
                continue
            chart = shape.chart
            try:
                cats = list(chart.plots[0].categories)
            except Exception:
                continue
            if cats == ["남성", "여성"]:
                cd = CategoryChartData()
                cd.categories = cats
                cd.add_series("성별", (target.male_total / 100, target.female_total / 100))
                chart.replace_data(cd)
                _force_percent_format(chart)
                filled.append("성별 파이차트")
            elif len(cats) == len(target.age_labels) and len(chart.plots[0].series) == 2:
                cd = CategoryChartData()
                cd.categories = cats  # 템플릿 라벨(예: '24세 이하') 그대로 유지
                cd.add_series("남", [v / 100 for v in target.male_by_age])
                cd.add_series("녀", [v / 100 for v in target.female_by_age])
                chart.replace_data(cd)
                _force_percent_format(chart)
                filled.append("연령대별 성별 막대차트")
    if not filled:
        confirm.add("PPT/팔로워 비중", "성별·연령 차트",
                    "전월 PPT에서 성별/연령 차트를 찾지 못해 채우지 못함(템플릿 구성 확인 필요)")


# ---------------------------------------------------------------------------
# 최상위: PPT 조합
# ---------------------------------------------------------------------------

def build_ppt(prev_ppt_file, wb, target_month, confirm, block_cache=None, wb_data=None,
              follower_target=None):
    prs = Presentation(prev_ppt_file)
    src = XlsxSource(wb, target_month, block_cache, wb_data=wb_data)
    partnership_filled = False

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            table = shape.table
            kind = _classify_table(table)
            if kind is None:
                continue
            try:
                if kind == "kpi":
                    kpi_results = fill_kpi_table(table, src, confirm)
                    fill_kpi_donut_and_text(slide, kpi_results, confirm)
                elif kind == "monthly_compare_dual":
                    fill_monthly_compare_dual_table(table, src, confirm)
                elif kind == "channel_compare":
                    row1_label = _row_texts(table, 1)[0] if len(table.rows) > 1 else ""
                    sheet = "블로그" if row1_label in ("이웃수",) else "인스타그램"
                    fill_channel_compare_table(table, sheet, src, confirm)
                elif kind == "ad_overview":
                    fill_ad_overview_table(table, src, confirm)
                elif kind == "budget_summary":
                    fill_budget_summary_table(table, src)
                elif kind == "content_instagram":
                    rows, avgs, totals = src.block_rows("인스타그램", max_col=17)
                    fill_data_table(table, INSTAGRAM_CONTENT_SPEC, rows, avgs, totals,
                                     url_idx=CONTENT_URL_IDX, url_ppt_col=2)
                    fill_content_count_label(slide, rows)
                elif kind == "content_blog":
                    rows, avgs, totals = src.block_rows("블로그", max_col=11)
                    fill_data_table(table, BLOG_CONTENT_SPEC, rows, avgs, totals,
                                     url_idx=CONTENT_URL_IDX, url_ppt_col=2)
                elif kind == "ad_participation":
                    rows, avgs, totals = src.block_rows("AD 데이터 (참여)", max_col=15)
                    fill_data_table(table, AD_PARTICIPATION_SPEC, rows, avgs, totals)
                elif kind == "ad_reach":
                    rows, avgs, totals = src.block_rows("AD 데이터 (도달)", max_col=15)
                    fill_data_table(table, AD_REACH_SPEC, rows, avgs, totals)
                elif kind == "ad_video":
                    rows, avgs, totals = src.block_rows("AD 데이터 (동영상조회)", max_col=16)
                    fill_data_table(table, AD_VIDEO_SPEC, rows, avgs, totals)
                elif kind == "content_top_reach":
                    fill_content_top_table(table, src, 8, _TOP_REACH_ROWS, confirm, "도달수 TOP3")
                    _delete_all_pictures(slide)
                    _rename_top_content_subtitle(slide, target_month)
                elif kind == "content_top_engagement":
                    fill_content_top_table(table, src, 15, _TOP_ENGAGEMENT_ROWS, confirm,
                                            "반응수 TOP3", exclude_event=True)
                    _delete_all_pictures(slide)
                    _rename_top_content_subtitle(slide, target_month)
                elif kind in ("partnership_insta", "partnership_blog"):
                    fill_partnership_table(table, kind, src, confirm)
                    partnership_filled = True
            except Exception as e:
                confirm.add("PPT", kind, f"자동 채움 실패({e}) - 수기 확인 필요")

    if partnership_filled:
        sections = _read_partnership_sections(wb, wb_data)
        insta_n = len(sections["insta"][0])
        blog_n = len(sections["blog"][0])
        for slide in prs.slides:
            if any(shape.has_table and _classify_table(shape.table) in
                   ("partnership_insta", "partnership_blog") for shape in slide.shapes):
                _rename_partner_count_label(slide, insta_n, blog_n)

    fill_follower_trend_charts(prs, wb, wb_data, target_month, confirm)

    if follower_target is not None:
        fill_gender_age_charts(prs, follower_target, confirm)

    confirm.add("PPT", "자동화 범위 안내",
                "운영 캘린더·검색 상위 노출·이벤트 소개 슬라이드는 스크린샷·서술형 데이터라"
                " 자동 채움 대상이 아닙니다(전월 내용이 그대로 남아있으니 직접 교체 필요). "
                "우수 콘텐츠 TOP3의 스크린샷 이미지도 직접 채워야 합니다.")
    return prs
