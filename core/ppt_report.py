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
from pptx import Presentation

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
        """월별 '열'에 값이 쌓이는 표(운영요약의 KPI/광고비 표)에서, 이번 달보다
        이전 달들은 data_only 스냅샷으로, 이번 달은 방금 계산한 파이썬 값으로
        더해 진짜 누적값을 만든다."""
        ws = self.wb[sheet_for_history]
        ws_data = self.wb_data[sheet_for_history] if self.wb_data else None
        total = 0
        for m in range(first_month, self.month_num):
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
        cum = self._cumulative("운영요약", metric_row, 3, 7, current_value)
        target = ws.cell(row=metric_row, column=10).value  # J열: 고정 목표치(항상 리터럴)
        rate = (cum / target) if isinstance(target, (int, float)) and target else None
        return {"cur": current_value, "cum": cum, "target": target, "rate": rate}

    # -- 채널(인스타/블로그) 월 요약 -----------------------------------
    def channel_summary(self, sheet, prev_month_num, cur_values):
        """인스타그램/블로그 시트의 월별 요약(행6~11)에서 전월 값은 data_only
        스냅샷으로 읽고, 당월 값은 이번 실행에서 계산한 파이썬 값을 그대로 쓴다.
        cur_values: {"발행수":.., "지표2":.., "지표3":.., "지표4":..} (없는 키는 None)"""
        ws = self.wb[sheet]
        ws_data = self.wb_data[sheet] if self.wb_data else None
        cols = {"발행수": 4, "지표2": 5, "지표3": 6, "지표4": 7}
        prev_row = 6 + (prev_month_num - 7)
        out = {}
        for name, c in cols.items():
            prev_v = _numval(ws, ws_data, prev_row, c)
            out[name] = (prev_v, cur_values.get(name))
        return out

    # -- 콘텐츠 발행 내역(인스타/블로그 공통), AD 데이터 참여/도달/동영상조회 공통 --
    def block_rows(self, sheet):
        """block_cache에 저장해 둔 이번 달 (데이터 행, 평균 행, 합계 행)을 그대로 준다."""
        entry = self.block_cache.get(sheet)
        if not entry:
            return [], None, None
        return entry["rows"], entry["avg"], entry["total"]

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
    (8, 10, "num"), (9, 11, "num"), (10, 12, "num"), (11, 13, "text"),
]
AD_REACH_SPEC = [
    (0, 0, "no"), (1, 3, "text"), (2, 4, "text"), (3, 5, "text"),
    (4, 6, "date"), (5, 7, "date"), (6, 8, "num"), (7, 9, "num"),
    (8, 10, "num"), (9, 12, "num"), (10, 13, "text"),
]
AD_VIDEO_SPEC = [
    (0, 0, "text"), (1, 1, "text"), (2, 2, "date"), (3, 3, "text"), (4, 4, "text"),
    (5, 5, "text"), (6, 6, "date"), (7, 7, "date"), (8, 8, "num"), (9, 9, "num"),
    (10, 10, "pct"), (11, 11, "num"), (12, 12, "num"), (13, 13, "num"), (14, 14, "text"),
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
        return
    col = header.index(src.target_month)
    cum_col, target_col, rate_col = header.index("누적"), header.index("목표"), header.index("달성률")
    instagram_totals = src.summary.get("instagram_totals", {})
    for r, metric_row in _KPI_METRIC_ROW.items():
        cur_value = instagram_totals.get(_KPI_CUR_KEY[metric_row])
        d = src.kpi_row(metric_row, cur_value)
        set_cell_text(table.rows[r].cells[col], _fmt_num(d["cur"]))
        set_cell_text(table.rows[r].cells[cum_col], _fmt_num(d["cum"]))
        set_cell_text(table.rows[r].cells[target_col], _fmt_num(d["target"]))
        set_cell_text(table.rows[r].cells[rate_col], _fmt_pct(d["rate"]))


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

    set_cell_text(table.rows[0].cells[1], prev_label)
    set_cell_text(table.rows[0].cells[2], src.target_month)

    for r in range(1, len(table.rows)):
        label = table.rows[r].cells[0].text.strip()
        # 기존 '당월' 열 값을 새 '전월' 열로 이동
        prev_display = table.rows[r].cells[2].text.strip()
        set_cell_text(table.rows[r].cells[1], prev_display)

        metric = _CHANNEL_ROW_SOURCE.get(label)
        if metric is None or cur_values.get(metric) is None:
            set_cell_text(table.rows[r].cells[2], CONFIRM_NEEDED)
            confirm.add("PPT/채널요약", f"{sheet} - {label}", "자동 산출 소스가 없어 확인 필요")
        else:
            prev_v, cur_v = summary[metric]
            set_cell_text(table.rows[r].cells[2], _fmt_num(cur_v))

        prev_num = _parse_num(prev_display)
        cur_num = _parse_num(table.rows[r].cells[2].text)
        _change_text_and_color(prev_num, cur_num, table.rows[r].cells[3])


def fill_ad_overview_table(table, src: XlsxSource, confirm):
    ws = src.wb["운영요약"]
    set_cell_text(table.rows[0].cells[1], ws.cell(row=36, column=3).value)
    set_cell_text(table.rows[0].cells[2], ws.cell(row=36, column=4).value)

    for r in range(1, len(table.rows)):
        row_num = 36 + r
        prev_v = ws.cell(row=row_num, column=3).value
        cur_v = ws.cell(row=row_num, column=4).value
        chg_v = ws.cell(row=row_num, column=5).value
        set_cell_text(table.rows[r].cells[1], _fmt_num(prev_v) if isinstance(prev_v, (int, float)) else (prev_v or ""))
        set_cell_text(table.rows[r].cells[2], _fmt_num(cur_v) if isinstance(cur_v, (int, float)) else (cur_v or ""))
        if isinstance(chg_v, (int, float)):
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
# 최상위: PPT 조합
# ---------------------------------------------------------------------------

def build_ppt(prev_ppt_file, wb, target_month, confirm, block_cache=None, wb_data=None):
    prs = Presentation(prev_ppt_file)
    src = XlsxSource(wb, target_month, block_cache, wb_data=wb_data)

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
                    fill_kpi_table(table, src, confirm)
                elif kind == "channel_compare":
                    row1_label = _row_texts(table, 1)[0] if len(table.rows) > 1 else ""
                    sheet = "블로그" if row1_label in ("이웃수",) else "인스타그램"
                    fill_channel_compare_table(table, sheet, src, confirm)
                elif kind == "ad_overview":
                    fill_ad_overview_table(table, src, confirm)
                elif kind == "budget_summary":
                    fill_budget_summary_table(table, src)
                elif kind == "content_instagram":
                    rows, avgs, totals = src.block_rows("인스타그램")
                    fill_data_table(table, INSTAGRAM_CONTENT_SPEC, rows, avgs, totals,
                                     url_idx=CONTENT_URL_IDX, url_ppt_col=2)
                elif kind == "content_blog":
                    rows, avgs, totals = src.block_rows("블로그")
                    fill_data_table(table, BLOG_CONTENT_SPEC, rows, avgs, totals,
                                     url_idx=CONTENT_URL_IDX, url_ppt_col=2)
                elif kind == "ad_participation":
                    rows, avgs, totals = src.block_rows("AD 데이터 (참여)")
                    fill_data_table(table, AD_PARTICIPATION_SPEC, rows, avgs, totals)
                elif kind == "ad_reach":
                    rows, avgs, totals = src.block_rows("AD 데이터 (도달)")
                    fill_data_table(table, AD_REACH_SPEC, rows, avgs, totals)
                elif kind == "ad_video":
                    rows, avgs, totals = src.block_rows("AD 데이터 (동영상조회)")
                    fill_data_table(table, AD_VIDEO_SPEC, rows, avgs, totals)
            except Exception as e:
                confirm.add("PPT", kind, f"자동 채움 실패({e}) - 수기 확인 필요")

    confirm.add("PPT", "자동화 범위 안내",
                "운영 캘린더·팔로워 추이/성별연령·우수 콘텐츠 TOP3·검색 상위 노출·이벤트/제휴 소개 슬라이드는 "
                "스크린샷·서술형 데이터라 자동 채움 대상이 아닙니다(전월 내용이 그대로 남아있으니 직접 교체 필요).")
    return prs
