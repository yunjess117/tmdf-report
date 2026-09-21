# -*- coding: utf-8 -*-
"""로우데이터(취합) 조합 로직 — 전월 최종본에 이번 달 블록을 이어붙인다.

핵심 원칙: 전월 최종본에 있는 기존 수식(합계/평균/누적/증감률 등)은 절대 값으로
덮어쓰지 않는다. 이번 달에 새로 채우는 집계 셀도 전월과 같은 수식 패턴(범위만
이번 달 데이터 위치로 조정)으로 써서, 파일을 열어보면 전월 행과 새 행이
서식·수식 패턴 모두 동일하게 보이도록 한다. 실제 값이 필요한 곳(PPT 등)은
셀을 다시 읽지 않고, 여기서 계산에 썼던 파이썬 값을 그대로 돌려준다(수식은
엑셀이 열어야 계산되고 openpyxl은 그 값을 읽을 수 없기 때문).
"""
import datetime as dt

from core.excel_block import (
    CONFIRM_NEEDED, find_last_label_row, insert_block, write_row,
    range_formula, copy_month_row_formula, col_letter, copy_row_style,
)
from core.formatters import normalize_url_key, roundup_10000
from core.matching import content_ad_values, match_ad_for_content, match_content_for_ad

NO_COL = 2       # B열: NO.
TOPIC_COL = 5    # E열: 콘텐츠 주제 (평균/합계 라벨이 찍히는 열, 기본값)


def _detect_topic_col(ws, header_row, default=TOPIC_COL, search_rows=40, col_range=(2, 12)):
    """'평균'/'합계' 라벨이 실제로 찍히는 열을 찾는다.

    인스타그램/AD 시트는 '콘텐츠 주제' 열(E)에 라벨이 찍히지만 블로그 시트는
    'URL' 열(F)에 찍히는 등 시트마다 다르므로 고정 상수 대신 탐지해서 쓴다.
    아직 데이터 블록이 없는(최초 실행) 시트라면 찾지 못하므로 default를 쓴다.
    """
    for r in range(header_row + 1, header_row + 1 + search_rows):
        for c in range(*col_range):
            if ws.cell(row=r, column=c).value in ("평균", "합계"):
                return c
    return default


def _find_last_data_row(ws, before_row, no_col=NO_COL):
    """before_row 위쪽에서 NO. 열에 정수가 들어있는 마지막(가장 아래) 데이터 행을 찾는다."""
    for r in range(before_row - 1, 0, -1):
        v = ws.cell(row=r, column=no_col).value
        if isinstance(v, (int, float)):
            return r
    return None


def _avg(values):
    vals = [v for v in values if isinstance(v, (int, float))]
    return sum(vals) / len(vals) if vals else None


def _sum(values):
    vals = [v for v in values if isinstance(v, (int, float))]
    return sum(vals) if vals else None


def _month_num(label: str) -> int:
    return int(str(label).rstrip("월"))


def _numval(ws_live, ws_data, row, col):
    """전월 최종본의 셀이 수식(예: '=SUM(...)')이면 openpyxl로는 계산된 숫자를
    못 읽으므로, 같은 파일을 data_only=True로 한 번 더 읽어둔 스냅샷(ws_data)에서
    캐시된 계산값을 대신 가져온다. 이번 실행에서 방금 써넣은 값은 이미 숫자
    리터럴이라 ws_live에서 바로 읽힌다."""
    v = ws_live.cell(row=row, column=col).value
    if isinstance(v, (int, float)):
        return v
    if ws_data is not None:
        v2 = ws_data.cell(row=row, column=col).value
        if isinstance(v2, (int, float)):
            return v2
    return None


def _restore_agg_row_style(ws, insert_at, avg_row, total_row, topic_col, min_col, max_col):
    """평균/합계 행은 데이터 행과 다른 서식(강조 배경색, 날짜 열은 숫자 형식으로
    되돌림 등)을 쓰는 경우가 많다. insert_block이 일괄로 데이터 행 서식을 복사해
    버리므로, 바로 위(전월)의 평균/합계 행 서식을 다시 덮어써 되돌린다. 전월 블록이
    없는 최초 실행이면(비교 대상이 없으면) 아무것도 하지 않는다."""
    old_avg_row, old_total_row = insert_at - 2, insert_at - 1
    if ws.cell(row=old_total_row, column=topic_col).value == "합계":
        copy_row_style(ws, old_avg_row, avg_row, min_col, max_col)
        copy_row_style(ws, old_total_row, total_row, min_col, max_col)


def _write_agg_rows(ws, insert_at, n_data, topic_col, numeric_cols, func_by_col=None):
    """평균/합계 행을 전월과 같은 수식 패턴(AVERAGE/SUM, 범위는 이번 달 데이터
    행 구간)으로 채운다. func_by_col: {col: '결과당비용'} 처럼 특수 처리가
    필요한 열은 여기서 뒤에 별도로 덮어쓴다(호출부에서 처리)."""
    data_start, data_end = insert_at, insert_at + n_data - 1
    avg_row, total_row = insert_at + n_data, insert_at + n_data + 1
    max_c = max(numeric_cols) if numeric_cols else 17
    _restore_agg_row_style(ws, insert_at, avg_row, total_row, topic_col, min_col=2, max_col=max_c)
    ws.cell(row=avg_row, column=topic_col, value="평균")
    ws.cell(row=total_row, column=topic_col, value="합계")
    for c in numeric_cols:
        ws.cell(row=avg_row, column=c, value=range_formula("AVERAGE", c, data_start, data_end))
        ws.cell(row=total_row, column=c, value=range_formula("SUM", c, data_start, data_end))
    return avg_row, total_row


# ---------------------------------------------------------------------------
# 인스타그램 콘텐츠 시트
# ---------------------------------------------------------------------------

def append_instagram_block(wb, target_month, pub_rows, content_perf, ad_rows, confirm):
    ws = wb["인스타그램"]
    header_row = None
    for r in range(1, 20):
        if ws.cell(row=r, column=NO_COL).value == "NO.":
            header_row = r
            break
    topic_col = _detect_topic_col(ws, header_row)
    last_total = find_last_label_row(ws, "합계", topic_col, search_from=header_row)
    insert_at = (last_total + 1) if last_total else (header_row + 1)
    style_row = _find_last_data_row(ws, insert_at) or header_row

    rows = [r for r in pub_rows if r.month == target_month]
    rows.sort(key=lambda r: (r.pub_date or dt.date.min, r.no or 0))
    n_data = len(rows)
    insert_block(ws, insert_at, n_data + 2, style_row, min_col=2, max_col=17)

    py_rows = []  # PPT 등에서 쓸 파이썬 계산값(수식이 아니라 실제 숫자)
    for i, pr in enumerate(rows):
        key = normalize_url_key(pr.url)
        perf = content_perf.get(key)
        if perf is None:
            confirm.add("인스타그램", f"{target_month} NO.{i+1} ({pr.title})",
                        "성과 원본에서 고유 링크 매칭 실패 - 조회/도달 등 확인 필요")
            perf = {k: CONFIRM_NEEDED for k in
                    ("조회", "도달", "팔로우", "좋아요", "댓글", "공유", "저장", "총반응")}
        ad_spend, ad_reach, _ = content_ad_values(pr.pub_date, pr.title, ad_rows)
        if ad_spend == CONFIRM_NEEDED:
            confirm.add("인스타그램", f"{target_month} NO.{i+1} ({pr.title})",
                        "발행일과 일치하는 광고 캠페인을 찾지 못함 - 광고비/광고도달 확인 필요")
        r = insert_at + i
        # 총 반응 = 전월 행과 동일한 수식 패턴(=SUM(좋아요:저장)) 유지
        total_reaction_formula = f"=SUM(M{r}:P{r})" if perf["총반응"] != CONFIRM_NEEDED else CONFIRM_NEEDED
        values = [i + 1, target_month, pr.pub_date, pr.title, pr.url, pr.content_type,
                  ad_spend, perf["조회"], perf["도달"], ad_reach, perf["팔로우"],
                  perf["좋아요"], perf["댓글"], perf["공유"], perf["저장"], total_reaction_formula]
        write_row(ws, r, 2, values)
        py_rows.append([i + 1, target_month, pr.pub_date, pr.title, pr.url, pr.content_type,
                         ad_spend, perf["조회"], perf["도달"], ad_reach, perf["팔로우"],
                         perf["좋아요"], perf["댓글"], perf["공유"], perf["저장"], perf["총반응"]])

    # range(8, 18) = H(광고비)~Q(총반응). G(유형)는 텍스트 열이라 합계/평균 대상에서 뺀다
    # (전월 최종본에도 유형 열은 평균/합계 행이 비어 있음).
    avg_row, total_row = _write_agg_rows(ws, insert_at, n_data, topic_col, range(8, 18))

    def col(idx):
        return [row[idx] for row in py_rows]

    py_avg = [None, None, None, "평균", None, None] + [_avg(col(i)) for i in range(6, 16)]
    py_total = [None, None, None, "합계", None, None] + [_sum(col(i)) for i in range(6, 16)]

    # 영상 조회수 = '조회' 컬럼 전체 합이 아니라, 유형이 '영상'인 콘텐츠의 조회만 합산
    # (전월 최종본의 SUMIFS(...,"영상") 수식과 동일한 정의).
    video_views = _sum([row[7] for row in py_rows if row[5] == "영상"])
    totals = {
        "발행수": n_data,
        "도달수": _sum(col(8)),
        "영상조회수": video_views,
        "참여수": _sum(col(15)),
    }
    # 상단 월별 요약표(6~11행)의 target 달 행은 전월(6행)과 같은 COUNTIF/SUMIF
    # 수식 패턴을 그대로 복제한다(고정 범위 $C$14:$C$100 등은 자동으로 유지됨).
    summary_row = 6 + (_month_num(target_month) - 7)
    copy_month_row_formula(ws, template_row=6, target_row=summary_row, min_col=4, max_col=7)

    return totals, {"rows": py_rows, "avg": py_avg, "total": py_total}


# ---------------------------------------------------------------------------
# 블로그 시트 (성과 원본 없음 -> 확인 필요로 채움)
# ---------------------------------------------------------------------------

def append_blog_block(wb, target_month, pub_rows, confirm):
    ws = wb["블로그"]
    header_row = None
    for r in range(1, 20):
        if ws.cell(row=r, column=NO_COL).value == "NO.":
            header_row = r
            break
    topic_col = _detect_topic_col(ws, header_row)
    last_total = find_last_label_row(ws, "합계", topic_col, search_from=header_row)
    insert_at = (last_total + 1) if last_total else (header_row + 1)
    style_row = _find_last_data_row(ws, insert_at) or header_row

    rows = [r for r in pub_rows if r.month == target_month]
    rows.sort(key=lambda r: (r.pub_date or dt.date.min, r.no or 0))
    n_data = len(rows)
    insert_block(ws, insert_at, n_data + 2, style_row, min_col=2, max_col=11)

    py_rows = []
    for i, pr in enumerate(rows):
        confirm.add("블로그", f"{target_month} NO.{i+1} ({pr.title})",
                    "블로그 성과 원본 파일이 입력되지 않아 조회/공감/댓글 확인 필요")
        values = [i + 1, target_month, pr.pub_date, pr.title, pr.url, pr.content_type,
                  CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED]
        write_row(ws, insert_at + i, 2, values)
        py_rows.append(list(values))

    avg_row = insert_at + n_data
    total_row = avg_row + 1
    _restore_agg_row_style(ws, insert_at, avg_row, total_row, topic_col, min_col=2, max_col=11)
    ws.cell(row=avg_row, column=topic_col, value="평균")
    ws.cell(row=total_row, column=topic_col, value="합계")
    for c in range(8, 12):
        ws.cell(row=avg_row, column=c, value=CONFIRM_NEEDED)
        ws.cell(row=total_row, column=c, value=CONFIRM_NEEDED)
    py_avg = [None, None, None, "평균", None, None, CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED]
    py_total = [None, None, None, "합계", None, None, CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED]

    # 상단 월별 요약표: 발행 수는 실제 COUNTIF 수식으로, 나머지(조회/총반응/참여)는
    # 성과 원본이 없어 전월 수식 패턴을 복제해도 값이 안 나오므로 확인 필요로 둔다.
    summary_row = 6 + (_month_num(target_month) - 7)
    copy_month_row_formula(ws, template_row=6, target_row=summary_row, min_col=4, max_col=4)
    for c in (5, 6, 7):
        ws.cell(row=summary_row, column=c, value=CONFIRM_NEEDED)

    return {"발행수": n_data}, {"rows": py_rows, "avg": py_avg, "total": py_total}


# ---------------------------------------------------------------------------
# AD 데이터 (참여/도달/동영상조회)
# ---------------------------------------------------------------------------

_AD_SHEETS = {
    "참여": {"sheet": "AD 데이터 (참여)", "has_vtr": False},
    "도달": {"sheet": "AD 데이터 (도달)", "has_vtr": False},
    "동영상조회": {"sheet": "AD 데이터 (동영상조회)", "has_vtr": True},
}


def append_ad_type_block(wb, ad_type, target_month, ad_rows, content_rows, confirm):
    info = _AD_SHEETS[ad_type]
    ws = wb[info["sheet"]]
    header_row = None
    for r in range(1, 20):
        if ws.cell(row=r, column=NO_COL).value == "NO.":
            header_row = r
            break
    topic_col = _detect_topic_col(ws, header_row)
    last_total = find_last_label_row(ws, "합계", topic_col, search_from=header_row)
    insert_at = (last_total + 1) if last_total else (header_row + 1)
    max_col = 16 if info["has_vtr"] else 15
    style_row = _find_last_data_row(ws, insert_at) or header_row

    typed = [a for a in ad_rows if a.ad_type == ad_type]
    matched = [(match_content_for_ad(a, content_rows), a) for a in typed]
    matched.sort(key=lambda pair: (pair[0].pub_date if pair[0] else (pair[1].anchor_date or dt.date.min)))
    n_data = len(matched)
    insert_block(ws, insert_at, n_data + 2, style_row, min_col=2, max_col=max_col)

    # 열 매핑 (절대 컬럼 번호). 헤더: NO,월,발행일,콘텐츠주제,유형,타깃,시작,종료,결과,결과당비용,[VTR,]노출,도달,지출,예산
    if info["has_vtr"]:
        cols = {"결과": 10, "결과당비용": 11, "VTR": 12, "노출": 13, "도달": 14, "지출": 15, "예산": 16}
    else:
        cols = {"결과": 10, "결과당비용": 11, "노출": 12, "도달": 13, "지출": 14, "예산": 15}

    py_rows = []
    for i, (content, a) in enumerate(matched):
        if content is None:
            confirm.add(info["sheet"], f"{target_month} NO.{i+1} ({a.ad_name})",
                        "발행일/제목이 일치하는 콘텐츠를 발행리스트에서 찾지 못함 - 발행일/콘텐츠 주제/유형 확인 필요")
            pub_date, title, ctype = CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED
        else:
            pub_date, title, ctype = content.pub_date, content.title, content.content_type
        camp_start = a.campaign_start if a.campaign_start else CONFIRM_NEEDED
        camp_end = a.campaign_end if a.campaign_end else CONFIRM_NEEDED
        # AD 데이터 시트의 '지출 금액'은 광고 원본 그대로(raw)를 쓴다. '광고예산'은 원본에
        # 없으므로 지출 금액을 만원 단위로 올림한 값을 그대로 쓴다(예: 49,415 -> 50,000).
        spend = a.spend if a.spend is not None else CONFIRM_NEEDED
        budget = roundup_10000(a.spend) if a.spend is not None else CONFIRM_NEEDED
        if info["has_vtr"]:
            vtr = (a.result / a.impressions) if a.result is not None and a.impressions else None
            values = [i + 1, target_month, pub_date, title, ctype, CONFIRM_NEEDED,
                      camp_start, camp_end, a.result, a.cost_per_result, vtr,
                      a.impressions, a.reach, spend, budget]
        else:
            values = [i + 1, target_month, pub_date, title, ctype, CONFIRM_NEEDED,
                      camp_start, camp_end, a.result, a.cost_per_result,
                      a.impressions, a.reach, spend, budget]
        write_row(ws, insert_at + i, 2, values)
        py_rows.append(list(values))
        confirm.add(info["sheet"], f"{target_month} NO.{i+1}", "타깃은 원본에 없어 확인 필요")

    data_start, data_end = insert_at, insert_at + n_data - 1
    avg_row, total_row = insert_at + n_data, insert_at + n_data + 1
    _restore_agg_row_style(ws, insert_at, avg_row, total_row, topic_col, min_col=2, max_col=max_col)
    ws.cell(row=avg_row, column=topic_col, value="평균")
    ws.cell(row=total_row, column=topic_col, value="합계")

    scale = 1000 if ad_type == "도달" else 1  # 전월 최종본의 기존 수식 관례(도달만 *1000)를 그대로 따름
    for name in ("결과", "노출", "도달", "지출", "예산"):
        c = cols[name]
        if n_data:
            ws.cell(row=avg_row, column=c, value=range_formula("AVERAGE", c, data_start, data_end))
            ws.cell(row=total_row, column=c, value=range_formula("SUM", c, data_start, data_end))
    # '결과당 비용' 평균 = 합계 행의 지출÷결과(가중평균). 전월 수식(N{합계}/J{합계}) 그대로 복제.
    cpr_col, spend_col, result_col = cols["결과당비용"], cols["지출"], cols["결과"]
    if n_data:
        spend_L, result_L = col_letter(spend_col), col_letter(result_col)
        formula = f"={spend_L}{total_row}/{result_L}{total_row}"
        if scale != 1:
            formula += f"*{scale}"
        ws.cell(row=avg_row, column=cpr_col, value=formula)
    ws.cell(row=total_row, column=cpr_col, value="-")
    if info["has_vtr"]:
        vtr_col, imp_col = cols["VTR"], cols["노출"]
        if n_data:
            # VTR 평균 = 평균행 자신의 결과÷노출(전월 수식 J{평균}/M{평균} 패턴)
            result_L, imp_L = col_letter(result_col), col_letter(imp_col)
            ws.cell(row=avg_row, column=vtr_col, value=f"={result_L}{avg_row}/{imp_L}{avg_row}")
        ws.cell(row=total_row, column=vtr_col, value="-")

    # 파이썬 계산값(평균/합계 실제 숫자) - PPT 등에서 사용
    def numcol(idx0):
        return [row[idx0] for row in py_rows]

    total_result = _sum(numcol(cols["결과"] - 2))
    total_spend = _sum(numcol(cols["지출"] - 2))
    avg_cpr = (total_spend / total_result * scale) if total_spend and total_result else None
    py_avg = [None] * (max_col - 1)
    py_total = [None] * (max_col - 1)
    py_avg[topic_col - 2] = "평균"
    py_total[topic_col - 2] = "합계"
    for name in ("결과", "노출", "도달", "지출", "예산"):
        idx0 = cols[name] - 2
        py_avg[idx0] = _avg(numcol(idx0))
        py_total[idx0] = _sum(numcol(idx0))
    py_avg[cpr_col - 2] = avg_cpr
    py_total[cpr_col - 2] = "-"
    if info["has_vtr"]:
        idx_vtr, idx_imp, idx_res = cols["VTR"] - 2, cols["노출"] - 2, cols["결과"] - 2
        py_avg[idx_vtr] = (py_avg[idx_res] / py_avg[idx_imp]) if py_avg[idx_imp] else None
        py_total[idx_vtr] = "-"

    # 상단 월별 요약표(5~10행)의 target 달 행: 전월(5행)과 같은 COUNTIF/SUMIF 패턴 복제.
    # 광고예산(H열)도 이제 각 행에 실제 값(지출 만원 올림)이 들어가므로 SUMIF가 그대로 합산한다.
    summary_row = 5 + (_month_num(target_month) - 7)
    copy_month_row_formula(ws, template_row=5, target_row=summary_row, min_col=4, max_col=8)

    return ({"진행수량": n_data, "총결과": total_result, "지출금액": total_spend},
            {"rows": py_rows, "avg": py_avg, "total": py_total})


# ---------------------------------------------------------------------------
# AD 데이터(전체) — 참여/도달/동영상조회 원본 광고 행을 한 표로 통합
# ---------------------------------------------------------------------------

def append_ad_all_block(wb, target_month, ad_rows, content_rows, confirm):
    ws = wb["AD 데이터(전체)"]
    header_row = None
    for r in range(1, 20):
        if ws.cell(row=r, column=NO_COL).value == "NO.":
            header_row = r
            break
    # AD 데이터(전체) 시트는 평균/합계 행이 없는 구조라 마지막 데이터 행(NO. 있는 행)
    # 바로 다음이 삽입 지점이다.
    last_data = _find_last_data_row(ws, ws.max_row + 1)
    insert_at = (last_data + 1) if last_data else (header_row + 1)
    style_row = last_data or header_row

    typed = [a for a in ad_rows if a.ad_type]
    matched = [(match_content_for_ad(a, content_rows), a) for a in typed]
    matched.sort(key=lambda pair: (pair[0].pub_date if pair[0] else (pair[1].anchor_date or dt.date.min)))
    n_data = len(matched)
    insert_block(ws, insert_at, n_data + 1, style_row, min_col=2, max_col=22)  # 전체 시트는 합계 행 없음(관찰된 구조)

    for i, (content, a) in enumerate(matched):
        if content is None:
            pub_date, title, ctype = CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED
        else:
            pub_date, title, ctype = content.pub_date, content.title, content.content_type
        camp_start = a.campaign_start if a.campaign_start else CONFIRM_NEEDED
        camp_end = a.campaign_end if a.campaign_end else CONFIRM_NEEDED
        is_video = a.ad_type == "동영상조회"
        vtr = (a.result / a.impressions) if is_video and a.result is not None and a.impressions else "-"
        v2s = a.result if is_video else "-"
        v2s_cost = a.cost_per_result if is_video else "-"
        spend = a.spend if a.spend is not None else CONFIRM_NEEDED  # raw 지출(콘텐츠 시트만 만원 올림 적용)
        budget = roundup_10000(a.spend) if a.spend is not None else CONFIRM_NEEDED  # 예산 = 지출 만원 올림
        values = [i + 1, target_month, pub_date, title, ctype, camp_start, camp_end,
                  a.result_type, a.result, a.cost_per_result, a.impressions, a.reach,
                  v2s, v2s_cost, vtr, a.link_clicks, a.cpc, a.ctr, a.cpm, spend,
                  budget]
        write_row(ws, insert_at + i, 2, values)

    # 월별 광고비 현황 행(5~10행)의 target 달 행: 전월(5행)의 SUMIF/SUMIFS 패턴을 그대로
    # 복제한다. 헤더: D=월, E=광고비(총합, SUMIF), F/G/H=참여/도달/동영상조회(SUMIFS).
    # E열은 월 텍스트가 수식 안에 하드코딩돼 있어("7월") 텍스트를 바꿔주고,
    # F/G/H열은 D열(같은 행)을 상대참조하므로 Translator로 행 번호만 옮기면 된다.
    r = 5 + (_month_num(target_month) - 7)
    prev_r = r - 1
    prev_label = ws.cell(row=prev_r, column=4).value
    prev_e = ws.cell(row=prev_r, column=5).value
    ws.cell(row=r, column=4, value=target_month)
    if isinstance(prev_e, str) and prev_e.startswith("=") and prev_label:
        ws.cell(row=r, column=5, value=prev_e.replace(f'"{prev_label}"', f'"{target_month}"'))
    copy_month_row_formula(ws, template_row=prev_r, target_row=r, min_col=6, max_col=8)
    # 이 수식들은 '예산'(V)열을 집계한다. 각 행의 예산을 지출 만원 올림 값으로 채워
    # 넣으므로(위 budget), 사람이 따로 채우지 않아도 자동으로 집계된다.

    # 파이썬 계산값(참여/도달/동영상조회 raw 지출 합) - PPT/운영요약에서 사용.
    # 누적/잔여비/총예산 행(11~13)은 이미 고정 수식(=SUM(E5:E10) 등)이라 손대지 않는다.
    totals_by_type = {}
    for t in ("참여", "도달", "동영상조회"):
        vals = [a.spend for _, a in matched if a.ad_type == t and a.spend is not None]
        totals_by_type[t] = _sum(vals) or 0
    return totals_by_type


# ---------------------------------------------------------------------------
# 제휴(인플루언서 체험단) 오픈 보고서 -> 신규 '제휴' 시트
# ---------------------------------------------------------------------------

_PARTNER_INSTA_HEADER = ["NO.", "오픈일", "인스타 아이디", "내용 요약", "팔로워 수", "조회 수", "좋아요 수", "댓글 수", "비고", "URL"]
_PARTNER_BLOG_HEADER = ["NO.", "오픈일", "블로거", "제목", "일평균 방문자수", "오픈일 방문자수", "PV수치", "댓글 수", "공감 수", "비고", "URL"]


def _ensure_partnership_sheet(wb):
    if "제휴" in wb.sheetnames:
        return wb["제휴"]
    import copy
    ws = wb.create_sheet("제휴")
    ws.cell(row=2, column=2, value="제휴(인플루언서 체험단) 데이터")
    ws.cell(row=2, column=2).font = copy.copy(wb["인스타그램"]["B2"].font)
    ws.cell(row=4, column=2, value="인스타그램")
    for j, h in enumerate(_PARTNER_INSTA_HEADER):
        ws.cell(row=5, column=2 + j, value=h)
    return ws


def append_partnership_block(wb, target_month, insta_rows, blog_rows):
    # 이 시트는 전월 최종본에 존재하지 않던(이번 자동화가 처음 만드는) 시트라
    # '기존 수식을 보존'할 대상이 없다 - 취합 시점의 값을 그대로 적는다.
    ws = _ensure_partnership_sheet(wb)

    # '제휴' 시트가 이미 인스타+블로그 섹션을 갖고 있는 상태(= PPT 패널에서 만든
    # 결과물을 '최종 로우데이터'로 다시 올려 재생성하는 경우)에서, 인스타 합계를
    # 찾는 범위를 시트 전체로 잡으면 블로그 섹션의 마지막 합계까지 걸려서 이번
    # 달 인스타 데이터가 블로그 섹션 '뒤'에 붙어버린다(인스타-블로그-인스타-블로그
    # 순으로 어긋남). 그러면 이후 읽을 때 '마지막 인스타 블록'을 블로그 섹션
    # 앞의 예전 블록으로 잘못 집어 PPT에 지난달 데이터가 그대로 남는 문제로
    # 이어진다 - 블로그 섹션 시작 전까지로 검색 범위를 제한한다.
    blog_marker_row = None
    for r in range(6, ws.max_row + 3):
        if ws.cell(row=r, column=2).value == "블로그":
            blog_marker_row = r
            break

    insta_search_to = (blog_marker_row - 1) if blog_marker_row else 200
    last_total = find_last_label_row(ws, "합계", 3, search_from=5, search_to=insta_search_to) or 5
    insert_at = last_total + 1 if last_total != 5 else 6

    n_new_insta_rows = len(insta_rows) + 1  # 데이터 행 + 합계 행
    if blog_marker_row is not None and insert_at <= blog_marker_row:
        # 블로그 섹션이 이미 인스타 섹션 바로 뒤에 있으면, 새 인스타 행이 들어갈
        # 자리를 진짜로 밀어서 만든다(그냥 다음 빈 줄에 쓰면 인스타/블로그
        # 섹션이 뒤섞인다).
        ws.insert_rows(insert_at, n_new_insta_rows)
        blog_marker_row += n_new_insta_rows

    for i, p in enumerate(insta_rows):
        r = insert_at + i
        write_row(ws, r, 2, [i + 1, p.open_date, p.handle, p.summary, p.followers, p.views,
                              p.likes, p.comments, p.note, p.url])
    total_row = insert_at + len(insta_rows)
    ws.cell(row=total_row, column=3, value="합계")
    for c, field in ((6, "followers"), (7, "views"), (8, "likes"), (9, "comments")):
        vals = [getattr(p, field) for p in insta_rows if isinstance(getattr(p, field), (int, float))]
        ws.cell(row=total_row, column=c, value=_sum(vals))

    blog_header_row = (blog_marker_row + 1) if blog_marker_row is not None else None
    if blog_header_row is None:
        ws.cell(row=total_row + 2, column=2, value="블로그")
        blog_header_row = total_row + 3
        for j, h in enumerate(_PARTNER_BLOG_HEADER):
            ws.cell(row=blog_header_row, column=2 + j, value=h)
        blog_insert_at = blog_header_row + 1
    else:
        last_blog_total = find_last_label_row(ws, "합계", 3, search_from=blog_header_row, search_to=ws.max_row) or blog_header_row
        blog_insert_at = last_blog_total + 1 if last_blog_total != blog_header_row else blog_header_row + 1

    for i, p in enumerate(blog_rows):
        r = blog_insert_at + i
        write_row(ws, r, 2, [i + 1, p.open_date, p.blogger, p.title, p.daily_visitors,
                              p.open_day_visitors, p.pv, p.comments, p.likes, p.note, p.url])
    blog_total_row = blog_insert_at + len(blog_rows)
    ws.cell(row=blog_total_row, column=3, value="합계")
    for c, field in ((5, "daily_visitors"), (6, "open_day_visitors"), (7, "pv"),
                     (8, "comments"), (9, "likes")):
        vals = [getattr(p, field) for p in blog_rows if isinstance(getattr(p, field), (int, float))]
        ws.cell(row=blog_total_row, column=c, value=_sum(vals))


# ---------------------------------------------------------------------------
# 인스타그램 팔로워 타깃(성별/연령) — 전월 최종본에 없던 신규 시트.
# PPT의 '팔로워 성별/연령 비중' 차트가 이 시트 값을 읽어간다. 여기 값만 고치면
# (수식이 아니라 일반 셀이라 수기 수정 가능) 다음에 PPT를 다시 만들 때 반영된다.
# ---------------------------------------------------------------------------

_TARGET_SHEET = "인스타그램 팔로워 타깃"


def write_follower_target(wb, target_month, target: "FollowerTarget"):
    if target is None or not target.age_labels:
        return
    if _TARGET_SHEET in wb.sheetnames:
        ws = wb[_TARGET_SHEET]
    else:
        import copy
        ws = wb.create_sheet(_TARGET_SHEET)
        ws.cell(row=2, column=2, value="인스타그램 팔로워 타깃(성별/연령)")
        ws.cell(row=2, column=2).font = copy.copy(wb["인스타그램"]["B2"].font)

    ws.cell(row=4, column=2, value="기준월")
    ws.cell(row=4, column=3, value=target_month)
    ws.cell(row=4, column=4, value="* 인스타그램 타깃(인사이트) CSV 기준. 필요하면 이 시트 값을 직접 고쳐도 됩니다.")

    ws.cell(row=6, column=2, value="성별 비중(%)")
    ws.cell(row=7, column=2, value="구분")
    ws.cell(row=7, column=3, value="남성")
    ws.cell(row=7, column=4, value="여성")
    ws.cell(row=8, column=2, value="비율")
    ws.cell(row=8, column=3, value=round(target.male_total, 1))
    ws.cell(row=8, column=4, value=round(target.female_total, 1))

    ws.cell(row=10, column=2, value="연령대별 성별 비중(%)")
    ws.cell(row=11, column=2, value="연령대")
    ws.cell(row=11, column=3, value="남성")
    ws.cell(row=11, column=4, value="여성")
    for i, age in enumerate(target.age_labels):
        r = 12 + i
        ws.cell(row=r, column=2, value=age)
        ws.cell(row=r, column=3, value=target.male_by_age[i])
        ws.cell(row=r, column=4, value=target.female_by_age[i])


# ---------------------------------------------------------------------------
# 운영요약 시트 — 이번 달 열만 채운다(다른 열/누적/달성률은 이미 살아있는 수식이라 건드리지 않음)
# ---------------------------------------------------------------------------

def update_operations_summary(wb, target_month, ad_totals_by_type):
    ws = wb["운영요약"]
    month_col = 3 + (_month_num(target_month) - 7)  # C=7월 ... H=12월
    month_L = col_letter(month_col)

    # 콘텐츠 KPI (행14~17): target 달 열만 '=인스타그램!{열}{그달행}' 수식으로 채운다.
    # 인스타그램 시트에서 발행수/도달수/영상조회수/참여수가 있는 열은 각각 D/E/F/G로 고정.
    kpi_src_col = {14: "D", 15: "E", 16: "F", 17: "G"}
    insta_row = 6 + (_month_num(target_month) - 7)
    for row, src_col in kpi_src_col.items():
        ws.cell(row=row, column=month_col, value=f"=인스타그램!{src_col}{insta_row}")

    # 연간 광고비 지출 현황 (행25~27): target 달 열만 "='AD 데이터(전체)'!{열}{그달행}" 수식.
    ad_all_src_col = {25: "F", 26: "G", 27: "H"}
    ad_all_row = 5 + (_month_num(target_month) - 7)
    for row, src_col in ad_all_src_col.items():
        ws.cell(row=row, column=month_col, value=f"='AD 데이터(전체)'!{src_col}{ad_all_row}")
    # 합계(행28) target 달 열 = 같은 열의 참여+도달+동영상조회 합(전월과 같은 수식 패턴)
    ws.cell(row=28, column=month_col, value=f"=SUM({month_L}25:{month_L}27)")

    # 예산 요약(행32)은 라벨(월)만 바꾼다. C/D/E열은 이미 '=I28', '=E32-C32', 고정 총예산이라
    # 그대로 두면 자동으로 최신 값을 반영한다.
    ws.cell(row=32, column=2, value=target_month)

    # 월간 광고 요약(행36 헤더 + 37~49)은 전월 최종본에서도 라이브 수식이 아니라 사람이
    # 채워 넣는 스냅샷 표라(증감만 수식) 여기서도 같은 방식으로 값을 옮기고 새 값을 채운다.
    prev_label = ws.cell(row=36, column=4).value
    prev_values = {r: ws.cell(row=r, column=4).value for r in range(37, 50)}
    ws.cell(row=36, column=3, value=prev_label)
    ws.cell(row=36, column=4, value=target_month)
    for r, v in prev_values.items():
        ws.cell(row=r, column=3, value=v)

    def _set(row, new_value):
        old = ws.cell(row=row, column=3).value
        ws.cell(row=row, column=4, value=new_value)
        if isinstance(old, (int, float)) and old and isinstance(new_value, (int, float)):
            ws.cell(row=row, column=5, value=(new_value - old) / old * 100)
        else:
            ws.cell(row=row, column=5, value=CONFIRM_NEEDED)

    layout = {"참여": (37, 38, 39, 40), "도달": (41, 42, 43, 44), "동영상조회": (45, 46, 47, 48)}
    for t, (r_cnt, r_spend, r_result, r_avg) in layout.items():
        d = ad_totals_by_type.get(t, {"진행수량": 0, "지출금액": 0, "총결과": 0})
        _set(r_cnt, d["진행수량"])
        _set(r_spend, d["지출금액"])
        _set(r_result, d["총결과"])
        scale = 1000 if t == "도달" else 1  # 도달 유형은 '결과당 비용'이 *1000 스케일(원본 규칙)
        avg = (d["지출금액"] / d["총결과"] * scale) if d.get("총결과") else None
        _set(r_avg, avg)
    total_spend_month = sum((ad_totals_by_type.get(t, {}).get("지출금액") or 0) for t in layout)
    _set(49, total_spend_month)
    # 행37 E열('=D37-C37')과 행49 D열('=C28')은 전월 최종본에서도 리터럴이 아니라
    # 살아있는 수식이었다(37행만 %가 아니라 건수 차이, 49행은 광고비 합계표를 참조).
    # 위 _set()이 두 칸 다 값으로 덮어썼으므로 같은 수식 패턴으로 되돌린다.
    ws.cell(row=37, column=5, value="=D37-C37")
    ws.cell(row=49, column=4, value=f"={month_L}28")


# ---------------------------------------------------------------------------
# 최상위 취합 함수
# ---------------------------------------------------------------------------

def _last_filled_month_num(wb, first_row=6, first_month=7):
    """방금 openpyxl로 새로 저장해 실제 엑셀로 한 번도 열어보지 않은 파일은
    새로 써넣은 수식 셀이 data_only 스냅샷에서 캐시값 없이 None으로 읽힌다.
    그래서 '숫자인지'가 아니라 '비어있지 않은지'로 채워졌는지를 판단한다 —
    라이브 워크북에서는 수식 문자열이, data_only 스냅샷에서는 캐시된 숫자가
    잡히므로 두 경우 모두 이 기준으로 정상 동작한다."""
    ws = wb["인스타그램"]
    last_filled = None
    for i in range(6):
        r = first_row + i
        v = ws.cell(row=r, column=4).value
        if v is not None and v != "":
            last_filled = first_month + i
    return last_filled


def detect_target_month(wb, first_row=6, first_month=7):
    """전월 최종본의 인스타그램 시트 요약 표에서 마지막으로 채워진 달의 다음 달을 찾는다.
    (로우데이터를 '새로' 만들 때 이번 달이 몇 월인지 판단하는 용도)"""
    last_filled = _last_filled_month_num(wb, first_row, first_month)
    if last_filled is None:
        raise ValueError("전월 최종본에서 채워진 월을 찾지 못했습니다.")
    return f"{last_filled + 1}월"


def detect_latest_filled_month(wb, first_row=6, first_month=7):
    """이미 완성된('최종') 로우데이터 파일에서 마지막으로 채워진 달 자체를 찾는다.
    (PPT를 독립적으로 만들 때, 그 파일이 어느 달 몫인지 판단하는 용도)
    라이브 워크북(data_only=False로 읽은 wb)을 넘겨야 한다 — 실제 엑셀로 열어
    저장하지 않은 파일은 수식 셀의 캐시값이 없어 data_only 스냅샷에서 비어보인다."""
    last_filled = _last_filled_month_num(wb, first_row, first_month)
    if last_filled is None:
        raise ValueError("로우데이터에서 채워진 월을 찾지 못했습니다.")
    return f"{last_filled}월"


def build_raw_data(prev_workbook_bytes, target_month, publish_lists, content_perf, ad_rows):
    """전월 최종본(bytes) + 이번달 입력들을 받아 취합된 워크북(openpyxl Workbook)과
    ConfirmLog, PPT용 파이썬 계산값 캐시(block_cache)를 반환한다.

    제휴(인플루언서 체험단) 데이터는 이 함수에서 다루지 않는다 — PPT 만들기 단계에서
    '최종 로우데이터'에 append_partnership_block()을 별도로 호출해 채운다."""
    import io
    import openpyxl
    from core.confirm import ConfirmLog

    raw_bytes = prev_workbook_bytes.read() if hasattr(prev_workbook_bytes, "read") else prev_workbook_bytes
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes))
    confirm = ConfirmLog()
    block_cache = {}

    instagram_totals, block_cache["인스타그램"] = append_instagram_block(
        wb, target_month, publish_lists["인스타그램"], content_perf, ad_rows, confirm)
    blog_totals, block_cache["블로그"] = append_blog_block(wb, target_month, publish_lists["블로그"], confirm)

    content_rows = [r for r in publish_lists["인스타그램"] if r.month == target_month]

    ad_totals_by_type = {}
    for ad_type in ("참여", "도달", "동영상조회"):
        totals, cache = append_ad_type_block(
            wb, ad_type, target_month, ad_rows, content_rows, confirm)
        ad_totals_by_type[ad_type] = totals
        block_cache[_AD_SHEETS[ad_type]["sheet"]] = cache
    append_ad_all_block(wb, target_month, ad_rows, content_rows, confirm)

    update_operations_summary(wb, target_month, ad_totals_by_type)

    # 운영요약 KPI/광고비 셀은 이제 전부 살아있는 수식이라 openpyxl로 값을 못 읽는다.
    # PPT 등에서 필요한 실제 숫자(이번 달 값)는 여기 파이썬 계산값을 그대로 넘긴다.
    block_cache["_summary"] = {
        "instagram_totals": instagram_totals,
        "blog_totals": blog_totals,
        "ad_totals_by_type": ad_totals_by_type,
    }

    return wb, block_cache, confirm
