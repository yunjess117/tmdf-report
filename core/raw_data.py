# -*- coding: utf-8 -*-
"""로우데이터(취합) 조합 로직 — 전월 최종본에 이번 달 블록을 이어붙인다."""
import datetime as dt

from core.excel_block import (
    CONFIRM_NEEDED, find_last_label_row, insert_block, write_row,
)
from core.formatters import normalize_url_key
from core.matching import content_ad_values, match_ad_for_content, match_content_for_ad

NO_COL = 2       # B열: NO.
TOPIC_COL = 5    # E열: 콘텐츠 주제 (평균/합계 라벨이 찍히는 열)


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

    built = []
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
        row_no = i + 1
        values = [row_no, target_month, pr.pub_date, pr.title, pr.url, pr.content_type,
                  ad_spend, perf["조회"], perf["도달"], ad_reach, perf["팔로우"],
                  perf["좋아요"], perf["댓글"], perf["공유"], perf["저장"], perf["총반응"]]
        write_row(ws, insert_at + i, 2, values)
        built.append({"perf": perf, "reach": pr, "ad_spend": ad_spend})

    def col(field):
        idx = {"조회": 9, "도달": 10, "광고도달": 11, "팔로우": 12,
               "좋아요": 13, "댓글": 14, "공유": 15, "저장": 16, "총반응": 17}[field]
        return [ws.cell(row=insert_at + i, column=idx).value for i in range(n_data)]

    avg_row, total_row = insert_at + n_data, insert_at + n_data + 1
    ws.cell(row=avg_row, column=topic_col, value="평균")
    ws.cell(row=total_row, column=topic_col, value="합계")
    for c in range(7, 18):
        vals = [ws.cell(row=insert_at + i, column=c).value for i in range(n_data)]
        ws.cell(row=avg_row, column=c, value=_avg(vals))
        ws.cell(row=total_row, column=c, value=_sum(vals))

    totals = {
        "발행수": n_data,
        "도달수": _sum(col("도달")),
        "영상조회수": _sum([ws.cell(row=insert_at + i, column=9).value
                        for i, pr in enumerate(rows) if pr.content_type == "영상"])
                    or _sum(col("조회")) if any(r.content_type == "영상" for r in rows) else 0,
        "참여수": _sum(col("총반응")),
    }
    # '영상 조회수'는 운영요약 정의상 '영상' 유형 콘텐츠의 조회수 합산이 아니라
    # 메타 인사이트 영상 조회수 합산이며, 이 시트의 '조회' 컬럼 자체가 그 값이다.
    totals["영상조회수"] = _sum(col("조회"))
    _update_month_summary_rows(ws, target_month, totals["발행수"], totals["도달수"],
                                totals["영상조회수"], totals["참여수"])
    return totals


def _update_month_summary_rows(ws, target_month, publish_cnt, reach, views, engage,
                                first_row=6, first_month=7):
    r = first_row + (_month_num(target_month) - first_month)
    ws.cell(row=r, column=4, value=publish_cnt)
    ws.cell(row=r, column=5, value=reach)
    ws.cell(row=r, column=6, value=views)
    ws.cell(row=r, column=7, value=engage)


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

    for i, pr in enumerate(rows):
        confirm.add("블로그", f"{target_month} NO.{i+1} ({pr.title})",
                    "블로그 성과 원본 파일이 입력되지 않아 조회/공감/댓글 확인 필요")
        values = [i + 1, target_month, pr.pub_date, pr.title, pr.url, pr.content_type,
                  CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED]
        write_row(ws, insert_at + i, 2, values)

    avg_row, total_row = insert_at + n_data, insert_at + n_data + 1
    ws.cell(row=avg_row, column=topic_col, value="평균")
    ws.cell(row=total_row, column=topic_col, value="합계")
    for c in range(8, 12):
        ws.cell(row=avg_row, column=c, value=CONFIRM_NEEDED)
        ws.cell(row=total_row, column=c, value=CONFIRM_NEEDED)

    r = 6 + (_month_num(target_month) - 7)
    ws.cell(row=r, column=4, value=n_data)
    for c in (5, 6, 7):
        ws.cell(row=r, column=c, value=CONFIRM_NEEDED)
    return {"발행수": n_data}


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

    for i, (content, a) in enumerate(matched):
        if content is None:
            confirm.add(info["sheet"], f"{target_month} NO.{i+1} ({a.ad_name})",
                        "발행일/제목이 일치하는 콘텐츠를 발행리스트에서 찾지 못함 - 발행일/콘텐츠 주제/유형 확인 필요")
            pub_date, title, ctype = CONFIRM_NEEDED, CONFIRM_NEEDED, CONFIRM_NEEDED
        else:
            pub_date, title, ctype = content.pub_date, content.title, content.content_type
        camp_start = a.campaign_start if a.campaign_start else CONFIRM_NEEDED
        camp_end = a.campaign_end if a.campaign_end else CONFIRM_NEEDED
        # AD 데이터 시트의 '지출 금액'은 광고 원본 그대로(raw)를 쓴다. 만원 단위 올림은
        # 콘텐츠 시트의 '광고비' 컬럼에만 적용하는 규칙(LX세미콘 문서 기준)이라 여기선 하지 않는다.
        spend = a.spend if a.spend is not None else CONFIRM_NEEDED
        if info["has_vtr"]:
            vtr = (a.result / a.impressions) if a.result is not None and a.impressions else None
            values = [i + 1, target_month, pub_date, title, ctype, CONFIRM_NEEDED,
                      camp_start, camp_end, a.result, a.cost_per_result, vtr,
                      a.impressions, a.reach, spend, CONFIRM_NEEDED]
        else:
            values = [i + 1, target_month, pub_date, title, ctype, CONFIRM_NEEDED,
                      camp_start, camp_end, a.result, a.cost_per_result,
                      a.impressions, a.reach, spend, CONFIRM_NEEDED]
        write_row(ws, insert_at + i, 2, values)
        confirm.add(info["sheet"], f"{target_month} NO.{i+1}", "타깃/광고예산은 원본에 없어 확인 필요")

    # 열 매핑 (절대 컬럼 번호). 헤더: NO,월,발행일,콘텐츠주제,유형,타깃,시작,종료,결과,결과당비용,[VTR,]노출,도달,지출,예산
    if info["has_vtr"]:
        cols = {"결과": 10, "결과당비용": 11, "VTR": 12, "노출": 13, "도달": 14, "지출": 15, "예산": 16}
    else:
        cols = {"결과": 10, "결과당비용": 11, "노출": 12, "도달": 13, "지출": 14, "예산": 15}

    avg_row, total_row = insert_at + n_data, insert_at + n_data + 1
    ws.cell(row=avg_row, column=topic_col, value="평균")
    ws.cell(row=total_row, column=topic_col, value="합계")
    for name, c in cols.items():
        vals = [ws.cell(row=insert_at + i, column=c).value for i in range(n_data)]
        vals = [v for v in vals if isinstance(v, (int, float))]
        if name in ("결과", "노출", "도달", "지출"):
            ws.cell(row=avg_row, column=c, value=_avg(vals))
            ws.cell(row=total_row, column=c, value=_sum(vals) if vals else None)
        elif name == "결과당비용":
            spend_vals = [ws.cell(row=insert_at + i, column=cols["지출"]).value for i in range(n_data)]
            spend_vals = [v for v in spend_vals if isinstance(v, (int, float))]
            result_vals = [ws.cell(row=insert_at + i, column=cols["결과"]).value for i in range(n_data)]
            result_vals = [v for v in result_vals if isinstance(v, (int, float))]
            scale = 1000 if ad_type == "도달" else 1  # 전월 최종본의 기존 수식 관례를 그대로 따름
            avg_cpr = (sum(spend_vals) / sum(result_vals) * scale) if spend_vals and sum(result_vals) else None
            ws.cell(row=avg_row, column=c, value=avg_cpr)
            ws.cell(row=total_row, column=c, value="-")
        elif name == "VTR":
            imp = [ws.cell(row=insert_at + i, column=cols["노출"]).value for i in range(n_data)]
            imp = [v for v in imp if isinstance(v, (int, float))]
            res = [ws.cell(row=insert_at + i, column=cols["결과"]).value for i in range(n_data)]
            res = [v for v in res if isinstance(v, (int, float))]
            avg_vtr = (sum(res) / sum(imp)) if imp and sum(imp) else None
            ws.cell(row=avg_row, column=c, value=avg_vtr)
            ws.cell(row=total_row, column=c, value="-")

    total_spend = _sum([ws.cell(row=insert_at + i, column=cols["지출"]).value for i in range(n_data)])
    total_result = _sum([ws.cell(row=insert_at + i, column=cols["결과"]).value for i in range(n_data)])
    scale = 1000 if ad_type == "도달" else 1
    avg_cost = (total_spend / total_result * scale) if total_spend and total_result else None
    r = 5 + (_month_num(target_month) - 7)
    ws.cell(row=r, column=4, value=n_data)
    ws.cell(row=r, column=5, value=total_result)
    ws.cell(row=r, column=6, value=avg_cost)
    ws.cell(row=r, column=7, value=total_spend)
    ws.cell(row=r, column=8, value=CONFIRM_NEEDED)  # 광고예산: 원본에 없음

    return {"진행수량": n_data, "총결과": total_result, "지출금액": total_spend}


# ---------------------------------------------------------------------------
# AD 데이터(전체) — 참여/도달/동영상조회 원본 광고 행을 한 표로 통합
# ---------------------------------------------------------------------------

def append_ad_all_block(wb, wb_data, target_month, ad_rows, content_rows, confirm):
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

    from core.formatters import roundup_10000

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
        values = [i + 1, target_month, pub_date, title, ctype, camp_start, camp_end,
                  a.result_type, a.result, a.cost_per_result, a.impressions, a.reach,
                  v2s, v2s_cost, vtr, a.link_clicks, a.cpc, a.ctr, a.cpm, spend,
                  CONFIRM_NEEDED]
        write_row(ws, insert_at + i, 2, values)

    totals_by_type = {}
    for t in ("참여", "도달", "동영상조회"):
        vals = [a.spend for _, a in matched if a.ad_type == t and a.spend is not None]
        # 캠페인 단위로 각각 만원 올림한 뒤 합산(LX세미콘 규칙: 실지출을 캠페인별로 올림)
        totals_by_type[t] = sum(roundup_10000(v) for v in vals) if vals else 0

    r = 5 + (_month_num(target_month) - 7)
    ws.cell(row=r, column=4, value=target_month)
    ws.cell(row=r, column=5, value=sum(totals_by_type.values()))
    ws.cell(row=r, column=6, value=totals_by_type["참여"])
    ws.cell(row=r, column=7, value=totals_by_type["도달"])
    ws.cell(row=r, column=8, value=totals_by_type["동영상조회"])

    # 누적 행(고정 라벨 '누적') 재계산
    cum_row = None
    for rr in range(5, r + 1):
        if ws.cell(row=rr, column=4).value == "누적":
            cum_row = rr
            break
    if cum_row:
        ws_data = wb_data["AD 데이터(전체)"] if wb_data else None
        month_rows = range(5, cum_row)
        for c in range(5, 9):
            vals = [_numval(ws, ws_data, mr, c) for mr in month_rows]
            ws.cell(row=cum_row, column=c, value=_sum(vals))
    return totals_by_type


# ---------------------------------------------------------------------------
# 제휴(인플루언서 체험단) 오픈 보고서 -> 신규 '제휴' 시트
# ---------------------------------------------------------------------------

_PARTNER_INSTA_HEADER = ["NO.", "오픈일", "인스타 아이디", "팔로워 수", "조회 수", "좋아요 수", "댓글 수", "비고"]
_PARTNER_BLOG_HEADER = ["NO.", "오픈일", "블로거", "제목", "일평균 방문자수", "오픈일 방문자수", "PV수치", "댓글 수", "공감 수", "비고"]


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
    ws = _ensure_partnership_sheet(wb)

    # 인스타 섹션: 헤더(5행) 다음부터 마지막 '합계' 뒤에 이어붙임
    last_total = find_last_label_row(ws, "합계", 3, search_from=5, search_to=200) or 5
    insert_at = last_total + 1 if last_total != 5 else 6
    for i, p in enumerate(insta_rows):
        r = insert_at + i
        write_row(ws, r, 2, [i + 1, p.open_date, p.handle, p.followers, p.views,
                              p.likes, p.comments, p.note])
    total_row = insert_at + len(insta_rows)
    ws.cell(row=total_row, column=3, value="합계")
    for c, field in ((5, "followers"), (6, "views"), (7, "likes"), (8, "comments")):
        vals = [getattr(p, field) for p in insta_rows if isinstance(getattr(p, field), (int, float))]
        ws.cell(row=total_row, column=c, value=_sum(vals))

    # 블로그 섹션: 인스타 섹션 합계 아래에 헤더를 새로 두고 이어붙임(최초 1회) 또는 기존 헤더 재사용
    blog_header_row = None
    for r in range(total_row + 1, ws.max_row + 3):
        if ws.cell(row=r, column=2).value == "블로그":
            blog_header_row = r + 1
            break
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
                              p.open_day_visitors, p.pv, p.comments, p.likes, p.note])
    blog_total_row = blog_insert_at + len(blog_rows)
    ws.cell(row=blog_total_row, column=3, value="합계")
    for c, field in ((5, "daily_visitors"), (6, "open_day_visitors"), (7, "pv"),
                     (8, "comments"), (9, "likes")):
        vals = [getattr(p, field) for p in blog_rows if isinstance(getattr(p, field), (int, float))]
        ws.cell(row=blog_total_row, column=c, value=_sum(vals))


# ---------------------------------------------------------------------------
# 운영요약 시트 — KPI/광고비 누적 표 갱신
# ---------------------------------------------------------------------------

def update_operations_summary(wb, wb_data, target_month, instagram_totals, ad_totals_by_type, rounded_spend_by_type):
    ws = wb["운영요약"]
    ws_data = wb_data["운영요약"] if wb_data else None
    month_col = 3 + (_month_num(target_month) - 7)  # C=7월 ... H=12월

    # 콘텐츠 KPI (행14~17: 발행수/도달수/영상조회수/참여수), I=누적 J=목표(고정) K=달성률
    kpi_rows = {14: instagram_totals["발행수"], 15: instagram_totals["도달수"],
                16: instagram_totals["영상조회수"], 17: instagram_totals["참여수"]}
    for row, value in kpi_rows.items():
        ws.cell(row=row, column=month_col, value=value)
        vals = [_numval(ws, ws_data, row, c) for c in range(3, 9)]
        cum = _sum(vals)
        ws.cell(row=row, column=9, value=cum)
        target = ws.cell(row=row, column=10).value
        if isinstance(target, (int, float)) and target:
            ws.cell(row=row, column=11, value=cum / target)

    # 상단 '이번달 요약' 누적수치(행9)/달성률(행10) - 콘텐츠지표 C~F
    for src_row, dst_col in ((14, 3), (15, 4), (16, 5), (17, 6)):
        cum = ws.cell(row=src_row, column=9).value
        ws.cell(row=9, column=dst_col, value=cum)
        target = ws.cell(row=8, column=dst_col).value
        if isinstance(target, (int, float)) and target:
            ws.cell(row=10, column=dst_col, value=cum / target)

    # 연간 광고비 지출 현황 (행25~28: 참여/도달/동영상조회/합계)
    # 이 표와 아래 '월간 광고 요약'은 캠페인별로 만원 단위 올림한 금액(=콘텐츠 시트
    # '광고비'와 같은 규칙)을 쓴다. AD 데이터 (참여/도달/동영상조회) 시트 자체의
    # '지출금액'은 원본 그대로(raw)라서 이 값과는 다르다.
    type_row = {"참여": 25, "도달": 26, "동영상조회": 27}
    for t, row in type_row.items():
        spend = rounded_spend_by_type.get(t) or 0
        ws.cell(row=row, column=month_col, value=spend)
    total_spend_month = sum((rounded_spend_by_type.get(t) or 0) for t in type_row)
    ws.cell(row=28, column=month_col, value=total_spend_month)
    for row in (25, 26, 27, 28):
        vals = [_numval(ws, ws_data, row, c) for c in range(3, 9)]
        ws.cell(row=row, column=9, value=_sum(vals))
    total_budget = _numval(ws, ws_data, 28, 11)  # K열 총 광고예산(고정값, 유지)
    cum_spend = ws.cell(row=28, column=9).value or 0
    if isinstance(total_budget, (int, float)):
        ws.cell(row=28, column=10, value=total_budget - cum_spend)

    # 광고비 지출 현황 요약 (행32, 최신 상태 1행으로 덮어씀)
    total_budget_fixed = _numval(ws, ws_data, 32, 5)  # 기존 총가용예산 유지
    ws.cell(row=32, column=2, value=target_month)
    ws.cell(row=32, column=3, value=cum_spend)
    if isinstance(total_budget_fixed, (int, float)):
        ws.cell(row=32, column=4, value=total_budget_fixed - cum_spend)
        ws.cell(row=32, column=5, value=total_budget_fixed)

    # 월간 광고 요약 (행36 헤더 + 37~49, 전월/이번월/증감률 2열 비교표를 최신으로 교체)
    prev_label = ws.cell(row=36, column=4).value  # 기존 D열(직전 실행의 '이번월')이 새 '전월'이 됨
    prev_values = {r: _numval(ws, ws_data, r, 4) for r in range(37, 50)}
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
        _set(r_spend, rounded_spend_by_type.get(t) or 0)
        _set(r_result, d["총결과"])
        scale = 1000 if t == "도달" else 1  # 도달 유형은 '결과당 비용'이 *1000 스케일(원본 규칙)
        avg = (d["지출금액"] / d["총결과"] * scale) if d.get("총결과") else None
        _set(r_avg, avg)
    _set(49, total_spend_month)


# ---------------------------------------------------------------------------
# 최상위 취합 함수
# ---------------------------------------------------------------------------

def detect_target_month(wb, first_row=6, first_month=7):
    """전월 최종본의 인스타그램 시트 요약 표에서 마지막으로 채워진 달의 다음 달을 찾는다."""
    ws = wb["인스타그램"]
    last_filled = None
    for i in range(6):
        r = first_row + i
        if isinstance(ws.cell(row=r, column=4).value, (int, float)):
            last_filled = first_month + i
    if last_filled is None:
        raise ValueError("전월 최종본에서 채워진 월을 찾지 못했습니다.")
    return f"{last_filled + 1}월"


def build_raw_data(prev_workbook_bytes, target_month, publish_lists, content_perf,
                    ad_rows, partnership):
    """전월 최종본(bytes) + 이번달 입력들을 받아 취합된 워크북(openpyxl Workbook)과
    ConfirmLog를 반환한다."""
    import io
    import openpyxl
    from core.confirm import ConfirmLog

    raw_bytes = prev_workbook_bytes.read() if hasattr(prev_workbook_bytes, "read") else prev_workbook_bytes
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes))
    # 전월 최종본에는 다른 시트를 참조하는 수식이 많이 남아 있어(예: 운영요약의
    # '=AD 데이터(전체)!F5'), openpyxl로는 계산된 값을 읽을 수 없다. data_only=True로
    # 한 번 더 읽어 엑셀이 마지막 저장 시 캐시해 둔 계산값을 따로 확보해 둔다.
    wb_data = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    confirm = ConfirmLog()

    instagram_totals = append_instagram_block(
        wb, target_month, publish_lists["인스타그램"], content_perf, ad_rows, confirm)
    append_blog_block(wb, target_month, publish_lists["블로그"], confirm)

    content_rows = [r for r in publish_lists["인스타그램"] if r.month == target_month]

    ad_totals_by_type = {}
    for ad_type in ("참여", "도달", "동영상조회"):
        ad_totals_by_type[ad_type] = append_ad_type_block(
            wb, ad_type, target_month, ad_rows, content_rows, confirm)
    rounded_spend_by_type = append_ad_all_block(wb, wb_data, target_month, ad_rows, content_rows, confirm)

    append_partnership_block(wb, target_month, partnership["인스타"], partnership["블로그"])

    update_operations_summary(wb, wb_data, target_month, instagram_totals, ad_totals_by_type, rounded_spend_by_type)

    return wb, wb_data, confirm
