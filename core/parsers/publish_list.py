# -*- coding: utf-8 -*-
"""발행리스트 (인스타그램/블로그 시트) 파서."""
from dataclasses import dataclass
import datetime as dt
import openpyxl

from core.formatters import parse_month_label, parse_pub_date


@dataclass
class PublishRow:
    no: int
    month: str          # '8월'
    pub_date: dt.date
    title: str
    url: str
    content_type: str   # 카드뉴스/영상 등


def _guess_year(pub_date_hint: int, month: int) -> int:
    return pub_date_hint


def parse_publish_list(file_like, year_hint: int) -> dict:
    """발행리스트 엑셀을 읽어 {'인스타그램': [PublishRow,...], '블로그': [...]}로 반환.

    시트에는 발행일 연도가 없어 year_hint(발행리스트를 업로드한 시점의 연도)를
    사용한다. 12월->1월로 넘어가는 회계연도 경계는 다루지 않는다.
    """
    wb = openpyxl.load_workbook(file_like, data_only=True)
    result = {}
    for sheet_name in ("인스타그램", "블로그"):
        if sheet_name not in wb.sheetnames:
            result[sheet_name] = []
            continue
        ws = wb[sheet_name]
        header_row = None
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=2).value == "NO.":
                header_row = r
                break
        rows = []
        if header_row:
            for r in range(header_row + 1, ws.max_row + 1):
                no = ws.cell(row=r, column=2).value
                month_raw = ws.cell(row=r, column=3).value
                date_raw = ws.cell(row=r, column=4).value
                title = ws.cell(row=r, column=5).value
                url = ws.cell(row=r, column=6).value
                ctype = ws.cell(row=r, column=7).value
                if no is None and title is None:
                    continue
                month = parse_month_label(month_raw)
                try:
                    pub_date = parse_pub_date(date_raw, year_hint)
                except ValueError:
                    pub_date = None
                rows.append(PublishRow(
                    no=no, month=month, pub_date=pub_date,
                    title=(str(title).strip() if title else ""),
                    url=(str(url).strip() if url else ""),
                    content_type=(str(ctype).strip() if ctype else ""),
                ))
        result[sheet_name] = rows
    return result


def latest_month_in(rows: list) -> str:
    """행 목록에서 가장 늦은 월을 '8월' 형태로 반환. year_hint 순서를 알고 있으므로
    월 숫자만으로 비교(연 경계를 넘는 경우는 다루지 않음)."""
    months = [r.month for r in rows if r.month]
    if not months:
        return ""
    from core.formatters import month_to_int
    return max(months, key=month_to_int)
