# -*- coding: utf-8 -*-
"""PPT 표 행 복제·하이퍼링크·글자색 엔진.

python-pptx는 표에 행을 추가/삭제하는 고수준 API가 없어서, 템플릿 행(<a:tr>)의
서식을 그대로 복제해 XML 레벨에서 행 개수를 이번 달 데이터 건수에 맞춘다.
"""
import copy
from pptx.util import Pt
from pptx.dml.color import RGBColor

UP_COLOR = RGBColor(0xFF, 0x00, 0x00)     # 상승 ▲ 빨강
DOWN_COLOR = RGBColor(0x00, 0x70, 0xC0)   # 하락 ▼ 파랑
NEUTRAL_COLOR = RGBColor(0x80, 0x80, 0x80)  # 확인 필요 등 중립 회색


def set_row_count(table, n_data_rows, template_row_idx=1, trailer_row_idxs=()):
    """표의 데이터 행 개수를 n_data_rows로 맞춘다.

    template_row_idx: 서식을 복제할 데이터 행(보통 헤더 바로 다음 행=1).
    trailer_row_idxs: 평균/합계처럼 표 맨 끝에 고정으로 남아있는 행들의
      '원래' 인덱스(헤더/기존 데이터 행 기준). 이 행들은 유지한 채 데이터 행만 늘리거나 줄인다.
    반환값: 실제 데이터 행이 시작하는 0-based row index 목록(header 제외).
    """
    tbl = table._tbl
    rows = list(tbl.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}tr"))
    header_row = rows[0]
    template_row = rows[template_row_idx]
    trailer_rows = [rows[i] for i in trailer_row_idxs]
    current_data_rows = [r for i, r in enumerate(rows) if i != 0 and i not in trailer_row_idxs]

    # 기존 데이터 행 전부 제거(트레일러 앞까지)
    for r in current_data_rows:
        tbl.remove(r)

    insert_before = trailer_rows[0] if trailer_rows else None
    new_rows = []
    for i in range(n_data_rows):
        new_row = copy.deepcopy(template_row)
        if insert_before is not None:
            insert_before.addprevious(new_row)
        else:
            tbl.append(new_row)
        new_rows.append(new_row)
    return new_rows


def set_cell_text(cell, text, keep_format=True):
    """셀의 첫 런(run) 서식은 유지한 채 텍스트만 교체."""
    tf = cell.text_frame
    if not tf.paragraphs or not tf.paragraphs[0].runs:
        tf.text = str(text)
        return
    p = tf.paragraphs[0]
    first_run = p.runs[0]
    first_run.text = str(text)
    for extra in list(p.runs[1:]):
        extra._r.getparent().remove(extra._r)
    for extra_p in list(tf.paragraphs[1:]):
        extra_p._p.getparent().remove(extra_p._p)


def set_cell_hyperlink(cell, url):
    p = cell.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].hyperlink.address = url


def set_run_color(cell, color: RGBColor):
    for p in cell.text_frame.paragraphs:
        for r in p.runs:
            r.font.color.rgb = color


def fill_row(row, values, url_col=None, url=None):
    """row: table.rows[i](python-pptx _Row). values: 셀 텍스트 리스트(왼쪽부터).
    url_col: 하이퍼링크를 걸 셀의 0-based 열 인덱스."""
    cells = list(row.cells)
    for i, v in enumerate(values):
        if i >= len(cells):
            break
        set_cell_text(cells[i], v)
        if url_col is not None and i == url_col and url:
            set_cell_hyperlink(cells[i], url)
