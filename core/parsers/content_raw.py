# -*- coding: utf-8 -*-
"""인스타그램 콘텐츠 성과 원본 파서.

두 가지 형식을 모두 지원한다.
  (a) 플랫폼(Meta Business Suite) 원본 CSV
      헤더: 게시물 ID,계정 ID,계정 사용자 이름,계정 이름,설명,기간(초),게시 시간,
            고유 링크,게시물 유형,데이터 댓글,날짜,조회,도달,팔로우,좋아요,댓글,공유,저장
  (b) 손으로 정리한 '정리본' xlsx (NO/발행리스트 열이 있음)
      헤더: NO,발행리스트,설명,게시 시간,고유 링크,게시물 유형,도달,조회,좋아요,댓글,저장,공유,팔로우

반환값은 매칭키(정규화된 고유 링크) -> 성과 dict.
"""
import csv
import io
import openpyxl

from core.formatters import normalize_url_key, total_engagement

# 인코딩 자동 판별용 시도 순서. LX 프로젝트 규칙과 동일하게
# UTF-8 BOM을 우선 판별하고, UTF-16이 UTF-8로 잘못 디코딩되는 함정을 피한다.
_ENCODINGS = ["utf-8-sig", "utf-16", "cp949", "utf-8"]


def _read_text(raw_bytes: bytes) -> str:
    if raw_bytes[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in raw_bytes[:200]:
        return raw_bytes.decode("utf-16")
    for enc in _ENCODINGS:
        try:
            return raw_bytes.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("CSV 인코딩을 판별할 수 없음")


def _parse_platform_csv(raw_bytes: bytes) -> dict:
    text = _read_text(raw_bytes)
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        return {}
    header = [h.strip() for h in rows[0]]
    idx = {h: i for i, h in enumerate(header)}

    def col(row, name):
        i = idx.get(name)
        return row[i] if i is not None and i < len(row) else None

    result = {}
    for row in rows[1:]:
        if not row or len(row) < 3:
            continue
        url = col(row, "고유 링크")
        key = normalize_url_key(url)
        if not key:
            continue

        def num(name):
            v = col(row, name)
            try:
                return int(float(v))
            except (TypeError, ValueError):
                return None

        likes, comments, shares, saves = num("좋아요"), num("댓글"), num("공유"), num("저장")
        result[key] = {
            "조회": num("조회"),
            "도달": num("도달"),
            "팔로우": num("팔로우"),
            "좋아요": likes,
            "댓글": comments,
            "공유": shares,
            "저장": saves,
            "총반응": total_engagement(likes, comments, shares, saves),
        }
    return result


def _parse_cleaned_xlsx(file_like) -> dict:
    wb = openpyxl.load_workbook(file_like, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row = None
    header = {}
    for r in range(1, min(ws.max_row, 10) + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if "고유 링크" in values:
            header_row = r
            header = {v: c for c, v in enumerate(values, start=1) if v}
            break
    if header_row is None:
        raise ValueError("정리본 헤더(고유 링크)를 찾을 수 없음")

    def get(row, name):
        c = header.get(name)
        return ws.cell(row=row, column=c).value if c else None

    result = {}
    for r in range(header_row + 1, ws.max_row + 1):
        url = get(r, "고유 링크")
        key = normalize_url_key(url)
        if not key:
            continue

        def num(name):
            v = get(r, name)
            try:
                return int(float(v))
            except (TypeError, ValueError):
                return None

        likes, comments, shares, saves = num("좋아요"), num("댓글"), num("공유"), num("저장")
        result[key] = {
            "조회": num("조회"),
            "도달": num("도달"),
            "팔로우": num("팔로우"),
            "좋아요": likes,
            "댓글": comments,
            "공유": shares,
            "저장": saves,
            "총반응": total_engagement(likes, comments, shares, saves),
        }
    return result


def parse_content_raw(file_like, filename: str) -> dict:
    """파일명 확장자로 CSV/xlsx 형식을 판별해 통합 dict를 반환."""
    if filename.lower().endswith(".csv"):
        raw = file_like.read()
        if isinstance(raw, str):
            raw = raw.encode("utf-8")
        return _parse_platform_csv(raw)
    return _parse_cleaned_xlsx(file_like)
