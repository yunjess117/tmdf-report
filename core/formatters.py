# -*- coding: utf-8 -*-
"""날짜/URL/금액 등 공통 정규화 유틸리티."""
import re
import datetime as dt

_WEEKDAY_RE = re.compile(r"\s*\([월화수목금토일]\)\s*$")


def parse_month_label(value) -> str:
    """'7월', 7, '07' 등을 '7월' 형태로 정규화."""
    if value is None:
        return ""
    s = str(value).strip()
    if s.endswith("월"):
        return s
    try:
        return f"{int(s)}월"
    except ValueError:
        return s


def month_to_int(label: str) -> int:
    s = str(label).strip().rstrip("월")
    return int(s)


def parse_pub_date(value, year_hint: int, month_hint: int = None):
    """발행리스트의 '7/3(금)' 문자열 또는 datetime 값을 date로 정규화.

    시트에 연도가 없으므로 year_hint를 사용한다. 12월->1월로 넘어가는
    회계연도 경계는 이 프로젝트 범위에서 발생하지 않는다고 가정한다.
    """
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    s = str(value).strip()
    s = _WEEKDAY_RE.sub("", s)
    m = re.match(r"^(\d{1,2})[/.](\d{1,2})$", s)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        return dt.date(year_hint, month, day)
    raise ValueError(f"날짜 형식을 해석할 수 없음: {value!r}")


_IG_SHORTCODE_RE = re.compile(r"instagram\.com/(?:p|reel|tv)/([A-Za-z0-9_-]+)")
_NAVER_BLOG_RE = re.compile(r"blog\.naver\.com/([^/?#]+)/([0-9]+)")


def normalize_url_key(url: str) -> str:
    """인스타/네이버블로그 URL에서 매칭용 고유 키(shortcode/postId)를 뽑는다.

    인스타그램은 /p/, /reel/, /tv/ 경로가 같은 게시물을 가리킬 수 있으므로
    경로 종류를 무시하고 shortcode만 비교 키로 쓴다.
    """
    if not url:
        return ""
    s = str(url).strip()
    m = _IG_SHORTCODE_RE.search(s)
    if m:
        return "ig:" + m.group(1)
    m = _NAVER_BLOG_RE.search(s)
    if m:
        return "blog:" + m.group(2)
    # 알 수 없는 형식은 querystring 제거 + 끝 슬래시 제거한 원본 사용
    s = s.split("?")[0].rstrip("/")
    s = re.sub(r"^https?://(www\.)?", "", s)
    return s


def roundup_10000(amount) -> int:
    """실지출을 만원 단위로 올림. 179,999 -> 180,000."""
    if amount is None or amount == "":
        return None
    import math
    return int(math.ceil(float(amount) / 10000.0) * 10000)


def total_engagement(likes, comments, shares, saves) -> int:
    """총 반응수 = 좋아요+공유+댓글+저장 (누락값은 0으로 취급하지 않고 전체를 None 처리)."""
    vals = [likes, comments, shares, saves]
    if any(v is None for v in vals):
        return None
    return int(sum(vals))
