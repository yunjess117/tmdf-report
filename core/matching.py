# -*- coding: utf-8 -*-
"""발행 콘텐츠 <-> 광고 원본 매칭.

실데이터로 검증해 보니 광고 소재명의 날짜 코드(YYMMDD_제목, anchor_date)는
광고를 만든 날짜이며, 콘텐츠 발행일보다 1~2일 늦게 생성되는 경우가 흔하다
(예: 8/26 발행 콘텐츠 2건에 대해 8/27, 8/28에 각각 광고를 만든 사례). 그래서
'날짜가 정확히 하나 일치하면 그걸 쓴다'는 방식은, 같은 날 발행된 다른 콘텐츠의
광고 날짜가 우연히 오늘 날짜와 같아지는 순간 엉뚱하게 매칭된다(실제로 발생:
8/26 발행된 '상인 스토리'가 날짜만으로는 같은 날 앵커를 가진 '메인' 영상
광고와 잘못 매칭됐었음).

그래서 날짜 근접도와 제목 텍스트 겹침을 함께 점수화해서(점수 = 겹치는 단어 수
*10 - 날짜 차이일수) 가장 높은 후보를 고른다. 텍스트가 겹치는 실제 콘텐츠가
날짜만 우연히 같은 다른 후보보다 항상 높은 점수를 받도록 하기 위함이다. 후보가
1건뿐이면 그대로 쓰고, 여러 건인데 유의미한 겹침(>=2단어)이 전혀 없으면 절대
추측하지 않고 '확인 필요'로 남긴다.
"""
import re
from difflib import SequenceMatcher
from core.excel_block import CONFIRM_NEEDED

_STRIP_RE = re.compile(r"[^0-9A-Za-z가-힣]+")  # \w는 밑줄(_)도 포함하므로 명시적으로 제외
_MIN_RATIO = 0.3
_WINDOW_DAYS = 3


def _normalize(text: str) -> str:
    """공백/기호를 모두 제거해 '시장 메뉴 추천'과 '시장메뉴추천' 같은 표기 차이를 흡수."""
    if not text:
        return ""
    return _STRIP_RE.sub("", text)


def _lcs_ratio(a: str, b: str) -> float:
    """최장 공통 '연속' 부분문자열 길이 / 두 문자열 중 짧은 쪽 길이.

    같은 시리즈에 접미어만 다른 제목들(예: '...메인' vs '...메인 인터뷰')을 비교할 때
    difflib의 표준 ratio()는 전체 길이 차이에 끌려가 엉뚱한 후보를 더 높게 채점하는
    경우가 있어(짧은 쪽이 우연히 길이가 더 비슷해서), 대신 공통 접두 구간의 길이를
    더 짧은 문자열 기준으로 정규화한다.
    """
    if not a or not b:
        return 0.0
    match = SequenceMatcher(None, a, b).find_longest_match(0, len(a), 0, len(b))
    return match.size / min(len(a), len(b))


def _ad_title_fragment(ad_name: str) -> str:
    if "_" in ad_name:
        return ad_name.split("_", 1)[1]
    return ad_name


def _best_candidate(anchor_date, norm_key, candidates, date_of, key_of):
    """candidates: 매칭 후보 리스트. date_of/key_of: 후보에서 날짜/제목 문자열을 뽑는 함수.

    점수 = 제목 유사도(0~1)*100 - 날짜 차이일수. 유사도가 낮아도 후보가 1건뿐이면
    (window이 ±3일로 좁으므로) 그대로 채택한다. 후보가 여럿인데 최고 유사도가
    낮으면(<0.3) 절대 추측하지 않고 확인 필요로 남긴다.
    """
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]
    scored = []
    for c in candidates:
        d = date_of(c)
        distance = (d - anchor_date).days if d and anchor_date else 99
        ratio = _lcs_ratio(norm_key, _normalize(key_of(c)))
        scored.append((ratio * 100 - abs(distance), ratio, c))
    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_ratio, _ = scored[0]
    if best_ratio < _MIN_RATIO:
        return None
    ties = [o for s, r, o in scored if s == best_score]
    return ties[0] if len(ties) == 1 else None


def match_ad_for_content(pub_date, title, ad_rows):
    window = [a for a in ad_rows
              if a.anchor_date and pub_date and -_WINDOW_DAYS <= (a.anchor_date - pub_date).days <= _WINDOW_DAYS]
    return _best_candidate(pub_date, _normalize(title), window,
                            date_of=lambda a: a.anchor_date,
                            key_of=lambda a: _ad_title_fragment(a.ad_name))


def match_content_for_ad(ad, content_rows):
    """AD 데이터 시트 행을 채우기 위해, 광고 1건에 대응하는 발행 콘텐츠(PublishRow)를 찾는다."""
    window = [c for c in content_rows
              if c.pub_date and ad.anchor_date and -_WINDOW_DAYS <= (c.pub_date - ad.anchor_date).days <= _WINDOW_DAYS]
    return _best_candidate(ad.anchor_date, _normalize(_ad_title_fragment(ad.ad_name)), window,
                            date_of=lambda c: c.pub_date,
                            key_of=lambda c: c.title)


def content_ad_values(pub_date, title, ad_rows):
    """콘텐츠 시트의 '광고비'/'광고도달' 값을 반환. 매칭 실패 시 확인 필요."""
    ad = match_ad_for_content(pub_date, title, ad_rows)
    if ad is None:
        return CONFIRM_NEEDED, CONFIRM_NEEDED, None
    from core.formatters import roundup_10000
    spend = roundup_10000(ad.spend) if ad.spend is not None else CONFIRM_NEEDED
    reach = ad.reach if ad.reach is not None else CONFIRM_NEEDED
    return spend, reach, ad
