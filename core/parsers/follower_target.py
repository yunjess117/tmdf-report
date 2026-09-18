# -*- coding: utf-8 -*-
"""인스타그램 '타깃' CSV(Meta 인사이트 내보내기 - 상위 도시/국가/연령 및 성별) 파서.

'연령 및 성별' 표만 쓴다(PPT의 팔로워 성별 파이차트·연령대별 성별 막대차트 원본).
성별 파이차트 값은 연령대별 값의 합으로 계산한다(원본에 성별 합계 행이 따로 없음).
"""
from dataclasses import dataclass, field
import csv
import io

AGE_BRACKETS = ["18-24", "25-34", "35-44", "45-54", "55-64", "65+"]


@dataclass
class FollowerTarget:
    age_labels: list = field(default_factory=list)   # CSV 원문 라벨(예: '18-24')
    female_by_age: list = field(default_factory=list)  # % (0~100)
    male_by_age: list = field(default_factory=list)
    female_total: float = None
    male_total: float = None


def _read_text(raw_bytes: bytes) -> str:
    if raw_bytes[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in raw_bytes[:200]:
        return raw_bytes.decode("utf-16")
    for enc in ("utf-8-sig", "cp949", "utf-8"):
        try:
            return raw_bytes.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("CSV 인코딩을 판별할 수 없음")


def parse_follower_target(file_like) -> FollowerTarget:
    raw = file_like.read()
    if isinstance(raw, str):
        raw = raw.encode("utf-8")
    text = _read_text(raw)
    reader = csv.reader(io.StringIO(text))
    rows = [r for r in reader]

    # '연령 및 성별' 섹션 찾기: 헤더 다음 줄이 ['', '여성', '남성'] 형태
    section_start = None
    for i, row in enumerate(rows):
        if row and row[0].strip() == "연령 및 성별":
            section_start = i
            break
    if section_start is None:
        raise ValueError("CSV에서 '연령 및 성별' 섹션을 찾을 수 없음")

    header = rows[section_start + 1]
    col_idx = {name.strip(): i for i, name in enumerate(header)}
    female_col, male_col = col_idx.get("여성"), col_idx.get("남성")
    if female_col is None or male_col is None:
        raise ValueError("'연령 및 성별' 섹션에서 여성/남성 열을 찾을 수 없음")

    result = FollowerTarget()
    r = section_start + 2
    while r < len(rows) and rows[r] and rows[r][0].strip():
        row = rows[r]
        age_label = row[0].strip()
        try:
            female_v = float(row[female_col])
            male_v = float(row[male_col])
        except (ValueError, IndexError):
            r += 1
            continue
        result.age_labels.append(age_label)
        result.female_by_age.append(female_v)
        result.male_by_age.append(male_v)
        r += 1

    if result.female_by_age:
        result.female_total = sum(result.female_by_age)
        result.male_total = sum(result.male_by_age)
    return result
