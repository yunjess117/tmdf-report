# -*- coding: utf-8 -*-
"""시트에 쓰인 '확인 필요' 항목을 모아 화면에 보여주기 위한 수집기."""
from dataclasses import dataclass, field


@dataclass
class ConfirmLog:
    items: list = field(default_factory=list)

    def add(self, sheet, location, reason):
        self.items.append({"sheet": sheet, "location": location, "reason": reason})

    def extend(self, other: "ConfirmLog"):
        self.items.extend(other.items)
