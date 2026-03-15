from __future__ import annotations

import re

from opencc import OpenCC


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text or "").lower()


def parse_target_list(raw_text: str) -> list[str]:
    parts = re.split(r"[\r\n,?;?|]+", raw_text or "")
    results: list[str] = []
    seen: set[str] = set()
    for part in parts:
        candidate = part.strip()
        normalized = normalize_text(candidate)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        results.append(candidate)
    return results


class TextNormalizer:
    def __init__(self) -> None:
        self.to_traditional = OpenCC("s2t")
        self.to_simplified = OpenCC("t2s")

    def forms(self, text: str) -> set[str]:
        cleaned = normalize_text(text)
        if not cleaned:
            return set()
        forms = {cleaned}
        forms.add(self.to_traditional.convert(cleaned))
        forms.add(self.to_simplified.convert(cleaned))
        return {item for item in forms if item}

    def matches(self, target: str, texts: list[str]) -> tuple[bool, str]:
        target_forms = self.forms(target)
        if not target_forms:
            return False, ""

        for text in texts:
            candidate_forms = self.forms(text)
            for candidate in candidate_forms:
                if any(target_form in candidate for target_form in target_forms):
                    return True, text

        merged_forms = self.forms("".join(texts))
        for candidate in merged_forms:
            if any(target_form in candidate for target_form in target_forms):
                return True, "".join(texts)

        return False, ""

    def matches_any(self, targets: list[str], texts: list[str]) -> tuple[bool, str, str]:
        for target in targets:
            matched, hit = self.matches(target, texts)
            if matched:
                return True, target, hit
        return False, "", ""
