"""Vietnamese anchor recognition, independent from OOXML manipulation."""
import re

QUESTION_RE = re.compile(r"^\s*Câu\s+(?P<number>\d+)\s*[.:]", re.IGNORECASE)
SECTION_RE = re.compile(r"^\s*(?:PHẦN|PHAN|PART)\s+(?:[IVXLCDM]+|\d+|[A-Z])(?:\s*[.:–—-]|\s|$)", re.IGNORECASE)
END_RE = re.compile(r"^\s*[-–—_=*.\s]*(?:HẾT|HET|END)(?:\s|[-–—_=*.]|$)", re.IGNORECASE)
SCORE_RE = re.compile(r"\(\s*(\d+(?:[,.]\d+)?)\s*điểm\s*\)", re.IGNORECASE)
MC_RE = re.compile(r"(?<!\S)(?P<label>[A-D])(?P<punct>[.)])(?=\s|$)")
TF_RE = re.compile(r"(?<!\S)(?P<label>[a-d])(?P<punct>\))(?=\s|$)")


def section_type(title: str):
    from .models import QuestionType
    title = title.casefold()
    if "tự luận" in title or "essay" in title:
        return QuestionType.ESSAY
    if "đúng" in title and "sai" in title or "true/false" in title:
        return QuestionType.TRUE_FALSE
    if "trắc nghiệm" in title or "multiple choice" in title or "multiple-choice" in title:
        return QuestionType.MULTIPLE_CHOICE
    return None
