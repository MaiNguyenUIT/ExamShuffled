from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class QuestionType(str, Enum):
    MULTIPLE_CHOICE = "multiple_choice"
    TRUE_FALSE = "true_false"
    ESSAY = "essay"


@dataclass(frozen=True)
class ParseIssue:
    message: str
    severity: str = "warning"
    question_id: str | None = None


@dataclass(frozen=True)
class AnswerOption:
    label: str
    block_index: int  # Relative to the question.
    label_start: int  # Character offset; labels may cross Word runs.
    label_end: int
    text: str
    colors: tuple[str, ...] = ()


@dataclass
class Question:
    id: str
    original_number: int
    section_id: str
    question_type: QuestionType
    start: int
    end: int
    blocks: list[Any]
    options: list[AnswerOption] = field(default_factory=list)
    option_layout: str = "none"
    option_shuffle_safe: bool = True
    score: str | None = None
    correct_answer: str | dict[str, bool] | None = None

    @property
    def labels(self) -> tuple[str, ...]:
        return tuple(option.label for option in self.options)

    @property
    def tables(self) -> list[Any]:
        from .document_blocks import W
        return [b for b in self.blocks if b.tag == W + "tbl"]

    @property
    def passages(self) -> list[Any]:
        stop = min((o.block_index for o in self.options), default=len(self.blocks))
        return self.blocks[1:stop]


@dataclass
class ExamSection:
    id: str
    title: str
    questions: list[Question] = field(default_factory=list)
    heading_index: int | None = None


@dataclass
class ParsedExam:
    source_bytes: bytes
    filename: str
    blocks: list[Any]
    sections: list[ExamSection]
    issues: list[ParseIssue]
    header_blocks: list[Any]
    footer_blocks: list[Any]

    @property
    def questions(self) -> list[Question]:
        return [q for section in self.sections for q in section.questions]

    @property
    def errors(self) -> list[ParseIssue]:
        return [issue for issue in self.issues if issue.severity == "error"]


@dataclass(frozen=True)
class ShuffleConfig:
    versions: int = 4
    starting_code: int = 101
    shuffle_questions: bool = True
    shuffle_options: bool = True
    shuffle_statements: bool = False
    shuffle_essays: bool = False
    seed: int = 42


@dataclass(frozen=True)
class PlannedQuestion:
    question_id: str
    section_id: str
    number: int
    option_order: tuple[str, ...]  # Original labels in their new order.


@dataclass(frozen=True)
class ExamVersion:
    code: int
    sections: dict[str, tuple[PlannedQuestion, ...]]

    @property
    def question_signature(self) -> tuple[tuple[str, ...], ...]:
        return tuple(tuple(q.question_id for q in qs) for qs in self.sections.values())


@dataclass(frozen=True)
class AnswerKey:
    answers: dict[str, str | dict[str, bool]]
    confirmed: bool = False


@dataclass(frozen=True)
class MappedAnswer:
    section_id: str
    number: int
    question_id: str
    answer: str | dict[str, bool]


@dataclass
class ExportBundle:
    files: dict[str, bytes]
    zip_bytes: bytes
    versions: list[ExamVersion]
    validation_messages: list[str]
