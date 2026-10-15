from __future__ import annotations

from .document_blocks import DocumentError
from .models import AnswerKey, ExamVersion, MappedAnswer, ParsedExam, QuestionType


def color_candidates(exam: ParsedExam) -> dict[str, tuple[str, ...]]:
    """Suggestions only; no AnswerKey is created or confirmed automatically."""
    return {q.id: tuple(o.label for o in q.options if o.colors)
            for q in exam.questions if any(o.colors for o in q.options)}


def validate_answer_key(exam: ParsedExam, key: AnswerKey) -> None:
    if not key.confirmed:
        raise DocumentError("Answer data must be reviewed and explicitly confirmed.")
    by_id = {q.id: q for q in exam.questions}
    for question_id, answer in key.answers.items():
        if question_id not in by_id:
            raise DocumentError(f"Answer key references unknown question {question_id}.")
        question = by_id[question_id]
        if question.question_type == QuestionType.MULTIPLE_CHOICE:
            if not isinstance(answer, str) or answer not in question.labels:
                raise DocumentError(f"{question_id}: correct answer must be one of {question.labels}.")
        elif question.question_type == QuestionType.TRUE_FALSE:
            if (not isinstance(answer, dict) or set(answer) != set(question.labels) or
                    any(type(value) is not bool for value in answer.values())):
                raise DocumentError(f"{question_id}: provide True/False for every statement.")
        else:
            raise DocumentError("Essay answers are not supported by this answer key editor.")


def remap_answers(exam: ParsedExam, version: ExamVersion, key: AnswerKey) -> list[MappedAnswer]:
    validate_answer_key(exam, key)
    by_id = {q.id: q for q in exam.questions}
    result = []
    for section_id, planned in version.sections.items():
        for item in planned:
            if item.question_id not in key.answers:
                continue
            question = by_id[item.question_id]
            answer = key.answers[item.question_id]
            old_to_new = dict(zip(item.option_order, question.labels))
            if isinstance(answer, str):
                mapped = old_to_new[answer]
            else:
                mapped = {new: answer[old] for old, new in zip(item.option_order, question.labels)}
            result.append(MappedAnswer(section_id, item.number, item.question_id, mapped))
    return result
