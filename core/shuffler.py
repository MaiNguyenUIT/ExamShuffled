"""Produce reproducible plans; source OOXML is never mutated here."""
from __future__ import annotations

import math
import random

from .document_blocks import DocumentError
from .models import ExamVersion, ParsedExam, PlannedQuestion, QuestionType, ShuffleConfig


def _unrank(items: list, rank: int) -> list:
    remaining = list(items)
    result = []
    for size in range(len(remaining), 0, -1):
        index, rank = divmod(rank, math.factorial(size - 1))
        result.append(remaining.pop(index))
    return result


def _distinct_ranks(rng: random.Random, capacity: int, count: int) -> list[int]:
    # Sparse Fisher-Yates works even when capacity exceeds sys.maxsize. Avoids
    # factorial enumeration and retry loops as the permutation space fills up.
    swaps: dict[int, int] = {}
    ranks = []
    for offset in range(min(count, capacity)):
        last = capacity - offset - 1
        position = rng.randrange(last + 1)
        ranks.append(swaps.get(position, position))
        swaps[position] = swaps.get(last, last)
    while len(ranks) < count:
        ranks.append(rng.randrange(capacity))
    return ranks


def create_versions(exam: ParsedExam, config: ShuffleConfig, *, reviewed: bool = False) -> list[ExamVersion]:
    if exam.errors:
        raise DocumentError("Parsing has blocking errors. Correct the document before generating exams.")
    if exam.issues and not reviewed:
        raise DocumentError("Review and acknowledge the parsing warnings before generating exams.")
    if not 1 <= config.versions <= 100:
        raise DocumentError("Number of exam versions must be between 1 and 100.")
    if not 0 <= config.starting_code <= 999999 or config.starting_code + config.versions - 1 > 999999:
        raise DocumentError("Exam codes must be between 0 and 999999.")
    if not isinstance(config.seed, int):
        raise DocumentError("The random seed must be an integer.")
    for question in exam.questions:
        enabled = config.shuffle_options if question.question_type == QuestionType.MULTIPLE_CHOICE else (
            config.shuffle_statements if question.question_type == QuestionType.TRUE_FALSE else False)
        if enabled and not question.option_shuffle_safe:
            raise DocumentError(f"{question.id}: disable option/statement shuffling for this document; its layout is unsupported.")
    rng = random.Random(config.seed)
    movable = {}
    capacities = {}
    for section in exam.sections:
        movable[section.id] = [q for q in section.questions if config.shuffle_questions and (
            q.question_type != QuestionType.ESSAY or config.shuffle_essays)]
        capacities[section.id] = math.factorial(len(movable[section.id]))
    capacity = math.prod(capacities.values())
    ranks = _distinct_ranks(rng, capacity, config.versions)
    versions = []
    for version_index, rank in enumerate(ranks):
        planned_sections = {}
        for section in exam.sections:
            rank, local_rank = divmod(rank, capacities[section.id])
            eligible = movable[section.id]
            eligible_ids = {q.id for q in eligible}
            shuffled = iter(_unrank(eligible, local_rank))
            ordered = [next(shuffled) if q.id in eligible_ids else q for q in section.questions]
            planned = []
            for number, question in enumerate(ordered, 1):
                labels = list(question.labels)
                shuffle = (question.question_type == QuestionType.MULTIPLE_CHOICE and config.shuffle_options or
                           question.question_type == QuestionType.TRUE_FALSE and config.shuffle_statements)
                if shuffle:
                    rng.shuffle(labels)
                planned.append(PlannedQuestion(question.id, section.id, number, tuple(labels)))
            planned_sections[section.id] = tuple(planned)
        versions.append(ExamVersion(config.starting_code + version_index, planned_sections))
    return versions
