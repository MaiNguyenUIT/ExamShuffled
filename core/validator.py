"""Validation independent of the generation plan's rendering implementation."""
from __future__ import annotations

from collections import Counter
from copy import deepcopy
from io import BytesIO

from docx import Document
from lxml import etree

from .document_blocks import DocumentError, R, W, block_text, canonical, read_package, replace_text_span
from .models import AnswerKey, ExamVersion, ParsedExam, Question, QuestionType
from .question_detector import QUESTION_RE


def validate_version(exam: ParsedExam, version: ExamVersion, key: AnswerKey | None = None) -> list[str]:
    sections = {section.id: section for section in exam.sections}
    if set(version.sections) != set(sections):
        raise DocumentError("Generated section set differs from the original.")
    all_ids = []
    for section_id, planned in version.sections.items():
        by_id = {q.id: q for q in sections[section_id].questions}
        ids = [q.question_id for q in planned]
        if Counter(ids) != Counter(by_id.keys()):
            raise DocumentError(f"{section_id}: missing, duplicated or misplaced questions.")
        if [q.number for q in planned] != list(range(1, len(planned) + 1)):
            raise DocumentError(f"{section_id}: question numbers are not sequential.")
        for item in planned:
            question = by_id[item.question_id]
            if item.section_id != section_id:
                raise DocumentError("Question moved between sections.")
            if Counter(item.option_order) != Counter(question.labels):
                raise DocumentError(f"{question.id}: missing or duplicated answer options.")
        all_ids.extend(ids)
    if Counter(all_ids) != Counter(q.id for q in exam.questions):
        raise DocumentError("Generated exam question set differs from the original.")
    if key is not None:
        from .answer_key import remap_answers
        for mapped in remap_answers(exam, version, key):
            question = next(q for q in exam.questions if q.id == mapped.question_id)
            item = next(p for p in version.sections[mapped.section_id] if p.question_id == mapped.question_id)
            answer = key.answers[mapped.question_id]
            if isinstance(answer, str):
                original = item.option_order[question.labels.index(mapped.answer)]
                if original != answer:
                    raise DocumentError("Incorrect multiple-choice answer remapping.")
            elif any(mapped.answer[new] != answer[old] for new, old in zip(question.labels, item.option_order)):
                raise DocumentError("Incorrect true/false answer remapping.")
    return ["Question sets, section membership, numbering and option permutations validated."]


def _normalized_paragraph(paragraph, span=None):
    clone = deepcopy(paragraph)
    if span:
        replace_text_span(clone, *span, "#")
    # A shorter number split across runs can leave empty w:t nodes. Normalize
    # textual run segmentation while preserving the sequence of text/style pairs.
    tokens = []
    properties = clone.find(W + "pPr")
    if properties is not None:
        tokens.append(("pPr", canonical(properties)))
    for child in clone:
        if child.tag == W + "pPr":
            continue
        if child.tag != W + "r":
            tokens.append(("xml", canonical(child)))
            continue
        props = child.find(W + "rPr")
        style = canonical(props) if props is not None else b""
        for node in child:
            if node.tag == W + "rPr":
                continue
            if node.tag == W + "t":
                text = node.text or ""
                if not text:
                    continue
                if tokens and tokens[-1][0] == "text" and tokens[-1][1] == style:
                    previous = tokens.pop()
                    tokens.append(("text", style, previous[2] + text))
                else:
                    tokens.append(("text", style, text))
            else:
                tokens.append(("xml", style, canonical(node)))
    return tokens


def _question_units(question: Question):
    """Separate stem and choices for content comparisons after moving choices."""
    options = question.options
    if question.option_layout == "blocks":
        starts = [o.block_index for o in options]
        stem = question.blocks[:starts[0]]
        units = [question.blocks[start:end] for start, end in zip(starts, starts[1:] + [len(question.blocks)])]
        return stem, units
    return question.blocks, []


def _compare_blocks(original, generated, *, number=False, label=False):
    if len(original) != len(generated):
        raise DocumentError("Question block count changed during export.")
    for index, (before, after) in enumerate(zip(original, generated)):
        if before.tag != after.tag:
            raise DocumentError("Question block types changed during export.")
        if index == 0 and (number or label):
            if number:
                before_match = QUESTION_RE.match(block_text(before))
                after_match = QUESTION_RE.match(block_text(after))
                if not before_match or not after_match:
                    raise DocumentError("Question anchor was lost during export.")
                before_span, after_span = before_match.span('number'), after_match.span('number')
            else:
                before_text, after_text = block_text(before), block_text(after)
                before_start = len(before_text) - len(before_text.lstrip())
                after_start = len(after_text) - len(after_text.lstrip())
                before_span, after_span = (before_start, before_start + 1), (after_start, after_start + 1)
            if _normalized_paragraph(before, before_span) != _normalized_paragraph(after, after_span):
                raise DocumentError("Question content or formatting changed beyond its number/label.")
        elif canonical(before) != canonical(after):
            raise DocumentError("A question paragraph, passage, table or object changed unexpectedly.")


def validate_docx(exam: ParsedExam, version: ExamVersion, data: bytes) -> list[str]:
    from .parser import parse_docx
    source_parts, source_root = read_package(exam.source_bytes)
    parts, root = read_package(data)
    if set(parts) != set(source_parts):
        raise DocumentError("Package entries changed; a relationship or embedded part may be missing.")
    if any(parts[name] != value for name, value in source_parts.items() if name != "word/document.xml"):
        raise DocumentError("A package part other than document.xml changed.")
    if [canonical(n) for n in source_root.iter(W + "sectPr")] != [canonical(n) for n in root.iter(W + "sectPr")]:
        raise DocumentError("Page size, margins or section properties changed.")
    relationships = etree.fromstring(parts["word/_rels/document.xml.rels"])
    ids = {node.get("Id") for node in relationships}
    for node in root.iter():
        for attribute, value in node.attrib.items():
            if attribute.startswith(R) and value not in ids:
                raise DocumentError(f"Dangling document relationship {value}.")
    try:
        Document(BytesIO(data))
    except Exception as exc:
        raise DocumentError("Exported DOCX could not be reopened.") from exc
    parsed = parse_docx(data)
    if parsed.errors:
        raise DocumentError("Exported DOCX has invalid question structure: " + parsed.errors[0].message)
    if len(parsed.sections) != len(exam.sections):
        raise DocumentError("Exported section count changed.")
    original_question_blocks = {i for q in exam.questions for i in range(q.start, q.end)}
    generated_question_blocks = {i for q in parsed.questions for i in range(q.start, q.end)}
    fixed_before = [canonical(b) for i, b in enumerate(exam.blocks) if i not in original_question_blocks]
    fixed_after = [b for i, b in enumerate(parsed.blocks) if i not in generated_question_blocks]
    code_indices = [i for i, block in enumerate(fixed_after) if block.tag == W + "p" and
                    block_text(block) == f"Mã đề: {version.code}"]
    if not any([canonical(b) for i, b in enumerate(fixed_after) if i != code_index] == fixed_before
               for code_index in code_indices):
        raise DocumentError("School information, headings, instructions, ending or exam code changed unexpectedly.")
    by_id = {q.id: q for q in exam.questions}
    for original_section, generated_section in zip(exam.sections, parsed.sections):
        planned = version.sections[original_section.id]
        if len(planned) != len(generated_section.questions):
            raise DocumentError("Exported question count changed.")
        for item, generated in zip(planned, generated_section.questions):
            original = by_id[item.question_id]
            if generated.original_number != item.number or generated.score != original.score:
                raise DocumentError("Question number or essay score changed incorrectly.")
            if original.option_layout == "blocks":
                before_stem, before_options = _question_units(original)
                after_stem, after_options = _question_units(generated)
                _compare_blocks(before_stem, after_stem, number=True)
                for new_index, old_label in enumerate(item.option_order):
                    _compare_blocks(before_options[original.labels.index(old_label)], after_options[new_index], label=True)
            elif original.option_layout == "inline" and item.option_order != original.labels:
                _validate_inline(original, generated, item.option_order)
            else:
                _compare_blocks(original.blocks, generated.blocks, number=True)
    return [f"Exam {version.code}: reopened; question content, passages, tables, scores and styles validated.",
            "All other package parts, media, relationships, headers/footers and section properties unchanged."]


def _validate_inline(original, generated, order):
    from .document_blocks import inline_fragments
    index = original.options[0].block_index
    if len(original.blocks) != len(generated.blocks):
        raise DocumentError("Inline question block count changed.")
    # Compare surrounding blocks, renumbering the question anchor if separate.
    for i, (before, after) in enumerate(zip(original.blocks, generated.blocks)):
        if i != index:
            _compare_blocks([before], [after], number=(i == 0))
    before_p, after_p = deepcopy(original.blocks[index]), deepcopy(generated.blocks[index])
    if index == 0:
        for paragraph in (before_p, after_p):
            match = QUESTION_RE.match(block_text(paragraph))
            replace_text_span(paragraph, *match.span('number'), "#")
        # Recompute offsets after normalizing a possibly multi-digit number.
        from .question_detector import MC_RE, TF_RE
        pattern = MC_RE if original.question_type == QuestionType.MULTIPLE_CHOICE else TF_RE
        before_cuts = [m.start('label') for m in pattern.finditer(block_text(before_p))]
        after_cuts = [m.start('label') for m in pattern.finditer(block_text(after_p))]
    else:
        before_cuts = [o.label_start for o in original.options]
        after_cuts = [o.label_start for o in generated.options]
    before_stem, before_choices, before_spaces = inline_fragments(before_p, before_cuts)
    after_stem, after_choices, after_spaces = inline_fragments(after_p, after_cuts)
    def paragraph(nodes, normalize=False):
        p = etree.Element(W + "p", nsmap=before_p.nsmap)
        p.extend(deepcopy(nodes))
        return _normalized_paragraph(p, (0, 1) if normalize else None)
    if paragraph(before_stem) != paragraph(after_stem):
        raise DocumentError("Inline stem content changed.")
    for new_index, old_label in enumerate(order):
        if paragraph(before_choices[original.labels.index(old_label)], True) != paragraph(after_choices[new_index], True):
            raise DocumentError("Inline choice content or style changed.")
    if [paragraph(nodes) for nodes in before_spaces] != [paragraph(nodes) for nodes in after_spaces]:
        raise DocumentError("Inline choice spacing changed.")
    before_props, after_props = before_p.find(W + "pPr"), after_p.find(W + "pPr")
    if (canonical(before_props) if before_props is not None else None) != (canonical(after_props) if after_props is not None else None):
        raise DocumentError("Inline paragraph formatting changed.")
