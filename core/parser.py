from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re

from docx import Document

from .document_blocks import (DocumentError, W, NS, block_text, colored_text,
                              read_package, split_paragraph_fragments)
from .models import AnswerOption, ExamSection, ParsedExam, ParseIssue, Question, QuestionType
from .question_detector import END_RE, MC_RE, QUESTION_RE, SCORE_RE, SECTION_RE, TF_RE, section_type


def _parse_options(question: Question, declared: QuestionType | None, issues: list[ParseIssue]) -> None:
    found = {QuestionType.MULTIPLE_CHOICE: [], QuestionType.TRUE_FALSE: []}
    table_labels = False
    for index, block in enumerate(question.blocks):
        if block.tag == W + "tbl":
            if any(MC_RE.match(block_text(p).lstrip()) or TF_RE.match(block_text(p).lstrip())
                   for p in block.iter(W + "p")):
                table_labels = True
            continue
        if block.tag != W + "p":
            continue
        text = block_text(block)
        for kind, pattern in ((QuestionType.MULTIPLE_CHOICE, MC_RE), (QuestionType.TRUE_FALSE, TF_RE)):
            matches = list(pattern.finditer(text))
            # Single inline labels in prose are not considered standalone options.
            if len(matches) == 1 and text[:matches[0].start()].strip():
                continue
            for j, match in enumerate(matches):
                end = matches[j + 1].start() if j + 1 < len(matches) else len(text)
                found[kind].append(AnswerOption(match['label'], index, match.start('label'),
                                               match.end('label'), text[match.end():end].strip(),
                                               colored_text(block)))
    inferred = [kind for kind, options in found.items() if options]
    kind = declared or (inferred[0] if len(inferred) == 1 else QuestionType.ESSAY)
    question.question_type = kind
    question.options = found.get(kind, [])
    if len(inferred) > 1 or (declared == QuestionType.ESSAY and inferred):
        issues.append(ParseIssue("Conflicting question/answer patterns; fix or simplify this question.", "error", question.id))
    if table_labels:
        question.option_shuffle_safe = False
        issues.append(ParseIssue("Answer labels occur in a table. Table content is preserved; table-choice parsing is unsupported.",
                                 "error", question.id))
    if kind == QuestionType.ESSAY:
        return
    expected = tuple("ABCD" if kind == QuestionType.MULTIPLE_CHOICE else "abcd")
    if question.labels != expected:
        issues.append(ParseIssue(f"Expected labels {', '.join(expected)} exactly once in order; found {question.labels}.",
                                 "error", question.id))
        question.option_shuffle_safe = False
        return
    indices = [o.block_index for o in question.options]
    if len(set(indices)) == 4 and all(not block_text(question.blocks[o.block_index])[:o.label_start].strip()
                                      for o in question.options):
        question.option_layout = "blocks"
        if any(b.tag == W + "tbl" for b in question.blocks[indices[0]:]):
            question.option_shuffle_safe = False
            issues.append(ParseIssue("A table follows an answer label; disable option shuffling to retain its position.", question_id=question.id))
    elif len(set(indices)) == 1:
        question.option_layout = "inline"
        paragraph = question.blocks[indices[0]]
        try:
            split_paragraph_fragments(paragraph, [o.label_start for o in question.options])
        except DocumentError:
            question.option_shuffle_safe = False
            issues.append(ParseIssue("Inline choices contain fields, links, drawings or complex XML. Disable option shuffling.", question_id=question.id))
    else:
        question.option_layout = "unsupported"
        question.option_shuffle_safe = False
        issues.append(ParseIssue("Choices span a mixed inline layout. Disable option shuffling; whole questions can still move.", question_id=question.id))
    # Paired XML ranges must not be broken by moving option groups separately.
    if question.option_layout == "blocks":
        starts = [o.block_index for o in question.options]
        for start_tag, end_tag in (("bookmarkStart", "bookmarkEnd"), ("commentRangeStart", "commentRangeEnd"),
                                   ("permStart", "permEnd")):
            owners = ({}, {})
            for index, block in enumerate(question.blocks):
                group = sum(index >= start for start in starts)
                for nodes, tag in zip(owners, (start_tag, end_tag)):
                    for node in block.iter(W + tag):
                        nodes[node.get(W + "id")] = group
            if any(owners[0].get(key) != owners[1].get(key) for key in owners[0].keys() | owners[1].keys()):
                question.option_shuffle_safe = False
                issues.append(ParseIssue(f"A {start_tag} range spans answer groups. Disable option shuffling.", question_id=question.id))
                break
    dependent = re.compile(r"(?:phương án|đáp án|mệnh đề|nhận định|ý)\s+[A-Da-d][.)]|(?:tất cả|cả)\s+(?:các\s+)?(?:đáp án|phương án|ý)\s+(?:trên|còn lại)", re.IGNORECASE)
    if any(dependent.search(option.text) for option in question.options):
        question.option_shuffle_safe = False
        issues.append(ParseIssue("An option appears to refer to another option or to 'all of the above'. Disable option shuffling and review.", question_id=question.id))


def parse_docx(source: bytes | str | Path, filename: str | None = None) -> ParsedExam:
    if not isinstance(source, bytes):
        path = Path(source)
        filename = filename or path.name
        source = path.read_bytes()
    parts, root = read_package(source)
    try:
        Document(BytesIO(source))
    except Exception as exc:
        raise DocumentError("python-docx could not reopen this Word package.") from exc
    blocks = list(root.find(W + "body"))
    issues: list[ParseIssue] = []
    sections: list[ExamSection] = []
    current_section = None
    current_question = None
    footer_start = len(blocks) - (1 if blocks and blocks[-1].tag == W + "sectPr" else 0)
    ended = False
    declared = {}
    for index, block in enumerate(blocks):
        text = block_text(block)
        is_paragraph = block.tag == W + "p"
        if is_paragraph and END_RE.match(text):
            if current_question:
                current_question.end = index
                current_question = None
            footer_start = index
            ended = True
        if ended or block.tag == W + "sectPr":
            continue
        if is_paragraph and SECTION_RE.match(text):
            if current_question:
                current_question.end = index
                current_question = None
            current_section = ExamSection(f"s{len(sections) + 1}", text.strip(), heading_index=index)
            sections.append(current_section)
            declared[current_section.id] = section_type(text)
            continue
        match = QUESTION_RE.match(text) if is_paragraph else None
        if match:
            if current_question:
                current_question.end = index
            if current_section is None:
                current_section = ExamSection("s1", "Unlabelled section")
                sections.append(current_section)
                declared[current_section.id] = None
                issues.append(ParseIssue("No section heading precedes the first question; an implicit section was created."))
            question_id = f"{current_section.id}-q{len(current_section.questions) + 1}"
            current_question = Question(question_id, int(match['number']), current_section.id,
                                        QuestionType.ESSAY, index, footer_start, [])
            current_section.questions.append(current_question)
        elif block.tag not in {W + "p", W + "tbl"}:
            issues.append(ParseIssue(f"Unsupported body block {block.tag.rsplit('}', 1)[-1]}; content is retained, but question boundaries cannot be verified.", "error"))
        if block.tag == W + "tbl" and any(QUESTION_RE.match(block_text(p)) for p in block.iter(W + "p")):
            issues.append(ParseIssue("Question starts inside a table are unsupported; convert the question label to a body paragraph.", "error"))
    questions = [q for section in sections for q in section.questions]
    if current_question:
        current_question.end = footer_start
    if not questions:
        issues.append(ParseIssue("No questions were detected. Expected body paragraphs beginning with 'Câu 1:' or 'Câu 1.'.", "error"))
    if questions and not ended:
        issues.append(ParseIssue("No ending marker was found. Review the final question: trailing body instructions may belong to it."))
    for section in sections:
        numbers = [q.original_number for q in section.questions]
        if not numbers:
            issues.append(ParseIssue(f"Section '{section.title}' contains no detected questions."))
        if len(set(numbers)) != len(numbers):
            issues.append(ParseIssue(f"Duplicate original question numbers in {section.id}.", "error"))
        for question in section.questions:
            question.blocks = blocks[question.start:question.end]
            score = SCORE_RE.search(block_text(question.blocks[0]))
            question.score = score[1] if score else None
            _parse_options(question, declared[section.id], issues)
            if any(b.find('.//' + W + 'sectPr') is not None for b in question.blocks):
                issues.append(ParseIssue("A section break is inside this question. Move the break to a section heading before generating.", "error", question.id))
            if any(b.xpath('.//w:numPr', namespaces=NS) for b in question.blocks):
                issues.append(ParseIssue("Automatic Word numbering occurs inside this question; labels may change when moved. Use literal labels.", "error", question.id))
    # Moves cannot safely separate paired range markers spanning question boundaries.
    ownership = {id(b): q.id for q in questions for b in q.blocks}
    for start_tag, end_tag in (("bookmarkStart", "bookmarkEnd"), ("commentRangeStart", "commentRangeEnd"),
                               ("permStart", "permEnd")):
        starts, ends = {}, {}
        for block in blocks:
            owner = ownership.get(id(block), "fixed")
            for node in block.iter(W + start_tag):
                starts[node.get(W + "id")] = owner
            for node in block.iter(W + end_tag):
                ends[node.get(W + "id")] = owner
        if any(starts.get(key) != ends.get(key) for key in starts.keys() | ends.keys()):
            issues.append(ParseIssue(f"A {start_tag} range crosses a question boundary; remove or localize the range before generating.", "error"))
    for tag, explanation in (("ins", "tracked insertions"), ("del", "tracked deletions"),
                              ("moveFrom", "tracked moves"), ("moveTo", "tracked moves"),
                              ("txbxContent", "text boxes"), ("altChunk", "embedded alternate content")):
        if root.find('.//' + W + tag) is not None:
            issues.append(ParseIssue(f"Document contains {explanation}; accept revisions or simplify these structures before generating.", "error"))
    if any(name.startswith("_xmlsignatures/") for name in parts):
        issues.append(ParseIssue("Digitally signed documents cannot be edited while retaining a valid signature.", "error"))
    if root.xpath('.//w:fldSimple | .//w:instrText', namespaces=NS):
        issues.append(ParseIssue("Word fields are retained. Cross-references and displayed page values may require refresh in a Word viewer."))
    if any(q.question_type == QuestionType.ESSAY and q.score is None for q in questions):
        issues.append(ParseIssue("Some essay questions have no detected parenthesized score; review their classification and content."))
    first = min((q.start for q in questions), default=len(blocks))
    return ParsedExam(source, filename or "exam.docx", blocks, sections, issues,
                      blocks[:first], blocks[footer_start:])
