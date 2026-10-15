"""Local Streamlit application. Run: python -m streamlit run app.py"""
from __future__ import annotations

import hashlib
import json
import logging
from dataclasses import asdict

import streamlit as st

from core.answer_key import color_candidates
from core.document_blocks import DocumentError, W, block_text
from core.exporter import generate_bundle
from core.localization import QUESTION_TYPE_LABELS, section_title, translate_diagnostic
from core.models import AnswerKey, ParsedExam, QuestionType, ShuffleConfig
from core.parser import parse_docx

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
logger = logging.getLogger(__name__)


def show_preview(exam: ParsedExam) -> None:
    st.subheader("2 · Phân tích và kiểm tra")
    columns = st.columns(4)
    columns[0].metric("Số phần", len(exam.sections))
    for column, kind, label in zip(columns[1:], QuestionType, ["Trắc nghiệm lựa chọn", "Đúng / Sai", "Tự luận"]):
        column.metric(label, sum(q.question_type == kind for q in exam.questions))
    for issue in exam.issues:
        (st.error if issue.severity == "error" else st.warning)(
            f"{issue.question_id + ': ' if issue.question_id else ''}{translate_diagnostic(issue.message)}")
    with st.expander("Thông tin trường và phần kết thúc đề"):
        st.text("\n".join(block_text(b) for b in exam.header_blocks if block_text(b).strip()))
        st.divider()
        st.text("\n".join(block_text(b) for b in exam.footer_blocks if block_text(b).strip()))
    for section in exam.sections:
        with st.expander(f"{section_title(section.title)} — {len(section.questions)} câu hỏi"):
            for question in section.questions:
                st.markdown(f"**Câu {question.original_number}** · {QUESTION_TYPE_LABELS[question.question_type]}")
                for block in question.blocks:
                    if block.tag == W + "tbl":
                        rows = [[block_text(cell) for cell in row.findall(W + "tc")]
                                for row in block.findall(W + "tr")]
                        st.table(rows)
                    elif block_text(block).strip():
                        st.text(block_text(block))
                st.caption(f"Mã câu: {question.id} · Số bảng: {len(question.tables)} · Điểm: {question.score or '—'}")
                st.divider()
    st.caption("Bản xem trước hiển thị văn bản và nội dung bảng. Hình ảnh, công thức và cách phân trang của Word cần được kiểm tra trong tệp DOCX tải về.")


def invalidate_confirmation(token: str) -> None:
    st.session_state[f"confirmed_{token}"] = False


def edit_answer_key(exam: ParsedExam, token: str) -> tuple[AnswerKey | None, list[str]]:
    answers = {}
    incomplete = []
    if not st.checkbox("Thêm bảng đáp án", key=f"key_enabled_{token}"):
        return None, incomplete
    st.caption("Nhập đáp án theo số câu và nhãn phương án trong ĐỀ GỐC. Câu để trống sẽ không được đưa vào bảng đáp án.")
    suggestions = color_candidates(exam)
    if st.checkbox("Xem gợi ý đáp án từ màu chữ", key=f"colors_{token}"):
        if suggestions:
            st.warning("Màu chữ chỉ là gợi ý chưa được xác minh. Hãy kiểm tra từng đáp án trước khi xác nhận.")
            st.write({question_id: ", ".join(labels) for question_id, labels in suggestions.items()})
            if st.button("Điền gợi ý cho câu có một phương án được tô màu", key=f"prefill_{token}"):
                for question in exam.questions:
                    labels = suggestions.get(question.id, ())
                    if question.question_type == QuestionType.MULTIPLE_CHOICE and len(labels) == 1:
                        st.session_state[f"answer_{token}_{question.id}"] = labels[0]
                st.session_state[f"confirmed_{token}"] = False
        else:
            st.info("Không tìm thấy phương án có màu chữ được đặt trực tiếp. Màu theo chủ đề có thể chỉ dùng để trang trí.")
    for section in exam.sections:
        questions = [q for q in section.questions if q.options]
        if not questions:
            continue
        with st.expander(f"Đáp án · {section_title(section.title)}"):
            for question in questions:
                if question.question_type == QuestionType.MULTIPLE_CHOICE:
                    value = st.selectbox(f"Câu {question.original_number}", ("", *question.labels),
                                         key=f"answer_{token}_{question.id}",
                                         on_change=invalidate_confirmation, args=(token,))
                    if value:
                        answers[question.id] = value
                else:
                    st.write(f"Câu {question.original_number}")
                    values = {}
                    for column, label in zip(st.columns(len(question.options)), question.labels):
                        choice = column.selectbox(f"{label})", ("", "Đúng", "Sai"),
                                                  key=f"answer_{token}_{question.id}_{label}",
                                                  on_change=invalidate_confirmation, args=(token,))
                        if choice:
                            values[label] = choice == "Đúng"
                    if values and len(values) != len(question.options):
                        incomplete.append(f"{question.id}: hãy nhập Đúng/Sai cho cả bốn nhận định hoặc để trống toàn bộ câu.")
                    elif values:
                        answers[question.id] = values
    for message in incomplete:
        st.warning(message)
    confirmed = st.checkbox("Tôi đã kiểm tra và xác nhận các đáp án đã nhập", key=f"confirmed_{token}")
    st.caption(f"Đã nhập đáp án cho {len(answers)}/{sum(bool(q.options) for q in exam.questions)} câu trắc nghiệm.")
    return AnswerKey(answers, confirmed), incomplete


def main() -> None:
    st.set_page_config(page_title="Exam Shuffler · Trộn đề thi", page_icon="📄", layout="wide")
    st.title("Exam Shuffler")
    st.caption("Trộn đề thi tiếng Việt · Xử lý trên máy của bạn · Giữ nguyên nội dung Word gốc")
    st.subheader("1 · Tải đề thi lên")
    uploaded = st.file_uploader("Chọn tệp đề thi Word (.docx)", type=["docx"], key="upload",
                                help="Hỗ trợ tệp .docx tối đa 30 MiB, bao gồm tên tệp và nội dung có dấu tiếng Việt.")
    if uploaded is None:
        st.info("Tải đề thi lên để phân tích các phần, câu hỏi và phương án trả lời.")
        return
    source = uploaded.getvalue()
    digest = hashlib.sha256(source).hexdigest()
    token = digest[:12]
    if st.session_state.get("source_digest") != digest:
        for key in list(st.session_state):
            if key in {"exam", "parse_error", "bundle", "generation_fingerprint"} or key.startswith(
                    ("answer_", "key_enabled_", "colors_", "prefill_", "confirmed_", "reviewed_")):
                del st.session_state[key]
        st.session_state.source_digest = digest
        try:
            st.session_state.exam = parse_docx(source, uploaded.name)
        except DocumentError as exc:
            st.session_state.parse_error = str(exc)
        except Exception:
            logger.exception("Unexpected DOCX analysis failure")
            st.session_state.parse_error = "Không thể phân tích tệp Word này. Hãy kiểm tra tệp .docx không được mã hóa và các câu hỏi nằm trong đoạn văn thông thường."
    if "parse_error" in st.session_state:
        st.error(translate_diagnostic(st.session_state.parse_error))
        return
    exam = st.session_state.exam
    st.success(f"Đã tải lên {uploaded.name} · {len(source):,} byte")
    show_preview(exam)
    reviewed = st.checkbox("Tôi đã kiểm tra các câu hỏi được nhận diện và các cảnh báo", key=f"reviewed_{token}")
    st.subheader("3 · Thiết lập trộn đề")
    columns = st.columns(3)
    versions = columns[0].number_input("Số lượng mã đề", min_value=1, max_value=100, value=4, step=1, key="version_count")
    start = columns[1].number_input("Mã đề bắt đầu", min_value=0, max_value=999900, value=101, step=1, key="starting_code")
    seed_text = columns[2].text_input("Giá trị khởi tạo ngẫu nhiên (seed)", value="42", key="random_seed",
                                    help="Nhập một số nguyên. Cùng đề gốc, thiết lập và giá trị này sẽ tạo lại cùng kết quả trộn.")
    columns = st.columns(2)
    shuffle_questions = columns[0].checkbox("Trộn thứ tự câu hỏi trong từng phần", value=True, key="shuffle_questions")
    shuffle_options = columns[1].checkbox("Trộn phương án trắc nghiệm A–D", value=True, key="shuffle_options")
    shuffle_statements = columns[0].checkbox("Trộn nhận định đúng/sai a–d", value=False, key="shuffle_statements")
    shuffle_essays = columns[1].checkbox("Trộn câu hỏi tự luận", value=False, disabled=not shuffle_questions, key="shuffle_essays")
    try:
        seed = int(seed_text.strip())
        config = ShuffleConfig(int(versions), int(start), shuffle_questions, shuffle_options,
                               shuffle_statements, shuffle_essays, seed)
    except ValueError:
        config = None
        st.error("Giá trị khởi tạo ngẫu nhiên phải là số nguyên.")
    with st.expander("Tùy chọn · Nhập và xác nhận đáp án"):
        key, incomplete = edit_answer_key(exam, token)
    fingerprint = hashlib.sha256(json.dumps({
        "language": "vi", "source": digest, "config": asdict(config) if config else None,
        "key": asdict(key) if key else None, "incomplete": incomplete,
    }, sort_keys=True).encode()).hexdigest()
    st.subheader("4 · Tạo đề thi")
    blocked = bool(exam.errors or not reviewed or config is None or incomplete or
                   (key and key.answers and not key.confirmed))
    if st.button("Tạo các mã đề", type="primary", disabled=blocked):
        progress = st.progress(0, text="Đang kiểm tra và tạo các tệp Word…")
        try:
            effective_key = key if key and key.answers else None
            bundle = generate_bundle(exam, config, effective_key, reviewed=reviewed,
                                     progress=lambda done, total: progress.progress(done / total,
                                         text=f"Đã tạo và kiểm tra {done}/{total} mã đề"))
            st.session_state.bundle = bundle
            st.session_state.generation_fingerprint = fingerprint
        except DocumentError as exc:
            st.error(translate_diagnostic(str(exc)))
        except Exception:
            logger.exception("Unexpected exam generation failure")
            st.error("Tạo đề không thành công. Chưa có tệp chưa hoàn chỉnh nào được cung cấp để tải về. Hãy kiểm tra đề gốc hoặc nhật ký ứng dụng trên máy.")
        finally:
            progress.empty()
    if not reviewed:
        st.info("Hãy xem lại các câu hỏi đã phân tích và đánh dấu xác nhận kiểm tra để bật chức năng tạo đề.")
    if key and key.answers and not key.confirmed:
        st.info("Hãy xác nhận các đáp án đã nhập để xuất kèm bảng đáp án.")
    if "bundle" not in st.session_state:
        return
    if st.session_state.get("generation_fingerprint") != fingerprint:
        st.info("Thiết lập hoặc đáp án đã thay đổi. Hãy tạo lại đề để cập nhật các tệp tải về.")
        return
    bundle = st.session_state.bundle
    st.success(f"Đã tạo và kiểm tra thành công {len(bundle.versions)} mã đề.")
    with st.expander("Kết quả kiểm tra"):
        for message in bundle.validation_messages:
            st.write("✓ " + translate_diagnostic(message))
    st.subheader("5 · Tải kết quả về")
    st.download_button("Tải tất cả tệp ZIP", bundle.zip_bytes, "Exam_Shuffler.zip", "application/zip", type="primary")
    columns = st.columns(3)
    for index, (name, data) in enumerate(bundle.files.items()):
        label = "Tải bảng đáp án" if name == "Answer_Key.docx" else f"Tải đề {name.removeprefix('Exam_').removesuffix('.docx')}"
        columns[index % 3].download_button(label, data, name, DOCX_MIME, key=f"download_{name}")
    st.caption("Đã kiểm tra định dạng và tính toàn vẹn của tệp. Cách phân trang có thể thay đổi sau khi trộn; hãy mở và xem lại từng đề trong trình đọc Word trước khi in.")


if __name__ == "__main__":
    main()
