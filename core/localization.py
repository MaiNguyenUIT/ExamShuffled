"""Vietnamese presentation of stable internal document diagnostics.

Source exam text is never translated. English core diagnostics stay stable for
API callers; only the application presentation is localized.
"""
from __future__ import annotations

import re

from .models import QuestionType

QUESTION_TYPE_LABELS = {
    QuestionType.MULTIPLE_CHOICE: "Trắc nghiệm nhiều lựa chọn",
    QuestionType.TRUE_FALSE: "Trắc nghiệm đúng/sai",
    QuestionType.ESSAY: "Tự luận",
}

_MESSAGES = {
    "DOCX exceeds the 30 MiB upload limit.": "Tệp DOCX vượt quá giới hạn tải lên 30 MiB.",
    "DOCX expands beyond the supported package limits.": "Dung lượng hoặc số thành phần sau khi giải nén của tệp DOCX vượt quá giới hạn hỗ trợ.",
    "DOCX contains duplicate ZIP entries.": "Tệp DOCX chứa các thành phần ZIP bị trùng tên.",
    "DTD declarations are unsupported in DOCX XML.": "Không hỗ trợ khai báo DTD trong XML của tệp DOCX.",
    "Unsupported Word XML namespace or missing document body.": "Định dạng XML của Word chưa được hỗ trợ hoặc tệp thiếu phần nội dung chính.",
    "Cannot read this DOCX. Upload an unencrypted Word .docx file.": "Không thể đọc tệp DOCX này. Vui lòng tải lên tệp Word .docx không được mã hóa.",
    "Invalid text replacement span.": "Vùng văn bản cần cập nhật không hợp lệ.",
    "A label overlaps a tab or break; cannot safely relabel it.": "Nhãn câu hỏi hoặc phương án chồng lên ký tự tab hay ngắt dòng; không thể đổi nhãn an toàn.",
    "Could not locate the complete question/option label.": "Không xác định được đầy đủ nhãn câu hỏi hoặc phương án.",
    "Inline choices contain complex XML; disable option shuffling.": "Các phương án cùng đoạn chứa XML phức tạp. Hãy tắt trộn phương án.",
    "Conflicting question/answer patterns; fix or simplify this question.": "Cấu trúc câu hỏi và phương án không thống nhất. Hãy chỉnh sửa hoặc đơn giản hóa câu hỏi này.",
    "Answer labels occur in a table. Table content is preserved; table-choice parsing is unsupported.": "Phát hiện nhãn phương án trong bảng. Nội dung bảng được giữ nguyên, nhưng chưa hỗ trợ phân tích phương án trong bảng.",
    "A table follows an answer label; disable option shuffling to retain its position.": "Có bảng nằm sau nhãn phương án. Hãy tắt trộn phương án để giữ vị trí của bảng.",
    "Inline choices contain fields, links, drawings or complex XML. Disable option shuffling.": "Các phương án cùng đoạn chứa trường Word, liên kết, hình vẽ hoặc XML phức tạp. Hãy tắt trộn phương án.",
    "Choices span a mixed inline layout. Disable option shuffling; whole questions can still move.": "Các phương án có bố cục đan xen trên nhiều đoạn. Hãy tắt trộn phương án; vẫn có thể trộn nguyên câu hỏi.",
    "An option appears to refer to another option or to 'all of the above'. Disable option shuffling and review.": "Một phương án có thể tham chiếu đến phương án khác hoặc có nội dung 'tất cả các phương án trên'. Hãy tắt trộn phương án và kiểm tra lại.",
    "python-docx could not reopen this Word package.": "Không thể mở lại tệp Word này bằng thư viện python-docx.",
    "No section heading precedes the first question; an implicit section was created.": "Không tìm thấy tiêu đề phần trước câu hỏi đầu tiên. Ứng dụng đã tạo một phần tạm; hãy kiểm tra cách chia phần.",
    "Question starts inside a table are unsupported; convert the question label to a body paragraph.": "Chưa hỗ trợ câu hỏi bắt đầu trong bảng. Hãy chuyển nhãn câu hỏi ra một đoạn văn bên ngoài bảng.",
    "No questions were detected. Expected body paragraphs beginning with 'Câu 1:' or 'Câu 1.'.": "Không tìm thấy câu hỏi. Câu hỏi cần bắt đầu ở đoạn văn chính bằng 'Câu 1:' hoặc 'Câu 1.'.",
    "No ending marker was found. Review the final question: trailing body instructions may belong to it.": "Không tìm thấy dấu kết thúc đề. Hãy kiểm tra câu cuối: hướng dẫn ở cuối tài liệu có thể đang được gộp vào câu này.",
    "A section break is inside this question. Move the break to a section heading before generating.": "Có ngắt phần (section break) trong câu hỏi này. Hãy chuyển ngắt phần đến tiêu đề phần trước khi tạo đề.",
    "Automatic Word numbering occurs inside this question; labels may change when moved. Use literal labels.": "Câu hỏi dùng đánh số tự động của Word nên nhãn có thể đổi khi di chuyển. Hãy nhập nhãn câu hỏi và phương án bằng văn bản thông thường.",
    "Digitally signed documents cannot be edited while retaining a valid signature.": "Không thể chỉnh sửa tài liệu đã ký số mà vẫn giữ chữ ký hợp lệ.",
    "Word fields are retained. Cross-references and displayed page values may require refresh in a Word viewer.": "Các trường Word được giữ nguyên. Tham chiếu chéo và số trang hiển thị có thể cần được cập nhật khi mở trong trình đọc Word.",
    "Some essay questions have no detected parenthesized score; review their classification and content.": "Một số câu tự luận chưa có điểm trong ngoặc được nhận diện. Hãy kiểm tra loại câu hỏi và nội dung.",
    "Parsing has blocking errors. Correct the document before generating exams.": "Phân tích tài liệu có lỗi cần xử lý. Hãy sửa tài liệu trước khi tạo đề.",
    "Review and acknowledge the parsing warnings before generating exams.": "Hãy kiểm tra và xác nhận các cảnh báo phân tích trước khi tạo đề.",
    "Number of exam versions must be between 1 and 100.": "Số lượng mã đề phải nằm trong khoảng từ 1 đến 100.",
    "Exam codes must be between 0 and 999999.": "Mã đề phải nằm trong khoảng từ 0 đến 999999.",
    "The random seed must be an integer.": "Giá trị khởi tạo ngẫu nhiên phải là số nguyên.",
    "Answer data must be reviewed and explicitly confirmed.": "Dữ liệu đáp án phải được kiểm tra và xác nhận rõ ràng.",
    "Essay answers are not supported by this answer key editor.": "Trình nhập đáp án này chưa hỗ trợ đáp án tự luận.",
    "Unsupported option layout.": "Bố cục phương án chưa được hỗ trợ.",
    "Question anchor is missing.": "Không tìm thấy nhãn bắt đầu câu hỏi.",
    "Cannot export a document with blocking parsing errors.": "Không thể xuất tài liệu khi còn lỗi phân tích cần xử lý.",
    "The requested batch exceeds the 256 MiB memory budget. Generate fewer versions at a time.": "Lượt tạo đề này vượt quá giới hạn bộ nhớ dự kiến 256 MiB. Hãy giảm số mã đề trong mỗi lượt.",
    "Some question orders repeat: the enabled question shuffles have fewer possible permutations than requested versions.": "Một số thứ tự câu hỏi bị trùng vì số hoán vị có thể tạo theo thiết lập hiện tại nhỏ hơn số mã đề yêu cầu.",
    "Generated section set differs from the original.": "Các phần trong đề đã tạo không khớp với đề gốc.",
    "Question moved between sections.": "Có câu hỏi bị chuyển sang phần khác.",
    "Generated exam question set differs from the original.": "Tập hợp câu hỏi trong đề đã tạo không khớp với đề gốc.",
    "Incorrect multiple-choice answer remapping.": "Ánh xạ đáp án trắc nghiệm sau khi trộn không chính xác.",
    "Incorrect true/false answer remapping.": "Ánh xạ đáp án đúng/sai sau khi trộn không chính xác.",
    "Question sets, section membership, numbering and option permutations validated.": "Đã kiểm tra tập hợp câu hỏi, vị trí trong từng phần, số thứ tự và hoán vị phương án.",
    "Question block count changed during export.": "Số khối nội dung của câu hỏi bị thay đổi khi xuất.",
    "Question block types changed during export.": "Loại khối nội dung của câu hỏi bị thay đổi khi xuất.",
    "Question anchor was lost during export.": "Nhãn bắt đầu câu hỏi bị mất khi xuất.",
    "Question content or formatting changed beyond its number/label.": "Nội dung hoặc định dạng câu hỏi bị thay đổi ngoài phần số thứ tự hoặc nhãn phương án.",
    "A question paragraph, passage, table or object changed unexpectedly.": "Đoạn văn, tư liệu, bảng hoặc đối tượng của câu hỏi bị thay đổi ngoài dự kiến.",
    "Package entries changed; a relationship or embedded part may be missing.": "Các thành phần trong tệp DOCX bị thay đổi; có thể thiếu liên kết hoặc nội dung nhúng.",
    "A package part other than document.xml changed.": "Một thành phần ngoài document.xml bị thay đổi.",
    "Page size, margins or section properties changed.": "Khổ giấy, lề trang hoặc thuộc tính phần bị thay đổi.",
    "Exported DOCX could not be reopened.": "Không thể mở lại tệp DOCX đã xuất.",
    "Exported section count changed.": "Số phần trong đề đã xuất bị thay đổi.",
    "School information, headings, instructions, ending or exam code changed unexpectedly.": "Thông tin trường, tiêu đề, hướng dẫn, phần kết thúc hoặc mã đề bị thay đổi ngoài dự kiến.",
    "Exported question count changed.": "Số câu hỏi trong đề đã xuất bị thay đổi.",
    "Question number or essay score changed incorrectly.": "Số thứ tự câu hỏi hoặc điểm tự luận bị thay đổi không chính xác.",
    "All other package parts, media, relationships, headers/footers and section properties unchanged.": "Các thành phần khác của DOCX, nội dung đa phương tiện, liên kết, đầu trang/chân trang và thuộc tính phần được giữ nguyên.",
    "Inline question block count changed.": "Số khối nội dung của câu hỏi có phương án cùng đoạn bị thay đổi.",
    "Inline stem content changed.": "Nội dung dẫn của câu hỏi có phương án cùng đoạn bị thay đổi.",
    "Inline choice content or style changed.": "Nội dung hoặc định dạng phương án cùng đoạn bị thay đổi.",
    "Inline choice spacing changed.": "Khoảng cách giữa các phương án cùng đoạn bị thay đổi.",
    "Inline paragraph formatting changed.": "Định dạng đoạn chứa các phương án bị thay đổi.",
}

_TEMPLATES = [
    (r"Expected labels (.+) exactly once in order; found (.+)\.",
     "Cần có đủ các nhãn {0}, mỗi nhãn xuất hiện một lần theo đúng thứ tự; đã tìm thấy {1}."),
    (r"A (.+) range spans answer groups\. Disable option shuffling\.",
     "Vùng đánh dấu {0} kéo dài qua nhiều phương án. Hãy tắt trộn phương án."),
    (r"A (.+) range crosses a question boundary; remove or localize the range before generating\.",
     "Vùng đánh dấu {0} kéo dài qua nhiều câu hỏi. Hãy xóa hoặc giới hạn vùng này trong một câu trước khi tạo đề."),
    (r"Unsupported body block (.+); content is retained, but question boundaries cannot be verified\.",
     "Khối nội dung {0} chưa được hỗ trợ. Nội dung được giữ nguyên nhưng không thể xác minh ranh giới câu hỏi."),
    (r"Section '(.+)' contains no detected questions\.", "Không tìm thấy câu hỏi trong phần '{0}'."),
    (r"Duplicate original question numbers in (.+)\.", "Số thứ tự câu hỏi trong phần {0} bị trùng."),
    (r"Answer key references unknown question (.+)\.", "Bảng đáp án tham chiếu đến câu hỏi không tồn tại: {0}."),
    (r"(.+): correct answer must be one of (.+)\.", "{0}: đáp án đúng phải là một trong các nhãn {1}."),
    (r"(.+): provide True/False for every statement\.", "{0}: hãy nhập Đúng/Sai cho tất cả các nhận định."),
    (r"(.+): this option layout cannot be shuffled safely\.", "{0}: không thể trộn an toàn các phương án có bố cục này."),
    (r"(.+): disable option/statement shuffling for this document; its layout is unsupported\.",
     "{0}: hãy tắt trộn phương án/nhận định vì bố cục này chưa được hỗ trợ."),
    (r"(.+): missing, duplicated or misplaced questions\.", "{0}: có câu hỏi bị thiếu, trùng hoặc nằm sai phần."),
    (r"(.+): question numbers are not sequential\.", "{0}: số thứ tự câu hỏi không liên tiếp."),
    (r"(.+): missing or duplicated answer options\.", "{0}: có phương án bị thiếu hoặc trùng."),
    (r"Dangling document relationship (.+)\.", "Liên kết {0} trong tài liệu không có thành phần đích hợp lệ."),
    (r"Exam (\d+): reopened; question content, passages, tables, scores and styles validated\.",
     "Mã đề {0}: đã mở lại thành công; nội dung câu hỏi, tư liệu, bảng, điểm và định dạng đã được kiểm tra."),
]
_PATTERNS = [(re.compile(pattern), translation) for pattern, translation in _TEMPLATES]
_FEATURES = {
    "tracked insertions": "nội dung chèn có theo dõi thay đổi",
    "tracked deletions": "nội dung xóa có theo dõi thay đổi",
    "tracked moves": "nội dung di chuyển có theo dõi thay đổi",
    "text boxes": "hộp văn bản",
    "embedded alternate content": "nội dung thay thế được nhúng",
}


def section_title(title: str) -> str:
    return "Phần chưa có tiêu đề" if title == "Unlabelled section" else title


def translate_diagnostic(message: str) -> str:
    if message in _MESSAGES:
        return _MESSAGES[message]
    prefix = "Exported DOCX has invalid question structure: "
    if message.startswith(prefix):
        return "Cấu trúc câu hỏi trong tệp DOCX đã xuất không hợp lệ: " + translate_diagnostic(message[len(prefix):])
    for feature, translation in _FEATURES.items():
        if message == f"Document contains {feature}; accept revisions or simplify these structures before generating.":
            return f"Tài liệu chứa {translation}. Hãy chấp nhận các thay đổi hoặc đơn giản hóa cấu trúc trước khi tạo đề."
    for pattern, translation in _PATTERNS:
        match = pattern.fullmatch(message)
        if match:
            return translation.format(*match.groups())
    return message
