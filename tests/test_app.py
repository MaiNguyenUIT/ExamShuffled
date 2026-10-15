from types import SimpleNamespace
from pathlib import Path

import streamlit as st
from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parents[1] / "app.py"


def checkbox(app, label):
    return next(widget for widget in app.checkbox if widget.label == label)


def test_empty_upload_screen():
    app = AppTest.from_file(APP).run()
    assert not app.exception
    assert app.title[0].value == "Exam Shuffler"
    assert "Tải đề thi lên" in app.info[0].value


def test_generation_downloads_and_stale_settings(monkeypatch, exam_bytes):
    upload = SimpleNamespace(name="exam.docx", getvalue=lambda: exam_bytes)
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: upload)
    app = AppTest.from_file(APP, default_timeout=20).run()
    assert not app.exception
    assert [metric.value for metric in app.metric] == ["3", "12", "4", "2"]
    assert app.button[0].disabled
    checkbox(app, "Tôi đã kiểm tra các câu hỏi được nhận diện và các cảnh báo").check().run()
    app.button[0].click().run()
    assert not app.exception
    assert len(app.get("download_button")) == 5  # ZIP plus four exams.
    assert len(app.session_state["bundle"].versions) == 4
    app.number_input[0].set_value(2).run()
    assert not app.exception
    assert len(app.get("download_button")) == 0
    assert any("Thiết lập hoặc đáp án đã thay đổi" in message.value for message in app.info)


def test_confirmed_answer_key_ui(monkeypatch, exam_bytes):
    upload = SimpleNamespace(name="exam.docx", getvalue=lambda: exam_bytes)
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: upload)
    app = AppTest.from_file(APP, default_timeout=20).run()
    checkbox(app, "Tôi đã kiểm tra các câu hỏi được nhận diện và các cảnh báo").check().run()
    checkbox(app, "Thêm bảng đáp án").check().run()
    assert len(app.selectbox) == 28
    app.selectbox[0].select("B").run()
    assert app.button[0].disabled
    checkbox(app, "Tôi đã kiểm tra và xác nhận các đáp án đã nhập").check().run()
    app.button[0].click().run()
    assert not app.exception
    assert "Answer_Key.docx" in app.session_state["bundle"].files
    assert len(app.get("download_button")) == 6
    app.selectbox[0].select("C").run()
    assert app.button[0].disabled
    assert not checkbox(app, "Tôi đã kiểm tra và xác nhận các đáp án đã nhập").value
    assert len(app.get("download_button")) == 0


def test_invalid_docx_and_seed(monkeypatch, exam_bytes):
    upload = SimpleNamespace(name="broken.docx", getvalue=lambda: b"broken")
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: upload)
    app = AppTest.from_file(APP).run()
    assert not app.exception
    assert "Không thể đọc" in app.error[0].value
    upload.getvalue = lambda: exam_bytes
    app.run()
    app.text_input[0].set_value("invalid").run()
    assert not app.exception
    assert any("số nguyên" in message.value for message in app.error)
    assert app.button[0].disabled
