import json
import logging

import pandas as pd
import pytest

import generator


@pytest.fixture
def tmp_config_files(tmp_path):
    """Provide helper paths for background/font assets."""
    background = tmp_path / "background.png"
    background.write_bytes(b"fake image content")
    font_file = tmp_path / "font.ttf"
    font_file.write_text("fake font data", encoding="utf-8")
    return background, font_file


def test_load_config_merges_files(tmp_path):
    first = tmp_path / "first.json"
    first.write_text(json.dumps({"a": 1, "shared": "initial"}), encoding="utf-8")
    second = tmp_path / "second.json"
    second.write_text(json.dumps({"b": 2, "shared": "override"}), encoding="utf-8")

    merged = generator.load_config([first, second])

    assert merged == {"a": 1, "shared": "override", "b": 2}


@pytest.mark.parametrize(
    "raw_value,expected_label,should_send",
    [
        ("Full", "Full", True),
        ("f", "Full", True),
        ("TRUE", "Full", True),
        ("Test", "Test", False),
        ("t", "Test", False),
        (False, "Test", False),
    ],
)
def test_resolve_debug_mode_accepts_variants(raw_value, expected_label, should_send):
    label, decision = generator.resolve_debug_mode({"debug_mode": raw_value})
    assert label == expected_label
    assert decision is should_send


def test_resolve_debug_mode_rejects_unknown_value():
    with pytest.raises(ValueError) as exc:
        generator.resolve_debug_mode({"debug_mode": "maybe"})

    assert "Unsupported debug_mode value" in str(exc.value)


def test_calculate_text_center():
    class DummyPDF:
        def get_string_width(self, text):
            return len(text)

    pdf = DummyPDF()
    position = generator.calculate_text_center(pdf, "abcd", page_width=100)
    assert position == pytest.approx(48)


def test_generate_certificate_creates_expected_pdf(monkeypatch, tmp_config_files):
    background, font_file = tmp_config_files

    monkeypatch.setattr(
        generator,
        "content_config",
        {
            "orientation": "L",
            "background_image": str(background),
            "font_path": str(font_file),
            "font_size": 32,
            "text_y": 107,
        },
        raising=False,
    )

    os_calls = {}

    def fake_makedirs(path, exist_ok):
        os_calls["path"] = path
        os_calls["exist_ok"] = exist_ok

    monkeypatch.setattr(generator.os, "makedirs", fake_makedirs)

    class DummyPDF:
        last_instance = None

        def __init__(self, orientation, unit, format):
            DummyPDF.last_instance = self
            self.orientation = orientation
            self.unit = unit
            self.format = format
            # Simulate A4 dimensions for orientation
            if orientation == "L":
                self.w, self.h = 297, 210
            else:
                self.w, self.h = 210, 297

        def set_auto_page_break(self, auto):
            self.auto_page_break = auto

        def set_margins(self, left, top, right):
            self.margins = (left, top, right)

        def add_page(self):
            self.pages = getattr(self, "pages", 0) + 1

        def get_string_width(self, text):
            return len(text)

        def image(self, path, x, y, w, h):
            self.image_args = (path, x, y, w, h)

        def add_font(self, name, style, path):
            self.font_args = (name, style, path)

        def set_font(self, name, style, size):
            self.set_font_args = (name, style, size)
            self.font_size = size

        def text(self, x, y, txt):
            self.text_args = (x, y, txt)

        def output(self, filename):
            self.output_file = filename
            return filename

    monkeypatch.setattr(generator, "FPDF", DummyPDF)

    result = generator.generate_certificate("Ada", "Lovelace", "ada@example.com")

    assert result == "./certificates/Ada_Lovelace.pdf"
    assert os_calls == {"path": "certificates", "exist_ok": True}
    pdf = DummyPDF.last_instance
    assert pdf is not None
    assert pdf.image_args[0] == str(background)
    assert pdf.font_args[2] == str(font_file)
    assert pdf.set_font_args == ("MyFont", "", 32)
    expected_x = generator.calculate_text_center(pdf, "Ada Lovelace", pdf.w)
    text_x, text_y, text_value = pdf.text_args
    assert text_x == pytest.approx(expected_x)
    assert text_y == pytest.approx(107)
    assert text_value == "Ada Lovelace"
    assert pdf.output_file == "./certificates/Ada_Lovelace.pdf"


def test_generate_certificate_raises_when_background_missing(monkeypatch, tmp_path):
    font_file = tmp_path / "font.ttf"
    font_file.write_text("fake", encoding="utf-8")

    class MinimalPDF:
        def __init__(self, *_, **__):
            self.w, self.h = 297, 210

        def set_auto_page_break(self, auto):
            self.auto_page_break = auto

        def set_margins(self, left, top, right):
            self.margins = (left, top, right)

        def add_page(self):
            self.page_added = True

        def image(self, *args, **kwargs):
            pass

    monkeypatch.setattr(
        generator,
        "content_config",
        {
            "orientation": "P",
            "background_image": str(tmp_path / "missing.png"),
            "font_path": str(font_file),
            "font_size": 20,
            "text_y": 100,
        },
        raising=False,
    )
    monkeypatch.setattr(generator, "FPDF", MinimalPDF)

    with pytest.raises(FileNotFoundError):
        generator.generate_certificate("Ada", "Lovelace", "ada@example.com")


def test_generate_certificate_raises_when_font_missing(monkeypatch, tmp_path):
    background = tmp_path / "background.png"
    background.write_bytes(b"fake image")

    class MinimalPDF:
        def __init__(self, *_, **__):
            self.w, self.h = 297, 210

        def set_auto_page_break(self, auto):
            self.auto_page_break = auto

        def set_margins(self, left, top, right):
            self.margins = (left, top, right)

        def add_page(self):
            self.page_added = True

        def image(self, *args, **kwargs):
            pass

    monkeypatch.setattr(
        generator,
        "content_config",
        {
            "orientation": "P",
            "background_image": str(background),
            "font_path": str(tmp_path / "missing.ttf"),
            "font_size": 20,
            "text_y": 100,
        },
        raising=False,
    )
    monkeypatch.setattr(generator, "FPDF", MinimalPDF)

    with pytest.raises(FileNotFoundError):
        generator.generate_certificate("Ada", "Lovelace", "ada@example.com")


def test_prepare_email_content_builds_message(monkeypatch, tmp_path):
    attachment = tmp_path / "Ada_Lovelace.pdf"
    attachment.write_bytes(b"pdf-content")

    monkeypatch.setattr(
        generator,
        "email_config",
        {"subject": "Certificate", "body": "<p>Hello {name}</p>"},
        raising=False,
    )
    monkeypatch.setattr(generator, "email_sender", "sender@example.com", raising=False)

    msg = generator.prepare_email_content("ada@example.com", "Ada", str(attachment))

    assert msg["From"] == "sender@example.com"
    assert msg["To"] == "ada@example.com"
    assert msg["Subject"] == "Certificate"
    payloads = msg.get_payload()
    assert len(payloads) == 2
    assert payloads[0].get_payload() == "<p>Hello Ada</p>"
    assert payloads[1].get_filename() == "Ada_Lovelace.pdf"


def test_send_email_uses_smtp(monkeypatch):
    class DummySMTP:
        def __init__(self, host, port):
            self.host = host
            self.port = port
            self.starttls_called = False
            self.login_args = None
            self.sent = []
            self.quit_called = False

        def starttls(self):
            self.starttls_called = True

        def login(self, email, password):
            self.login_args = (email, password)

        def sendmail(self, sender, receiver, message):
            self.sent.append((sender, receiver, message))

        def quit(self):
            self.quit_called = True

    smtp_instance = {}

    def smtp_factory(host, port):
        smtp_instance["client"] = DummySMTP(host, port)
        return smtp_instance["client"]

    monkeypatch.setattr(generator.smtplib, "SMTP", smtp_factory)
    monkeypatch.setattr(generator, "smtp_server", "smtp.example.com", raising=False)
    monkeypatch.setattr(generator, "smtp_port", 587, raising=False)
    monkeypatch.setattr(generator, "email_sender", "sender@example.com", raising=False)
    monkeypatch.setattr(generator, "email_password", "secret", raising=False)

    class DummyMessage:
        def as_string(self):
            return "raw message"

    generator.send_email("ada@example.com", DummyMessage())

    client = smtp_instance["client"]
    assert client.host == "smtp.example.com"
    assert client.port == 587
    assert client.starttls_called
    assert client.login_args == ("sender@example.com", "secret")
    assert client.sent == [("sender@example.com", "ada@example.com", "raw message")]
    assert client.quit_called


def test_send_email_logs_errors(caplog, monkeypatch):
    def failing_smtp(*_, **__):
        raise RuntimeError("SMTP offline")

    monkeypatch.setattr(generator.smtplib, "SMTP", failing_smtp)
    monkeypatch.setattr(generator, "smtp_server", "smtp.example.com", raising=False)
    monkeypatch.setattr(generator, "smtp_port", 25, raising=False)
    monkeypatch.setattr(generator, "email_sender", "sender@example.com", raising=False)
    monkeypatch.setattr(generator, "email_password", "secret", raising=False)

    class DummyMessage:
        def as_string(self):
            return "raw message"

    with caplog.at_level(logging.ERROR):
        generator.send_email("ada@example.com", DummyMessage())

    assert "Error when sending email to: ada@example.com" in caplog.text


def build_participant_csv(tmp_path, rows):
    csv_path = tmp_path / "participants.csv"
    if rows:
        df = pd.DataFrame(rows)
    else:
        df = pd.DataFrame(columns=["FirstName", "LastName", "Email"])
    df.to_csv(csv_path, index=False)
    return csv_path


def test_process_csv_handles_empty_file(tmp_path, caplog, monkeypatch):
    csv_path = build_participant_csv(tmp_path, [])

    calls = {"generate": 0}

    def track_generate(*args, **kwargs):
        calls["generate"] += 1
        return None

    monkeypatch.setattr(generator, "generate_certificate", track_generate)
    monkeypatch.setattr(generator, "prepare_email_content", lambda *args, **kwargs: None)
    monkeypatch.setattr(generator, "send_email", lambda *args, **kwargs: None)

    with caplog.at_level(logging.WARNING):
        generator.process_csv(csv_path, "Test", False)

    assert calls["generate"] == 0
    assert f"No participants found in {csv_path}" in caplog.text


def test_process_csv_generates_without_sending(tmp_path, capsys, monkeypatch):
    csv_path = build_participant_csv(
        tmp_path,
        [
            {"FirstName": "Ada", "LastName": "Lovelace", "Email": "ada@example.com"},
            {"FirstName": "Grace", "LastName": "Hopper", "Email": "grace@example.com"},
        ],
    )

    generated = []

    def fake_generate(name, surname, email):
        generated.append((name, surname, email))
        return str(tmp_path / f"{name}_{surname}.pdf")

    monkeypatch.setattr(generator, "generate_certificate", fake_generate)
    monkeypatch.setattr(generator, "prepare_email_content", lambda *args, **kwargs: None)

    emails_sent = []
    monkeypatch.setattr(generator, "send_email", lambda *args, **kwargs: emails_sent.append(args))

    generator.process_csv(csv_path, "Test", False)

    assert generated == [
        ("Ada", "Lovelace", "ada@example.com"),
        ("Grace", "Hopper", "grace@example.com"),
    ]
    assert emails_sent == []

    out = capsys.readouterr().out
    assert "Debug Mode: Test" in out
    assert "Progress: 2/2" in out


def test_process_csv_sends_emails_when_enabled(tmp_path, monkeypatch):
    csv_path = build_participant_csv(
        tmp_path,
        [
            {"FirstName": "Ada", "LastName": "Lovelace", "Email": "ada@example.com"},
        ],
    )

    monkeypatch.setattr(generator, "generate_certificate", lambda *_, **__: "certificate.pdf")

    prepared_messages = []

    def fake_prepare(receiver_email, name, attachment_path):
        msg = f"message-for-{receiver_email}"
        prepared_messages.append((receiver_email, name, attachment_path, msg))
        return msg

    monkeypatch.setattr(generator, "prepare_email_content", fake_prepare)

    sent_emails = []
    monkeypatch.setattr(generator, "send_email", lambda receiver, msg: sent_emails.append((receiver, msg)))

    generator.process_csv(csv_path, "Full", True)

    assert prepared_messages == [
        ("ada@example.com", "Ada", "certificate.pdf", "message-for-ada@example.com")
    ]
    assert sent_emails == [("ada@example.com", "message-for-ada@example.com")]
