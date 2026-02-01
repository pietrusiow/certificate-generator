from email.mime.application import MIMEApplication

import pandas as pd
import smtplib
import os
import json
import logging
import time

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from email import encoders
from fpdf import FPDF

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logging.log'),  # Info log file
    ]
)
# Silence verbose fontTools logs emitted during font subsetting.
logging.getLogger("fontTools").setLevel(logging.WARNING)

# Function to load and merge multiple configuration files
def load_config(config_files):
    config = {}
    for config_file in config_files:
        with open(config_file, "r", encoding="utf-8") as file:
            current_config = json.load(file)
            # Merge the current config with the main config
            config.update(current_config)
    return config


def resolve_debug_mode(debug_mode_config):
    """
    Normalize the debug mode setting so the caller can decide whether to send emails.
    Accepts legacy values like 'T'/'F' alongside the new 'Test'/'Full' strings.
    """
    raw_value = debug_mode_config.get("debug_mode")
    normalized = str(raw_value).strip().lower()

    mapping = {
        "full": ("Full", True),
        "f": ("Full", True),
        "true": ("Full", True),
        "test": ("Test", False),
        "t": ("Test", False),
        "false": ("Test", False),
    }

    if normalized in mapping:
        return mapping[normalized]

    raise ValueError(
        f"Unsupported debug_mode value: {raw_value}. Expected 'Test' or 'Full'."
    )

def calculate_text_center(pdf, text, page_width):
    """Calculate the X position to center text."""
    text_width = pdf.get_string_width(text)
    x_position = (page_width - text_width) / 2  # Center horizontally
    return x_position


def resolve_text_baseline(pdf, config):
    """
    Determine the baseline Y-coordinate (in mm) for the recipient name.
    Uses the explicit text_y value when present; otherwise falls back to the font height.
    """
    raw_baseline = config.get("text_y")
    if raw_baseline is not None:
        try:
            return float(raw_baseline)
        except (TypeError, ValueError):
            logging.warning("Invalid text_y value in content config: %s", raw_baseline)

    logging.warning("Falling back to baseline equal to font height; configure text_y for precise placement.")
    return pdf.font_size


def _safe_int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _safe_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _count_name_characters(full_name):
    """Length of the recipient name excluding whitespace."""
    return sum(1 for char in full_name if not char.isspace())


def should_split_full_name(full_name, config):
    """
    Determine whether the recipient name should be rendered across two lines.
    Falls back to a default threshold of 24 characters (excluding whitespace)
    when the configuration omits the split limit.
    """
    threshold_raw = config.get("split_name_threshold", 24)
    threshold = _safe_int(threshold_raw)
    if threshold is None:
        logging.warning(
            "Invalid split_name_threshold value '%s'; expected integer. Skipping multi-line split for '%s'.",
            threshold_raw,
            full_name,
        )
        return False

    return _count_name_characters(full_name) > threshold


def resolve_split_line_gap(pdf, config):
    """
    Decide how far apart (in mm) to place multi-line names.
    Uses the configured value if present; otherwise derives spacing from the font height.
    """
    gap_raw = config.get("split_name_line_gap")
    gap = _safe_float(gap_raw)
    if gap is not None:
        return gap
    return pdf.font_size * 0.85


def resolve_split_style(base_font_size, baseline_override, config):
    """
    Determine the font size and baseline override for split (multi-line) name rendering.
    Falls back to the already-resolved values when the configuration omits the overrides.
    """
    font_size = base_font_size
    raw_font_size = config.get("split_name_font_size")
    if raw_font_size is not None:
        override_font = _safe_float(raw_font_size)
        if override_font is None:
            logging.warning(
                "Invalid split_name_font_size value '%s'; using resolved font size %s.",
                raw_font_size,
                base_font_size,
            )
        else:
            font_size = override_font

    baseline = baseline_override
    raw_baseline = config.get("split_name_text_y")
    if raw_baseline is not None:
        override_baseline = _safe_float(raw_baseline)
        if override_baseline is None:
            logging.warning(
                "Invalid split_name_text_y value '%s'; using resolved baseline %s.",
                raw_baseline,
                baseline_override,
            )
        else:
            baseline = override_baseline

    return font_size, baseline


def resolve_name_style(full_name, config):
    """
    Compute the font size and optional baseline override for the given name.
    Falls back to the default font settings when the long-name configuration is missing or invalid.
    """
    base_font_size = _safe_float(config.get("font_size"))
    if base_font_size is None:
        raise ValueError("Invalid or missing font_size in content configuration.")

    threshold_raw = config.get("long_name_threshold")
    threshold = _safe_int(threshold_raw)
    if threshold_raw is not None and threshold is None:
        logging.warning("Invalid long_name_threshold value '%s'; expected integer.", threshold_raw)

    use_alternate = threshold is not None and _count_name_characters(full_name) > threshold

    if use_alternate:
        alt_font_size = _safe_float(config.get("long_name_font_size"))
        if alt_font_size is None:
            logging.warning(
                "long_name_font_size is missing or invalid; defaulting to primary font_size for '%s'.",
                full_name,
            )
            alt_font_size = base_font_size
        text_y = config.get("long_name_text_y", config.get("text_y"))
        return alt_font_size, text_y

    return base_font_size, config.get("text_y")


def parse_hex_color(color_code):
    if not color_code:
        return None

    value = str(color_code).strip()
    if value.startswith("#"):
        value = value[1:]

    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)

    if len(value) != 6:
        logging.warning("Invalid text_color value '%s'; expected #RRGGBB or #RGB.", color_code)
        return None

    try:
        r = int(value[0:2], 16)
        g = int(value[2:4], 16)
        b = int(value[4:6], 16)
        return r, g, b
    except ValueError:
        logging.warning("Invalid text_color value '%s'; expected hexadecimal digits.", color_code)
        return None


def apply_text_color(pdf, color_code):
    rgb = parse_hex_color(color_code)
    if rgb:
        pdf.set_text_color(*rgb)

def generate_certificate(name, surname, email):

    orientation = content_config.get("orientation", "L").upper()
    pdf = FPDF(orientation=orientation, unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    pdf.set_margins(0, 0, 0)
    pdf.add_page()
    page_width, page_height = pdf.w, pdf.h

    background_image = content_config["background_image"]
    if not os.path.exists(background_image):
        logging.error("Background image not found at %s while creating certificate for %s", background_image, email)
        raise FileNotFoundError(f"Background image not found: {background_image}")
    pdf.image(background_image, x=0, y=0, w=page_width, h=page_height)

    font_path = content_config["font_path"]
    if not os.path.exists(font_path):
        logging.error("Font file not found at %s while creating certificate for %s", font_path, email)
        raise FileNotFoundError(f"Font file not found: {font_path}")
    pdf.add_font("MyFont", "", font_path)

    full_name = f"{name} {surname}"
    first_line = name.strip()
    second_line = surname.strip()
    use_split = (
        should_split_full_name(full_name, content_config)
        and bool(first_line)
        and bool(second_line)
    )

    font_size_pt, text_y_override = resolve_name_style(full_name, content_config)
    if use_split:
        font_size_pt, text_y_override = resolve_split_style(font_size_pt, text_y_override, content_config)

    pdf.set_font("MyFont", "", font_size_pt)
    apply_text_color(pdf, content_config.get("text_color"))
    baseline_config = {"font_size": font_size_pt, "text_y": text_y_override}
    baseline_y = resolve_text_baseline(pdf, baseline_config)
    if use_split:
        gap = resolve_split_line_gap(pdf, content_config)
        first_line_y = baseline_y - gap
        if first_line:
            first_x = calculate_text_center(pdf, first_line, page_width)
            pdf.text(first_x, first_line_y, first_line)
        second_x = calculate_text_center(pdf, second_line, page_width)
        pdf.text(second_x, baseline_y, second_line)
    else:
        name_x = calculate_text_center(pdf, full_name, page_width)
        pdf.text(name_x, baseline_y, full_name)
    filename = f"./certificates/{name}_{surname}.pdf"
    os.makedirs("certificates", exist_ok=True)

    pdf.output(filename)
    logging.info(f"generated certificate for {email}")
    return filename


def prepare_email_content(receiver_email, name, attachment_path):
    # Create the MIMEMultipart message
    msg = MIMEMultipart()
    msg['From'] = formataddr(("Eletive", email_sender))
    msg["To"] = receiver_email
    msg["Subject"] = email_config["subject"]


    # Email body
    body = email_config["body"].format(name=name)
    msg.attach(MIMEText(body, "html"))

    filename = os.path.basename(attachment_path)  # Get the file name (e.g., "certificate.pdf")
    part = MIMEApplication(open(attachment_path, "rb").read())
    part.add_header('Content-Disposition', 'attachment', filename=filename)

    # Encode the file in base64 to send it over email
    encoders.encode_base64(part)
    msg.attach(part)
    return msg

def send_email(receiver_email, msg):
    try:
        # Establish a connection to the SMTP server and send the email
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        server.login(email_sender, email_password)  # Log into the SMTP server
        server.sendmail(email_sender, receiver_email, msg.as_string())  # Send the email
        server.quit()  # Logout and close the connection

        # Log the success
        logging.info(f"E-mail sent to: {receiver_email}")

    except Exception as e:
        # Log any errors
        logging.error(f"Error when sending email to: {receiver_email}: {e}")


# Load data from the CSV and generate certificates
def process_csv(file_path, debug_mode_label, should_send_email):
    data = pd.read_csv(file_path)
    total = len(data.index)

    if total == 0:
        logging.warning("No participants found in %s", file_path)
        return

    print(f"Debug Mode: {debug_mode_label}")

    for position, (_, row) in enumerate(data.iterrows(), start=1):
        name, surname, receiver_email = row["FirstName"], row["LastName"], row["Email"]
        pdf_path = generate_certificate(name, surname, receiver_email)
        logging.info(
            "Progress: %d/%d (%.1f%%) certificates prepared",
            position,
            total,
            (position / total) * 100,
        )
        print(f"Progress: {position}/{total} ({(position / total) * 100:.1f}%) certificates prepared")

        if should_send_email:
            print(f"Sending email to: {receiver_email}")
            msg = prepare_email_content(receiver_email, name, pdf_path)
            send_email(receiver_email, msg)

if __name__ == "__main__":
    content_config_files = ["./config/content_config.json"]
    email_config_files = ["./config/email_config.json"]
    smtp_config_files = ["./config/smtp_config.json"]
    debug_mode_files = ["./config/debug_mode.json"]

    content_config = load_config(content_config_files)
    email_config = load_config(email_config_files)
    smtp_config = load_config(smtp_config_files)
    debug_mode = load_config(debug_mode_files)
    debug_mode_label, should_send_email = resolve_debug_mode(debug_mode)

    smtp_server = smtp_config["smtp_server"]
    smtp_port = smtp_config["smtp_port"]
    email_sender = smtp_config["email_sender"]
    email_password = smtp_config["email_password"]

    csv_file = "participants.csv"
    process_csv(csv_file, debug_mode_label, should_send_email)

