from email.mime.application import MIMEApplication

import pandas as pd
import smtplib
import os
import json
import logging

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
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

# Function to load and merge multiple configuration files
def load_config(config_files):
    config = {}
    for config_file in config_files:
        with open(config_file, "r", encoding="utf-8") as file:
            current_config = json.load(file)
            # Merge the current config with the main config
            config.update(current_config)
    return config

def calculate_text_center(pdf, text, page_width):
    """Calculate the X position to center text."""
    text_width = pdf.get_string_width(text)
    x_position = (page_width - text_width) / 2  # Center horizontally
    return x_position

def generate_certificate(name, surname, email):

    if content_config["orientation"] == "L":
        pdf = FPDF(orientation="L")
        page_width, page_height = 297, 210  # A4 Landscape
    else:
        pdf = FPDF()
        page_width, page_height = 210, 297  # A4 Portrait

    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    background_image = content_config["background_image"]
    pdf.image(background_image, x=0, y=0, w=page_width, h=page_height)

    pdf.add_font("MyFont", "", content_config["font_path"], uni=True)
    pdf.set_font("MyFont", "", content_config["font_size"])

    full_name = f"{name} {surname}"
    name_x = calculate_text_center(pdf, full_name, page_width)
    pdf.set_x(name_x)  # Centered horizontally
    pdf.cell(pdf.get_string_width(full_name),  content_config["text_height"], full_name)
    filename = f"./certificates/{name}_{surname}.pdf"
    os.makedirs("certificates", exist_ok=True)

    pdf.output(filename)
    logging.info(f"generated certificate for {email}")
    return filename


def prepare_email_content(receiver_email, name, attachment_path):
    # Create the MIMEMultipart message
    msg = MIMEMultipart()
    msg["From"] = email_sender
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
def process_csv(file_path):
    data = pd.read_csv(file_path)
    for index, row in data.iterrows():
        name, surname, receiver_email = row["FirstName"], row["LastName"], row["Email"]
        pdf_path = generate_certificate(name, surname, receiver_email)
        if debug_mode["debug_mode"] != "F":
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

    smtp_server = smtp_config["smtp_server"]
    smtp_port = smtp_config["smtp_port"]
    email_sender = smtp_config["email_sender"]
    email_password = smtp_config["email_password"]

    csv_file = "participants.csv"
    process_csv(csv_file)

