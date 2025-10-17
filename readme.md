# Certificate Generator - User Guide

## Overview

This script generates personalized certificates for participants and automatically sends them via email. The configuration is handled through JSON files, making it easy to use without any coding knowledge.

## Prerequisites

Before running the script, ensure you have the following installed on your system:

- Python 3.x
- Required Python libraries (install using the command below)
- A properly formatted CSV file with participant details

## Installation

### 1. Download and Install Python

If you don’t have Python installed, download and install it from [python.org](https://www.python.org/downloads/).

### 2. Create and Activate a Virtual Environment

Creating an isolated Python environment keeps the project dependencies separate from your global Python installation.

```
python -m venv .venv
```

Activate the environment before installing packages:

- **Windows (PowerShell):**
  ```
  .\.venv\Scripts\Activate.ps1
  ```
- **macOS/Linux (bash/zsh):**
  ```
  source .venv/bin/activate
  ```

### 3. Install Required Libraries

With the virtual environment active, install the dependencies:

```
pip install pandas fpdf2
```

### 4. Folder Structure

Ensure your project folder contains the following structure:

```
project_folder/
├── config/
│   ├── smtp_config.json
│   ├── email_config.json
│   ├── debug_mode.json
│   ├── content_config.json
├── example_config/
│   ├── smtp_config.json
│   ├── email_config.json
│   ├── debug_mode.json
│   ├── content_config.json
├── fonts/
│   ├── Lato2OFL/
│   │   ├── Lato-Black.ttf
├── background/
│   ├── fancyone.jpg
├── certificates/ (This will be auto-created when running the script)
├── participants.csv
├── script.py
```

## Using the Example Configuration

The `example_config` folder provides a ready-to-use starting point for all required configuration files. If you are setting up the project for the first time, copy these files into the `config` directory so the script can find them:

```
copy example_config\*.json config\
```

- Update the copied files with your own SMTP credentials, email content, and certificate layout.
- The example files include safe defaults (e.g., `debug_mode` is set to `"T"`), so you can run a dry run before sending real emails.
- The script always reads files from the `config` directory; keeping `example_config` untouched makes it easy to reset or compare your custom settings later.

## Configuration

The script is fully configurable through JSON files located in the `config` folder.

### 1. SMTP Configuration (`config/smtp_config.json`)

This file contains SMTP server settings to send emails.

```json
{
  "smtp_server": "smtp.email.com",
  "smtp_port": "587",
  "email_sender": "name@domain.com",
  "email_password": "your_password"
}
```

**Important:** Replace `your_password` with the actual email password.

### 2. Email Content Configuration (`config/email_config.json`)

Customize the subject and body of the email. (use html body)

```json
{
  "subject": "Your Participation Certificate",
  "body": "<h2>Hello {name},</h2><p>Your <strong>participation certificate</strong> is attached.</p><p>Best regards!</p>"
}
```

The `{name}` placeholder will be replaced with the participant’s name.

### 3. Debug Mode (`config/debug_mode.json`)

Enables or disables email sending for testing purposes.

```json
{
  "debug_mode": "T"
}
```

- **"T"** (Test Mode) – Certificates will be generated but not sent.
- **"F"** (Full Mode) – Certificates will be generated and sent via email.

### 4. Certificate Design (`config/content_config.json`)

Controls the certificate appearance.

```json
{
    "font_path": "./fonts/Lato2OFL/Lato-Black.ttf",
    "font_size": 32,
    "text_height": 160,
    "background_image": "./background/fancyone.jpg",
    "orientation": "L"
}
```

- `font_path`: Path to the font used.
- `font_size`: Size of the text on the certificate.
- `text_height`: Position of the text on the certificate.
- `background_image`: Background image for the certificate.
- `orientation`: "L" for landscape, "P" for portrait.

## CSV File Format

The script reads participant data from a CSV file (`participants.csv`). The file must follow this format:

```
FirstName,LastName,Email
John,Doe,john.doe@example.com
Jane,Smith,jane.smith@example.com
```

## Running the Script

1. Open a terminal or command prompt.
2. Navigate to the project directory:
   
   ```
   cd /path/to/project_folder
   ```

3. Run the script:
   
   ```
   python generator.py
   ```

## Troubleshooting

### Common Issues

- **SMTP Authentication Error:**
  - Ensure the email and password in `smtp_config.json` are correct.
  - Some email providers require an app password or less secure app access enabled.
- **Missing Fonts or Background Image:**
  - Check if the font file and background image exist in the specified paths.
- **Certificates Not Generated:**
  - Ensure the CSV file is formatted correctly.
- **Emails Not Sent:**
  - If `debug_mode.json` is set to "T", change it to "F" and retry.
