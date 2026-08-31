# Vectorys Automated Planification

## Overview

**Vectorys Automated Planification** is an internship project designed to automate the processing of logistics orders received by email.

The application retrieves emails, analyzes their content using a local AI model, identifies whether they are related to an order, extracts relevant information, processes Excel attachments, and generates structured output for the logistics workflow and Sage WMS environment.

The preferred local AI model is **`qwen3:14b`**, running through Ollama.

## Main Features

- Retrieve and process emails.
- Analyze email content using a local LLM.
- Classify emails and identify order-related messages.
- Detect new orders, updates, cancellations, and other requests.
- Extract structured JSON from email content.
- Process Excel attachments.
- Transform extracted order data.
- Generate Sage WMS-related output files.
- Store processed orders locally.
- Maintain JSONL registers and audit data.
- Detect missing or ambiguous information.
- Flag cases requiring human intervention.
- Allow manual corrections and reprocessing through Streamlit.

## Technologies Used

- Python
- Streamlit
- Ollama
- Qwen3 (`qwen3:14b`)
- LangChain
- Pandas
- OpenPyXL
- IMAP / SMTP
- JSON / JSONL
- Python Dotenv

## Requirements

The project dependencies are listed in `requirements.txt`:

```text
langchain-core
langchain-ollama
python-dotenv
pandas
openpyxl
ollama
python-dateutil
streamlit
flask
```

## Installation

### 1. Create a Python virtual environment

```bash
python -m venv venv
```

### 2. Activate the environment

#### Windows

```bash
venv\Scripts\activate
```

#### Linux / macOS

```bash
source venv/bin/activate
```

### 3. Install the dependencies

```bash
pip install -r requirements.txt
```

## Ollama Installation

The application uses Ollama to run the AI model locally.

Install Ollama from the official website:

https://ollama.com/

After installation, verify that Ollama is available:

```bash
ollama --version
```

## Installing the AI Model

The preferred model is:

```text
qwen3:14b
```

Pull the model with:

```bash
ollama pull qwen3:14b
```

Verify that it is installed:

```bash
ollama list
```

You can also test it manually:

```bash
ollama run qwen3:14b
```

## Gmail Configuration

The application currently works **only with Gmail**.

A `.env` file is required to provide the Gmail account credentials used by the application.

### 1. Create the `.env` file

Create a file named `.env` in the root directory of the project:

```text
Vectorys Automated Planification/
├── .env
├── main.py
├── requirements.txt
├── README.md
└── ...
```

The `.env` file should contain:

```env
EMAIL=your_email@gmail.com
APP_PASSWORD=your_gmail_app_password
```

* `EMAIL` → The Gmail address used by the application.
* `APP_PASSWORD` → The **Google App Password** associated with this Gmail account.

### 2. Gmail App Password

The application should **not** use the normal Gmail account password.

Instead, a **Google App Password** should be generated for the Gmail account used by the application.

The Gmail account used for the application should ideally be a **dedicated company account**, rather than an employee's personal account.

For example:

```text
EMAIL=orders@company.com
APP_PASSWORD=xxxxxxxxxxxxxxxx
```

The exact account used depends on the company's email configuration.

### 3. Gmail Requirement

At the moment, the email ingestion system is designed specifically for **Gmail**.

Other email providers are **not currently supported** by the application.


## Configuration

The application configuration manages values such as:

- Working directory.
- Orders directory.
- Excel directory.
- Prompt template paths.
- Sage WMS template path.
- Processing delay.
- Automatic reply configuration.
- AI model.
- Processing mode.

Example:

```json
{
  "paths": {
    "mail_reader_template": "templates/mail_template.txt",
    "xls_reader_template": "templates/excel_template.txt",
    "sage_form_template.xlsm": "templates/sage_form_template.xlsm",
    "main_directory_path": "C:/..../Vectorys Planification",
    "orders_directory_path": "C:/..../Vectorys Planification/Orders",
    "excel_files_directory_path": "C:/..../Vectorys Planification/Excel Files"
  },
  "delay": 5,
  "auto_reply": false,
  "model": "qwen3:14b"
}
```


## Running the Application

After installing the dependencies and ensuring that Ollama and `qwen3:14b` are available, run:

```bash
streamlit run main.py
```

## General Workflow

```text
Incoming Email
      |
      v
Retrieve Email
      |
      v
Extract Metadata and Email Body
      |
      v
Check for Excel Attachment
      |
      v
AI Analysis with qwen3:14b
      |
      v
Structured JSON Result
      |
      +----------------------------+
      |                            |
      v                            v
Valid / Complete             Missing or Ambiguous
      |                            |
      v                            v
Process Excel               Flag for Human Review
      |                            |
      v                            v
Transform Order Data        Edit / Add Missing Data
      |                            |
      v                            |
Generate Sage Output        |
      |                            |
      v                            |
Store Order                 |
      |                            |
      +----------------------------+
                   |
                   v
             Reprocessing
```

## AI Processing

The local AI model analyzes email information and converts it into structured data.

The analysis can include:

- Whether the email is related to an order.
- The intent of the email.
- Client information.
- Company information.
- Order references.
- Requested actions.
- Missing information.
- Ambiguous instructions.
- Whether human intervention is required.

The AI is instructed to return structured JSON that can be processed by the Python application.

## Excel Processing

When an email contains an Excel attachment, the application can process the file and extract required order information.

This can include:

- Order header information.
- Dates.
- Client information.
- Product or article lines.
- Quantities.
- Other required order details.

The extracted data is transformed into the structure required by the Sage WMS workflow.

## Human Intervention

Some emails cannot safely be processed automatically.

Human intervention can be required for:

- Unknown clients.
- Unknown stockers.
- Missing information.
- Ambiguous instructions.
- Invalid or incomplete Excel files.
- Contradictory requests.
- AI output that requires verification.

An operator can then add missing master data, correct the order, or modify the required information before the order is processed again.

## Data Storage

The application stores processed information locally.

```text
vectorys planification/
├── excel_files/
├── orders/
│   └── order_{local_order_id}_{date}/
│       ├── mail_{local_mail_id}.json
│       ├── order_{local_order_id}.json
│       ├── date_{local_order_id}.json
│       ├── details_{local_order_id}.json
│       └── sage_form_{local_order_id}.xlsm
├── all_mails.jsonl
├── clients.jsonl
├── non_orders.jsonl
├── orders.jsonl
├── temp_mail.json
└── temp_order.json
```

## Processing Modes

### Fast Stop

Processes one unread email and then stops.

Useful for debugging, testing, and reviewing one email at a time.

### Slow Stop

Processes all available unread emails sequentially until the processing queue is empty.

Useful for batch processing.

## Recommended System Setup

Because the preferred model is `qwen3:14b`, the project benefits from sufficient GPU memory and RAM.

A setup with:

- **16 GB VRAM**
- **32 GB RAM**

is suitable for running the model locally alongside the application.

Actual performance depends on the operating system, Ollama configuration, and other running applications.

## Project Goal

The goal of **Vectorys Automated Planification** is to automate repetitive logistics order processing while maintaining human control over uncertain cases.

The project combines:

- Automated email retrieval.
- Local AI analysis.
- Structured JSON extraction.
- Excel processing.
- Sage WMS transformation.
- Local data storage.
- Human validation when necessary.

The objective is to automate repetitive work and direct human attention toward cases that genuinely require intervention.

## Notes

- Preferred AI model: `qwen3:14b`
- Ollama must be installed before running AI analysis.
- Python dependencies are installed through `requirements.txt`.
- The application is started with `streamlit run main.py`.
- Local paths and configuration values may need to be adjusted depending on the machine.
