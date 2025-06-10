# ------------------------------------------------------------------------------------
# PYTHON SCRIPT FOR QUALTRICS DATA, "ALL-IN-ONE" LLM ANALYSIS, AND REPORTING
# ------------------------------------------------------------------------------------

# --- 1. IMPORT PACKAGES ---
import os
import sys
import time
import requests
import pandas as pd
import json
from datetime import datetime
from io import BytesIO, StringIO
from zipfile import ZipFile
from dotenv import load_dotenv
from tqdm import tqdm
import pypandoc

# Imports for Email Functionality
import smtplib
import ssl
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# --- 2. QUALTRICS API CONFIGURATION AND DATA FETCHING ---
# Load environment variables from .env file
load_dotenv()
api_key = os.getenv("QUALTRICS_API_KEY")
base_url = os.getenv("QUALTRICS_BASE_URL")
if not api_key or not base_url:
    print("Qualtrics credentials not found in .env file. Halting.", file=sys.stderr)
    sys.exit()
headers = {"X-API-TOKEN": api_key}
api_v3_url = f"https://{base_url.replace('https://', '')}/API/v3/"

def fetch_survey(survey_id: str, use_labels: bool = True):
    """ Fetches survey data and the question text map from the Qualtrics export. """
    print(f"Starting export for Survey ID: {survey_id}...")
    export_payload = {"format": "csv", "useLabels": use_labels}
    start_url = api_v3_url + f"surveys/{survey_id}/export-responses"
    try:
        start_response = requests.post(start_url, headers=headers, json=export_payload)
        start_response.raise_for_status()
        progress_id = start_response.json()["result"]["progressId"]
    except requests.exceptions.RequestException as e:
        print(f"Error starting export: {e}", file=sys.stderr)
        return None, None
    progress_url = f"{start_url}/{progress_id}"
    while True:
        progress_response = requests.get(progress_url, headers=headers)
        progress_response.raise_for_status()
        progress_status = progress_response.json()["result"]
        if progress_status["status"] == "complete":
            file_id = progress_status["fileId"]
            break
        time.sleep(2)
    print("Export complete. Downloading file...")
    download_url = f"{start_url}/{file_id}/file"
    download_response = requests.get(download_url, headers=headers, stream=True)
    download_response.raise_for_status()
    try:
        with ZipFile(BytesIO(download_response.content)) as zf:
            csv_filename = zf.namelist()[0]
            with zf.open(csv_filename) as f:
                csv_content = f.read().decode('utf-8')
                data_df = pd.read_csv(StringIO(csv_content), header=0, skiprows=[1, 2])
                question_text_df = pd.read_csv(StringIO(csv_content), header=0, nrows=1)
                question_map = question_text_df.iloc[0].to_dict()
                return data_df, question_map
    except Exception as e:
        print(f"Failed to unzip or process the survey data: {e}", file=sys.stderr)
        return None, None

# --- Script Execution Starts Here ---
QUALTRICS_SURVEY_ID = "SV_3xuiPokGeKBWmc6"
survey_data, question_map = fetch_survey(survey_id=QUALTRICS_SURVEY_ID)
if survey_data is None: sys.exit("Failed to fetch survey data.")
print("Successfully extracted question text map.")

# --- 3. SCRIPT CONFIGURATION ---
LOCAL_LLM_API_URL = "http://127.0.0.1:1234/v1/chat/completions"
LOCAL_LLM_MODEL_NAME = "Meta-Llama-3-8B-Instruct-Q5_K_M.gguf"
# ** USER ACTION: Define which questions from your survey to analyze. **
QUESTIONS_TO_ANALYZE = ["Q1", "Q2", "Q3"]
ID_COLUMN_NAME = "ResponseId"
# ** USER ACTION: Set the number of raw data rows to show in the appendix. **
NUM_ROWS_TO_DISPLAY_IN_REPORT = 25


def call_llm(system_prompt, user_prompt, request_timeout=300, retry_attempts=2):
    """Generic function to call the LLM."""
    request_body = { "model": LOCAL_LLM_MODEL_NAME, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_prompt}], "temperature": 0.1, "stream": False }
    for _ in range(retry_attempts):
        try:
            response = requests.post(LOCAL_LLM_API_URL, json=request_body, timeout=request_timeout, headers={"Content-Type": "application/json"})
            response.raise_for_status()
            return response.json()['choices'][0]['message']['content'].strip()
        except requests.exceptions.RequestException as e:
            time.sleep(5)
    return f"Error: Failed to get LLM response after {retry_attempts} attempts."

# --- 4. "ALL-IN-ONE" THEMATIC ANALYSIS AND REPORTING ---
print("\n--- Starting Thematic Analysis ---")
all_narrative_summaries = []
report_sections = []

system_prompt_analyst = """You are an expert qualitative research analyst. You are precise, data-driven, and follow formatting instructions exactly. Your language is concise and accessible. You do not invent or paraphrase content not present in the provided data."""
analysis_prompt_template = """Your task is to analyze a collection of workshop "exit ticket" responses for a single question and generate a structured report section.

**INPUT DATA:**
- **Question Asked:** "{question_text}"
- **Total Number of Responses:** {response_count}
- **Raw Responses:**
{response_list}

**YOUR TASK:**
1.  **Identify Key Themes:** Read all responses to identify 3-5 key themes. Capture both common and uncommon ideas.
2.  **Analyze Themes:** For each theme, you must:
    a. Create a concise theme name.
    b. Write a 1-2 sentence description summarizing the theme.
    c. Select one or two **verbatim** responses that best exemplify the theme.
    d. Count the exact number of responses that mention the theme (Frequency).
    e. Calculate the percentage of total responses this represents (Relative Frequency), rounded to a whole number.
3.  **Summarize Findings:** Write a 3-5 sentence narrative summary of the key takeaways, patterns, and any standout responses.
4.  **Format Output:** Assemble your entire analysis into the format below **EXACTLY**. Do not add any other conversational text or explanations.

**OUTPUT FORMAT:**
### {question_text}

**Summary of Responses**

[Your 3-5 sentence narrative summary goes here.]

**Thematic Table**

| Theme | Description | Illustrative Example(s) | Frequency | Relative Frequency |
|---|---|---|---|---|
| [Theme Name 1] | [Theme Description 1] | - "[Verbatim quote 1]" | [Frequency 1] | [Relative Frequency 1]% |
| [Theme Name 2] | [Theme Description 2] | - "[Verbatim quote 2]"<br>- "[Verbatim quote 3]" | [Frequency 2] | [Relative Frequency 2]% |
"""

for text_column in QUESTIONS_TO_ANALYZE:
    if text_column not in survey_data.columns:
        print(f"Warning: Column '{text_column}' not found. Skipping.", file=sys.stderr)
        continue

    full_question_text = question_map.get(text_column, text_column)
    print(f"\nAnalyzing responses for: \"{full_question_text}\"")

    responses = survey_data[text_column].dropna().astype(str).str.strip().tolist()
    responses = [res for res in responses if res]
    
    if not responses:
        report_sections.append(f"### {full_question_text}\n\n*No responses were submitted for this question.*")
        continue

    response_list_str = "\n".join([f"{i+1}. {res}" for i, res in enumerate(responses)])
    final_prompt = analysis_prompt_template.format(question_text=full_question_text, response_count=len(responses), response_list=response_list_str)
    
    analysis_section = call_llm(system_prompt_analyst, final_prompt)
    
    try:
        summary_text = analysis_section.split("**Summary of Responses**")[1].split("**Thematic Table**")[0].strip()
        all_narrative_summaries.append(summary_text)
    except IndexError:
        print(f"Warning: Could not parse narrative summary for '{text_column}'.")

    report_sections.append(analysis_section)

# --- 5. ASSEMBLE, SAVE, AND SEND FINAL REPORT ---
# Create the Executive Summary
executive_summary = ""
if all_narrative_summaries:
    print("\nAsking LLM to generate the Executive Summary...")
    summary_system_prompt = "You are a senior research director. Synthesize the following individual analyses from a workshop feedback report into a single, cohesive summary paragraph (3-5 sentences) highlighting the most important cross-cutting takeaways."
    summary_user_prompt = "Please synthesize these points into a single paragraph:\n\n---\n\n" + "\n\n".join(all_narrative_summaries)
    executive_summary = call_llm(summary_system_prompt, summary_user_prompt)

# Assemble the final report body
report_content = ["# Daily Feedback Thematic Analysis"]
if executive_summary and not executive_summary.startswith("Error:"):
    report_content.extend(["### Executive Summary", executive_summary, "---"])
report_content.extend(report_sections)
methodology_content = [
    "---", "### Methodology", f"- **Report Generated On:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
    f"- **Qualtrics Survey ID:** {QUALTRICS_SURVEY_ID}", f"- **Analysis Model:** `{LOCAL_LLM_MODEL_NAME}`",
    "- **Process:** This report was generated by providing all responses for a given question to an LLM for a full thematic analysis in a single step."
]
report_content.extend(methodology_content)

# MODIFICATION: Add the raw data appendix
appendix_content = []
columns_to_display_in_appendix = [ID_COLUMN_NAME] + QUESTIONS_TO_ANALYZE
# Ensure all columns exist before trying to select them
columns_to_display_in_appendix = [col for col in columns_to_display_in_appendix if col in survey_data.columns]

if columns_to_display_in_appendix:
    data_appendix_df = survey_data[columns_to_display_in_appendix]
    total_rows = len(data_appendix_df)
    rows_to_show = min(total_rows, NUM_ROWS_TO_DISPLAY_IN_REPORT)
    
    appendix_content.append("---")
    appendix_content.append("### Appendix: Raw Data Sample")
    appendix_content.append(f"*Showing the first {rows_to_show} of {total_rows} total responses.*")
    appendix_content.append(data_appendix_df.head(rows_to_show).to_markdown(index=False))

report_content.extend(appendix_content)
final_markdown = "\n\n".join(report_content)

# Save the final report as Markdown and convert to DOCX
md_filename = "daily_feedback_report.md"
docx_filename = "daily_feedback_report.docx"
with open(md_filename, "w", encoding="utf-8") as f: f.write(final_markdown)
print(f"\nSUCCESS: Markdown report has been saved to '{md_filename}'")

try:
    print("Converting report to DOCX...")
    pypandoc.convert_file(md_filename, 'docx', outputfile=docx_filename)
    print(f"SUCCESS: DOCX report has been saved to '{docx_filename}'")
except Exception as e:
    print(f"ERROR: Could not convert to DOCX. The Markdown file was saved. Error: {e}", file=sys.stderr)
    print("Please ensure Pandoc is installed and in your system's PATH.", file=sys.stderr)
    docx_filename = None

# Send the email with the DOCX attachment
if docx_filename:
    print("\n--- Attempting to Send Report via Email ---")
    facilitator_emails = ["jrosenb8@utk.edu"] # Replace with your facilitator emails
    sender_email = os.getenv("GMAIL_SENDER")
    password = os.getenv("GMAIL_APP_PASSWORD")

    if not sender_email or not password:
        print("\nWARNING: Gmail credentials not found in .env file. Skipping email.", file=sys.stderr)
    else:
        message = MIMEMultipart("alternative")
        message["Subject"] = f"Daily Feedback Report: {datetime.now().strftime('%Y-%m-%d')}"
        message["From"] = sender_email
        message["To"] = ", ".join(facilitator_emails)
        message.attach(MIMEText("Hi team, here's today's feedback!", "plain"))
        try:
            with open(docx_filename, "rb") as f:
                attachment = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.wordprocessingml.document")
                attachment.add_header('Content-Disposition', 'attachment', filename=docx_filename)
                message.attach(attachment)
            context = ssl.create_default_context()
            print(f"Connecting to Gmail SMTP server to send email to: {', '.join(facilitator_emails)}...")
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.send_message(message)
            print("\nSUCCESS: The report has been sent successfully via email.")
        except Exception as e:
            print(f"\nERROR: Failed to send email. Check .env credentials, Gmail settings, and file path. {e}", file=sys.stderr)

print("\nScript finished.")