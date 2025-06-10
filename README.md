# qualtrics-survey-local-llm-analysis

# Automated Qualtrics Report with LLM Thematic Analysis

This project automates the entire workflow of generating a professional, qualitative report from survey data. It fetches data directly from the Qualtrics API, uses a locally-run Large Language Model (LLM) via LM Studio to perform a sophisticated thematic analysis, and generates a `.docx` report complete with narrative summaries, tables, and an executive summary, which is then automatically emailed to a list of facilitators.

## Key Features

-   **Direct Qualtrics Integration**: Fetches survey data on-demand using the Qualtrics API.
-   **Local & Private LLM Analysis**: Performs all AI-powered analysis locally using LM Studio, ensuring data privacy.
-   **"All-in-One" Thematic Analysis**: For each survey question, it sends all responses to the LLM at once for a holistic analysis that identifies themes, writes descriptions, provides illustrative examples, and calculates frequencies.
-   **AI-Generated Executive Summary**: After analyzing individual questions, it uses the LLM to synthesize the findings into a high-level executive summary.
-   **Professional Report Generation**: Assembles a clean, well-structured report in Markdown and automatically converts it to a professional `.docx` document using Pandoc.
-   **Automated Email Distribution**: Sends the final `.docx` report as an email attachment to a predefined list of recipients using Gmail SMTP.

## How It Works: The Workflow

The script executes the following steps automatically:

1.  **Fetch Data**: Connects to the Qualtrics API and downloads the latest data for a specified survey, including the full question text.
2.  **Analyze Responses**: For each question specified in the configuration, the script sends all text responses to the LLM. The LLM is prompted to return a complete, formatted Markdown section containing a narrative summary and a thematic table with descriptions, examples, and frequencies.
3.  **Create Executive Summary**: The individual narrative summaries are collected and sent back to the LLM one final time to generate a cohesive, high-level summary.
4.  **Assemble and Convert Report**: The script assembles the executive summary, individual analyses, a raw data appendix, and a methodology section into a single Markdown file. It then uses **Pandoc** to convert this file into a polished `.docx` document.
5.  **Email Report**: Finally, it connects to Gmail's SMTP server and emails the `.docx` report to the specified facilitators.

## Setup and Installation

Follow these steps to get the project running on your local machine.

### 1. External Tools Setup

This script relies on two external applications:

-   **LM Studio**: For running the local LLM.
    -   [Download and install LM Studio](https://lmstudio.ai/).
    -   Inside LM Studio, download a model. The **`Meta-Llama-3-8B-Instruct-Q5_K_M.gguf`** model is highly recommended for this script.
    -   Go to the local server tab (`<->`) and start the server.
-   **Pandoc**: For converting the report to a `.docx` file.
    -   [Download and install Pandoc](https://pandoc.org/installing.html).
    -   For macOS, download the `.pkg` installer. If you get a security warning, go to `System Settings` > `Privacy & Security` to "Allow" the installation.
    -   After installing, you may need to add it to your system's PATH. See the [Troubleshooting](#troubleshooting) section if you encounter errors.

### 2. Project Setup

1.  **Clone the Repository**:
    ```bash
    git clone <your-repository-url>
    cd <your-repository-name>
    ```

2.  **Create a Virtual Environment**:
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate
    ```

3.  **Install Python Dependencies**: Create a file named `requirements.txt` with the content below, then run `pip install -r requirements.txt`.

    **`requirements.txt`**:
    ```
    requests
    pandas
    python-dotenv
    tqdm
    pypandoc
    ```

    **Installation command**:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Create Credentials File**:
    -   In the project's main directory, create a file named `.env`.
    -   Add your secret keys and credentials to this file. **This file should never be committed to Git.**

    **`.env` template**:
    ```ini
    # --- Qualtrics Credentials ---
    QUALTRICS_API_KEY=your_actual_qualtrics_api_key_goes_here
    QUALTRICS_BASE_URL=your_datacenterid.qualtrics.com

    # --- Gmail Credentials for Automated Email ---
    # This must be a 16-character Google App Password, not your regular password.
    GMAIL_SENDER=your-email@gmail.com
    GMAIL_APP_PASSWORD=thesixteencharacterapppassword
    ```

5.  **Create `.gitignore` File**: To ensure your `.env` file remains private, create a `.gitignore` file in the same directory with the following content:
    ```
    # Environment variables
    .env

    # Python cache
    __pycache__/

    # Generated reports
    *.md
    *.docx
    ```

## Configuration

To tailor the script to your needs, edit the variables in the **`--- 3. SCRIPT CONFIGURATION ---`** section of the Python file:

-   `LOCAL_LLM_MODEL_NAME`: Must match the model file name you have loaded in LM Studio.
-   `QUALTRICS_SURVEY_ID`: The `SV_...` ID for the survey you want to analyze.
-   `QUESTIONS_TO_ANALYZE`: A list of the question IDs (e.g., `["Q1", "Q2"]`) you want to include in the report.
-   `NUM_ROWS_TO_DISPLAY_IN_REPORT`: The number of raw data rows to show in the appendix at the end of the report.

## Running the Script

Once everything is set up:
1.  Ensure your LM Studio server is running with the correct model loaded.
2.  Activate your virtual environment (`source .venv/bin/activate`).
3.  Run the script from your terminal:
    ```bash
    python your_script_name.py
    ```

The script will print its progress to the console and, upon completion, will have generated the report files and sent the email.

## Troubleshooting

-   **Pandoc Error / `FileNotFoundError`**: If you get an error during the `.docx` conversion, it means the script can't find the `pandoc` program. Ensure you have installed it correctly and that its location is in your system's `PATH` variable. On macOS, the standard installer typically handles this, but you may need
