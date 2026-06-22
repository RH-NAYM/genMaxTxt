# WAV Transcript Generator and Bangla Text Normalizer

This project turns `.wav` audio files into an Excel transcript file, then creates a second Excel file with Bangla-friendly normalized text for text-to-speech or dataset cleanup.

It has two main scripts:

- `generate_text.py`: transcribes `.wav` files with Google Cloud Speech-to-Text.
- `normalize_excel_with_gemini.py`: normalizes transcript text with Gemini 2.5 Flash on Vertex AI.
- `rename_audio_files_from_excel.py`: renames WAV files from the normalized workbook and writes a new Excel file with the updated names.

## Features

- Batch transcribes all `.wav` files in a folder.
- Can transcribe a single `.wav` file.
- Skips audio files that are 3 seconds or shorter.
- Supports 16-bit PCM WAV and 32-bit float WAV conversion.
- Saves progress into Excel after every processed file or row.
- Normalizes numbers, phone numbers, emails, URLs, dates, times, currency, abbreviations, acronyms, social handles, and mixed Bangla-English text.
- Creates a simple 3-column normalized Excel file:
  - `audio file`
  - `old text`
  - `normalize text`

## Requirements

- Python 3.10 or newer
- Google Cloud service account JSON file
- Google Cloud Speech-to-Text API enabled
- Vertex AI API enabled
- Service account permissions for Speech-to-Text and Vertex AI

Install Python dependencies:

```bash
pip install -r requirements.txt
```

## Google Cloud Setup

1. Create or select a Google Cloud project.
2. Enable these APIs:
   - Cloud Speech-to-Text API
   - Vertex AI API
3. Create a service account.
4. Give it the needed roles, for example:
   - `Cloud Speech Client`
   - `Vertex AI User`
5. Download the service account JSON key.
6. Save it in the project root as:

```text
service_account.json
```

Do not commit this file. It is already ignored by `.gitignore`.

## Project Structure

```text
.
├── generate_text.py
├── normalize_excel_with_gemini.py
├── rename_audio_files_from_excel.py
├── requirements.txt
├── service_account.json        # local only, do not commit
├── wav_text.xlsx               # generated transcript file
├── wav_text_normalized_3_columns.xlsx
└── wavs/
    ├── 1.wav
    └── ...
```

## Step 1: Generate Transcripts

Put your `.wav` files inside the `wavs` folder.

Run:

```bash
python generate_text.py
```

Default output:

```text
wav_text.xlsx
```

The generated Excel file has:

```text
Audio Id | Text
```

### Useful Transcription Commands

Transcribe a custom folder:

```bash
python generate_text.py --dir wavs
```

Transcribe one file:

```bash
python generate_text.py --file wavs/294.wav
```

Use a custom output file:

```bash
python generate_text.py --output my_transcripts.xlsx
```

Use a different language:

```bash
python generate_text.py --language en-US
```

## Step 2: Normalize Text With Gemini

Run:

```bash
python normalize_excel_with_gemini.py --file wav_text.xlsx
```

Default output:

```text
wav_text_normalized_3_columns.xlsx
```

The normalized Excel file has exactly:

```text
audio file | old text | normalize text
```

### Progress Bar

The normalizer shows progress like this:

```text
Progress: [##########--------------------] 42/2071 (  2.0%) elapsed 1m 12s eta 58m 30s DONE 123.wav
```

It shows:

- processed rows
- total rows
- percentage
- elapsed time
- estimated time remaining
- row status
- current audio file

The script first checks each sentence with regular expressions. If nothing looks like a number, date, URL, email, handle, abbreviation, acronym, or other normalization target, the original sentence is copied directly into the output workbook without calling Gemini.

### Useful Normalization Commands

Run a small test first:

```bash
python normalize_excel_with_gemini.py --file wav_text.xlsx --limit 10 --output sample_normalized.xlsx
```

Use a different Vertex AI location:

```bash
python normalize_excel_with_gemini.py --file wav_text.xlsx --location us-central1
```

Use a different Google Cloud project:

```bash
python normalize_excel_with_gemini.py --file wav_text.xlsx --project your-project-id
```

Use a different service account file:

```bash
python normalize_excel_with_gemini.py --file wav_text.xlsx --credentials path/to/service_account.json
```

Handle quota/rate-limit errors:

```bash
python normalize_excel_with_gemini.py --file wav_text.xlsx --retry-delay 30 --max-retry-delay 600
```

By default, retryable Gemini errors such as `429 RESOURCE_EXHAUSTED` are retried forever with exponential backoff. Existing `ERROR:` rows in the output workbook are retried on the next run.

Resume after stopping:

```bash
python normalize_excel_with_gemini.py --file wav_text.xlsx
```

The script skips audio files already present in the output file with a normalized value.

## Output Files

`wav_text.xlsx` is produced by `generate_text.py`.

`wav_text_normalized_3_columns.xlsx` is produced by `normalize_excel_with_gemini.py`.

`wav_text_renamed.xlsx` is produced by `rename_audio_files_from_excel.py`.

You can safely choose new output names with `--output`.

## Developer Notes

### Dependencies

Runtime dependencies are listed in `requirements.txt`:

```text
google-cloud-speech
google-genai
openpyxl
```

Python standard library modules are used for argument parsing, file paths, WAV parsing, JSON, timing, and environment setup.

### Script Design

`generate_text.py`:

- Reads WAV metadata directly.
- Converts 32-bit float WAV data to 16-bit LINEAR16 for Google Speech-to-Text.
- Writes or updates rows in the transcript workbook.
- Saves the workbook after every processed audio file.

`normalize_excel_with_gemini.py`:

- Reads `Audio Id` and `Text` from the transcript workbook.
- Uses Vertex AI Gemini through the service account.
- Writes a new workbook with only the 3 user-facing columns.
- Saves the workbook after every processed row by default.
- Can resume work by skipping rows already present in the output workbook.

### Code Style

- Keep defaults relative to the project root so the scripts work on Windows, macOS, and Linux.
- Keep credentials out of source control.
- Prefer adding new CLI flags instead of hardcoding paths or project-specific values.
- Keep generated Excel and audio files out of commits unless they are intentionally shared sample data.

## Troubleshooting

### `Missing dependency`

Install requirements:

```bash
pip install -r requirements.txt
```

### `Google credentials file not found`

Make sure `service_account.json` exists in the project root, or pass:

```bash
python generate_text.py --credentials path/to/service_account.json
```

For normalization:

```bash
python normalize_excel_with_gemini.py --credentials path/to/service_account.json
```

### `Permission denied` or API errors

Check that:

- the service account JSON is valid
- the project has Speech-to-Text enabled
- the project has Vertex AI enabled
- the service account has the required roles
- billing is enabled for the Google Cloud project

### Output file already exists with different columns

The normalizer expects its output file to use the 3-column format. Use a new output file:

```bash
python normalize_excel_with_gemini.py --file wav_text.xlsx --output new_normalized.xlsx
```

### Rename audio files from the normalized workbook

Run:

```bash
python rename_audio_files_from_excel.py
```

Default behavior:

- reads `wav_text_normalized_3_columns.xlsx`
- checks whether each `audio file` exists in `wavs`
- renames matching files
- writes `wav_text_renamed.xlsx` with updated names like `records_3rd_session_1.wav`, `records_3rd_session_2.wav`, `records_3rd_session_3.wav`

Use a different base name:

```bash
python rename_audio_files_from_excel.py --base-name myRecords
```

If you want to rename from workbook text instead of sequential names, pass:

```bash
python rename_audio_files_from_excel.py --name-source "normalize text"
```

## Security

Never commit:

- `service_account.json`
- `.env`
- private audio datasets
- generated files that contain private transcripts

Review generated transcripts before sharing, because audio data may contain personal or sensitive information.
