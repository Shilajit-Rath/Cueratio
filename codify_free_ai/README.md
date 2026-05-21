# CI Benchmarking → CODIFY Converter — Free Local AI Version

This version uses **Ollama** as the free/local AI engine instead of a paid API.

## What it does

- Upload a benchmarking Excel workbook.
- AI maps features into existing CODIFY sheets.
- New/extra benchmarking features automatically become new CODIFY sheets.
- Python validates the AI output so the workbook structure stays controlled.
- Every generated sheet includes `raw_text` for verification.
- `00_AI_Mapping_Audit` shows how every benchmarking feature was grouped.

## 1. Install Python

Install Python 3.10 or newer.

## 2. Install Ollama

Download and install Ollama from:

https://ollama.com

Then open Terminal / Command Prompt and run:

```bash
ollama pull llama3.1:8b
```

For weaker laptops, you can try:

```bash
ollama pull llama3.2:3b
```

Then set the model in `.env` or as an environment variable:

```bash
OLLAMA_MODEL=llama3.2:3b
```

## 3. Run the app

### Windows

Double click:

```text
run_windows.bat
```

### Mac / Linux

```bash
chmod +x run_mac_linux.sh
./run_mac_linux.sh
```

Then open:

```text
http://127.0.0.1:5000
```

## Notes

- Ollama must be running before conversion.
- First conversion may be slower because the local model loads into memory.
- Local AI quality depends on the model. `llama3.1:8b` is a good free default.
- If AI fails to produce valid JSON, the code falls back to rule-based grouping so the app does not completely fail.

## Files

- `app.py` — Flask web UI.
- `codify.py` — reader, free AI mapper/extractor, validator, workbook writer.
- `requirements.txt` — Python dependencies.
- `.env.example` — optional model settings.
