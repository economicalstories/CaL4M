# CAL4M: Local LLM Integration in Excel

## Overview

Imagine unlocking a new dimension of spreadsheet power—where your cells can not only perform calculations but execute entire AI-driven workflows locally. From automating custom data summaries and forecasting trends, to orchestrating multi-step analyses, all without recurring API bills or risking exposure of sensitive information, CAL4M empowers you to harness AI securely on your own hardware.

**CAL4M** is an Excel macro (User-Defined Function) that lets you **C**all **A** **L**ocal **L**arge **L**anguage **M**odel (CAL4M) directly from any cell. It uses [Ollama](https://ollama.com) as the backend LLM server, running on your machine (CPU or GPU). CAL4M auto-sizes its reply to the width of the calling cell, caches answers in-memory for the duration of the Excel session, and is non-volatile (it recalculates only when the prompt argument changes).

Once set up, simply type:

```excel
=CAL4M("Give me a type of tree")
```
or
```excel
=CAL4M(A2)
```
or
```excel
=CAL4M("What is the colour of " & A2)
```

in a cell, and Excel will return the model’s response (with line breaks and sanitized text to prevent formula injection).

## Files in This Repository

- **CAL4M.bas**: Exported VBA module containing the `CAL4M(prompt)` function.  
- **CAL4M_demo.xlsm**: Macro-enabled Excel workbook preconfigured with the CAL4M function.

## Prerequisites

1. **Ollama (Local LLM Server)**  
   - Provides the HTTP API on `http://localhost:11434`.  
2. **Excel for Windows**  
   - Excel 2016 or later (supports VBA macros).  
3. **Windows Users**: you must install **Ubuntu** (WSL2) first, then follow the Linux instructions inside your WSL terminal.

## Installation

### 1. Install Ollama

#### On Linux or macOS

```bash
curl -fsSL https://ollama.com/install.sh | sh
```

This installs `ollama` in `/usr/local/bin` and sets up a systemd service on Linux.

#### On Windows (with WSL2)

1. Install **Windows Subsystem for Linux (WSL2)** and Ubuntu from the Microsoft Store.  
2. Open Ubuntu and run the same Linux command above.

### 2. Pull a Model

Choose and download a model, for example:

```bash
ollama pull tinyllama:latest
```

### 3. Serve the Model

Start the Ollama HTTP server:

```bash
ollama serve
```

Confirm it’s running and GPU-enabled by checking the logs and `nvidia-smi`.

## Setting Up Excel

There are two ways to get the CAL4M function into your workbook:

Preconfigured demo: Open **CAL4M_demo.xlsm**, enable macros, and you’re ready to use =CAL4M(...)—no further setup required.

Manual import: In a blank or existing .xlsm workbook:

1. Press `Alt+F11` to open the VBA editor.  
2. Import **CAL4M.bas** (`File → Import File…`).  
3. Save and close the editor, ensure macros are enabled.

## Usage

In any cell, enter:

```excel
=CAL4M(A2)
```

Where `A2` contains your prompt. The function will:

1. Poll the Ollama server until ready (up to 30 s).  
2. Send a chat-completion request with a system instruction for terse output.  
3. Replace literal \n with Excel line breaks.  
4. Sanitize leading characters to prevent formula injection.

### How sizing works
* CAL4M reads `Application.Caller.ColumnWidth`.
* It sets `max_tokens` ≈ (column-width × 0.9 ÷ 4).
* Answers longer than the cell width are automatically truncated by the model.


## Further Improvements

- **Caching**: Store recent responses in VBA or in a hidden sheet to avoid duplicate API calls.  
- **Batch Processing**: Use Python + [XLWings](https://xlwings.org) for reading/writing whole columns at once.  
- **Model Selection**: Add an optional parameter to choose between multiple pulled models.  
- **Timeout Tuning**: Expose `maxWait` as an optional argument or global setting.  
- **Non-Blocking UI**: Provide a button-triggered Sub to avoid freezing Excel during calls.  
- **Alternative Backends**: Support `llama.cpp`, HF TGI, or other local LLM servers via configurable endpoints.  
- **Error Logging**: Write errors to a hidden worksheet or external log file for debugging.

---

Feel free to fork, submit issues, or propose enhancements on GitHub!

## Author

PC Hubbard

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
