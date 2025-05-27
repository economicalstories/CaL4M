# CAL4M: Call A Local Low-Latency Lamnguage Model in Excel

## Overview

Unlock a whole new tier of spreadsheet intelligence—use a local LLM just like any built-in Excel function. **CAL4M** is an User-Defined Function for Excel that lets you **C**all **A** **L**ocal **L**ow **L**atency **L**anguage **M**odel (CAL4M) directly from any cell. 

With CAL4M, your cells can call your LLM directly in formulas, fire off hundreds of prompts in parallel via Excel’s recalculation engine, and chain results through dependent workflows and pivots—no chat window required. 

It uses [Ollama](https://ollama.com) as the backend LLM server, running on your machine (GPU highly recommended). Everything runs on a small, quantized model on your own machine, so you get sub-second, zero-cost inference with full data privacy and no recurring API bills. 

CAL4M auto-sizes its reply to the width of the calling cell, caches answers in-memory for the duration of the Excel session, and is non-volatile (it recalculates only when the prompt argument changes). The LLM has access only to the data provided in the formula, not the whole spreadsheet.

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

Choose and download a model (model choice is important here - models that are good at verbose chat aren't necessarily good at rapid and terse excel-style responses), for example:

```bash
ollama pull phi3.5:3.8b
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
* Answers longer than the cell width are automatically truncated by the model.


---

Feel free to fork, submit issues, or propose enhancements on GitHub!

## Author

PC Hubbard

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
