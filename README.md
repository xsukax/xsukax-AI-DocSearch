# xsukax AI DocSearch

**Private, offline-first, multilingual document search with local GGUF AI and incremental Retrieval-Augmented Generation (RAG).**

[Repository](https://github.com/xsukax/xsukax-AI-DocSearch) · [Report an issue](https://github.com/xsukax/xsukax-AI-DocSearch/issues) · [GPL-3.0 license](#license)

xsukax AI DocSearch turns folders on your computer into a searchable local document library. It extracts text from common document formats, builds a persistent SQLite full-text index, and provides a lightweight web interface for searching, filtering, previewing, and opening the original files. For deeper research, it can run **GGUF language models locally** to answer questions, summarize documents, and draft source-referenced reports.

An optional **persistent semantic RAG index** adds embedding-based retrieval without rebuilding the entire collection after every change. Search and inference stay on your machine; an internet connection is needed only when you explicitly download models or install dependencies.

> [!IMPORTANT]
> **No cloud AI account, API key, or external document-upload service is required.** AI functions are optional. The regular document search works without downloading a language model.

## Table of contents

- [Overview](#overview)
- [Features](#features)
- [Supported file formats](#supported-file-formats)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [How it works](#how-it-works)
- [Project structure](#project-structure)
- [Performance and limitations](#performance-and-limitations)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [Security and privacy](#security-and-privacy)
- [License](#license)
- [Author and maintainers](#author-and-maintainers)
- [Contact and support](#contact-and-support)

## Overview

xsukax AI DocSearch is a **single-file Python application** with an embedded HTML/CSS/JavaScript interface, a Flask server bound to the local computer, and SQLite databases for persistent indexing. It is designed for researchers, professionals, developers, students, and anyone who needs to search or analyze a collection of local files without uploading them to a third-party search or AI service.

The application offers three complementary workflows:

1. **Document search:** fast, exact/Boolean text search using SQLite FTS5, with an additional trigram index for selected CJK and Thai text.
2. **Local AI research:** use a compatible GGUF chat model to analyze keyword-matched documents or a broader indexed collection and generate source-referenced responses.
3. **Hybrid semantic RAG:** build a separate local vector index with a dedicated GGUF embedding model, then combine semantic passage retrieval with keyword results for evidence-grounded answers.

The primary index is persistent, and routine rescans reuse files whose size and modification timestamp have not changed. The separate semantic index also reuses unchanged vectors where possible.

## Features

| Area | Capabilities |
| --- | --- |
| Local-first operation | Runs on your computer at `127.0.0.1`; no cloud inference or hosted database required. |
| Single-file app | Python backend and browser-based UI are bundled in one `.py` script. |
| Recursive indexing | Add multiple non-overlapping folders; recursively discover eligible files and report progress and errors. |
| Incremental updates | Reuse unchanged document records, update modified files, remove deleted entries after successful scans, and rescan automatically. |
| Full-text search | SQLite FTS5, relevance ranking, `AND` / `OR` / `NOT`, parentheses, quoted phrases, and trailing-`*` prefix queries. |
| Advanced filtering | Filter by indexed folder, filename, extension, or path; sort by relevance, filename, modification date, size, or type. |
| Document preview | Read extracted content, highlight matches, jump between occurrences, copy paths, and open original documents with local applications. |
| Multilingual content | Handles Unicode, common text encodings, and right-to-left content; specialized trigram search supports CJK/Thai terms. |
| File selection | Choose built-in extension groups, add custom text extensions, skip hidden items, and exclude directory names. |
| Local AI | Load/unload compatible GGUF chat models and request answers, summaries, or reports with source IDs. |
| GGUF management | Discover manually copied models in `gguf/`, or optionally browse and download public GGUF files from Hugging Face. |
| Persistent semantic RAG | Store chunk embeddings in `rag.sqlite3`, resume interrupted work, reuse unchanged chunks, and perform incremental updates. |
| Hybrid retrieval | Combine semantic matches with SQLite keyword hits, including recently indexed documents whose embeddings are not yet ready. |
| Resource controls | Configure CPU threads, GPU layers, context size, chunk size, overlap, and embedding delays; pause or retry embedding jobs. |
| Privacy protections | Loopback-only web server, local request token, origin/host checks, and restrictive HTTP response headers. |

### Search syntax

The search box supports operators and field-specific queries:

| Example | Meaning |
| --- | --- |
| `climate adaptation` | Match both terms in **All words** mode. |
| `climate OR agriculture` | Match either term. |
| `climate NOT agriculture` | Include one term and exclude the other. |
| `"annual report"` | Search for a contiguous phrase. |
| `(energy OR water) AND policy` | Group Boolean conditions. |
| `report*` | Match words beginning with `report`. |
| `filename:budget` | Search file names. |
| `ext:pdf` | Restrict results to PDF files. |
| `path:research` | Restrict results to paths containing `research`. |

The **Any word** and **Exact phrase** dropdown modes provide alternatives to the default **All words** behavior. Additional folder, extension, filename, path, and sort controls are available below the search field.

## Supported file formats

The following extensions are offered in the application. **Not every format uses the same extractor**; files for which text extraction is not supported or fails may still appear by filename.

| Category | Extensions | Extraction notes |
| --- | --- | --- |
| Documents | `.pdf`, `.docx`, `.doc`, `.odt`, `.rtf`, `.txt`, `.md` | PDF, DOCX, ODT, and RTF have dedicated extractors; legacy DOC needs an optional local converter. |
| Spreadsheets | `.xlsx`, `.xls`, `.csv`, `.ods` | XLSX/XLS and ODS are parsed; CSV is handled as text. |
| Presentations | `.pptx`, `.ppt`, `.odp` | PPTX and ODP are supported; legacy PPT requires LibreOffice. |
| Web and structured data | `.html`, `.htm`, `.xml`, `.json`, `.yaml`, `.yml` | HTML text is extracted without scripts/styles; the others are searched as text. |
| Programming | `.py`, `.js`, `.ts`, `.java`, `.c`, `.cpp`, `.h`, `.cs`, `.go`, `.rs`, `.php`, `.sh`, `.sql` | Text-based source code. |
| Other text | `.log`, `.rst`, `.ini`, `.cfg`, `.tex`, `.srt` | Text-based extraction with encoding detection. |

Only a subset of these extensions is enabled by default; you can change the selection in **File extensions**. Custom extensions are useful for additional text formats. Arbitrary proprietary or binary files do not automatically become text-searchable when their extensions are selected.

**Important limitations:** scanned/image-only PDFs require OCR before meaningful text search; the application does not include an OCR engine. Encrypted PDFs and some malformed/unsupported files cannot be extracted. Very large files may be indexed by filename only.

## Prerequisites

- **Python 3.10 or later**, with `pip` and the standard `sqlite3` module.
- **SQLite 3.34 or later**, built with **FTS5 and trigram tokenizer support**. A compatible SQLite build is required even if you do not use AI features.
- **Windows, macOS, or Linux**, with a modern browser such as Edge, Chrome, Firefox, or Safari.
- Sufficient disk space for extracted text, SQLite indexes, and (if enabled) GGUF models and embedding vectors.
- Read permission for the folders being indexed and write permission for the configured data location and the application’s `gguf/` directory.

### Python packages

Required for the document-search application:

```bash
python -m pip install flask pymupdf python-docx openpyxl python-pptx xlrd striprtf odfpy beautifulsoup4 charset-normalizer
```

Optional packages for AI and model management:

```bash
python -m pip install llama-cpp-python huggingface-hub
```

`llama-cpp-python` is needed for both local chat generation and GGUF embeddings. `huggingface-hub` is needed only for downloading models through the UI. `numpy` is optional; if installed, vector similarity calculations can use it for better performance. Installing `llama-cpp-python` may require platform-specific build tools or GPU backend setup; see its official installation guidance if a wheel is unavailable.

### Optional legacy Office converters

- **LibreOffice (`libreoffice` or `soffice`)**: required to extract text from legacy `.ppt` files and can also convert `.doc` files.
- **antiword**: an alternative local extractor for `.doc` files.

These external converters are not necessary for modern `.docx` or `.pptx` files.

To check your Python and SQLite versions:

```bash
python --version
python -c "import sqlite3; print('SQLite:', sqlite3.sqlite_version)"
```

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/xsukax/xsukax-AI-DocSearch.git
cd xsukax-AI-DocSearch
```

Alternatively, download the Python application from the repository and place it in a writable folder. The commands below use the filename **`xsukax-AI-DocSearch.py`**, as provided with the project source. If your repository uses another filename, substitute it in the commands.

### 2. Create and activate a virtual environment

**macOS / Linux:**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

**Windows PowerShell:**

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If your Windows PowerShell policy prevents environment activation, you can use `.\.venv\Scripts\python.exe` directly in place of `python` for subsequent commands.

### 3. Install the required dependencies

```bash
python -m pip install --upgrade pip
python -m pip install flask pymupdf python-docx openpyxl python-pptx xlrd striprtf odfpy beautifulsoup4 charset-normalizer
```

Install the AI packages **only if you plan to use local GGUF models or download them through the interface**:

```bash
python -m pip install llama-cpp-python huggingface-hub
```

### 4. Start the application

```bash
python xsukax-AI-DocSearch.py
```

The application attempts to open your default browser automatically. If it does not, navigate to:

**http://127.0.0.1:9511**

Use `Ctrl+C` in the terminal to stop the server. On first launch, the application creates its local databases and the `gguf/` directory if necessary.

## Configuration

### Command-line options

```bash
python xsukax-AI-DocSearch.py --help
```

| Option | Default | Purpose |
| --- | --- | --- |
| `--port PORT` | `9511` | Choose the local HTTP port (`1024`–`65535`). |
| `--data-dir PATH` | Directory containing the Python script | Set the directory used for `index.sqlite3` and `rag.sqlite3`. |
| `--no-browser` | Off | Do not automatically open a browser. |

For example:

```bash
python xsukax-AI-DocSearch.py --port 9611 --data-dir "$HOME/DocSearchData" --no-browser
```

With this example, open `http://127.0.0.1:9611`. On Windows, replace the shell path with an appropriate local folder.

> [!NOTE]
> `--data-dir` changes the location of the **SQLite databases**, not the model directory. The **`gguf/` folder and its Hugging Face caches always live beside the Python script**.

### Indexed folder settings

Use the sidebar to configure:

| Setting | Default | Description |
| --- | --- | --- |
| File extensions | Common document/text formats | Choose which formats to index; add custom extensions. |
| Include hidden files and folders | Enabled | Include or skip names beginning with `.`. |
| Follow internal symbolic links | Enabled | Follow links whose resolved destinations remain inside the indexed root; avoid recursive loops. |
| Show filename-only results | Enabled | Display files with no extracted content in search results. |
| Maximum file size | `40 MB` | Larger files are indexed by filename, not full content. Configurable from `1` to `200 MB`. |
| Automatic rescan | `180 seconds` | Rescan periodically. Configurable from `30` to `86,400` seconds. |
| Excluded directory names | None | Skip named directories such as `.git`, `.venv`, or `node_modules`. |

Use **Rescan all** after changing files, **Cancel** to stop a running scan, or **Rebuild entire index** to regenerate the main document index from scratch. Removing a folder from the index does **not** delete its original files.

### GGUF model setup

The application uses **two different model roles**:

| Model type | Used for | Required for |
| --- | --- | --- |
| **Chat / instruction GGUF** | Writing answers, document summaries, and reports | AI answer generation in every AI evidence mode. |
| **Embedding GGUF** | Converting document chunks and queries into semantic vectors | Hybrid semantic + keyword RAG mode. |

To add models, either copy `.gguf` files into `gguf/` (including subfolders) and select **Refresh models**, or open **Local AI / RAG** and download a public GGUF from Hugging Face. Model downloads require internet access; after the models and Python packages are installed, inference does not require it.

Choose compatible GGUF models that work with your device’s CPU/GPU and memory. A normal chat model is not necessarily a valid sentence-embedding model; the embedding model must support pooled embeddings in `llama-cpp-python`. Multilingual results, including Arabic, depend on the selected models and source-document quality.

### Recommended low-resource RAG settings

The application initially configures semantic indexing with these conservative defaults:

| Setting | Default | Effect |
| --- | --- | --- |
| Embedding threads | `2` | Limits CPU concurrency during embedding. |
| Embedding GPU layers | `0` | CPU-only embedding. |
| Chunk size | `850` characters | Controls the approximate length of indexed passages. |
| Chunk overlap | `100` characters | Preserves context across adjacent passages. |
| Delay per chunk | `250 ms` | Introduces pauses to reduce continuous load. |
| Auto-update RAG | Enabled | Keeps vectors updated after document index changes. |
| Pooling | `auto` | Uses the model/library pooling default; other options include `last` and `mean`. |

The RAG worker waits while a normal document scan is running, pauses during AI answer generation, and releases its embedding model when idle. You can **Pause**, **Retry failed files**, or **Clear RAG vectors** without erasing the primary document index. Changing the embedding model or chunk configuration invalidates previous semantic vectors and requires rebuilding them.

## Usage

### Search your documents

1. Launch the application and select **Choose folder…** or paste an absolute folder path.
2. Click **+ Add and index**. Add additional folders if needed; nested or overlapping roots are not accepted.
3. In **File extensions**, select the formats you want and click **Save extensions**.
4. Allow the initial scan to finish. The sidebar shows discovered, updated, skipped, and failed files.
5. Enter a query such as `"annual report"`, `energy OR water`, or `ext:pdf`.
6. Filter and sort results as needed, select a document, and use the highlighted text preview to navigate matches.
7. Click **Open original** (or double-click a result) to open the file using your local desktop application.

Search works without an AI model. You can continue to search previously indexed files while background scans are running.

### Ask questions with a local AI model

1. Install `llama-cpp-python` if you have not already done so.
2. Open **Local AI / RAG** and add a compatible chat/instruction GGUF model.
3. Select it, adjust **Context tokens**, **GPU layers**, and **CPU threads** if needed, then choose **Load model**.
4. Enter a question, summary request, or report instruction, for example:

   ```text
   Summarize the main findings about renewable energy across my indexed documents.
   Highlight disagreements and cite the relevant source IDs.
   ```

5. Choose an **Evidence scope** and, optionally, a folder and maximum document count.
6. Click **Generate answer / report**. Review the generated text, referenced evidence documents, and coverage information. Use **Copy report** to copy the result.

#### Available AI evidence scopes

| Mode | Behavior | Best for |
| --- | --- | --- |
| **Hybrid semantic + keyword** | Retrieves embedded semantic passages and supplements them with keyword full-text hits. Requires a built embedding index. | Questions where the relevant text may not use the exact words in the question. |
| **Keywords only** | Searches indexed documents for relevant terms and supplies bounded evidence excerpts to the chat model. | Focused questions without an embedding model. |
| **All indexed documents** | Processes indexed text in batches and combines summaries into a final response. | Broad collection-level summaries and reports. |

For hybrid mode, first select a dedicated embedding GGUF in the **Persistent semantic RAG** panel, save its settings, and click **Build / update RAG**. The first build can take time; subsequent builds primarily process new or changed content. A stopped or paused build can resume from persisted checkpoints.

**Example prompts:**

```text
What policy changes are mentioned in the documents about regional water management?

Compare the arguments made in the 2024 and 2025 reports. Cite the document IDs.

Prepare a structured research brief on renewable energy projects, identifying
supporting sources and any conflicting information.
```

> [!CAUTION]
> AI output is not a substitute for reading the original documents. The model receives selected excerpts or passages rather than every byte of every file; source-ID references help verification but do not guarantee completeness or factual accuracy. Check important claims against the originals.

## How it works

```text
Selected local folders
        |
        v
Recursive file discovery and text extraction
        |
        v
index.sqlite3  (file records + SQLite FTS5 indexes)
        |                                |
        |                                +--> Local keyword search and preview
        |
        +--> Incremental document-chunk embedding (optional)
                    |
                    v
              rag.sqlite3  (vectors + source offsets)
                    |
User question --> Semantic retrieval + keyword retrieval
                    |
                    v
          Evidence passages and source IDs
                    |
                    v
          Local GGUF chat / instruction model
                    |
                    v
             Answer, summary, or report
```

**Index design:** the main database stores extracted text, metadata, paths, and full-text indexes. RAG uses a separate SQLite database containing vector embeddings, content hashes, and character offsets into the main database, rather than a second full copy of all extracted documents. Modified chunks may reuse their prior vectors when the chunk text is unchanged; deleted or deselected files are pruned on completed update passes.

**Retrieval design:** hybrid mode performs cosine-like similarity ranking over normalized vectors and augments results with FTS5 hits. Answers use selected source passages, and the local chat model is instructed to identify source IDs, uncertainty, and conflicts. The application does **not** let model-generated text execute arbitrary SQL.

## Project structure

The supplied application has a **single-file codebase**. The following illustrates the source file and runtime files it creates; `README.md`, `LICENSE`, and `.gitignore` are repository documentation/configuration files, not additional application modules.

```text
xsukax-AI-DocSearch/
├── xsukax-AI-DocSearch.py   # Flask app, indexing, extractors, RAG, and embedded web UI
├── README.md               # Project documentation (this file)
├── LICENSE                 # GNU GPL v3.0 full license text
├── .gitignore              # Recommended: exclude local databases, models, and env files
├── gguf/                   # Created at runtime; local GGUF models and HF caches
│   └── .hf-home/           # Hugging Face download/cache data (when used)
├── index.sqlite3           # Created at runtime; primary document index (default location)
├── rag.sqlite3             # Created at runtime; semantic vectors (default location)
└── .venv/                  # Optional local Python virtual environment
```

SQLite may also create temporary `-wal`, `-shm`, or journal companion files. When `--data-dir` is specified, both SQLite databases and their companion files are stored **there**, while `gguf/` remains alongside the script. The original files you index remain in their selected folders; they are not moved into the app directory.

### Suggested `.gitignore`

Prevent accidental commits of extracted document contents, local models, and generated environments:

```gitignore
.venv/
__pycache__/
*.py[cod]
index.sqlite3*
rag.sqlite3*
gguf/
.DS_Store
```

**Do not commit `index.sqlite3`, `rag.sqlite3`, or private GGUF files** to a public repository. SQLite index data can include extracted contents of confidential documents.

## Performance and limitations

- **Incremental, not constant-time:** regular rescans still traverse selected folders, but unchanged files are not re-extracted; RAG also saves vectors to reuse later.
- **Resource use:** large collections and local GGUF models can consume substantial disk space, RAM, and CPU. Start with modest thread counts, a smaller model, and a limited document scope.
- **Extraction limits:** extracted content is capped at approximately **4 million characters per file**, with up to **2,000 PDF pages** processed. The default content-extraction file size limit is **40 MB**.
- **Search constraints:** the application validates query length and syntax; the web UI paginates results. Search is not a universal binary-file scanner or OCR engine.
- **Semantic indexing:** requires a *dedicated embedding-capable* GGUF. Changes to model/chunk settings require re-embedding; model compatibility varies.
- **Source coverage:** AI responses are constrained by extraction success, retrieval quality, context windows, model capabilities, and limits selected in the interface.
- **Local desktop use:** the server is deliberately bound to loopback and is not intended as a multi-user or internet-facing web service.
- **No built-in document editing:** the app indexes, searches, previews, and opens originals; it does not modify source documents.

## Troubleshooting

| Problem | What to check |
| --- | --- |
| `ModuleNotFoundError` | Activate the correct virtual environment and install the required Python packages. |
| `SQLite ... does not include FTS5` or trigram errors | Install a Python distribution built against a compatible SQLite version with FTS5/trigram support. |
| Web interface does not open | Visit `http://127.0.0.1:9511` manually, verify the chosen port, and check the terminal for errors. |
| A document is indexed by filename only | Verify its extension, format, extraction permissions, file-size limit, and optional converter availability. Scanned PDFs need OCR separately. |
| `.doc` / `.ppt` extraction fails | Install LibreOffice; for DOC, `antiword` can also help. |
| RAG reports missing/invalid embeddings | Use a **pooled sentence-embedding GGUF** compatible with `llama-cpp-python`, then save RAG settings and rebuild. |
| GGUF model is missing from the list | Place a valid `.gguf` file under `gguf/`, then click **Refresh models**. |
| Model fails to load or runs out of memory | Try a smaller quantized model, lower context size, fewer GPU layers, or CPU-only mode. |
| AI answer misses relevant material | Check indexing status, selected file types and folder, evidence scope, max documents, and RAG completion status. |
| Database grows large | Exclude irrelevant folders/extensions, use a sensible file-size limit, and back up or rebuild indexes as appropriate. |

Include your OS, Python/SQLite versions, application command, relevant error message, and safe reproduction steps when filing a bug. **Remove private document text, absolute sensitive paths, credentials, and GGUF download tokens from logs or screenshots before sharing.**

## Contributing

Contributions, bug reports, documentation improvements, tests, and thoughtful feature requests are welcome.

1. Search [existing issues](https://github.com/xsukax/xsukax-AI-DocSearch/issues) before opening a duplicate.
2. Open an issue to discuss substantial changes or compatibility-impacting features.
3. Fork the repository and create a descriptive branch (for example, `feature/improve-pdf-extraction`).
4. Keep changes focused. Follow **PEP 8**, use descriptive names, preserve existing input validation and local-only security boundaries, and avoid introducing mandatory online services.
5. Test the changed workflows on supported operating systems when possible. For search or indexing changes, cover Unicode, malformed inputs, incremental updates, and relevant file formats.
6. Update the README when you change installation steps, user-facing behavior, configuration options, or dependencies.
7. Open a pull request to the default branch, describe the change and motivation, reference related issues, and include reproducible test results.

For a basic local source check:

```bash
python -m py_compile xsukax-AI-DocSearch.py
python xsukax-AI-DocSearch.py --help
```

These commands verify syntax and command-line parsing; they are **not** a substitute for end-to-end testing. A comprehensive automated test suite is not established by the supplied single-file source.

## Security and privacy

xsukax AI DocSearch is intended for **trusted, local desktop use**. Its security-related implementation includes:

- A Flask server listening on **`127.0.0.1` only**.
- Verification of request host and origin, plus a per-process token required by local API requests.
- HTTP response headers restricting framing, content types, referrers, and resource loading.
- Parameterized database queries and validation of search/query inputs.
- Limits on HTTP request size, document extraction, and supported archive expansion.
- Validation of GGUF model paths; models are loaded from the local model directory.
- Local AI inference without submitting document text to an external LLM API.

**Best practices:**

- Do not publish the Flask port to a LAN, tunnel, or the public internet; there is no remote-user authentication system.
- Do not share SQLite databases or logs containing sensitive document excerpts.
- Protect your device and data directory with appropriate operating-system permissions and encrypted backups where needed.
- Install Python dependencies and GGUF models from trustworthy sources; treat downloaded models and untrusted documents as potentially risky inputs.
- Keep Python, SQLite, parsers, and `llama-cpp-python` updated with relevant security fixes.
- Review generated answers and citations rather than assuming model output is authoritative.

### Reporting a vulnerability

**Please do not disclose exploitable security issues in a public issue.** If GitHub private vulnerability reporting is enabled, use the repository's **Security → Report a vulnerability** option. Otherwise, contact the maintainer privately through the contact information published on the [xsukax GitHub profile](https://github.com/xsukax). Include reproduction steps, affected versions/commit, likely impact, and a minimal proof of concept without real private documents. Public bug reports should avoid credentials and sensitive data.

## License

This project is distributed under the **GNU General Public License, version 3.0 (GPL-3.0)**. See the repository's [`LICENSE`](LICENSE) file for the complete license text and conditions.

The GPL permits use, study, modification, and redistribution subject to its terms, including applicable obligations when distributing modified versions. Model weights and third-party dependencies may be subject to **separate licenses**; review those licenses before redistribution.

## Author and maintainers

**xsukax** — project author and maintainer.

- GitHub: [@xsukax](https://github.com/xsukax)
- Project repository: [xsukax/xsukax-AI-DocSearch](https://github.com/xsukax/xsukax-AI-DocSearch)

Contributors are credited through the repository's commit and pull-request history.

## Contact and support

For help, feature requests, and non-sensitive bug reports, open a [GitHub issue](https://github.com/xsukax/xsukax-AI-DocSearch/issues). Include an informative title, the steps needed to reproduce an issue, and relevant environment details. For potential security vulnerabilities, follow the [private reporting guidance](#reporting-a-vulnerability) instead.

---

Made by **[xsukax](https://github.com/xsukax)** · **Offline-first · Open source · GPL-3.0**
