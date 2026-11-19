#  AuraChat - Offline Document Intelligence

**AuraChat** is a lightweight, 100% offline desktop Document Q&A application built with Python and Tkinter. It allows you to upload documents (`.pdf`, `.docx`, `.txt`) and ask natural-language questions about their contents without needing an internet connection, external APIs, or heavy AI model downloads.

---

##  Features

* **100% Offline & Private:** Runs completely locally using a custom TF-IDF + keyword-overlap NLP engine. Your documents never leave your machine.
* **Multi-Format Support:** Easily index and query `.txt`, `.pdf`, and `.docx` files.
* **Smart Answer Synthesis:** Detects question intent (e.g., *who, when, where, how many, yes/no*) to generate natural, context-aware responses.
* **Source Citations & Confidence Scores:** Displays exact source snippets from your document alongside a confidence indicator (*High*, *Medium*, or *Low*) for every answer.
* **Modern Dark UI:** Custom-styled Tkinter interface featuring smooth typing animations, thinking indicators, and a clean sidebar.
* **Chat Export:** Save your entire Q&A session to a timestamped `.txt` file.

---

##  Installation & Setup

### Prerequisites
* Python 3.8 or higher
* `tkinter` (usually included with standard Python installations)

### Optional Dependencies
AuraChat works out-of-the-box for `.txt` files using only Python's standard library. To enable **PDF** and **Word (`.docx`)** support, install the following packages:

```bash
pip install pypdf python-docx
```

---

##  Usage

1. Clone or download this repository.
2. Run the application from your terminal:
   ```bash
   python AuraChat.py
   ```
3. Click **📂 UPLOAD** (or press `Ctrl + O`) to load a document.
4. Type your question in the input box and press **Enter** or click **ASK ›**.

---

##  Keyboard Shortcuts

| Shortcut | Action |
| :--- | :--- |
| `Ctrl + O` | Upload a document (`.pdf`, `.docx`, `.txt`) |
| `Ctrl + E` | Export chat history to a `.txt` file |
| `Ctrl + L` | Clear current chat history |
| `Enter` | Submit your question |
| `Esc` | Cancel typing animation / Close popups |

---

##  How It Works

Instead of relying on cloud LLMs, AuraChat uses a fast, deterministic local retrieval pipeline (`LocalAIEngine`):
1. **Sentence Segmentation:** Cleans line breaks and splits text into sentences while preserving common abbreviations.
2. **TF-IDF & Overlap Scoring:** Computes Term Frequency–Inverse Document Frequency (TF-IDF) alongside direct token overlap, giving a slight positional boost to introductory and concluding sentences.
3. **Deduplication:** Filters out redundant sentences using Jaccard similarity.
4. **Synthesis:** Formats the top-scoring sentences based on the detected question type.
