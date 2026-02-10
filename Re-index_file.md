# Audit Document Search Engine - Complete Build Guide

## Overview

This document contains all steps to build an Audit Document Search Engine from scratch. The software allows auditors to search through thousands of documents (PDF, Word, Excel, Text) using keyword and fuzzy search, with advanced NLP capabilities for intelligent analysis.

---

## Table of Contents

1. [Project Summary](#1-project-summary)
2. [Prerequisites](#2-prerequisites)
3. [Installation Steps](#3-installation-steps)
4. [Project Structure](#4-project-structure)
5. [How to Use](#5-how-to-use)
6. [Features](#6-features)
7. [Next Steps - Future Phases](#7-next-steps---future-phases)
8. [Known Challenges & Unsolved Problems](#8-known-challenges--unsolved-problems)
9. [Troubleshooting](#9-troubleshooting)

---

## 1. Project Summary

| Item | Description |
|------|-------------|
| **Purpose** | Search engine for audit documents with NLP capabilities |
| **Technology** | Python, Streamlit, pypdf, python-docx, openpyxl |
| **Supported Files** | PDF, Word (.docx), Excel (.xlsx), Text (.txt), Email (.eml) |
| **Search Types** | Simple keyword search, Fuzzy/partial match search |
| **NLP Features** | NER, Topic Modelling, Communication Analysis, Sentiment Analysis |
| **Interface** | Web-based (runs in browser) |

---

## 2. Prerequisites

### Hardware Requirements
- RAM: 8 GB minimum (16 GB recommended)
- Storage: 512 GB
- Processor: Any modern CPU (Intel i5/i7 or AMD Ryzen)
- GPU: Optional (speeds up NLP processing)

### Software Requirements
- Windows 10 or 11
- Python 3.11 or 3.12
- Web browser (Chrome, Edge, Firefox)
- Java 11+ (for Apache Solr - Phase 2)

---

## 3. Installation Steps

### Step 3.1: Install Python

1. Go to: https://www.python.org/downloads/
2. Download Python 3.11 or 3.12
3. Run the installer
4. **IMPORTANT**: Check the box **"Add Python to PATH"**
5. Click "Install Now"

**Verify installation:**
```
python --version
```

### Step 3.2: Create Project Folder

Create this folder structure:
```
AuditSearchEngine/
├── app/
├── documents/
├── models/
└── solr_data/
```

### Step 3.3: Create Virtual Environment

Open Command Prompt or PowerShell and run:

```powershell
cd C:\Users\user\Documents\code\AI_P900\Software_Development\AuditSearchEngine

python -m venv venv
```

### Step 3.4: Activate Virtual Environment

**PowerShell:**
```powershell
.\venv\Scripts\activate
```

**Command Prompt:**
```cmd
venv\Scripts\activate
```

You should see `(venv)` at the beginning of your command line.

### Step 3.5: Install Required Packages

Run these commands one by one:

```powershell
pip install streamlit
pip install pypdf
pip install python-docx
pip install openpyxl
```

### Step 3.6: Create the Application File

Create file: `app/search_app.py`

(See source code in the app folder)

### Step 3.7: Run the Application

```powershell
python app/search_app.py
```

Or use Streamlit directly:
```powershell
streamlit run app/search_app.py
```

### Step 3.8: Open in Browser

Go to: http://localhost:8501

---

## 4. Project Structure

```
AuditSearchEngine/
│
├── app/
│   └── search_app.py          # Main application code
│
├── documents/                  # Put your documents here
│   ├── example.pdf
│   ├── report.docx
│   └── data.xlsx
│
├── venv/                       # Python virtual environment
│
├── search_index.json           # Auto-generated search index
│
├── requirements.txt            # Python dependencies
├── SETUP.bat                   # Windows setup script
├── RUN_APP.bat                 # Windows run script
├── GETTING_STARTED.txt         # Quick start guide
└── Re-index_file.md            # This documentation
```

---

## 5. How to Use

### 5.1: Add Documents

**Option A: Copy files manually**
1. Open File Explorer
2. Go to: `AuditSearchEngine\documents`
3. Copy your PDF, Word, Excel, or Text files here

**Option B: Upload through the app**
1. Open the app in browser (http://localhost:8501)
2. Look at the left sidebar
3. Use "Upload Documents" section
4. Select files to upload

### 5.2: Index Documents

1. Open the app in browser
2. Look at the left sidebar
3. Click **"Re-index All Documents"**
4. Wait for confirmation message

### 5.3: Search Documents

1. Type your search query in the search box
2. Select search type:
   - **Simple**: Exact keyword match
   - **Fuzzy**: Partial match (good for typos)
3. View results with:
   - Document name
   - Number of matches
   - Preview snippet

---

## 6. Features

### Current Features (Phase 1) - COMPLETED

| Feature | Description | Status |
|---------|-------------|--------|
| PDF Support | Extract and search text from PDF files | Done |
| Word Support | Extract and search text from .docx files | Done |
| Excel Support | Extract and search text from .xlsx files | Done |
| Text Support | Search plain text and markdown files | Done |
| Simple Search | Exact keyword matching | Done |
| Fuzzy Search | Partial matching for poor data quality | Done |
| Case Sensitivity | Option for case-sensitive search | Done |
| Result Preview | Shows text snippet around matches | Done |
| Match Count | Shows number of matches per document | Done |
| Upload Feature | Upload documents directly through browser | Done |

---

## 7. Next Steps - Future Phases

### Phase 2: Apache Solr Integration

**Purpose:** Faster search, better scalability, advanced features

| Feature | Description | Status |
|---------|-------------|--------|
| Faceted Search | Filter by date, type, department | Planned |
| Result Highlighting | Highlight matching terms | Planned |
| Spell Checking | Suggest corrections for typos | Planned |
| Synonyms Support | "irregularity" = "violation" | Planned |
| Stemming Search | Word-root matching (audit/auditing/audited) | Planned |
| Wildcard Search | Search with * and ? patterns | Planned |
| Boolean Queries | AND, OR, NOT operators | Planned |
| Exact Block Matching | Match exact text blocks | Planned |

**Tools:**
- Apache Solr 9.x
- Java 11+
- Apache Tika (document extraction)
- pysolr (Python library)

---

### Phase 3: NLP - Named Entity Recognition (NER)

**Purpose:** Identify and classify named entities in unstructured data

| Feature | Description | Status |
|---------|-------------|--------|
| Person Identification | Identify individuals mentioned in documents | Planned |
| Organization Detection | Identify companies, departments, agencies | Planned |
| Location Extraction | Identify geographic locations | Planned |
| Date Recognition | Extract dates and time periods | Planned |
| Monetary Values | Identify amounts, budgets, figures | Planned |
| Entity Summaries | Frequency reports of key entities | Planned |
| Entity Network | Visualize relationships between entities | Planned |

**Tools:**
- spaCy with en_core_web_lg model
- Hugging Face Transformers
- Custom NER models for audit domain

---

### Phase 4: Communication Analysis

**Purpose:** Analyze email correspondence and communication flows

| Feature | Description | Status |
|---------|-------------|--------|
| Email Parsing | Extract sender, receiver, date, subject, body | Planned |
| Communication Flow Visualization | Network graph of who communicates with whom | Planned |
| Key Actor Identification | Find central figures in communication network | Planned |
| Relationship Mapping | Map relationships between individuals | Planned |
| Timeline Analysis | Visualize communication over time | Planned |
| Thread Reconstruction | Group related emails into conversations | Planned |

**Tools:**
- Python email library
- NetworkX (network analysis)
- Plotly/PyVis (visualization)

---

### Phase 5: Topic Modelling

**Purpose:** Identify themes and categorize documents automatically

| Feature | Description | Status |
|---------|-------------|--------|
| LDA Topic Modelling | Latent Dirichlet Allocation algorithm | Planned |
| Topic Distribution | Show topics per document | Planned |
| Document Clustering | Group similar documents | Planned |
| Topic Filtering | Search within specific topics | Planned |
| Key Document Identification | Find assembly protocols, contracts, etc. | Planned |
| Topic Visualization | Word clouds, topic charts | Planned |
| Subset Search | Restrict search to document categories | Planned |

**Tools:**
- Gensim (LDA implementation)
- scikit-learn
- pyLDAvis (visualization)

---

### Phase 6: Sentiment Analysis

**Purpose:** Analyze tone and sentiment in documents

| Feature | Description | Status |
|---------|-------------|--------|
| Sentiment Scoring | Positive/negative/neutral classification | Planned |
| Risk Flagging | Identify high-risk or concerning content | Planned |
| Opinion Mining | Extract opinions and judgments | Planned |
| Audit Finding Tone | Analyze severity of findings | Planned |

**Tools:**
- Hugging Face Transformers
- TextBlob
- VADER Sentiment

---

### Phase 7: Machine Translation

**Purpose:** Translate documents between languages

| Feature | Description | Status |
|---------|-------------|--------|
| Multi-language Support | Translate to/from English, Urdu, others | Planned |
| Real-time Translation | Translate search results | Planned |
| Cross-language Search | Search in one language, find in another | Planned |

**Tools:**
- Hugging Face Transformers (MarianMT, mBART)
- Google Translate API (optional)
- OpenAI/Anthropic APIs (optional)

---

### Phase 8: Urdu Language Support

**Purpose:** Full support for Urdu language documents

| Feature | Description | Status |
|---------|-------------|--------|
| Urdu Text Extraction | Read Urdu from PDF/Word | Planned |
| Urdu Search | Search in Urdu text | Planned |
| RTL Display | Right-to-left text rendering | Planned |
| Urdu NER | Named entities in Urdu | Planned |
| Urdu Topic Modelling | Topics in Urdu documents | Planned |

**Tools:**
- urduhack library
- CAMeL Tools
- Multilingual BERT models

---

### Phase 9: Advanced AI Features

**Purpose:** Intelligent audit assistance

| Feature | Description | Status |
|---------|-------------|--------|
| Document Summarization | Auto-generate summaries | Planned |
| Question Answering | Ask questions, get answers from documents | Planned |
| Similar Document Finder | Find related documents | Planned |
| Audit Guidance | AI suggestions for audit process | Planned |
| Anomaly Detection | Flag unusual patterns | Planned |

**Tools:**
- LangChain / LlamaIndex
- Ollama (local LLM)
- OpenAI/Anthropic API (optional)

---

## 8. Known Challenges & Unsolved Problems

### 8.1 Data Quality Issues

| Problem | Description | Potential Solutions |
|---------|-------------|---------------------|
| Poor OCR Quality | Scanned documents with low quality text | Better OCR preprocessing, image enhancement |
| Inconsistent Formatting | Documents with varied structures | Robust parsing rules, ML-based extraction |
| Missing Metadata | Documents without proper dates/authors | Metadata inference from content |
| Encoding Issues | Special characters, mixed encodings | Unicode normalization, encoding detection |

---

### 8.2 Handwritten Notes - MAJOR UNSOLVED CHALLENGE

| Problem | Description | Current Status |
|---------|-------------|----------------|
| Handwriting Recognition | Convert handwritten notes to searchable text | **NOT SOLVED** |
| Varied Handwriting Styles | Different people write differently | Challenging |
| Mixed Content | Handwritten + printed text | Requires segmentation |
| Languages | Handwriting in English, Urdu, others | Language-specific models needed |

**Potential Approaches to Explore:**
- Microsoft Azure Handwriting Recognition API
- Google Cloud Vision API
- TrOCR (Transformer-based OCR)
- Custom trained models on audit-specific handwriting

**Why It's Difficult:**
- Handwriting varies significantly between individuals
- Quality of scans affects recognition
- Domain-specific terminology may not be in training data
- Urdu handwriting is particularly challenging (RTL, connected script)

---

### 8.3 NLP Quality Challenges

| Problem | Description | Potential Solutions |
|---------|-------------|---------------------|
| Domain Vocabulary | Audit-specific terms not in standard models | Fine-tune models on audit data |
| Context Understanding | NLP misunderstanding audit context | Custom training data |
| Multi-language Content | Documents mixing English and Urdu | Multilingual models |
| Abbreviations | Agency-specific abbreviations | Custom abbreviation dictionary |

---

### 8.4 Areas for Future Research

| Area | Description | Techniques to Explore |
|------|-------------|----------------------|
| **Improved Input Quality** | Enhance document quality before processing | Image preprocessing, denoising, super-resolution |
| **Transformer-based Translation** | Better machine translation | mBART, NLLB, GPT-based translation |
| **Sentiment in Audit Context** | Understanding sentiment in formal documents | Fine-tuned sentiment models |
| **Cross-lingual NER** | Entity recognition across languages | XLM-RoBERTa, multilingual BERT |

---

## 9. Troubleshooting

### Problem: "Python is not recognized"

**Solution:** Reinstall Python with "Add Python to PATH" checked

### Problem: "streamlit is not recognized"

**Solution:**
1. Make sure virtual environment is activated: `.\venv\Scripts\activate`
2. Install streamlit: `pip install streamlit`

### Problem: "ModuleNotFoundError: No module named 'streamlit'"

**Solution:**
1. Activate virtual environment: `.\venv\Scripts\activate`
2. Reinstall packages:
   ```
   pip install streamlit pypdf python-docx openpyxl
   ```

### Problem: Browser shows "localhost refused to connect"

**Solution:**
1. Check PowerShell/Command Prompt for error messages
2. Make sure the app is running (you should see "You can now view your Streamlit app")
3. Try: http://127.0.0.1:8501 instead of localhost

### Problem: "No documents found" after search

**Solution:**
1. Add documents to the `documents` folder
2. Click "Re-index All Documents" in the sidebar
3. Try a different search term

### Problem: PDF text not extracting properly

**Solution:**
- Some PDFs are image-based (scanned). These require OCR.
- OCR support will be added in future phases using Tesseract.

---

## Quick Reference Commands

```powershell
# Navigate to project folder
cd C:\Users\user\Documents\code\AI_P900\Software_Development\AuditSearchEngine

# Activate virtual environment
.\venv\Scripts\activate

# Install all packages
pip install streamlit pypdf python-docx openpyxl

# Run the application
streamlit run app/search_app.py

# Or run with Python directly
python app/search_app.py
```

---

## Development Roadmap Summary

| Phase | Feature | Priority | Difficulty | Status |
|-------|---------|----------|------------|--------|
| 1 | Basic Document Search | High | Easy | **DONE** |
| 2 | Apache Solr Integration | High | Medium | Planned |
| 3 | Named Entity Recognition | High | Medium | Planned |
| 4 | Communication Analysis | Medium | Medium | Planned |
| 5 | Topic Modelling (LDA) | Medium | Medium | Planned |
| 6 | Sentiment Analysis | Medium | Medium | Planned |
| 7 | Machine Translation | Medium | Hard | Planned |
| 8 | Urdu Language Support | High | Hard | Planned |
| 9 | Advanced AI Features | Low | Hard | Planned |

---

## Contact & Resources

- **Streamlit Documentation:** https://docs.streamlit.io/
- **pypdf Documentation:** https://pypdf.readthedocs.io/
- **Apache Solr:** https://solr.apache.org/
- **spaCy:** https://spacy.io/
- **Gensim (Topic Modelling):** https://radimrehurek.com/gensim/
- **Hugging Face:** https://huggingface.co/
- **NetworkX:** https://networkx.org/

---

*Document created: 2026-02-10*
*Version: 1.0 - Phase 1 Complete*
*Next Update: After Phase 2 implementation*
