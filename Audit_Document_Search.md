# Audit Document Search Engine

A powerful search tool for auditors with Named Entity Recognition (NER), Topic Modelling, Sentiment Analysis, and Advanced AI Features.

## Features

### Core Search Capabilities
- **Simple Search** - Basic keyword search across all documents
- **Boolean Search** - Support for AND, OR, NOT operators
- **Wildcard Search** - Use `*` for pattern matching (e.g., `procur*`)
- **Fuzzy Search** - Find partial matches and similar terms

### Document Support
- PDF files (with OCR for scanned documents)
- Word documents (.docx, .doc)
- Excel spreadsheets (.xlsx, .xls)
- Text files (.txt, .md)

### Named Entity Recognition (NER)
- **Persons** - Names and titles
- **Organizations** - Companies, government bodies, institutions
- **Locations** - Cities, countries (Pakistani and international)
- **Dates** - Various date formats
- **Money** - Currency amounts (USD, PKR, Rs., etc.)
- **Emails & Phone Numbers**
- **Percentages**

### PDF Table Extraction
- Extract tables from multiple PDFs
- Export to consolidated Excel file
- Preview extracted tables in the app

### Topic Modelling (LDA)
- Automatically discover themes in documents
- Assign documents to topics
- Filter documents by topic
- Configurable number of topics and words

### Sentiment & Risk Analysis
- Positive/Neutral/Negative sentiment detection
- Risk level assessment (High/Medium/Low)
- Risk keyword detection:
  - **High Risk**: fraud, embezzlement, violation, corruption, theft
  - **Medium Risk**: delay, overdue, pending, error, deviation
  - **Attention**: urgent, critical, priority, investigate

### AI Assistant (Phase 9)

#### Document Summarization
- Auto-generate summaries using T5 model
- Summaries saved for quick access
- Works with any indexed document

#### Q&A Chat (RAG)
- Ask questions in natural language
- Retrieves relevant document chunks
- Uses FAISS for fast vector search
- Sentence-transformers for embeddings

#### Anomaly Detection
- Extracts monetary amounts from documents
- Uses Isolation Forest algorithm
- Flags statistical outliers
- Supports multiple currency formats

---

## Installation

### Prerequisites
- Python 3.10+
- pip package manager

### Install Dependencies

```bash
# Core dependencies
pip install streamlit pypdf python-docx openpyxl

# NLP dependencies
pip install spacy textblob scikit-learn
python -m spacy download en_core_web_sm

# OCR dependencies (optional)
pip install pytesseract pdf2image pillow
# Also install Tesseract OCR and Poppler

# Table extraction
pip install pdfplumber

# AI Features (Phase 9)
pip install transformers torch sentence-transformers faiss-cpu
```

---

## Usage

### Running the Application

```bash
cd AuditSearchEngine
streamlit run app/search_app.py
```

The app will open at `http://localhost:8501`

### Indexing Documents

1. Place documents in the `documents/` folder
2. Click "Re-index with NER" in the sidebar
3. Wait for indexing to complete

### Using Search

1. Go to the **Search** tab
2. Enter your query
3. Select search type (Simple, Boolean, Wildcard, or Fuzzy)
4. View results with highlighted matches
5. Export to Excel if needed

### Using AI Features

#### Summarization
1. Go to **AI Assistant** > **Summarization**
2. Select a document
3. Click "Generate Summary"
4. View and save the summary

#### Q&A Chat
1. Go to **AI Assistant** > **Q&A Chat**
2. Click "Index Embeddings" (first time only)
3. Type your question
4. View relevant document excerpts

#### Anomaly Detection
1. Go to **AI Assistant** > **Anomaly Detection**
2. Click "Run Anomaly Detection"
3. Review flagged amounts

---

## Project Structure

```
AuditSearchEngine/
├── app/
│   └── search_app.py      # Main application
├── documents/             # Place documents here
├── search_index.json      # Document index
├── entities_index.json    # Extracted entities
├── topics_index.json      # Topic model results
├── sentiment_index.json   # Sentiment analysis results
├── summaries_index.json   # Generated summaries
├── faiss_index.bin        # Vector embeddings index
├── faiss_data.json        # Chunk data for FAISS
└── Audit_Document_Search.md
```

---

## Configuration

### File Paths
All data files are stored in the `AuditSearchEngine/` directory.

### OCR Setup (Windows)
1. Install Tesseract OCR from https://github.com/tesseract-ocr/tesseract
2. Install Poppler from https://github.com/osber/poppler-windows
3. Place Poppler in `C:\poppler\`

### Model Downloads
On first use, the following models will be downloaded:
- **T5-small** (~250MB) - For summarization
- **all-MiniLM-L6-v2** (~90MB) - For embeddings

---

## Tabs Overview

| Tab | Description |
|-----|-------------|
| **Search** | Main search interface with multiple search modes |
| **Entities (NER)** | View extracted persons, organizations, locations, etc. |
| **Statistics** | Document statistics and word counts |
| **Table Extraction** | Extract and export tables from PDFs |
| **Topics** | LDA topic modelling results |
| **Sentiment** | Risk and sentiment analysis |
| **AI Assistant** | Summarization, Q&A, and Anomaly Detection |

---

## Dependencies

### Core
- streamlit
- pypdf
- python-docx
- openpyxl

### NLP
- spacy
- textblob
- scikit-learn

### OCR (Optional)
- pytesseract
- pdf2image
- pillow

### AI Features
- transformers
- torch
- sentence-transformers
- faiss-cpu

---

## Troubleshooting

### "Summarization not available"
Install transformers and torch:
```bash
pip install transformers torch
```

### "Q&A Chat not available"
Install sentence-transformers and faiss:
```bash
pip install sentence-transformers faiss-cpu
```

### OCR not working
1. Verify Tesseract is installed
2. Check Poppler path in `C:\poppler\`
3. Ensure pdf2image is installed

### ChromaDB Pydantic Error
Use FAISS instead (already configured in Phase 9):
```bash
pip install faiss-cpu
```

---

## Version History

- **Phase 1-3**: Core search, NER, document support
- **Phase 4**: PDF table extraction
- **Phase 5**: Topic modelling (LDA)
- **Phase 6**: Sentiment and risk analysis
- **Phase 7**: OCR support for scanned PDFs
- **Phase 9**: AI features (Summarization, Q&A, Anomaly Detection)

---

## License

Internal audit tool - For authorized use only.

---

## Support

For issues or feature requests, contact the development team.
