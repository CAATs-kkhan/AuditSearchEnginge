# Audit Document Search Engine

A powerful search tool for auditors with Named Entity Recognition (NER), Topic Modelling, Sentiment Analysis, and Advanced AI Features.

## Features

| Feature | Description |
|---------|-------------|
| **Search** | Simple, Boolean, Wildcard, and Fuzzy search modes |
| **NER** | Extract persons, organizations, locations, dates, money, emails |
| **Table Extraction** | Extract tables from PDFs to Excel |
| **Topic Modelling** | LDA-based theme discovery |
| **Sentiment Analysis** | Risk level assessment and keyword detection |
| **AI Summarization** | Auto-generate document summaries (T5 model) |
| **Q&A Chat** | Ask questions about your documents (RAG) |
| **Anomaly Detection** | Flag unusual amounts using Isolation Forest |

## Supported Documents

- PDF files (with OCR for scanned documents)
- Word documents (.docx, .doc)
- Excel spreadsheets (.xlsx, .xls)
- Text files (.txt, .md)

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

# Table extraction
pip install pdfplumber

# OCR (optional)
pip install pytesseract pdf2image pillow

# AI Features
pip install transformers torch sentence-transformers faiss-cpu
```

## Usage

### Running the Application

```bash
cd AuditSearchEngine
streamlit run app/search_app.py
```

The app opens at `http://localhost:8501`

### Quick Start

1. Place documents in the `documents/` folder
2. Click "Re-index with NER" in the sidebar
3. Start searching!

### AI Features

| Feature | How to Use |
|---------|------------|
| **Summarization** | AI Assistant > Summarization > Select document > Generate |
| **Q&A Chat** | AI Assistant > Q&A Chat > Index Embeddings > Ask questions |
| **Anomaly Detection** | AI Assistant > Anomaly Detection > Run |

## Project Structure

```
AuditSearchEngine/
├── app/
│   └── search_app.py      # Main application
├── documents/             # Place documents here
├── .gitignore
├── README.md
└── Audit_Document_Search.md
```

## Tabs Overview

| Tab | Description |
|-----|-------------|
| Search | Main search interface with multiple modes |
| Entities (NER) | View extracted entities |
| Statistics | Document statistics and word counts |
| Table Extraction | Extract and export PDF tables |
| Topics | LDA topic modelling results |
| Sentiment | Risk and sentiment analysis |
| AI Assistant | Summarization, Q&A, Anomaly Detection |

## Version History

- **Phase 1-3**: Core search, NER, document support
- **Phase 4**: PDF table extraction
- **Phase 5**: Topic modelling (LDA)
- **Phase 6**: Sentiment and risk analysis
- **Phase 7**: OCR support for scanned PDFs
- **Phase 9**: AI features (Summarization, Q&A, Anomaly Detection)

## License

Internal audit tool - For authorized use only.

## Support

For issues or feature requests, contact the development team.
