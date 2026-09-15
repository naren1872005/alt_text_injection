# PDF Accessibility Checker

A powerful web-based tool for extracting figures and alternative text from PDF accessibility structure trees and Excel workbooks, with visual matching and alt text injection capabilities.

## Features

### 📄 PDF Figure Extraction
- **StructTreeRoot Parsing**: Extracts all `/Figure` tags from PDF accessibility structure tree
- **High-Resolution Crops**: Renders figure crops at 150 DPI for detailed inspection
- **Bounding Box Detection**: Computes precise bounding boxes from MCID (Marked Content IDs)
- **Alt Text Status**: Identifies figures with existing `/Alt` attributes vs. missing alt text
- **Page & MCID Tracking**: Maps each figure to its page number and content IDs

### 📊 Excel Manifest Processing
- **Drawing Extraction**: Automatically extracts all embedded images from Excel workbooks
- **Authoritative Alt Text**: Parses alt text from Excel image metadata
- **Manifest Parsing**: Processes complete image records with sr. no., filenames, and alt text
- **Image Gallery**: Creates an interactive visual gallery of all extracted images

### 🔄 Visual Matching
- **Automatic Alignment**: Uses visual similarity to match PDF figures with Excel images
- **Confidence Scores**: Provides matching confidence percentages for each match
- **Side-by-Side Inspection**: View matched pairs side-by-side for validation
- **Smart Assignment**: Automates alt text assignment from Excel to matched PDF figures

### ⚡ Alt Text Injection
- **Batch Injection**: Inject alt text for all matched figures in one operation
- **Single Figure Injection**: Inject alt text for individual figures
- **PDF Regeneration**: Creates accessible PDFs with injected alt text
- **Progressive Enhancement**: Preserves existing alt text while adding missing ones

### 📥 Download & Export
- **ZIP Exports**: Download all figures with CSV manifest
- **JSON Export**: Full session data export for integration
- **Excel Records Download**: Get all extracted Excel images with alt text metadata
- **Accessible PDF Download**: Download the enhanced PDF with injected alt text

## Technology Stack

### Backend
- **FastAPI** - Modern Python web framework
- **PyPDF2** - PDF manipulation and StructTreeRoot parsing
- **PIL/Pillow** - Image processing and crop rendering
- **OpenPyXL** - Excel workbook parsing
- **OpenCV** - Visual similarity matching

### Frontend
- **HTML5** - Semantic markup
- **CSS3** - Modern styling with Grid & Flexbox
- **Vanilla JavaScript** - No framework dependencies
- **Responsive Design** - Works on desktop and tablet

## Installation

### Requirements
- Python 3.8+
- pip package manager

### Setup

1. **Clone/Download the project**
   ```bash
   cd D:\pdf_alt_textconverter
   ```

2. **Create a virtual environment** (recommended)
   ```bash
   python -m venv venv
   source venv/Scripts/activate  # On Windows
   # or
   source venv/bin/activate      # On macOS/Linux
   ```

3. **Install dependencies**
   ```bash
   pip install fastapi uvicorn python-multipart pillow opencv-python openpyxl pypdf
   ```

4. **Run the server**
   ```bash
   python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload
   ```

5. **Open in browser**
   Navigate to `http://localhost:8000`

## Usage

### Workflow 1: Extract Figures from PDF

1. **Upload a PDF** via drag-and-drop or file browser
2. **View Extracted Figures**:
   - See all `/Figure` tags found in the structure tree
   - View high-resolution crops of each figure
   - Check alt text status (has alt / missing alt)
3. **Download Results**:
   - ZIP with all figure PNGs + CSV manifest
   - JSON export of full extraction data

### Workflow 2: Process Excel Manifest

1. **Upload an Excel Workbook** (.xlsx or .xlsm)
2. **Browse Extracted Images**:
   - View all embedded drawing images
   - Read authoritative alt text from Excel
   - Paginate through large manifests (4,848+ images)
3. **Download Results**:
   - ZIP with all extracted images + CSV manifest
   - Excel records with alt text metadata

### Workflow 3: Match & Inject

1. **Upload PDF** to extract figures
2. **Upload Excel Manifest** with alt text
3. **Automatic Matching**:
   - Visual matcher aligns PDF figures with Excel images
   - Displays confidence scores for each match
4. **Inject Alt Text**:
   - Inject all matched alt text at once
   - Or selectively inject per-figure
   - Download the enhanced accessible PDF

### Workflow 4: Use Sample Data

- **Sample PDF**: Loads Plesha Engineering Mechanics Chapter 15
- **Sample Excel**: Loads associated Excel manifest with 4,848 image records
- **Quick Test**: "Load Sample (Chap 15 PDF + Excel Manifest)" button combines both

## API Endpoints

### PDF Operations
- `POST /api/upload-pdf` - Upload and extract figures from PDF
- `POST /api/load-sample` - Load sample PDF
- `GET /api/figure-image/{session_id}/{filename}` - Retrieve figure crop image
- `GET /api/download-zip/{session_id}` - Download ZIP with all figures + CSV

### Excel Operations
- `POST /api/upload-excel` - Upload and process Excel manifest
- `POST /api/load-sample-excel` - Load sample Excel manifest
- `GET /api/excel-image/{session_id}/{filename}` - Retrieve Excel image
- `GET /api/excel-records/{session_id}` - Paginated Excel records with search
- `GET /api/download-excel-zip/{session_id}` - Download ZIP with Excel images + CSV

### Alt Text Injection
- `POST /api/inject-alt/{session_id}` - Inject all alt texts into PDF
- `POST /api/inject-single-alt/{session_id}` - Inject single figure alt text
- `GET /api/download-injected-pdf/{session_id}` - Download accessible PDF with injected alt

### Utility
- `GET /api/health` - Health check endpoint
- `GET /api/download-json/{session_id}` - Export session as JSON

## Project Structure

```
pdf_alt_textconverter/
├── app.py                 # FastAPI application & endpoints
├── extractor.py          # PDF figure extraction & alt injection
├── excel_parser.py       # Excel image & metadata extraction
├── matcher.py            # Visual similarity matching algorithm
├── static/
│   ├── index.html        # Frontend markup
│   ├── app.js            # Frontend logic & UI interactions
│   └── style.css         # Styling & responsive design
├── sessions/             # Runtime session directories
│   └── {session_id}/
│       ├── input.pdf
│       ├── figures/      # Extracted PDF figure crops
│       ├── manifest.xlsx # Uploaded Excel file
│       └── excel_images/ # Extracted Excel images
└── README.md             # This file
```

## Key Classes & Modules

### FigureExtractor (`extractor.py`)
Handles PDF parsing and figure extraction:
- `extract_figures_with_crops()` - Extracts all `/Figure` tags and renders crops
- `inject_alt_texts()` - Injects alt text into StructTreeRoot

### ExcelParser (`excel_parser.py`)
Processes Excel workbooks:
- `parse_all_records()` - Extracts all image records with alt text
- Supports drawing object extraction from Excel sheets

### VisualMatcher (`matcher.py`)
Matches PDF figures to Excel images:
- Computes visual similarity using image hashing or feature matching
- Returns confidence scores for each potential match
- Enables intelligent alt text assignment

## Session Management

The application uses in-memory sessions to track:
- **Session ID**: Unique identifier for each user workflow
- **Figures**: Extracted PDF figures with metadata
- **Excel Records**: Extracted Excel images & alt text
- **Matches**: Visual matches between PDF and Excel images
- **Metrics**: Alt text coverage statistics

Sessions are stored in `sessions/{session_id}/` with:
- Input PDF file
- Extracted figure crops (PNG)
- Uploaded Excel manifest
- Extracted Excel images
- Generated accessible PDFs

## Configuration & Deployment

### Development
```bash
python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload
```

### Production
```bash
python -m uvicorn app:app --host 0.0.0.0 --port 8000 --workers 4
```

### Docker (Optional)
Create a `Dockerfile` for containerized deployment:
```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["python", "-m", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
```

## Troubleshooting

### Common Issues

**PDF Not Extracting Figures**
- Ensure the PDF has a `/StructTreeRoot` with `/Figure` tags
- PDF must be tagged for accessibility (UA flag must be set)
- Check console for detailed error messages

**Excel Images Not Showing**
- Verify the Excel file contains embedded drawing objects
- Supported formats: .xlsx, .xlsm
- Check that images are actual embedded objects, not linked

**Visual Matching Low Accuracy**
- Ensure PDF figures and Excel images are similar visual content
- Different scales or rotations may reduce match confidence
- Manual inspection and override is always available

### Debug Mode
Set environment variables for more verbose output:
```bash
set DEBUG=True
python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload
```

## Performance Notes

- **Large PDFs** (100+ pages): May take 30-60 seconds to extract all figures
- **Excel Manifests** (4,000+ images): Pagination is recommended for browsing
- **Visual Matching**: Slower for high-resolution images; adjust DPI if needed
- **Memory**: Runs in-memory; consider the session cleanup for production

## Browser Support

- Chrome/Edge 90+
- Firefox 88+
- Safari 14+
- Mobile browsers (iOS Safari, Chrome Mobile)

## License

This project is provided as-is for educational and professional use.

## Author

Created for PDF accessibility enhancement and figure-to-alt-text automation workflows.

---

For questions or issues, review the debug console output or examine the FastAPI OpenAPI docs at `http://localhost:8000/docs`
