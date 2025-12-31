# Workers' Compensation Processor

A web-based application for processing workers' compensation reports from Foundation Software payroll exports.

## Features

- Upload and process ASR and ArmorPro wage reports
- Automatic class code validation and correction
- Wage-based dual-wage classification (California 6-digit codes)
- Drive time reclassification to lower wage codes
- Employee-specific code corrections
- Combined report generation
- Excel template population
- Download processed files (CSV and Excel)

## Quick Start

### Local Development

1. Clone the repository:
```bash
git clone https://github.com/YOUR_USERNAME/wc-processor-web.git
cd wc-processor-web
```

2. Create a virtual environment:
```bash
python -m venv venv
venv\Scripts\activate  # Windows
# or
source venv/bin/activate  # Linux/Mac
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python app.py
```

5. Open http://localhost:5000 in your browser

### Deployment to Render

1. Push to GitHub
2. Connect repository to Render
3. Deploy using the `render.yaml` configuration

## Usage

1. **Upload ASR Report** - Required wage report from Foundation Software
2. **Upload ArmorPro Report** (Optional) - Second company report if applicable
3. **Select Pay Period** - End date of the pay period
4. **Process Reports** - Click to process and generate output files
5. **Download Files** - Download CSV reports and Excel file

## Class Code System

The application automatically converts 4-digit NCCI codes to California 6-digit codes and validates wage classifications:

| Trade | High Wage Code | Low Wage Code | Threshold |
|-------|---------------|---------------|-----------|
| Carpentry | 543221 | 540321 | $41/hr |
| Wallboard | 544715 | 544615 | $41/hr |
| Painting | 548234 | 547434 | $32/hr |
| Plastering | 548515 | 548415 | $38/hr |
| Sheet Metal | 554222 | 553823 | $33/hr |
| Roofing | 555311 | 555211 | $31/hr |
| Excavation | 622038 | 621837 | $40/hr |

## File Structure

```
wc-processor-web/
├── app.py                 # Flask application
├── processing/            # Data processing modules
│   ├── wage_processor.py  # Core wage processing
│   ├── report_combiner.py # Report combining
│   └── excel_exporter.py  # Excel output
├── templates/             # HTML templates
│   └── index.html
├── static/               # CSS and JavaScript
│   ├── css/
│   └── js/
├── requirements.txt      # Python dependencies
├── render.yaml          # Render deployment config
└── README.md
```

## Technologies

- **Backend**: Flask (Python)
- **Frontend**: Bootstrap 5, vanilla JavaScript
- **Data Processing**: Pandas, OpenPyXL
- **Deployment**: Render.com

## License

Copyright 2025 ASR Construction. All rights reserved.
