# GoReport - Gophish Campaign Reporting Tool

**A fork of the original [GoReport](https://github.com/chrismaddalena/GoReport) project with added functionality for multi-campaign and timeline reporting.**

GoReport connects to Gophish servers via API to extract and analyze phishing campaign data. It generates comprehensive reports with detailed statistics, user behavior tracking, and timeline analysis. This fork focuses on improving multi-campaign workflows with better timeline reports, proper event tracking, and flexible output options.

## What This Tool Does

GoReport processes Gophish campaign results to provide:
- **Campaign Statistics**: Total events vs unique recipients for clicks, opens, and submissions
- **Timeline Reconstruction**: Chronological event logs with timestamps
- **Multi-Campaign Analysis**: Compare and analyze multiple campaigns together

## Whats different?

### Timeline Reports
- Each campaign sheet includes three organized sections:
  1. **Campaign Details**: Subject line, phishing URL, and launch time
  2. **Click Statistics**: Per-user click counts in a summary table
  3. **Timeline Data**: Detailed chronological event log
- Default behavior creates separate sheets for each campaign
- `--combine-campaigns` option places all campaigns sequentially in one sheet

### Custom RID
- If you have modified the default RID in your gophish instance, then you can specify a custom one when parsing results with `--rid-field`.

### Added CLI arguments
- `--timeline`: Generate focused timeline reports
- `--combine-campaigns`: Control how multiple campaigns are displayed
- `--output`: Specify output directoy/filename for the generated report.
- `--rid-field`: Specify custom reference ID if you have modified the default RID in your gophish instance.


## Quick Start with uv

```bash
# Install uv if you haven't already
curl -LsSf https://astral.sh/uv/install.sh | sh

# Clone and setup the project
git clone <repository-url>
cd Goreport

# Install dependencies with uv
uv pip install -r requirements.txt

# Copy and configure your settings
cp gophish.config.sample gophish.config
# Edit gophish.config with your API key and server details

# Run a basic report
uv run python GoReport.py --id 1 --format excel

# Generate timeline report with multiple campaigns
uv run python GoReport.py --id 1,2,3 --format excel --timeline
```

## Requirements

* Python 3.10+
* Gophish server with API access
* Dependencies managed via `pyproject.toml`:
  - gophish (API client)
  - xlsxwriter (Excel reports)
  - python-docx (Word reports)
  - click (CLI interface)
  - requests, user-agents, python-dateutil

## Setup

1. **Configure API Access**:
   ```bash
   cp gophish.config.sample gophish.config
   # Edit gophish.config with your credentials
   ```

2. **SSH Port Forwarding** (for remote servers):
   ```bash
   ssh -L 3333:localhost:3333 user@gophish-server
   ```

3. **Word Template** (optional):
   Place a `template.docx` file in the project root with your preferred styles

## Usage Examples

### Basic Reports

```bash
# Single campaign Excel report
uv run python GoReport.py --id 26 --format excel

# Multiple campaigns (comma-separated or ranges)
uv run python GoReport.py --id 26,29-33,54 --format excel

# Quick terminal output for campaign status
uv run python GoReport.py --id 26 --format quick

# Word document report
uv run python GoReport.py --id 26 --format word
```

### Timeline Reports

```bash
# Single campaign timeline
uv run python GoReport.py --id 1 --format excel --timeline

# Multiple campaigns - separate sheets (default)
uv run python GoReport.py --id 1,2,3 --format excel --timeline

# Multiple campaigns - combined in one sheet
uv run python GoReport.py --id 1,2,3 --format excel --timeline --combine-campaigns
```

### Additional Options

```bash
# Custom output path
uv run python GoReport.py --id 26 --format excel --output reports/phishing_Q4.xlsx

# Mark campaigns as complete after reporting
uv run python GoReport.py --id 26,27 --format excel --complete

# Use custom reference ID field
uv run python GoReport.py --id 26 --format excel --rid-field "employee_id"

# Use alternate config file
uv run python GoReport.py --id 26 --format excel --config production.config
```

## Configuration

### Config File Structure

```ini
[Gophish]
gp_host: https://127.0.0.1:3333
api_key: <YOUR_API_KEY>

[ipinfo.io]
# Optional: Free tier allows 1000 requests/day
ipinfo_token: <IPINFO_TOKEN>

[Google]
# Optional: For enhanced geolocation ($0.005/request)
geolocate_key: <GOOGLE_API_KEY>
```

### Multiple Configurations

Manage multiple Gophish servers or accounts:

```bash
# Create separate config files
cp gophish.config production.config
cp gophish.config testing.config

# Use specific config
uv run python GoReport.py --id 26 --format excel --config production.config
```

## Output Formats

### Excel Reports (.xlsx) - Recommended
Generate comprehensive workbooks with multiple worksheets:

**Standard Report Contents:**
- Campaign overview and settings
- Detailed recipient results with outcomes
- Browser and OS statistics
- IP addresses and geolocation data
- Complete event timeline

**Timeline Report Mode (`--timeline`):**
- Focused view on user interactions
- Three sections per campaign:
  - Campaign details (subject, URL, launch time)
  - Click statistics table (user email + click count)
  - Chronological event timeline
- Separate sheets per campaign (default)
- Sequential sections in one sheet (`--combine-campaigns`)

### Word Reports (.docx)
- Professional formatted documents
- Requires `template.docx` with "Goreport" table style
- Suitable for executive presentations
- All statistics and summaries included

### Quick Reports (Terminal)
- Instant campaign status check
- Basic statistics output
- No file generation
- Useful for monitoring active campaigns

## Command-Line Options

| Option | Description | Example |
|--------|-------------|---------|
| `--id` | Campaign ID(s) | `--id 1,2,5-10` |
| `--format` | Output format | `--format excel` |
| `--timeline` | Timeline-focused report | `--timeline` |
| `--combine-campaigns` | Merge campaigns into single sheet | `--combine-campaigns` |
| `--combine` | Merge campaigns into one report | `--combine` |
| `--complete` | Mark campaigns complete | `--complete` |
| `--output` | Custom output path | `--output reports/Q4.xlsx` |
| `--rid-field` | Custom reference ID field | `--rid-field employee_id` |
| `--config` | Alternate config file | `--config prod.config` |
| `--google` | Use Google Maps API | `--google` |

## Project Structure

```
Goreport/
├── GoReport.py           # Main CLI entry point
├── lib/
│   ├── goreport.py      # Core reporting logic
│   └── banners.py       # CLI banners
├── pyproject.toml       # Project metadata and dependencies
├── requirements.txt     # Python dependencies
├── gophish.config.sample # Configuration template
├── template.docx        # Word report template (optional)
└── output/             # Default report output directory
```
