# Consultant Performance Evaluation System

This system automates the process of collecting, formatting, and distributing consultant performance evaluations at Devoteam.

## Features

- Automated data collection from Google Sheets
- AI-powered evaluation formatting using GPT-4
- PDF report generation
- Automated email distribution to relevant stakeholders
- Special handling for Digital Infrastructure (DI) team evaluations

## Prerequisites

- Python 3.8+
- Google Sheets API access
- OpenAI API key
- Gmail account with App Password

## Installation

1. Clone the repository:
```bash
git clone [repository-url]
cd consultants_performance_form
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Set up environment variables:
   - Create a `.env` file with the following variables:
   ```
   OPENAI_API_KEY=your_openai_api_key
   EMAIL_SENDER=your_email
   EMAIL_PASSWORD=your_app_password
   SPREADSHEET_ID=your_spreadsheet_id
   ```

4. Set up Google Sheets API:
   - Create a service account and download the credentials JSON file
   - Rename it to `parser_SA.json` and place it in the project root

## Usage

Run the main script:
```bash
python main.py
```

The script will:
1. Fetch evaluation data from Google Sheets
2. Format evaluations using GPT-4
3. Generate PDF reports
4. Send emails to relevant stakeholders

## Project Structure

- `main.py`: Main application logic
- `test.py`: Testing utilities
- `parser_SA.json`: Google Sheets API credentials
- `df_evaluation.csv`: Temporary data storage

## Security Notes

- Never commit sensitive credentials to the repository
- Keep API keys and service account credentials secure
- Use environment variables for sensitive information

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request 