# python-filtering-to-csv

This project processes RSS feed data from Feedly, filters articles based on keywords, and exports the results to Excel files.

## Setup Instructions

1. **Clone the Repository**
   ```bash
   git clone https://github.com/Yampanis/python-filtering-to-csv.git
   cd python-filtering-to-csv
   ```

2. **Create Virtual Environment**
   ```bash
   python -m venv venv
   ```

3. **Activate Virtual Environment**
   - Windows:
   ```bash
   .\venv\Scripts\activate
   ```
   - Unix/MacOS:
   ```bash
   source venv/bin/activate
   ```

4. **Install Dependencies**
   ```bash
   pip install setuptools
   pip install -r requirements.txt
   ```

5. **Configure Chrome WebDriver**
   - Download ChromeDriver matching your Chrome browser version from:
     https://sites.google.com/chromium.org/driver/
   - Add ChromeDriver to your system PATH or project directory

6. **Set Up Environment Variables**
   Create a `.env` file in the project root:
   ```plaintext
   CHROMEDRIVER_PATH=path/to/your/chromedriver.exe
   ```

7. **Initialize Excel Files**
   The script will automatically create these Excel files:
   - `Rory Testing Sheet 2024.xlsx`: Main output. Append new 'cleaned' tities. Truncated by user.
   - `titles_to_check.xlsx`: Running list of titles used to remove duplicates from new titles.
   - `negative_titles.xlsx`: Titles filtered from out as 'negative'.
   - `negatives.xlsx`: List of words that establish title as negative.

## Usage

Run the script:
```bash
python Feedly.py (from within virtual environment)
```
- Windows:
```bash
runFeedly.bat (from cmd or scheduler)
```

## Project Structure
```
python-filtering-to-csv/
├── Feedly.py           # Main script
├── requirements.txt    # Dependencies
├── .env               # Environment variables
└── README.md         # Documentation
```

## Dependencies
- Python 3.8+
- Selenium
- Pandas
- openpyxl
- python-dotenv
- webdriver_manager

## Notes
- Make sure Chrome browser is installed
- Keep ChromeDriver version compatible with Chrome
- Clear browser cache if encountering connection issues
