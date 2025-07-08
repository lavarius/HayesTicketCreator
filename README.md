# HayesTicketCreator

# Hayes gethelphss Ticket Automation

This repository provides a Python Selenium automation script to bulk-create help desk tickets in the Hayes gethelphss system based on data from an Excel file. It is designed to streamline repetitive ticket creation for devices like projectors, desktops / laptops, ensuring consistency and saving time for IT staff. Automate creation of several tickets without API, since an API wasn't available.

# Features

- Reads device and tag data from an Excel spreadsheet

- Automates browser actions to:

  - Log in via Google SSO (with Chrome profile support)

  - Navigate the multi-step ticket creation UI (Technology → Devices → Projectors / Speakers)

  - Fill out all required ticket fields, including dynamic summary and description

  - Assign tickets to a specific user and submitter

  - Submits tickets reliably, handling form validation and overlays

# Prerequisites

Python 3.8+

Google Chrome installed

ChromeDriver matching your Chrome version (in your System Variables' PATH)

Selenium, pandas, python-dotenv, openpyxl installed in your environment

```bash
pip install selenium pandas openpyxl python-dotenv
```

# Setup

## 1. Clone the Repository

```bash
git clone https://github.com/your-org/hayes-ticket-automation.git
cd hayes-ticket-automation
```

## 2. Prepare Your Chrome Profile

Create a dedicated Chrome user profile for automation, or use your existing one.

The script expects a persistent profile directory (see .env below).

## 3. Configure Environment Variables

Create a `.env` file in your project root:

```plaintext
USER_EMAIL=your.email@domain.com
PASSWORD=your_google_password
PROFILE_BASE=C:/Users/yourusername/selenium_profile

```

- Do not use USERNAME as a variable name; it conflicts with system variables.

- Make sure `.env` is in your `.gitignore` to keep credentials private.

## 4. Prepare Your Excel File

- Place your device/tag Excel file in the information/ directory.

- The file should have at least these columns: Projector Name, Tag.

Example:

```text
| Projector Name   | Tag    |
|------------------|--------|
| BUR-RM01-PROJ    | 769452 |
| BUR-RM02-PROJ    | 769453 |

```

# Usage

## 1. Activate Your Virtual Environment

```bash
python -m venv venv
# On Windows:
venv\Scripts\activate
# On Mac/Linux:
source venv/bin/activate
```

## 2. Install Dependencies

```bash
pip install -r requirements.txt
```

## 3. Run the Script

```bash
python main.py
```

- The script will launch Chrome, log in via SSO, and begin creating tickets as specified in your Excel file.

- Progress and any errors will be printed to the console.

# Customization

- Ticket Fields:
  Edit the script to adjust summary, description, or other fields as needed.

- Selectors:
  If the UI changes, update XPaths or IDs in the script for robustness.

- Rate Limiting:
  Adjust time.sleep() values if you encounter rate-limiting or want to throttle submissions.

# Troubleshooting

- Element Not Found/Clickable:

  - Ensure ChromeDriver matches your Chrome version.

  - Check that all form fields are being filled as required.

  - Use screenshots (the script can save them) to debug UI state.

- SSO Issues:

  - Make sure your Chrome profile is set up and not in use elsewhere.

  - Credentials in .env must be accurate and complete.

# Contributing

1. Fork this repository
2. Create a new branch (git checkout -b feature-branch)
3. Make your changes and add tests if applicable
4. Commit and push (git commit -am 'Add new feature')
5. Open a Pull Request

# Security & Privacy

- Never commit `.env` or credentials to the repository.
- Use a dedicated Chrome profile for automation to avoid interfering with personal browsing.

# License

This project is licensed under the MIT License.

# Credits

- Inspired by real-world IT automation needs for K-8 device management / ticket handling.
- Uses Selenium, pandas, and python-dotenv for robust, maintainable automation.

# Support

For issues, open a GitHub issue or contact the repository maintainer. Contributions and improvements are welcome!

# Required file

`.env` Structure

```bash
USER_EMAIL=
password=
PROFILE_PATH=
```
