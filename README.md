# Outlook Miner

![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

A powerful email forwarding automation tool for Microsoft Outlook that helps you automatically forward emails from your Sent Items folder based on configurable filters. Useful for forwarding sent email to documentation systems where auto-forwarding is not available to the user.

## Features

- **Smart Email Filtering**: Filter emails by subject keywords, date ranges, and file number prefixes
- **Automated Forwarding**: Automatically forward matching emails to designated recipients
- **Duplicate Prevention**: Track forwarded emails to avoid sending duplicates
- **File Number Extraction**: Extract file numbers from attachments or email subjects
- **Preview Mode**: Preview matching emails before forwarding
- **Multi-threaded Operation**: Responsive GUI with background processing
- **Configuration Management**: Save and manage multiple recipient configurations
- **Comprehensive Logging**: Detailed logging with timestamps for audit trails
- **Rate Limiting**: Configurable delays between forwarded emails

## Requirements

- Windows OS (tested on Windows 10/11)
- Microsoft Outlook installed and configured
- Python 3.7 or higher

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/royalpayne/OutlookMiner.git
   cd OutlookMiner
   ```

2. **Install required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Optional: Place your custom icon:**
   - Add a `myicon.ico` file to the project directory for custom branding

## Usage

### Starting the Application

Run the application using Python:

```bash
python outlook_miner.py
```

Or use the compiled executable (if available):

```bash
outlook_miner.exe
```

### Configuration Steps

1. **Enter Forward To Email Address**:
   - Type or select the recipient email address from the dropdown

2. **Set Subject Keyword**:
   - Enter the keyword to search for in email subjects (e.g., "BILLING INVOICE")

3. **Select Date Range**:
   - Choose start and end dates for the email search

4. **Optional File Number Prefixes**:
   - Enter comma-separated numeric prefixes (e.g., "759,123")
   - Only emails with matching file numbers will be processed

5. **Configure Options**:
   - **Require Attachments**: Only forward emails with attachments
   - **Skip Previously Forwarded Emails**: Avoid forwarding duplicates
   - **Delay (Sec.)**: Add delay between forwarded emails

6. **Save Configuration**:
   - Click "Save Config" to store your settings for future use

### Operations

#### Preview Mode
Click **Preview** to see a list of emails that match your criteria without forwarding them.

#### Scan and Forward
Click **Scan and Forward** to automatically forward all matching emails to the configured recipient.

#### Cancel Operation
Click **Cancel** to stop an ongoing search or forward operation.

## How It Works

### Email Scanning Process

When you click **Preview** or **Scan and Forward**, the application performs the following steps:

1. **Connect to Outlook**: Establishes a COM connection to Microsoft Outlook using the Windows API
2. **Access Sent Items**: Opens your Sent Items folder and retrieves the email count
3. **Apply Subject Filter**: Uses Outlook's MAPI filter to find emails containing your subject keyword (case-insensitive)
4. **Date Range Filtering**: Checks each email's sent date against your specified date range
5. **File Number Extraction**: If prefixes are specified, extracts file numbers from attachment filenames or email subjects
6. **Duplicate Check**: If "Skip Previously Forwarded" is enabled, checks the database for previously forwarded file numbers
7. **Attachment Check**: If "Require Attachments" is enabled, skips emails without attachments

### Forward Process

When forwarding emails, the application:

1. Creates a forward copy of the matching email
2. Sets the recipient to your configured "Forward To" address
3. **Replaces the subject line** with the extracted file number (if found) or keeps the original subject
4. Sends the forwarded email
5. Logs the file number to the database to prevent future duplicates
6. Applies the configured delay between emails (minimum 3 seconds for date ranges > 8 days)

### Configuration Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| **Forward To** | Email address where matching emails will be forwarded | Required |
| **Subject Keyword** | Text to search for in email subjects (case-insensitive) | "BILLING INVOICE" |
| **Start Date** | Beginning of the date range to search | Today |
| **End Date** | End of the date range to search | Today |
| **File Number Prefixes** | Comma-separated numeric prefixes (e.g., "759,123") to filter and extract file numbers | Empty (all emails) |
| **Delay (Sec.)** | Seconds to wait between forwarding each email | 0 |
| **Require Attachments** | Only forward emails that have attachments | Checked |
| **Skip Previously Forwarded** | Skip emails with file numbers already in the tracking database | Checked |

### File Number Extraction

File numbers are extracted using the following priority:

1. **From Attachments**: Scans attachment filenames for patterns matching your prefixes (e.g., "759-12345.pdf" extracts "759-12345")
2. **From Subject**: If no attachment match, searches the email subject for matching patterns

The extracted file number becomes the new subject line when forwarding, making it easy to identify and sort forwarded emails.

### Rate Limiting

To prevent overwhelming the mail server:

- **Manual Delay**: Configure any delay in the Configuration dialog
- **Automatic Delay**: A 3-second minimum delay is automatically applied when the date range exceeds 8 days
- **Recommended**: Use 1-3 second delays for large batches of emails

## Database

The application uses SQLite database (`minerdb.db`) with two tables:

### Clients Table
Stores configuration for each recipient:
- recipient (email address)
- start_date, end_date
- file_number_prefix
- subject_keyword
- require_attachments, skip_forwarded
- delay_seconds

### ForwardedEmails Table
Tracks forwarded emails to prevent duplicates:
- file_number
- recipient
- forwarded_at (timestamp)

## Logging

The application maintains several log files:

- **GUI Log Tab**: Real-time logging visible in the application
- **outlook_miner_startup.log**: Startup messages before GUI initialization
- **forwarded_emails.log**: Record of all forwarded emails with timestamps

All timestamps use US/Eastern timezone.

## Safety Features

- **Thread-Safe Operations**: All database and GUI operations are thread-safe
- **Retry Logic**: Automatic retry for Outlook connection failures
- **Input Validation**: Email format and date validation
- **Filter Sanitization**: Protection against MAPI filter injection
- **Performance Monitoring**: Warns if operations take too long

## Date Range Behavior

⚠️ **Important**: If you select a date range exceeding 8 days, the application automatically applies a 3-second delay between forwarded emails to prevent Outlook throttling.

## Troubleshooting

### Outlook Connection Issues
- Ensure Microsoft Outlook is installed and configured
- Try closing and reopening Outlook
- Check if Outlook is set as the default mail client

### No Emails Found
- Verify the subject keyword matches emails in Sent Items
- Adjust the date range to include relevant emails
- Check if "Require Attachments" should be unchecked
- Disable "Skip Previously Forwarded Emails" to resend

### Slow Performance
- Reduce the date range for faster searching
- Use more specific subject keywords
- Close other applications to free up resources

## Building Executable

To build a standalone executable using PyInstaller:

```bash
pyinstaller outlook_miner.spec
```

The executable will be created in the `dist` folder.

## Project Structure

```
OutlookMiner/
├── outlook_miner.py       # Main application file
├── outlook_miner.spec     # PyInstaller specification
├── myicon.ico             # Application icon
├── requirements.txt       # Python dependencies
├── README.md             # This file
├── minerdb.db            # SQLite database (created on first run)
└── *.log                 # Log files (created during operation)
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see below for details:

```
MIT License

Copyright (c) 2024 Royal Payne

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## Author

**Royal Payne**

## Acknowledgments

- Built with [pywin32](https://github.com/mhammond/pywin32) for Outlook integration
- GUI built with [tkinter](https://docs.python.org/3/library/tkinter.html)
- Date picker from [tkcalendar](https://github.com/j4321/tkcalendar)

## Support

For issues, questions, or suggestions, please open an issue on the [GitHub repository](https://github.com/royalpayne/OutlookMiner/issues).

---

**Note**: This tool is designed for Windows environments with Microsoft Outlook. It accesses your Outlook Sent Items folder and requires appropriate permissions to send emails on your behalf.
