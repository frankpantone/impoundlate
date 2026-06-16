# Outlook Email Automation

This project allows you to automate sending emails from your Microsoft Outlook account using Python. It uses the Outlook COM interface, which works with your existing Outlook installation.

## Prerequisites

- **Microsoft Outlook** must be installed and running on your Windows machine
- **Python 3.6+** installed
- An active email account configured in Outlook

## Installation

1. Install the required Python package:
```bash
pip install -r requirements.txt
```

Or install directly:
```bash
pip install pywin32
```

## Quick Start

### Send a Test Email

Run the simple test script:
```bash
python send_test_email.py
```

This will:
1. Connect to your running Outlook application
2. Display your current email account
3. Allow you to send a test email to yourself or another address

### Using the Email Sender Class

```python
from outlook_email_sender import OutlookEmailSender

# Create email sender instance
email_sender = OutlookEmailSender()

# Send a simple email
email_sender.send_email(
    to_recipients=["recipient@example.com"],
    subject="Hello from Python!",
    body="This email was sent using Python automation."
)

# Send an email with CC, BCC, and attachments
email_sender.send_email(
    to_recipients=["primary@example.com"],
    cc_recipients=["cc@example.com"],
    bcc_recipients=["bcc@example.com"],
    subject="Email with attachments",
    body="Please find the attached files.",
    attachments=["C:/path/to/file.pdf", "C:/path/to/image.jpg"]
)

# Send HTML email
html_content = """
<h1>HTML Email</h1>
<p>This is an <b>HTML formatted</b> email.</p>
<ul>
    <li>Item 1</li>
    <li>Item 2</li>
</ul>
"""

email_sender.send_email(
    to_recipients=["recipient@example.com"],
    subject="HTML Email Test",
    body=html_content,
    html_body=True
)
```

## Features

- ✅ Send emails using your existing Outlook account
- ✅ Support for TO, CC, and BCC recipients
- ✅ File attachments
- ✅ HTML and plain text email bodies
- ✅ Error handling and logging
- ✅ Get current user's email address
- ✅ Interactive test script

## API Reference

### OutlookEmailSender Class

#### `__init__()`
Initializes connection to the Outlook application.

#### `send_email(to_recipients, subject, body, cc_recipients=None, bcc_recipients=None, attachments=None, html_body=False)`
Sends an email through Outlook.

**Parameters:**
- `to_recipients` (List[str]): List of primary recipient email addresses
- `subject` (str): Email subject line
- `body` (str): Email body content
- `cc_recipients` (List[str], optional): List of CC recipient email addresses
- `bcc_recipients` (List[str], optional): List of BCC recipient email addresses
- `attachments` (List[str], optional): List of file paths to attach
- `html_body` (bool, optional): Whether the body is HTML formatted (default: False)

**Returns:** `bool` - True if email was sent successfully

#### `send_test_email(test_recipient)`
Sends a predefined test email to verify the setup.

#### `get_current_user_email()`
Returns the email address of the current Outlook user.

## Troubleshooting

### Common Issues

1. **"Failed to connect to Outlook"**
   - Make sure Microsoft Outlook is running
   - Ensure you have an email account configured in Outlook

2. **"Permission denied" or COM errors**
   - Run your Python script as Administrator
   - Check Windows security settings for COM applications

3. **"Module not found: win32com"**
   - Install pywin32: `pip install pywin32`
   - On some systems, you may need: `python -m pip install pywin32`

4. **Emails not sending**
   - Check your internet connection
   - Verify your Outlook account is properly configured
   - Check Outlook's Outbox for stuck emails

### Security Considerations

- This script uses the same permissions as your Outlook application
- Emails are sent through your configured Outlook account
- No credentials are stored or transmitted by this script
- The script operates within Outlook's security context

## Examples

See `send_test_email.py` for a complete interactive example that demonstrates:
- Connecting to Outlook
- Sending test emails
- Custom email composition
- Error handling

## License

This project is provided as-is for educational and automation purposes.







