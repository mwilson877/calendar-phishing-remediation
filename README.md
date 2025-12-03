# Calendar Event Remediation Tool

A Python-based remediation tool designed to help security teams and IT administrators quickly identify and remove malicious calendar invites from user mailboxes. This tool was primarily developed to combat phishing attacks that leverage calendar invitations to bypass email security filters.

## Primary Purpose

Phishing attacks increasingly use calendar invitations as a vector because:
- Calendar invites often bypass traditional email security filters
- Users may be more likely to trust calendar notifications
- Malicious links in calendar events can lead to credential theft or malware

This tool provides a streamlined way to search for and remove suspicious calendar events across user mailboxes, making incident response faster and more efficient.

## Features

- **Targeted Search**: Filter calendar events by date range, sender email, and subject keywords
- **Recurring Event Support**: Properly handles both single events and recurring meeting series
- **Batch Remediation**: Review multiple suspicious events before taking action
- **User Confirmation**: Requires explicit confirmation before deleting any calendar event
- **Chronological Display**: Shows events sorted by time for easier review
- **Interactive CLI**: Simple command-line interface for easy operation

## Other Use Cases

While designed for security remediation, this tool can also be used for:
- Bulk calendar cleanup
- Removing outdated or cancelled meetings
- Managing calendar events for offboarded users
- General calendar maintenance tasks

## Prerequisites

- Python 3.6 or higher
- Microsoft Graph API access with appropriate permissions
- Azure AD application registration with client credentials

### Required Microsoft Graph API Permissions

- `Calendars.ReadWrite` (Application permission)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/YOUR_USERNAME/calendar-remediation-tool.git
cd calendar-remediation-tool
```

2. Install required dependencies:
```bash
pip install requests
```

3. Create a `config.ini` file in the project directory:
```ini
[credentials]
client_id = YOUR_CLIENT_ID
client_secret = YOUR_CLIENT_SECRET
tenant_id = YOUR_TENANT_ID
```

## Azure AD Application Setup

1. Register a new application in [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Create a client secret under **Certificates & secrets**
5. Grant **Calendars.ReadWrite** application permission under **API permissions**
6. Have an administrator grant consent for the permissions

## Usage

Run the script:
```bash
python calendar_remediation.py
```

You will be prompted for:
- **Mailbox**: The user's email address (e.g., user@company.com)
- **Start Date**: Beginning of date range in YYYY-MM-DD format
- **End Date**: End of date range in YYYY-MM-DD format
- **Sender** (optional): Filter by organizer email address
- **Subject Contains** (optional): Filter by text in event subject

The tool will display matching events with:
- Event type (Single Event or Recurring Meeting)
- Time range in UTC
- Sender email
- Subject line

Select an event by number to review and confirm deletion.

## Example Workflow

```
Enter mailbox (user@domain.com): victim@company.com
Enter start date (yyyy-mm-dd): 2025-12-01
Enter end date (yyyy-mm-dd): 2025-12-05
Enter Sender (user@domain.com) [optional]: attacker@malicious.com
Enter text to search in subject [optional]: urgent

Calendar Events:
1. [Recurring Meeting] 9:00AM-9:30AM UTC | Sender: attacker@malicious.com | Subject: URGENT: Update Required
2. [Single Event] 2:00PM-2:30PM UTC | Sender: attacker@malicious.com | Subject: urgent action needed

Enter the number of the event you want to delete (or 'q' to quit): 1

You are about to delete this calendar invite:
Type: Recurring Meeting
Time: 9:00AM-9:30AM UTC
Sender: attacker@malicious.com
Subject: URGENT: Update Required
Do you want to continue? (y/n): y
Successfully deleted the calendar event!
```

## Security Considerations

- **Never commit `config.ini`** to version control
- Store credentials securely and rotate them regularly
- Use least-privilege access - only grant necessary permissions
- Log all remediation actions for audit purposes
- Verify the sender/subject before deletion to avoid removing legitimate events

## .gitignore Recommendation

Add to your `.gitignore`:
```
config.ini
*.ini
__pycache__/
*.pyc
```

## License

MIT License - See LICENSE file for details

## Contributing

Contributions welcome! Please open an issue or submit a pull request.

## Disclaimer

This tool deletes calendar events permanently. Always verify events before deletion. The authors are not responsible for data loss due to misuse of this tool.
