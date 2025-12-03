import configparser
import requests
from datetime import datetime

def get_access_token(tenant_id, client_id, client_secret):
    """
    Authenticates to Microsoft Graph API using client credentials grant and returns an access token.

    :param tenant_id: The Azure AD tenant ID
    :param client_id: The application (client) ID
    :param client_secret: The client secret for authentication
    :return: Access token string for Microsoft Graph API requests
    """

    # OAuth2 token endpoint for the specified tenant
    authority_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    # Client credentials grant request payload
    payload = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }

    # Request access token from Azure AD
    response = requests.post(authority_url, data=payload)
    response.raise_for_status()

    # Extract and return the access token
    token = response.json().get('access_token')
    return token

def delete_calendar_event(mailbox, access_token, delete_id):
    # Graph API endpoint for deleting calendar event
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/calendar/events/{delete_id}"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.delete(url, headers=headers)
    response.raise_for_status()

    return response.status_code

def get_calendar_events(mailbox, access_token, start_date, end_date, sender=None, subject_contains=None, delete_event_callback=None):
    # Graph API endpoint using calendarView to get recurring event instances
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/calendarView?startDateTime={start_date}T00:00:00&endDateTime={end_date}T23:59:59"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()

    data = response.json()

    # Store filtered meetings
    meetings = []
    for meeting in data['value']:
        organizer_email = meeting.get('organizer', {}).get('emailAddress', {}).get('address')
        subject = meeting.get('subject', '')
        series_master_id = meeting.get('seriesMasterId')

        # Extract start and end times
        start_time_str = meeting.get('start', {}).get('dateTime')
        end_time_str = meeting.get('end', {}).get('dateTime')

        # Parse and format times
        if start_time_str and end_time_str:
            start_time = datetime.fromisoformat(start_time_str)
            end_time = datetime.fromisoformat(end_time_str)
            time_range = f"{start_time.strftime('%I:%M%p').lstrip('0')}-{end_time.strftime('%I:%M%p').lstrip('0')} UTC"
        else:
            time_range = "Time N/A"

        # If sender is specified, filter by it
        if sender and organizer_email != sender:
            continue

        # If subject_contains is specified, filter by it
        if subject_contains and subject_contains.lower() not in subject.lower():
            continue

        # Determine if recurring and which ID to use for deletion
        is_recurring = series_master_id is not None
        delete_id = series_master_id if is_recurring else meeting.get('id')

        meetings.append({
            'id': meeting.get('id'),
            'delete_id': delete_id,
            'subject': subject,
            'organizer_email': organizer_email,
            'is_recurring': is_recurring,
            'time_range': time_range,
            'start_time': start_time if start_time_str else None
        })

    # Sort meetings by start time
    meetings.sort(key=lambda x: x['start_time'] if x['start_time'] else datetime.max)

    # Display all meetings
    if not meetings:
        print("No calendar events found matching your criteria.")
        return

    print("\nCalendar Events:")
    for idx, meeting in enumerate(meetings, 1):
        recurring_label = "Recurring Meeting" if meeting['is_recurring'] else "Single Event"
        print(f"{idx}. [{recurring_label}] {meeting['time_range']} | Sender: {meeting['organizer_email']} | Subject: {meeting['subject']}")

    # Get user selection
    while True:
        selection = input("\nEnter the number of the event you want to delete (or 'q' to quit): ").strip()

        if selection.lower() == 'q':
            print("Ok")
            return

        try:
            selected_idx = int(selection) - 1
            if selected_idx < 0 or selected_idx >= len(meetings):
                print("Invalid selection. Please try again.")
                continue

            selected_meeting = meetings[selected_idx]
            break

        except ValueError:
            print("Invalid input. Please enter a number.")
            continue

    # Confirm deletion
    recurring_label = "Recurring Meeting" if selected_meeting['is_recurring'] else "Single Event"
    print(f"\nYou are about to delete this calendar invite:")
    print(f"Type: {recurring_label}")
    print(f"Time: {selected_meeting['time_range']}")
    print(f"Sender: {selected_meeting['organizer_email']}")
    print(f"Subject: {selected_meeting['subject']}")

    choice = input("Do you want to continue? (y/n): ").strip().lower()

    if choice == 'y':
        try:
            status_code = delete_event_callback(mailbox, access_token, selected_meeting['delete_id'])
            if status_code == 204:
                print(f"Successfully deleted the calendar event!")
            else:
                print(f"Event deletion returned unexpected status code: {status_code}")
        except Exception as e:
            print(f"Error deleting event: {e}")
    elif choice == 'n':
        print("Canceling Request")
    else:
        print("Invalid choice.")


if __name__ == "__main__":

    config = configparser.ConfigParser()
    config.read('config.ini')

    client_id = config['credentials']['client_id']
    client_secret = config['credentials']['client_secret']
    tenant_id = config['credentials']['tenant_id']

    mailbox = input("Enter mailbox (user@domain.com): ")

    start_date = input("Enter start date (yyyy-mm-dd): ")
    end_date = input("Enter end date (yyyy-mm-dd): ")

    sender = input("Enter Sender (user@domain.com) [optional]: ")
    subject_contains = input("Enter text to search in subject [optional]: ")

    access_token = get_access_token(tenant_id, client_id, client_secret)
    get_calendar_events(mailbox, access_token, start_date, end_date,
                       sender if sender else None,
                       subject_contains if subject_contains else None,
                       delete_calendar_event)