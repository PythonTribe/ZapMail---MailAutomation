import win32com.client


def get_senders():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Access the inbox
    inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox folder
    messages = inbox.Items

    # Get unique sender names
    senders = set()
    for message in messages:
        try:
            senders.add(message.SenderName)
        except Exception:
            continue  # Skip any messages that might cause an error

    return list(senders)


def get_emails_by_sender(sender_name):
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Access the inbox
    inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox folder
    messages = inbox.Items

    # Filter emails by sender name
    emails = []
    for message in messages:
        try:
            if message.SenderName == sender_name:
                emails.append({
                    "Subject": message.Subject,
                    "Sender": message.SenderName,
                    "ReceivedTime": message.ReceivedTime,
                    "Body": message.Body[:100],  # Get the first 100 characters of the body
                    "To": message.To,
                    "CC": message.CC,
                    "BCC": message.BCC,
                    "Attachments": [attachment.FileName for attachment in
                                    message.Attachments] if message.Attachments.Count > 0 else [],
                    "Importance": message.Importance,
                    "Categories": message.Categories
                })
        except Exception:
            continue  # Skip any messages that might cause an error

    # Sort emails by ReceivedTime (newest first)
    emails.sort(key=lambda x: x["ReceivedTime"], reverse=True)
    return emails


def save_to_file(email_details, filename="email_output.txt"):
    with open(filename, 'w', encoding='utf-8') as f:
        if isinstance(email_details, list):
            for email in email_details:
                for key, value in email.items():
                    f.write(f"{key}: {value}\n")
                f.write("\n")  # Add a newline between emails
        else:
            f.write(email_details)


# Generate list of senders and prompt user for input
senders = get_senders()
print("Available senders:")
for index, sender in enumerate(senders):
    print(f"{index}: {sender}")

# Select sender by index
try:
    sender_index = int(input("Select the sender index to filter emails: "))
    sender_name = senders[sender_index]
except (IndexError, ValueError):
    print("Invalid index selected.")
    exit()

# Get emails by selected sender
email_details = get_emails_by_sender(sender_name)

# Print all emails from the selected sender
if email_details:
    print(f"\nEmails from {sender_name}:")
    for index, email in enumerate(email_details):
        print(f"{index}: {email['Subject']} (Received: {email['ReceivedTime']})")

    # Select an email by index
    try:
        email_index = int(input("Select the email index to fetch details: "))
        selected_email = email_details[email_index]
    except (IndexError, ValueError):
        print("Invalid index selected.")
        exit()

    # Print selected email details
    print("\nSelected Email Details:")
    for key, value in selected_email.items():
        print(f"{key}: {value}")

    # Save the output to a text file
    save_to_file([selected_email])  # Save the selected email details
else:
    print("No emails found for this sender.")
