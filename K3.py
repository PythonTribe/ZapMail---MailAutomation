import streamlit as st
import win32com.client
import streamlit_authenticator as stauth

# Dummy user data for authentication (Replace with a secure storage solution)
users = {
    "user1": {
        "password": "password1",  # Replace with hashed passwords in production
        "security_answer": "your_answer"  # Replace with actual answer
    }
}


# Function to connect to Outlook
def connect_outlook():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace
    except Exception as e:
        st.error(f"Error connecting to Outlook: {e}")
        return None


def get_senders(namespace):
    inbox = namespace.GetDefaultFolder(6)  # Inbox folder
    messages = inbox.Items

    senders = set()
    for message in messages:
        try:
            senders.add(message.SenderName)
        except Exception:
            continue

    return list(senders)


def get_emails_by_sender(namespace, sender_name):
    inbox = namespace.GetDefaultFolder(6)  # Inbox folder
    messages = inbox.Items

    emails = []
    for message in messages:
        try:
            if message.SenderName == sender_name:
                emails.append({
                    "Subject": message.Subject,
                    "Sender": message.SenderName,
                    "ReceivedTime": message.ReceivedTime,
                    "Body": message.Body[:100],
                })
        except Exception:
            continue

    emails.sort(key=lambda x: x["ReceivedTime"], reverse=True)
    return emails


# Streamlit application
st.title("Secure Email Extractor")

# User authentication
if 'authenticator' not in st.session_state:
    st.session_state['authenticator'] = stauth.Authenticate(
        users,
        'my_cookie_name',
        'my_signature_key',
        cookie_expiry_days=30
    )

if st.session_state['authenticator'].login('Login', 'main'):
    st.success("Logged in successfully!")

    # Security Question Input
    user = st.session_state['authenticator'].get_user()
    security_answer = st.text_input("What is your favorite color?", type="text")

    if st.button("Verify Answer"):
        if security_answer.lower() == users[user]['security_answer'].lower():
            st.success("Answer verified successfully!")

            # Connect to Outlook and fetch senders
            namespace = connect_outlook()
            if namespace:
                senders = get_senders(namespace)
                st.session_state.senders = senders

                # Select sender by index
                sender_index = st.number_input("Select the sender index to filter emails:", min_value=0,
                                               max_value=len(senders) - 1, step=1)
                sender_name = senders[sender_index]

                # Get emails by selected sender
                email_details = get_emails_by_sender(namespace, sender_name)

                # Display email details
                if email_details:
                    st.write(f"\nEmails from {sender_name}:")
                    for index, email in enumerate(email_details):
                        st.write(f"{index}: {email['Subject']} (Received: {email['ReceivedTime']})")
                else:
                    st.warning("No emails found for this sender.")
        else:
            st.error("Invalid answer.")
else:
    st.warning("Please log in.")

# Logout button
if st.session_state['authenticator'].logout('Logout', 'main'):
    st.session_state['authenticator'] = None
