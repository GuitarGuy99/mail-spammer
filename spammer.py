#!/usr/bin/python3

import smtplib
import imaplib
import email
import time
import argparse
from email.mime.text import MIMEText
import schedule
import datetime

ct = datetime.datetime.now()

# Function to handle command-line arguments
def parse_arguments():
    parser = argparse.ArgumentParser(description="Email automation script with a cache-clearing option")
    parser.add_argument('-c', '--clear-cache', action='store_true', help="Clear cache of received emails from the recipient")
    return parser.parse_args()

# Function to send email
def send_email(sender, recipient, subject, body):
    sender_email = outlook_username
    sender_password = outlook_password

    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient

    smtp_server = "smtp-mail.outlook.com"
    server = smtplib.SMTP(smtp_server, 587)
    server.starttls()

    try:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [recipient], msg.as_string())
        print(f"Email sent to {recipient}. Time: {datetime.datetime.now()}")

    except Exception as e:
        print(f"Error sending email: {e}")

    finally:
        server.quit()

# Function to check and clear email cache for specific recipient
def check_incoming_email(username, password, sender_email, recipient_email, clear_cache):
    outlook_server = "outlook.office365.com"
    connection = imaplib.IMAP4_SSL(outlook_server, 993)

    try:
        connection.login(username, password)
        print("Successfully logged in to Outlook account.")

        connection.select("inbox")

        if clear_cache:
            # Clear cache for specific recipient by marking their emails as seen
            print(f"Marking all emails from {recipient_email} as read.")
            typ, data = connection.search(None, f'FROM "{recipient_email}"')
            if typ == "OK":
                for num in data[0].split():
                    # Mark each email from the recipient as read (seen)
                    connection.store(num, '+FLAGS', '\\Seen')
                print(f"Marked all emails from {recipient_email} as read.")

        # Send an initial email
        print("Sending initial email...")
        send_email(sender_email, recipient_email, "Your Inbox Shall be Full", "Hello.\nThis is a Mail Spam\nPlease respond,or the torture shall continue.\nMany thanks\nGuitarGuy.")
        print("Initial email sent!")

        # Schedule email sending every 3 minutes for testing purposes
        schedule.every(0.05).minutes.do(
            send_email, sender_email, recipient_email, "Your Inbox Shall be Full", "Hello.\nThis is a Mail Spam\nPlease respond,or the torture shall continue.\nMany thanks\nGuitarGuy."
        )

        while True:
            schedule.run_pending()
            time.sleep(5)

            print("Checking inbox for new messages...")
            # Search for unseen (new) emails from the recipient
            typ, messages = connection.select("inbox")
            if typ == "OK":
                num_msgs = int(messages[0])
                print(f"Number of messages: {num_msgs}")

            typ, messages = connection.search(None, f'(UNSEEN FROM "{recipient_email}")')
            if typ == "OK":
                if messages[0]:  # Check if there are any unseen emails
                    num_msgs = messages[0].split()[-1]  # Get the latest unseen email
                    print(f"New unseen email found from {recipient_email}, processing...")

                    # Fetch and display the latest unseen email
                    typ, msg_data = connection.fetch(num_msgs, "(RFC822)")
                    if typ == "OK":
                        msg = email.message_from_bytes(msg_data[0][1])
                        sender_address = msg.get("From")
                        print(f"Email received from: {sender_address}")
                        if recipient_email in sender_address:
                            print(f"Email received from {sender_address}. Stopping the script.")
                            break

    except Exception as e:
        print(f"Error: {e}")

    finally:
        connection.logout()

# Main entry point
if __name__ == "__main__":
    args = parse_arguments()

    # Replace these with your Outlook and Live account details
    outlook_username = input("What is your email?:")
    outlook_password = input("What is your password?:")
    recipient_email = input("Recipient Email:")
    sender_email = outlook_username

    imaplib._MAXLINE = 10000000  # Set IMAP timeout

    # Start checking for incoming emails
    print(f"Checking for incoming emails... Time: {datetime.datetime.now()}")
    check_incoming_email(outlook_username, outlook_password, sender_email, recipient_email, args.clear_cache)
