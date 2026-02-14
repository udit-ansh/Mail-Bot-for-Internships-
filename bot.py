import pandas as pd
import smtplib
import imaplib
import email
import random
import time
import os
from datetime import datetime, timedelta
from email.message import EmailMessage

# ================= CONFIG =================
YOUR_EMAIL = "Your-email"
APP_PASSWORD = "Your - app password"
CV_PATH = "CV.pdf"

DAILY_LIMIT = 100
MIN_DELAY = 180      # 3 minutes
MAX_DELAY = 480      # 7 minutes
FOLLOWUP_AFTER_DAYS = 7
# ==========================================


# ---------------- UTILITIES ----------------

def load_sent_log():
    if os.path.exists("sent_log.csv"):
        log = pd.read_csv("sent_log.csv", dtype=str)
        return set(log["Email"])
    return set()


def update_sent_log(email_id):
    if os.path.exists("sent_log.csv"):
        log = pd.read_csv("sent_log.csv", dtype=str)
    else:
        log = pd.DataFrame(columns=["Email", "Date"])

    log.loc[len(log)] = [email_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    log.to_csv("sent_log.csv", index=False)


def generate_subject():
    subjects = [
        "Research Internship Opportunity Inquiry",
        "Research Internship / Project Inquiry"
    ]
    return random.choice(subjects)


def generate_body(research):
    return f"""Dear Sir,

    
I hope this email finds you well.

I am Your name, a first-year BS-MS student at College. Over the past few months, I have been exploring different research areas and found myself particularly drawn to your work in {research}.

I am currently strengthening my foundations in mathematics, physics, Python (including scientific libraries),C++ and Julia and I am eager to transition from classroom learning to a more hands-on research environment.

I would be grateful for the opportunity to contribute to your lab , even in a supporting role where I can learn the research process under your guidance.

I have attached my CV for your consideration. I would sincerely appreciate the opportunity to discuss any possible openings in your group.

Thank you very much for your time.

Sincerely,


Your Name

"""


# ---------------- MAIN MAILER ----------------

def send_initial_emails():

    df = pd.read_excel("professors.xlsx", dtype=str)
    df.columns = df.columns.str.strip()
    df.fillna("", inplace=True)

    already_sent = load_sent_log()

    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login(YOUR_EMAIL, APP_PASSWORD)

    sent_today = 0

    for index, row in df.iterrows():

        if sent_today >= DAILY_LIMIT:
            print("Daily limit reached.")
            break

        email_id = row["Email"].strip()

        if email_id in already_sent:
            continue

        if row["Status"] in ["Sent", "Followed Up", "Replied"]:
            continue

        msg = EmailMessage()
        msg["From"] = YOUR_EMAIL
        msg["To"] = email_id
        msg["Subject"] = generate_subject()
        msg.set_content(generate_body(row["Research"]))

        with open(CV_PATH, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename=os.path.basename(CV_PATH)
            )

        try:
            server.send_message(msg)
            print(f"Sent to {row['Name']}")

            df.at[index, "Status"] = "Sent"
            df.at[index, "Last Sent"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            update_sent_log(email_id)
            already_sent.add(email_id)

            sent_today += 1

            delay = random.randint(MIN_DELAY, MAX_DELAY)
            print(f"Waiting {delay} seconds...")
            time.sleep(delay)

        except Exception as e:
            print(f"Error sending to {row['Name']}: {e}")

    df.to_excel("professors.xlsx", index=False)
    server.quit()
    print("Initial emails done.")


# ---------------- FOLLOW-UP SYSTEM ----------------

def send_followups():

    df = pd.read_excel("professors.xlsx", dtype=str)
    df.columns = df.columns.str.strip()
    df.fillna("", inplace=True)

    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login(YOUR_EMAIL, APP_PASSWORD)

    for index, row in df.iterrows():

        if row["Status"] != "Sent":
            continue

        if row["Last Sent"] == "":
            continue

        last_sent = datetime.strptime(row["Last Sent"], "%Y-%m-%d %H:%M:%S")

        if datetime.now() - last_sent >= timedelta(days=FOLLOWUP_AFTER_DAYS):

            msg = EmailMessage()
            msg["From"] = YOUR_EMAIL
            msg["To"] = row["Email"]
            msg["Subject"] = "Following up on previous email"

            msg.set_content(f"""Dear,

I just wanted to gently follow up on my previous message in case it was missed.

I remain very interested in the possibility of contributing to your research group and would be grateful for any opportunity to assist.

Warm regards,
Uditansh Gupta
25MS002
IISER Kolkata
+91 7380430340
""")

            try:
                server.send_message(msg)
                print(f"Follow-up sent to {row['Name']}")
                df.at[index, "Status"] = "Followed Up"
                df.at[index, "Last Sent"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                time.sleep(random.randint(MIN_DELAY, MAX_DELAY))

            except Exception as e:
                print(f"Follow-up failed: {e}")

    df.to_excel("professors.xlsx", index=False)
    server.quit()
    print("Follow-ups completed.")


# ---------------- REPLY TRACKER ----------------

def check_replies():

    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(YOUR_EMAIL, APP_PASSWORD)
    mail.select("inbox")

    status, messages = mail.search(None, 'UNSEEN')
    mail_ids = messages[0].split()

    if not mail_ids:
        print("No new replies.")
        return

    df = pd.read_excel("professors.xlsx", dtype=str)
    df.columns = df.columns.str.strip()
    df.fillna("", inplace=True)

    for num in mail_ids:
        status, msg_data = mail.fetch(num, "(RFC822)")

        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                sender = email.utils.parseaddr(msg["From"])[1]

                print("Reply received from:", sender)

                df.loc[df["Email"] == sender, "Status"] = "Replied"

    df.to_excel("professors.xlsx", index=False)
    mail.logout()


# ---------------- RUN MENU ----------------

if __name__ == "__main__":

    print("1. Send Initial Emails")
    print("2. Send Follow-ups")
    print("3. Check Replies")

    choice = input("Choose option: ")

    if choice == "1":
        send_initial_emails()
    elif choice == "2":
        send_followups()
    elif choice == "3":
        check_replies()
    else:
        print("Invalid option")
