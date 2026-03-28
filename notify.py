import os
import smtplib
import pandas as pd
from email.message import EmailMessage

OUTPUT_FILE = "data.xlsx"

def build_email():
    df = pd.read_excel(OUTPUT_FILE, sheet_name="raw_data")

    df["report_month_date"] = pd.to_datetime(df["report_month"], format="%B %Y", errors="coerce")
    latest_month_date = df["report_month_date"].max()

    if pd.isna(latest_month_date):
        raise RuntimeError("Could not determine latest report month.")

    latest_df = (
        df[df["report_month_date"] == latest_month_date]
        .sort_values("category")
        .reset_index(drop=True)
    )

    latest_month_label = latest_df["report_month"].iloc[0]
    file_url = os.environ["FILE_URL"]

    lines = []
    lines.append(f"AEM update detected: {latest_month_label}")
    lines.append("")
    lines.append("Latest values:")
    lines.append("")

    for _, row in latest_df.iterrows():
        lines.append(
            f"- {row['category']}: "
            f"current month = {int(row['current_month_units'])}, "
            f"beginning inventory = {int(row['beginning_inventory'])}"
        )

    lines.append("")
    lines.append(f"Open workbook: {file_url}")

    subject = f"AEM update: {latest_month_label}"
    body = "\n".join(lines)

    return subject, body


def send_email(subject: str, body: str):
    smtp_host = os.environ["SMTP_HOST"]
    smtp_port = int(os.environ["SMTP_PORT"])
    smtp_username = os.environ["SMTP_USERNAME"]
    smtp_password = os.environ["SMTP_PASSWORD"]
    email_to = os.environ["EMAIL_TO"]
    email_from = os.environ["EMAIL_FROM"]

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = email_from
    msg["To"] = email_to
    msg.set_content(body)

    with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.send_message(msg)


def main():
    subject, body = build_email()
    send_email(subject, body)
    print("Notification sent.")


if __name__ == "__main__":
    main()
