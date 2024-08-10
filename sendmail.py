import smtplib
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

def write_message(participant_name ,event_name):
    message =f"""
Dear {participant_name},

Congratulations!

We are thrilled to announce that you have successfully completed the {event_name}. Your dedication and effort have truly paid off, and it is our pleasure to present you with a certificate of completion.

Please find your certificate attached to this email. We hope that this accomplishment serves as a testament to your hard work and commitment to continuous learning.

As a part of the Microsoft Learn Student Ambassadors community, your participation and enthusiasm make a significant impact. We encourage you to continue exploring new opportunities and advancing your skills.

If you have any questions or need further assistance, please do not hesitate to reach out.

Once again, congratulations and best wishes for your future endeavors!

Warm regards,

Tarun Nayaka R
Microsoft Learn Student Ambassador
"""
    return message


def send_email(participant_name,to_email, eventname,attachment_path):
    smtp_server = 'smtp.gmail.com'  # Gmail SMTP server
    smtp_port = 465  # Port for SSL
    smtp_user = 'r.tarunnayaka25042005@gmail.com'  # Your Gmail address
    smtp_password = 'zpji gmrm oqli kvyy'
    body = write_message(participant_name,eventname)
    subject =f"Congratulations on Completing the {eventname}!"
    try:
        # Create a multipart message
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = to_email
        msg['Subject'] = subject

        # Attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))

        # Open the file to be sent
        if os.path.exists(attachment_path):
            with open(attachment_path, "rb") as attachment:
                # Instance of MIMEBase and named as part
                part = MIMEBase('application', 'octet-stream')
                # To change the payload into encoded form
                part.set_payload(attachment.read())
                # Encode into base64
                encoders.encode_base64(part)
                # Add header with the attachment file name
                part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
                # Attach the instance 'part' to instance 'msg'
                msg.attach(part)
        else:
            print(f"Attachment path '{attachment_path}' does not exist. Skipping attachment.")

        # Create a secure SSL context and send the email
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, to_email, msg.as_string())
            print(f"Email sent to {to_email}")

    except Exception as e:
        print(f"Failed to send email to {to_email}. Error: {str(e)}")






# # Example usage
# if __name__ == "__main__":
#     smtp_server = 'smtp.gmail.com'  # Gmail SMTP server
#     smtp_port = 465  # Port for SSL
#     smtp_user = 'r.tarunnayaka25042005@gmail.com'  # Your Gmail address
#     smtp_password = 'zpji gmrm oqli kvyy'  # Your Gmail password or App Password

#     csv_file_path = 'p2.csv'  # Path to your CSV file

#     # Read the CSV file
#     try:
#         with open(csv_file_path, newline='') as csvfile:
#             reader = csv.DictReader(csvfile)
#             for row in reader:
#                 to_email = row['email']
#                 subject = row['subject']
#                 body = row['body']
#                 attachment_path = row['attachment']

#                 print(f"Processing email for {to_email} with attachment {attachment_path}")

#                 send_email(to_email, subject, body, attachment_path, smtp_server, smtp_port, smtp_user, smtp_password)
#     except Exception as e:
#         print(f"Failed to read CSV file. Error: {str(e)}")


