import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_bulk_emails(df_top15, sender_email, app_password):
    """
    Sends a personalized email to each student in the top 15 list.
    df_top15 should have columns: candidate_name, candidate_email, candidate_score
    """
    try:
        
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, app_password)

        for _, row in df_top15.iterrows():
            name = row.get("candidate_name", "Candidate")
            email = row.get("candidate_email")
            score = row.get("candidate_score", 0)

            if not email:
                continue  

            subject = "Congratulations — You’re among the Top Performers!"
            body = f"""
            Dear {name},

            Congratulations! 🎉
            You have been ranked among the **Top 15 candidates** based on your overall profile analysis.

            Your overall score: {score}

            Keep up the great work and continue building your GitHub projects, solving problems, and maintaining an ATS-friendly resume!

            Best regards,
            Profile Analyzer Team
            """

            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = email
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))

            server.sendmail(sender_email, email, msg.as_string())
            print(f" Email sent to {name} ({email})")

        server.quit()
        print("All emails sent successfully!")

    except Exception as e:
        print(f"Error while sending emails: {e}")
