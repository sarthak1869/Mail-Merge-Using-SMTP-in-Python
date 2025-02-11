# Mail-Merge-Using-SMTP-in-Python
This project automates the process of sending personalized bulk emails using Python. It reads recipient details from an Excel file, customizes an email template, and sends emails via Gmail SMTP. The script dynamically replaces placeholders like with actual data from the spreadsheet, ensuring each recipient receives a unique, tailored message.

# 🚀 Features

1.  Excel Integration: Reads recipient details (name, email, offer, etc.) from an Excel file.
2. Email Template Support: Uses a text-based template with placeholders for customization.
3. SMTP-Based Email Sending: Utilizes Gmail SMTP to send emails securely.
4. Personalized Content: Dynamically inserts recipient-specific details in each email.
5. Custom Subject Line: Extracts and personalizes the subject from the template.
6. Security with App Passwords: Uses Google App Password instead of storing plain credentials.

# 🛠 Technologies Used

1. Python (for automation)
2. Pandas (to handle Excel data)
3. smtplib & email.mime (for sending emails)

# Project Structure
 Your file should contain following files:
 - Python File
 - Txt File that will contain your template
 - Excel file that will contain the details of recipients

# 📝 Setup Instructions

1️⃣ Clone the Repository

2️⃣ Install Dependencies

     Make sure you have Python installed. Then, install Pandas:

3️⃣ Configure Your Email Credentials

      Go to Google App Passwords and generate a password.Open email_sender.py and update these lines with your Gmail credentials:

     EMAIL_ADDRESS = "your_email@gmail.com"
     EMAIL_PASSWORD = "your_app_password"

4️⃣ Prepare Your Email Template & Excel File

     Modify email_template.txt with placeholders (<<Name>>, <<Offer>>, etc.).Update contacts.xlsx with recipient data (Name, Email, Offer, etc.).

5️⃣ Run the Script
          
     Execute the script to send emails:

#  Potential Enhancements

     Send HTML emails with rich formatting.📎 Attach PDFs, images, or documents to emails.📩 Track email opens and responses.📅 Schedule emails for later delivery.

#  Contributing

     Feel free to submit issues and pull requests to improve the project! 😊
 

