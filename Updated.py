import pandas as pd
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# Your Gmail credentials here
GMAIL_ID = 'patilyashvardhan033@gmail.com'
GMAIL_PWD = 'klmy iouj vaos asyv'

# Define a function for sending email
def sendEmail(to, cc, sub, msg):
    # Connection to Gmail
    gmail_obj = smtplib.SMTP('smtp.gmail.com', 587)
    # Starting the session
    gmail_obj.starttls()
    # Login using credentials
    gmail_obj.login(GMAIL_ID, GMAIL_PWD)
    
    recipients = [to]  # List of recipients
    if cc:
        recipients.append(cc)
        
    # Sending email
    gmail_obj.sendmail(GMAIL_ID, recipients, msg.as_string())
    # Quit the session
    gmail_obj.quit()
    print("Email sent to " + str(to) + " and " + str(cc) + " with subject " + str(sub))

# Function to send emails
def send_emails(fixed_email):
    # Local file paths
    excel_file_path = "C:\\Users\\abhip\\OneDrive\\Pictures\\Downloads\\Automated Birthday Wisher\\Birthdate1.xlsx"

    image_file_path = "image1.jpg"

    # Read the Excel sheet having all the details
    dataframe = pd.read_excel(excel_file_path)
    # Today's date in format: DD-MM
    today = datetime.datetime.now().strftime("%d-%m")
    # Current year in format: YY
    yearNow = datetime.datetime.now().strftime("%Y")
    # Write index list
    writeInd = []
    
    for index, item in dataframe.iterrows():
        msg = MIMEMultipart("related")
        msg['From'] = GMAIL_ID
        msg['To'] = item['Email']
        msg['Subject'] = "Happy Birthday"

        # HTML message with embedded image
        html = f"""
        <html>
            <head>
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                    }}
                    .container {{
                        max-width: 600px;
                        margin: 0 auto;
                        padding: 20px;
                    }}
                    .header {{
                        background-color: #f5f5f5;
                        padding: 20px;
                        text-align: center;
                    }}
                    .content {{
                        padding: 20px;
                        background-color: #ffffff;
                    }}
                    .image-container {{
                        text-align: center;
                        margin-top: 20px;
                    }}
                    .footer {{
                        background-color: #f5f5f5;
                        padding: 20px;
                        text-align: center;
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h2>Happy Birthday, {item['Name']}!</h2>
                    </div>
                    <div class="content">
                        <p>On this special day, we wanted to take a moment to wish you a very happy birthday! May your day be filled with joy, laughter, and wonderful memories.</p>
                        <p>Thank you for being an important part of our company. Your hard work and dedication are truly appreciated. We value your contributions and the positive impact you make in your role.</p>
                        <p>Wishing you another fantastic year ahead!</p>
                        <div class="image-container">
                            <img src="cid:image1" alt="Birthday Image" style="max-width: 100%; height: auto;">
                        </div>
                    </div>
                    <div class="footer">
                        <p>Best regards,</p>
                        <p>Your Company</p>
                    </div>
                </div>
            </body>
        </html>
        """

        # Attach the HTML message
        msg.attach(MIMEText(html, 'html'))

        # Read the image file
        with open(image_file_path, "rb") as img_file:
            image_data = img_file.read()

        # Create a MIMEImage object
        image = MIMEImage(image_data, _subtype="jpeg")
        image.add_header('Content-ID', '<image1>')

        # Attach the image to the email
        msg.attach(image)

        # Stripping the birthday in Excel sheet as: DD-MM
        bday = item['Birthdate'].strftime("%d-%m")

        # Condition checking
        if (today == bday) and yearNow not in str(item['Year']):
            # Calling the sendEmail function with fixed_email as cc parameter
            sendEmail(item['Email'], fixed_email, "Happy Birthday", msg)
            writeInd.append(index)

    for i in writeInd:
        yr = dataframe.loc[i, 'Year']
        # This will record the years in which the email has been sent
        dataframe.loc[i, 'Year'] = str(yr) + ',' + str(yearNow)

    # Save the modified dataframe to the Excel file
    dataframe.to_excel(excel_file_path, index=False)

# Your fixed email ID here
fixed_email_id = 'yashvardhan.patil20@pccoepune.org'

# Call the send_emails() function
send_emails(fixed_email_id)
