import pandas as pd
from postmarker.core import PostmarkClient
import random
import datetime
import os
import base64

postmark = PostmarkClient(server_token='XXXXX-XXXX-415c-a50a-8a7b57b9f6ea')
input_path = r"C:\Users\AbdhulRahimSheikh.M\Downloads\ICE futures\inputs\BirthdaySheet.xlsx"
df = pd.read_excel(input_path)
EmployeeName = df['NAME'].tolist()
EmployeeDOB = df['DOB'].tolist()
EmployeeEmail = df['Email'].tolist()



def generate_wishes(name):
    birthday_wishes = [
        "Happy Birthday! {name}, Wishing you a day filled with love and cheer. May your special day be as amazing as you are, bringing joy and happiness that lasts throughout the year.",
        "Happy Birthday! {name}, May your birthday be filled with many happy hours and countless wonderful memories. Celebrate this day knowing that you are cherished and loved by many.",
        "Happy Birthday! {name}, Hope your special day brings you all that your heart desires! Here’s wishing you a day full of pleasant surprises and unforgettable moments with your loved ones.",
        "Happy Birthday! {name}, Wishing you a beautiful day with good health and happiness forever. May your year ahead be filled with endless opportunities and blessings.",
        "Happy Birthday! {name}, It’s your birthday! May all the best things in the world happen in your life because you are definitely one of the best people too. Enjoy every moment!",
        "Happy Birthday! {name}, Sending your way a bouquet of happiness to wish you a very happy birthday! May your day be filled with smiles, laughter, and all the joy you deserve.",
        "Happy Birthday! {name}, Wishing you a day that is as special in every way as you are. May your life be blessed with happiness, success, and prosperity.",
        "Happy Birthday! {name}, Your birthday is the perfect time to remind you what a wonderful person you are. Have a great day filled with love, joy, and all your heart's desires.",
        "Happy Birthday! {name}, Wishing you a year filled with the same joy you bring to others. May your life be blessed with happiness, health, and love.",
        "Happy Birthday! {name}, May your birthday be the start of a year filled with good luck, good health, and much happiness. Celebrate your day with those who matter most.",
        "Happy Birthday! {name}, Count your life by smiles, not tears. Count your age by friends, not years. May your day be filled with joy and your year with countless blessings.",
        "Happy Birthday! {name}, On your birthday, I wish you happiness, health, and love. May your special day be filled with laughter and wonderful memories.",
        "Happy Birthday! {name}, Wishing you a day filled with love, laughter, and joy. May your life be as extraordinary as you are, and may your dreams come true.",
        "Happy Birthday! {name}, May this special day bring you endless joy and tons of precious memories! Celebrate with those who love and cherish you.",
        "Happy Birthday! {name}, May your day be filled with joy, laughter, and unforgettable moments. Wishing you all the happiness in the world on your special day.",
        "Happy Birthday! {name}, Another year older, another year wiser. May your birthday bring you new adventures, growth, and fulfillment in all aspects of life.",
        "Happy Birthday! {name}, Wishing you a fantastic day and an even better year ahead. May your dreams take flight and your heart be filled with joy.",
        "Happy Birthday! {name}, May your birthday be the start of a wonderful year full of happiness and success. Enjoy your day to the fullest, surrounded by love.",
        "Happy Birthday! {name}, Wishing you a birthday that's just as wonderful as you are! May your life be filled with blessings and your year ahead with joy.",
        "Happy Birthday! {name}, May your birthday be the beginning of a year filled with good luck, good health, and much happiness. Enjoy every moment of your special day!"

    ]
    wish = random.choice(birthday_wishes).format(name=name)
    return wish

def send_email(wish, email):
    text_body = wish
    images_folder = r'C:\Users\AbdhulRahimSheikh.M\PycharmProjects\pythonProject\BirthdayWishes\Images'
    image_files = [os.path.join(images_folder, f) for f in os.listdir(images_folder) if os.path.isfile(os.path.join(images_folder, f))]

    if image_files:
        random_image = random.choice(image_files)
        with open(random_image, 'rb') as f:
            image_content = f.read()

        encoded_image = base64.b64encode(image_content).decode('utf-8')
        content_type = 'image/jpeg' if random_image.lower().endswith('.jpg') else 'image/png'
        content_id = 'cid:myimagecid'

        attachments = [{
            'Name': os.path.basename(random_image),
            'Content': encoded_image,
            'ContentType': content_type,
            'ContentID': content_id
        }]
    else:
        attachments = []

    html_body = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                background-color: #f4f4f4;
                padding: 20px;
            }}
            .birthday-wish {{
                font-style: italic;
                font-size: 16px;
                color: #333;
                background-color: #fff;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            }}
            .birthday-image {{
                display: block;
                margin: 20px auto;
                max-width: 100%;
                border-radius: 10px;
            }}
        </style>
    </head>
    <body>
        <div class="birthday-wish">{text_body}</div>
        <img src="cid:myimagecid" class="birthday-image"/>
    </body>
    </html>
    """

    email_data = {
        "From": "itsupport@exchange-data.in",
        "To": email,
        "Subject": "Birthday Wishes",
        "Tag": "inline",
        "HtmlBody": html_body,
        "TextBody": text_body,
        "Attachments": attachments
    }

    try:
        postmark.emails.send(**email_data)
    except Exception as e:
        print(f"Failed to send email to {email}: {e}")

def check_birthday():
    today = datetime.datetime.now().date()
    for emp_name, dob, email in zip(EmployeeName, EmployeeDOB, EmployeeEmail):
        if dob.month == today.month and dob.day == today.day:
            wish = generate_wishes(emp_name)
            send_email(wish, email)

check_birthday()
