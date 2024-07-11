
## Birthday Wishes Email Automation Script

This Python script automates the process of sending personalized birthday wishes via email to employees based on data fetched from an Excel sheet. It integrates with Postmark for email delivery and includes options for attaching random images to enhance the birthday message.

### Features

- **Email Sending**: Automatically sends birthday wishes emails using Postmark API.
- **Personalization**: Generates random birthday messages and includes employee-specific information.
- **Attachment Handling**: Optionally attaches a random image from a specified folder to the email.
- **Logging**: Logs any errors encountered during email sending to the console.

### Dependencies

- Python 3.x
- pandas
- postmarker (install via `pip install postmarker`)
- datetime
- base64 (part of Python standard library)

### Configuration

1. **Postmark Configuration**:
   - Ensure you have a Postmark account and generate a server token.
   - Replace `server_token` in the script with your actual Postmark server token.

2. **Excel Input**:
   - Provide the path to your Excel file containing employee data (`BirthdaySheet.xlsx`) in `input_path`.

3. **Images Folder**:
   - Place birthday-themed images in the `Images` folder located at `C:\Users\AbdhulRahimSheikh.M\PycharmProjects\pythonProject\BirthdayWishes\Images` or adjust `images_folder` variable accordingly.

### Usage

1. **Run the Script**:
   - Execute `birthday_wishes.py` to automatically check if any employees have birthdays today and send them personalized emails.

2. **Output**:
   - The script logs messages to the console indicating successful email sending or any errors encountered.

### Example

Assume `BirthdaySheet.xlsx` contains columns `NAME`, `DOB`, and `Email` with employee information. On each employee's birthday, the script will:
- Generate a random birthday message.
- Optionally attach a random image from the `Images` folder.
- Send a personalized email via Postmark to the employee's specified email address.

### Notes

- Ensure the script is run daily to send birthday wishes on the actual birthdate.
- Check the console output for any errors or exceptions during email sending.

