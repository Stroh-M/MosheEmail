## 🔧 Environment Setup

Before running the script, create a `.env` file in the root directory of the project with the following environment variables:

```env
EMAIL_ADDRESS=your_email@gmail.com
EMAIL_PASSWORD=your_app_password

EMAIL_FROM_1=example_sender_1@example.com
EMAIL_FROM_2=example_sender_2@example.com

RECIPIENT_1=recipient1@example.com
RECIPIENT_2=recipient2@example.com

TSV_FILE_PATH=/path/to/data.tsv
EXCEL_FILE_PATH=/path/to/output.xlsx
SHEET_NAME=YourSheetName
ERROR_EXCEL_PATH=/path/to/error_log.xlsx

SHIPPING_TXT_FILE=/path/to/shipping_output.txt
WALMART_ORDER_EXCEL_FILE=/path/to/walmart_orders.xlsx
```

### 📌 Notes:
- Use [Google App Passwords](https://support.google.com/accounts/answer/185833) if 2FA is enabled on your Gmail account.
- All file paths should be absolute or correctly relative to the working directory.
- Never commit your `.env` file to source control. Use `.gitignore` to exclude it.
