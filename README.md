# Invoice RPA Processor

Hey there! 👋 This is a simple but effective RPA (Robotic Process Automation) system I built to handle the tedious task of processing invoice PDFs from Gmail. Instead of manually downloading and extracting data from hundreds of invoices, this tool does it all automatically.

## What This Does

Ever found yourself drowning in invoice PDFs from your Gmail? This tool will:
- 🔍 **Search your Gmail** for invoice-related emails
- 📥 **Download PDF attachments** automatically
- 📊 **Extract key data** from each invoice (numbers, dates, amounts, vendors)
- 📈 **Generate Excel reports** with all the data organized neatly
- 📋 **Create summaries** so you can see totals at a glance

## How It Works

The process is pretty straightforward:

1. **Gmail Connection**: Uses your Gmail credentials to search for invoice emails
2. **PDF Download**: Grabs all PDF attachments and saves them locally
3. **Data Extraction**: Reads through each PDF and pulls out important info like:
   - Invoice numbers
   - Vendor names
   - Invoice dates
   - Amounts
   - Due dates
   - Payment status
4. **Report Generation**: Creates an Excel file with everything organized in sheets

## Getting Started

### Prerequisites
- Python 3.7 or higher
- A Gmail account
- Gmail App Password (for security)

### Installation

1. **Clone or download** this repository
2. **Install the requirements**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up your Gmail credentials** in `config.py`:
   ```python
   EMAIL_CONFIG = {
       'sender_email': 'your-email@gmail.com',
       'sender_password': 'your-app-password'  # Use App Password, not regular password
   }
   ```

   *Note: You'll need to create an App Password in your Google Account settings for this to work.*

### Running the Tool

Just run:
```bash
python main.py
```

The tool will:
- Show you what it's doing step by step
- Tell you how many files it downloaded
- Show you what data it extracted
- Create an Excel report in the `output/` folder

## What You'll Get

The system creates an Excel file with two sheets:

**Invoice_Data Sheet:**
- All your invoice information in a clean table
- Columns for Invoice Number, Vendor, Date, Amount, Due Date, and Status
- Empty cells for any data that couldn't be found

**Summary Sheet:**
- Total number of invoices processed
- Total dollar amount across all invoices
- Number of unique vendors

## Project Structure

```
├── main.py              # The main script that runs everything
├── email_downloader.py  # Handles Gmail connection and downloads
├── config.py           # Your settings and credentials
├── requirements.txt    # Python packages needed
├── input/             # Where downloaded PDFs are stored
└── output/            # Where Excel reports are saved
```

## Troubleshooting

**"Gmail connection failed"**
- Make sure you're using an App Password, not your regular Gmail password
- Check that your Gmail account has 2FA enabled

**"No invoices found"**
- The tool looks for emails with "invoice" in the subject line
- Make sure your invoice emails have relevant subject lines

**"PDF extraction failed"**
- Some PDFs might be image-based and harder to read
- The tool will skip problematic files and continue with the rest

## Customization

Want to change what the tool looks for? You can modify:
- **Email search criteria** in `email_downloader.py`
- **Data extraction patterns** in `main.py`
- **Excel report format** in the main processing function

## Why I Built This

I got tired of manually processing invoices and wanted something simple that could handle the repetitive work. This tool is designed to be:
- **Simple**: No complex setup or configuration
- **Reliable**: Handles errors gracefully and continues processing
- **Transparent**: Shows you exactly what it's doing
- **Flexible**: Easy to modify for different invoice formats

## Contributing

Found a bug or want to add a feature? Feel free to:
- Open an issue to report problems
- Submit a pull request with improvements
- Suggest new features that would be helpful

## License

This project is open source and available under the MIT License.

---

*Built with Python, pandas, and a lot of patience for invoice processing! 📄✨* 
