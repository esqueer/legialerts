import os.path
import gspread
import openai

from google.oauth2 import service_account
from googleapiclient import discovery


# Set up your OpenAI API key
openai.api_key = "openai_api_key"
gpt_model_id = "gpt-4" #placeholder until GPT4 api access is available

# Set up Google API credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/documents"]
SERVICE_ACCOUNT_FILE = "filename=f"{curr_path}/legialerts.json")"

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

sheets_instance = discovery.build("sheets", "v4", credentials="credentials")
docs_instance = discovery.build("docs", "v1", credentials="credentials")

# Google Sheet ID and range

spreadsheet_id = "your_spreadsheet_id"
sheet_range = "your_sheet_name_and_range"  

# Fetch legislation data from the Google Sheet

sheet = sheets_instance.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sheet_range).execute()
data = sheet.get("values", [])

# Analyze the bill using OpenAI

def analyze_bill(title, summary, text):
    prompt = f"Please provide an analysis of the following legislation, what it will do, and how it would affect LGBTQ people in the state:\n\nTitle: {title}\nSummary: {summary}\nText: {text}\n\nAnalysis:"
    
    response = openai.Completion.create(
        engine="GPT-4", #placeholder until we get GPT4 access
        prompt=prompt,
        max_tokens=25000,
        n=1,
        stop=None,
        temperature=0.5,
    )

    message = response.choices[0].text.strip()
    return message

#Process each row of data

for i, row in enumerate(data):
    if len(row) < 4:
        bill_title, bill_summary, bill_text_base64 = row
        bill_text = base64.b64decode(bill_text_base64).decode('utf-8')
        analysis = analyze_bill(bill_title, bill_summary, bill_text)

    # Create a new Google Doc with the analysis
        doc_name = f"Analysis - {bill_title}"
        body = {
            "title": doc_name
        }
        doc = docs_instance.documents().create(body=body).execute()
        doc_link = f"https://docs.google.com/document/d/{doc['documentId']}"

    # Add the analysis to the Google Doc

          requests = [
        {
            "insertText": {
                "location": {
                    "index": 1
                },
                "text": f"Title: {bill_title}\n\nSummary: {bill_summary}\n\nAnalysis:\n{analysis}\n"
            }
        }
    ]
    result = docs_instance.documents().batchUpdate(documentId=doc['documentId'], body={"requests": requests}).execute()

# format the Google Doc

def format_document(doc_id, docs_instance):
    requests = [

        # Format bill title
        {
            "updateTextStyle": {
                "range": {
                    "startIndex": 21,
                    "endIndex": 21 + len(bill_title)
                },
                "textStyle": {
                    "italic": True,
                    "fontSize": {
                        "magnitude": 12,
                        "unit": "PT"
                    }
                },
                "fields": "italic,fontSize"
            }
        },
        # Format "Summary"
        {
            "updateTextStyle": {
                "range": {
                    "startIndex": 21 + len(bill_title) + 2,
                    "endIndex": 21 + len(bill_title) + 2 + 8
                },
                "textStyle": {
                    "bold": True,
                    "fontSize": {
                        "magnitude": 12,
                        "unit": "PT"
                    }
                },
                "fields": "bold,fontSize"
            }
        },
        # Format "Analysis"
        {
            "updateTextStyle": {
                "range": {
                    "startIndex": 21 + len(bill_title) + 2 + len(bill_summary) + 2,
                    "endIndex": 21 + len(bill_title) + 2 + len(bill_summary) + 2 + 9
                },
                "textStyle": {
                    "bold": True,
                    "fontSize": {
                        "magnitude": 12,
                        "unit": "PT"
                    }
                },
                "fields": "bold,fontSize"
            }
        }
    ]
    docs_instance.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()

    # Update the Google Sheet with the link to the analysis Google Doc
    update_range = f"{sheet_range.split('!')[0]}!D{i+1}"
    update_body = {
        "range": update_range,
        "values": [[doc_link]],
        "majorDimension": "ROWS"
    }
    update_request = sheets_instance.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=update_range,
        body=update_body,
        valueInputOption="RAW"
    ).execute()
