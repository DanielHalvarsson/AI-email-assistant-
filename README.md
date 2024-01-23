# AI email assistant

## Description
This Python script integrates with Microsoft Outlook to read the latest email in your inbox and uses OpenAI's GPT-3.5 model to generate a draft response. The script then creates a draft email in Outlook with this AI-generated response.

## Features
- Retrieves the latest email from your Outlook inbox.
- Generates an AI response using OpenAI's GPT-3.5 model.
- Creates a draft email in Outlook with the AI response.

## Setup and Installation
1. **Install Python**: Ensure you have Python installed on your machine. [Download Python](https://www.python.org/downloads/)

2. **Install Dependencies**: Install the required Python libraries by running:
   ```bash
   pip install win32com.client openai python-dotenv

3. **Configure OpenAI API Key:** Set your OpenAI API key in an `.env` file:
    ```bash
    OPENAI_API_KEY=your_api_key_here

4. **Outlook Setup:** Make sure you have Outlook installed and configured with your email account.

5. Run the Script:** Execute the script with Python to start processing emails.
    ```bash
    python ai_email_assistant.py    

## Usage 

Once setup is complete, execute the script. It will:

1. Access the latest email in your inbox.
2. Use the email content to generate a draft response via OpenAI's model.
3. Create a new draft email in Outlook with the suggested response.

## Documentation
Below is a brief overview of each main function:
`get_latest_email_content()`: Retrieves the subject, body, and sender of the latest email from the Outlook inbox.
`generate_ai_response(email_subject, email_body)`: Generates an AI response using the provided email subject and body.
`create_draft_email(response, recipient, email_subject)`: Creates a draft email in Outlook with the specified response, recipient, and subject.

## Licensing

This project is released under the MIT License. This permissive license allows for free use, modification, and distribution of the software, provided that the original author is credited. For full license text, see [LICENSE.md](LICENSE.md) in this repository.

## Security Considerations

When using this AI Email Assistant, consider the following security aspects:

- **Email Privacy**: The script interacts with your email data. Ensure you have the necessary permissions and understand the privacy implications of using such scripts with your email account.
  
- **API Key Security**: Your OpenAI API key is sensitive information. Keep it secure and never expose it in public repositories or unsecured locations.

- **Data Handling**: While the script does not store email content or responses, it does process potentially sensitive information. Be cautious of where and how the data is processed, and ensure that the systems running the script are secure.
