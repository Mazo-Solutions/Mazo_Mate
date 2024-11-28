# MazoMate - Interview Question Generator

MazoMate is an interactive web application built using Streamlit that allows users to generate structured interview questions and answers for various professional domains. Powered by Gemini AI, it supports multiple domains, including programming languages (e.g., Python, Java, Data Science, HR), and provides options to export the generated content in both Excel and Word formats.

## Features

- **Generate Interview Questions and Answers**: The app uses Gemini AI to generate structured interview questions and answers for a specific domain or area, based on the user's input.
- **Customizable Configuration**: Users can input their desired domain/area, years of experience, and complexity of the interview questions.
- **Export Options**: After generating the content, users can download the questions and answers in both Excel and Word formats.

## Technologies Used

- **Streamlit**: For building the interactive web app interface.
- **Gemini API**: For generating interview questions and answers using Google's generative AI.
- **Pandas**: For data manipulation and exporting to Excel.
- **python-docx**: For exporting the interview questions and answers to a Word document.
- **openpyxl**: For handling Excel files.
- **python-dotenv**: For loading environment variables like the API key.

## Requirements

To run the project locally, ensure you have the following dependencies:

- streamlit
- pandas
- openpyxl
- python-docx
- google-generativeai
- python-dotenv

### Install dependencies

You can install all dependencies by running:
```bash
pip install -r requirements.txt



How to Use
Clone the repository or download the files.

Create a .env file and add your Gemini API Key


```plaintext
GEMINI_API_KEY="your_actual_api_key"

Run the Streamlit app:
``bash streamlit run app.py

Enter the desired configuration (domain, experience level, and complexity) in the sidebar.

Click "Generate Questions" to generate a set of interview questions and answers.

Download the results in Excel or Word format using the export buttons.


