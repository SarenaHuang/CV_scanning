# CV_scanning
Utilize NLP to automatically scan CV
1.	Initialization and Library Imports
•	Import necessary Python libraries such as spacy, openai, pandas, docx, openpyxl, json, os, and pathlib.
•	Load the English NLP model using spaCy.
•	Set up the OpenAI API key.

2.	Load API Key
•	Load the OpenAI API key from a JSON file

3.	Read Resume Files
•	Use the python-docx library to read and extract paragraph text from Word documents, this project use Word document for demonstrate.

4.	Extract Resume Information Using GPT
•	Send the full resume text to GPT and request the extraction of the following fields in JSON format, including Name, Email, Phone, Location, Education, Experience, Skills, Autobiography.

5.	Analyze Personality
•	Use GPT to analyze personality traits based on the autobiography section.

6.	Get User (HR) Requirements
•	Prompt the user to enter candidate requirements in natural language (e.g., skills, location, years of experience).

7.	Match Resume Against User Requirements
•	Use GPT to compare the resume data with the user’s demands and determine if the candidate is a match.

8.	Export to Excel
Use pandas to export all the resume data and matching results to an Excel file.
![image](https://github.com/user-attachments/assets/10a18253-ac70-4abc-8260-0c3cf0faa9c9)
