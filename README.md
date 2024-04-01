# Excel AI Chatbot

## Introduction

The Excel AI Chatbot is an application designed to simplify the management of Excel data through conversational AI. Leveraging technologies such as Flask, Google Cloud, and Dialogflow, this chatbot enables users to view, add, modify, and delete rows in an Excel file using natural language commands.

## Features

- **View Excel Data**: Users can request to see the current data stored in their Excel file.
- **Add Rows**: New data can be added as rows to the Excel file.
- **Modify Rows**: Existing data can be easily modified based on row number and specified changes.
- **Delete Rows**: Users can delete rows from the Excel file by specifying the row number.

The chatbot understands natural language inputs for these operations, making it intuitive to use without needing to understand the underlying technical implementation.

## Technology Stack

- **Flask**: A lightweight WSGI web application framework used to serve the chatbot's interface and handle API requests.
- **Google Cloud Storage**: Hosts the Excel file, ensuring data is securely stored and accessible in the cloud.
- **Dialogflow**: Powers the AI and natural language processing capabilities, allowing the chatbot to understand and respond to user queries.
- **Google Cloud Functions**: Hosts the application and integrates Dialogflow with Flask API endpoints, enabling seamless communication between the chatbot interface and the Excel file stored in Google Cloud.


## Accessing the Chatbot

You can interact with the Excel AI Chatbot by visiting [https://excel-chatbot-2.uc.r.appspot.com/](https://excel-chatbot-2.uc.r.appspot.com/). Example commands include:

- "Show me the data"
- "I want to add a row {'Column1': 'Value1', 'Column2': 'Value2'}"
- "Please delete row 1"
- "Modify row 2 to {'Column1': 'NewValue1', 'Column2': 'NewValue2'}"


The chatbot supports a variety of commands for interacting with the Excel data. Users are advised to follow the format `{'column name': 'value', 'column name': 'value' ...}` for adding or modifying data. For deletion, specifying the row number is sufficient.
