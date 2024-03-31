# Excel Chatbot

This project provides a simple API for managing an Excel file. Users can add, view, modify, and delete rows in the Excel file through RESTful endpoints. This service is designed to be interacted with through an AI chatbot interface, simplifying operations for non-technical users.

## Getting Started

Follow these instructions to get the project up and running on your local machine for development and testing purposes.

### Prerequisites

Ensure you have the following installed:

- Python 3.8 or newer
- Flask
- pandas
- openpyxl
- gunicorn (optional, for deployment)

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/karthik-25/excel-chatbot
   ```

2. Navigate to the project directory:
   ```bash
   cd <your directory>
   ```

3. Create and activate a virtual environment:
   - Windows: 
     ```bash
     python -m venv venv
     .\venv\Scripts\activate
     ```
   - Unix/MacOS:
     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```

4. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

Use the API below to run the application

## API Usage

The API supports operations for adding, viewing, modifying, and deleting rows in the Excel file.

### Add Row

https://excel-chatbot-2024.uc.r.appspot.com/add

Add a new row to the Excel file.

**Request:**

```http

POST /add
Content-Type: application/json

{
  "name": "Jane Doe",
  "email": "jane.doe@example.com"
}
```

**Response:**

```json
{
  "message": "Row added successfully"
}
```

### View Rows

https://excel-chatbot-2024.uc.r.appspot.com/view

View all rows in the Excel file.

**Request:**

```http
GET /view
```

**Response:**

```json
[
  {
    "id": 1,
    "name": "John Doe",
    "email": "john.doe@example.com"
  },
  {
    "id": 2,
    "name": "Jane Doe",
    "email": "jane.doe@example.com"
  }
]
```

### Modify Row

https://excel-chatbot-2024.uc.r.appspot.com/modify

Modify an existing row by ID.

**Request:**

```http
PUT /modify
Content-Type: application/json

{
  "id": 2,
  "name": "Jane Smith",
  "email": "jane.smith@example.com"
}
```

**Response:**

```json
{
  "message": "Row modified successfully"
}
```

### Delete Row

https://excel-chatbot-2024.uc.r.appspot.com/delete

Delete a row by ID.

**Request:**

```http
DELETE /delete?id=2
```

**Response:**

```json
{
  "message": "Row deleted successfully"
}
```
