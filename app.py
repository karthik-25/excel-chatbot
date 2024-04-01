from flask import Flask, request, jsonify, render_template
import pandas as pd
from google.cloud import storage
from io import BytesIO
import json
import ast

bucket_name = "excel-chatbot"
file_name = "Book1.xlsx"

app = Flask(__name__)

@app.route('/view', methods=['POST'])
def view_rows():
    # Logic to view rows from Excel
    df = read_excel_from_gcs(bucket_name, file_name)
    return jsonify({
        'fulfillment_response': {
        'messages': [
            {
                'text': {
                    'text': [
                        df.to_json(orient='records')
                    ],
                },
            },
        ],
        }
    })

@app.route('/add', methods=['POST'])
def add_row():
    # Load the Excel file
    df = read_excel_from_gcs(bucket_name, file_name)
    # Get data from request
    data = request.get_json()
    params = data.get('sessionInfo', {}).get('parameters', {}).get('data', {})
    # Append the new row
    params_dict = ast.literal_eval(params)
    new_row = pd.DataFrame([params_dict])
    df = pd.concat([df, new_row], ignore_index=True)
    # Save back to Excel
    save_excel_to_gcs(df, bucket_name, file_name)
    return jsonify({
        'fulfillment_response': {
        'messages': [
            {
                'text': {
                    'text': [
                        "Row added successfully"
                    ],
                },
            },
        ],
        }
    })

@app.route('/modify', methods=['POST'])
def modify_row():
    # Load the Excel file
    df = read_excel_from_gcs(bucket_name, file_name)
    # Get data from request
    data = request.get_json()
    row_index = int(data.get('sessionInfo', {}).get('parameters', {}).get('row', {}))
    new_data = data.get('sessionInfo', {}).get('parameters', {}).get('data', {})
    new_data_dict = ast.literal_eval(new_data)
    # Check if index is valid
    if row_index is not None and row_index < len(df):
        for column, value in new_data_dict.items():
            df.at[row_index, column] = value
        save_excel_to_gcs(df, bucket_name, file_name)
        return jsonify({
            'fulfillment_response': {
            'messages': [
                {
                    'text': {
                        'text': [
                            "Row modified successfully"
                        ],
                    },
                },
            ],
            }
        })

@app.route('/delete', methods=['POST'])
def delete_row():
    # Load the Excel file
    df = read_excel_from_gcs(bucket_name, file_name)
    # Get index from request args
    data = request.get_json()
    row_index = int(data.get('sessionInfo', {}).get('parameters', {}).get('row', {}))
    # Check if index is valid
    if row_index is not None and row_index < len(df):
        df = df.drop(df.index[row_index])
        save_excel_to_gcs(df, bucket_name, file_name)
        return jsonify({
            'fulfillment_response': {
            'messages': [
                {
                    'text': {
                        'text': [
                            "Row deleted successfully"
                        ],
                    },
                },
            ],
            }
        })
        

def read_excel_from_gcs(bucket_name, file_name):
    storage_client = storage.Client()
    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(file_name)
    data = blob.download_as_bytes()
    df = pd.read_excel(BytesIO(data))
    return df

def save_excel_to_gcs(df, bucket_name, file_name):
    storage_client = storage.Client()
    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(file_name)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    blob.upload_from_string(output.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/')
def home():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)