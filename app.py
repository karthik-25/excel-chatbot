from flask import Flask, request, jsonify
import pandas as pd
from google.cloud import storage
from io import BytesIO
import os
import requests

bucket_name = "excel-chatbot-bucket"
file_name = "Book1.xlsx"

app = Flask(__name__)

@app.route('/view', methods=['GET'])
def view_rows():
    # Logic to view rows from Excel
    df = read_excel_from_gcs(bucket_name, file_name)
    return df.to_json(orient='records')

@app.route('/add', methods=['POST'])
def add_row():
    try:
        # Load the Excel file
        df = read_excel_from_gcs(bucket_name, file_name)
        # Get data from request
        data = request.get_json()
        # Append the new row
        new_row = pd.DataFrame([data])
        df = pd.concat([df, new_row], ignore_index=True)
        # Save back to Excel
        save_excel_to_gcs(df, bucket_name, file_name)
        return jsonify({"message": "Row added successfully"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/modify', methods=['PUT'])
def modify_row():
    try:
        # Load the Excel file
        df = read_excel_from_gcs(bucket_name, file_name)
        # Get data from request
        data = request.get_json()
        row_index = data.get("index")
        new_data = data.get("newData")
        # Check if index is valid
        if row_index is not None and row_index < len(df):
            for column, value in new_data.items():
                df.at[row_index, column] = value
            save_excel_to_gcs(df, bucket_name, file_name)
            return jsonify({"message": "Row modified successfully"}), 200
        else:
            return jsonify({"error": "Invalid row index"}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/delete', methods=['DELETE'])
def delete_row():
    try:
        # Load the Excel file
        df = read_excel_from_gcs(bucket_name, file_name)
        # Get index from request args
        row_index = request.args.get('index', type=int)
        # Check if index is valid
        if row_index is not None and row_index < len(df):
            df = df.drop(df.index[row_index])
            save_excel_to_gcs(df, bucket_name, file_name)
            return jsonify({"message": "Row deleted successfully"}), 200
        else:
            return jsonify({"error": "Invalid row index"}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500

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


if __name__ == '__main__':
    app.run(debug=True)
