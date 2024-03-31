from flask import Flask, request, jsonify
import pandas as pd

app = Flask(__name__)

@app.route('/view', methods=['GET'])
def view_rows():
    # Logic to view rows from Excel
    df = pd.read_excel('Book1.xlsx')
    return df.to_json(orient='records')

@app.route('/add', methods=['POST'])
def add_row():
    try:
        # Load the Excel file
        df = pd.read_excel('Book1.xlsx')
        # Get data from request
        data = request.get_json()
        # Append the new row
        new_row = pd.DataFrame([data])
        df = pd.concat([df, new_row], ignore_index=True)
        # Save back to Excel
        df.to_excel('Book1.xlsx', index=False)
        return jsonify({"message": "Row added successfully"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/modify', methods=['PUT'])
def modify_row():
    try:
        # Load the Excel file
        df = pd.read_excel('Book1.xlsx')
        # Get data from request
        data = request.get_json()
        row_index = data.get("index")
        new_data = data.get("newData")
        # Check if index is valid
        if row_index is not None and row_index < len(df):
            for column, value in new_data.items():
                df.at[row_index, column] = value
            df.to_excel('Book1.xlsx', index=False)
            return jsonify({"message": "Row modified successfully"}), 200
        else:
            return jsonify({"error": "Invalid row index"}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/delete', methods=['DELETE'])
def delete_row():
    try:
        # Load the Excel file
        df = pd.read_excel('Book1.xlsx')
        # Get index from request args
        row_index = request.args.get('index', type=int)
        # Check if index is valid
        if row_index is not None and row_index < len(df):
            df = df.drop(df.index[row_index])
            df.to_excel('Book1.xlsx', index=False)
            return jsonify({"message": "Row deleted successfully"}), 200
        else:
            return jsonify({"error": "Invalid row index"}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)
