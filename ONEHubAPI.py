from flask import Flask, request, jsonify
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# Load credentials from the JSON key file
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('ai-chatbot-training-7796cf9d571e.json', scope)
client = gspread.authorize(creds)

# Spreadsheet key
spreadsheet_key = "1MlJyXDxq4y2Q4e_Nckkyplb-3vs0YBoTQS-jKGpCxbI" # Replace with your Spreadsheet ID
spreadsheet = client.open_by_key(spreadsheet_key)

@app.route('/insert', methods=['POST'])
def insert_into_sheet():
    try:
        payload = request.json
        sheet_name = payload['sheet_name']
        data_list = payload['data']
        worksheet = spreadsheet.worksheet(sheet_name)

        for data in data_list:
            row = [data['name'], data['age'], data['email']]
            worksheet.append_row(row)

        return jsonify({'message': 'Inserted successfully'}), 200

    except KeyError as e:
        return jsonify({'message': f'Missing key: {str(e)}'}), 400
    except Exception as e:
        return jsonify({'message': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)