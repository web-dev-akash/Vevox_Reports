import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from flask import Flask, render_template, request

app = Flask(__name__)

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("sheet.json", scope)
client = gspread.authorize(creds)


@app.get('/')
def upload():
    return render_template('upload-excel.html')


@app.post('/view')
def view():
    sheet = client.open("Test Sheet").get_worksheet_by_id(116082223)
    sheetData = sheet.get_all_records()
    # Read the File using Flask request
    file = request.files['file']
    # save file in local directory
    df = pd.DataFrame(sheetData)
    # print("Sheet Data : ", df)
    # Parse the data as a Pandas DataFrame type
    sessionName = pd.read_excel(
        file, sheet_name='Attendance', header=4, usecols=[1], nrows=1)
    data = pd.read_excel(file, sheet_name='Attendance',
                         header=10, usecols=[0, 1, 2, 3])
    newData = data[~(data['Email'].isnull())]
    newData['First Joined'] = pd.to_datetime(newData['First Joined'])
    newData['First Joined'] = newData['First Joined'].dt.strftime('%d %B %Y').astype(str)
    pollingData = pd.read_excel(
        file, sheet_name='Polling Results', header=8, usecols=[0, 1, 2])
    newPollingData = pollingData[
        pollingData[['First Name', 'Last Name']].apply(
            lambda row: (row['First Name'], row['Last Name']
                         ) in newData[['First Name', 'Last Name']].values,
            axis=1
        )
    ]
    newPollingData['Total answers correct'] = newPollingData['Total answers correct'].fillna(
        0)

    totalPolledData = pd.read_excel(
        file, sheet_name='Polling Results', header=3, usecols='D:BA', nrows=0)
    totalPolledColumns = totalPolledData.columns.tolist()
    polledCount = sum(
        1 for column in totalPolledColumns if 'Not Polled' not in column)

    totalAttempted = pd.read_excel(
        file, sheet_name='Polling Results', header=8, usecols="A:BA")
    newTotalAttempted = totalAttempted[
        totalAttempted[['First Name', 'Last Name']].apply(
            lambda row: (row['First Name'], row['Last Name']
                         ) in newData[['First Name', 'Last Name']].values,
            axis=1
        )
    ]
    finalData = pd.merge(newPollingData, newData[['First Name', 'Email', 'First Joined']], on='First Name',
                         how='left')
    finalData['Session ID'] = sessionName.values[0][0]
    finalData['Total Attempted'] = 0
    finalData['Total Polled'] = polledCount
    index = 0
    for i in newTotalAttempted.values:
        attemptCount = -2
        for j in i:
            if str(j) == 'nan' or str(j).isnumeric():
                continue
            else:
                attemptCount += 1
        finalData.loc[index,
                      'Total Attempted'] = attemptCount if attemptCount > 0 else 0
        index += 1
    filteredFinalData = finalData.drop_duplicates(
        subset=['Email', 'Session ID'])
    uniqueRows = filteredFinalData[~filteredFinalData.isin(
        df.to_dict(orient='list')).all(axis=1)]

    data = uniqueRows.values.tolist()
    print(data)
    if len(data) > 0:
        sheet.append_rows(data)

    return uniqueRows.to_html()
