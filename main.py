import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials
from googleapiclient import discovery
from googleapiclient.discovery import build


def scrape_page(url, title):
    data = [[title]]
    try:
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, features='html.parser')
        trs = soup.find('table', {'data-test': 'financials'}).find_all('tr')
        ths = trs[0].find_all('th')
        a = []
        for th in ths:
            a.append(th.text.strip())
        data.append(a)
        for tr in trs[1::]:
            tds = tr.find_all('td')
            a = []
            for td in tds:
                v = td.text.strip().replace(',', '').replace('%', '')
                try:
                    a.append(float(v))
                except:
                    a.append(v)
            data.append(a)
    except Exception as e:
        print(e)
    return data


def append_values(spreadsheet_id, range_, values):
    credentials = Credentials.from_service_account_file("prod.json")
    service = discovery.build("sheets", "v4", credentials=credentials)

    request = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_,
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": values}
    )
    response = request.execute()
    return response


def clear_document(spreadsheet_id):
    credentials = Credentials.from_service_account_file("prod.json")
    service = build('sheets', 'v4', credentials=credentials)
    names = ['Income Statement (Annual)', 'Income Statement (Quarterly)', 'Balance Sheet (Annual)',
             'Balance Sheet (Quarterly)', 'Cash Flow Statement (Annual)', 'Cash Flow Statement (Quarterly)']

    for sheet in names:
        sheet_id_to_delete = None
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        for sh in sheet_metadata['sheets']:
            if sh['properties']['title'] == sheet:
                sheet_id_to_delete = sh['properties']['sheetId']
                break
        request_body = {
            'requests': [
                {
                    'deleteSheet': {
                        'sheetId': sheet_id_to_delete
                    }
                }
            ]
        }
        try:
            response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request_body).execute()
        except:
            print("Delete err")
        request_body = {
            'requests': [
                {
                    'addSheet': {
                        'properties': {
                            'title': sheet
                        }
                    }
                }
            ]
        }
        try:
            spreadsheet = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request_body).execute()
        except:
            print('Creation err')

headers = {
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
}

stock_name = input('Введите название тикера: ')
income_annual = f'https://stockanalysis.com/stocks/{stock_name}/financials/'
income_quarterly = income_annual + "?period=quarterly"
balance_annual = income_annual + "balance-sheet/"
balance_quarterly = balance_annual + "?period=quarterly"
cash_flow_annual = income_annual + "cash-flow-statement/"
cash_flow_quarterly = cash_flow_annual + "?period=quarterly"

requests_list = [income_annual, income_quarterly, balance_annual, balance_quarterly, cash_flow_annual,
                 cash_flow_quarterly]
names = ['Income Statement (Annual)', 'Income Statement (Quarterly)', 'Balance Sheet (Annual)',
         'Balance Sheet (Quarterly)', 'Cash Flow Statement (Annual)', 'Cash Flow Statement (Quarterly)']

spreadsheet_id = "1FtFoVNZgF5xD4s-e9KfxNMrSSSUQgmnvHwC5I9CFxBw"
alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
index = 0
clear_document(spreadsheet_id)
for url in requests_list:
    values = scrape_page(url, names[index])
    range_ = f"{names[index]}!A1"  # example range
    append_values(spreadsheet_id, range_, values)
    index += 1




