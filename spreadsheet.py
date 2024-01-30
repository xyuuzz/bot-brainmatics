import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials


class Spreadsheet:
    worksheet = None
    client = None
    gc = None
    mapping_penawaran_colums = {
        'Nomor Penawaran': '1',
        'Jenis Training': '2',
        'Nama Training': '3',
        'Instansi': '4',
        'PIC Marketing': '5',
        'PIC Eksternal': '6',
        'Status': '7',
    }

    mapping_registrasi_columns = {
        'nomor_registrasi': '1',
        'jenis_training': '2',
        'nama_training': '3',
        'instansi': '4',
        'pic_internal': '5',
        'pic_eksternal': '6',
        'status': '7',
    }

    def __init__(self, spreadsheet_url, sheet):
        self.gc = gspread.service_account('service_account.json')

        # Open a sheet from a spreadsheet in one go
        self.worksheet = self.gc.open_by_url(spreadsheet_url).worksheet(sheet)

    def get_data(self):
        data = self.worksheet.get_all_values()
        headers = data.pop(0)
        df = pd.DataFrame(data, columns=headers)
        return df

    def update_data(self, type, col, row, data):
        col_index = None
        if type == 'penawaran':
            col_index = self.mapping_penawaran_colums[col]
        elif type == 'registrasi':
            col_index = self.mapping_registrasi_columns[col]

        self.worksheet.update_cell(row, col_index, data)

    def add_data(self, data):
        self.worksheet.append_row(data)

    def get_value_last_row(self):
        last_row = len(self.worksheet.get_all_values())
        return self.worksheet.acell(f'A{last_row}').value

    def get_last_row(self):
        last_row = len(self.worksheet.get_all_values())
        return last_row
