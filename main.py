import smtplib
import email.message
import gspread
from oauth2client.service_account import ServiceAccountCredentials


class googleHelper():

    def __init__(self):
        self.scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                      "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
            'python.json', self.scope)
        self.client = gspread.authorize(self.creds)

    def insertRow(self, rowinsert):

        if self.sheet.insert_row(rowinsert, 4):
            return True

        return False

    def _setsheet(self, sheet):
        self.sheet = sheet

    def _getall(self):
        data = self.sheet.get_all_records()
        return data

    def _remove(self, index):

        if self.sheet.delete_row(index):
            return True

        return False

    def _update(self, index, newrow):

        if self.sheet.delete_row(index):
            if self.sheet.insert_row(newrow, index):
                return True

        return False

    def _get(self, index):
        row = self.sheet.row_values()
        return row

    def _updatecell(self, columnindex, rowindex, newvalue):

        if self.sheet.update_cell(rowindex, columnindex, newvalue):
            return True

        return False

    def _createuuid(self):
        # generate a new uuid kkey
        return

    def _updatetable(self, sheet, data):
        try:

            for element in data:
                sheet.insert_row(element, 4)

        except Exception as ex:
            print(ex)

    def _send(self, subject, destination, information):

        try:

            msg = email.message.Message()
            password = "newMAN2022"
            mensagem = information
            msg['subject'] = subject
            msg['from'] = "adaptableman@gmail.com"
            msg['to'] = destination
            msg.add_header('Content-Type', 'text/html')
            msg.set_payload(mensagem)

            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()

            s.login(msg['from'], password)
            s.sendmail(msg['from'], msg['to'], msg.as_string().encode('utf-8'))

            print('Enviado com sucesso!')

        except Exception as ex:
            print(ex)


newrow = ["Ele mesmo", 14, 14, 14, 14, "Admitido"]
helper = googleHelper()
sheet = helper.client.open("simpleone").sheet1
helper._setsheet(sheet)
# helper._update(11, newrow)
helper._send("teste", "adaptableman@gmail.com",
             "<h1>Simpleinformation here and</h1>")
