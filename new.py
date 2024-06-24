import os 
import re
import pytz
import gspread
from datetime import *
from threading import Thread

PATH = os.getcwd()
class GoogleSheet():
    def __init__(self, config, spreads_id, namePage='Fllow Notrip'):
        self.config, self.spreads_id = config, spreads_id
        self.namePage = namePage
        self.wks = self.connect_sheet()

    def connect_sheet(self):
        for i in range(10):
            try:
                gs = gspread.service_account(self.config)
                table = gs.open_by_key(self.spreads_id)
                wks = table.worksheet(self.namePage)
                return wks
            except:
                time.sleep(5)

    def get_data_sheet(self):
        data = self.wks.get_all_values()
        data_from_A_to_F = [[row[i] for i in range(10)] for row in data[1:]]
        return data_from_A_to_F

    def add_data_to_sheet(self, data: list):
        self.wks.append_row(data)

    def edit_data_in_sheet(self, data: str, column:str, row: int):
        self.wks.update_acell(f'{column}{row}', data)

    def delete_data_from_sheet(self, row: int):
        self.wks.delete_rows(row)

class Tracking:
    def __init__(self) -> None:
        self.config = f"{PATH}\\config\\config-sheet.json"
        self.spreadsheet_id = '10GRaFwtnU7scsdU-gCnKz_bGV6FNnEmS03aLdHNK3jM'
        self.sheet = GoogleSheet(self.config, self.spreadsheet_id)

    def convertTimeVN(self):
        vn_tz = pytz.timezone('Asia/Ho_Chi_Minh')
        current_time_vn = datetime.now(vn_tz)
        return current_time_vn
    def removeDuplicates(self, arrData):
        seen = {}
        result = []
        for item in arrData:
            key = item['SapID']
            time_str = item['CutofffDay']
            time_obj = datetime.strptime(time_str, "%d/%m")
            if key not in seen or seen[key][1] < time_obj:
                seen[key] = (item, time_obj)
                # print(seen[key][1])
        result = [v[0] for v in seen.values()]
        return result
    def extractDate(self, input_string):
        date_pattern = r'\b\d{2}-[A-Za-z]{3}\b'
        match = re.search(date_pattern, input_string)
        if match:
            date_string = match.group(0)
            date_object = datetime.strptime(date_string, "%d-%b")
            return date_string, date_object
        else:
            return None, None
    def getDataNotrip(self):
        account = self.sheet.get_data_sheet()
        arrData = []
        currentDay = datetime.strptime(str(self.convertTimeVN()).split(' ')[0].replace('-','/'), "%Y/%m/%d")
        fiveDayAgo = currentDay - timedelta(days=5)
        for idx, data in enumerate(account):
            cutofffDay = datetime.strptime(data[5], "%d/%m").replace(year=datetime.now().year)
            if cutofffDay < fiveDayAgo: continue
            arrData.append({
                'index'        : idx,
                'SapID'        : data[0],
                'Count'        : data[1],
                'Name'         : data[2],
                'Tel'          : data[3],
                'Notrip'       : data[4],
                'CutofffDay'   : data[5],
                'DayOn'        : data[6],
                'DayOff'       : data[7],
                'StatusVehicle': data[8],
                'Note'         : data[9],
            })
        filteredData = self.removeDuplicates(arrData)
        print(filteredData)
        return filteredData

    def trackingDataNotrip(self):
        arrData = self.getDataNotrip()
        for data in arrData:
            index = data['index']
            dayOn = data['DayOn']
            dayOff = data['DayOff']
            statusVehicle = data['StatusVehicle']
            note = data['Note']
            notrip = int(re.findall(r'\d', data['Notrip'])[0])
            if dayOn == '' and dayOff == '' and note == '' and statusVehicle == '': 
                if notrip >= 5:
                    Thread(target=self.sheet.edit_data_in_sheet, args=(f'{notrip}-NO', 'K', index+2)).start()
                else:
                    Thread(target=self.sheet.edit_data_in_sheet, args=(f'{notrip}-YES', 'K', index+2)).start()
            elif dayOn != '' and dayOff != '' and note == '' and statusVehicle == '':
                dayOffCV = datetime.strptime(dayOff, "%d-%b")
                dayOnCV = datetime.strptime(dayOn, "%d-%b")
                if (dayOffCV - dayOnCV).days > 10:
                    Thread(target=self.sheet.edit_data_in_sheet, args=(f'{notrip}-NO', 'K', index+2)).start()
                else:
                    Thread(target=self.sheet.edit_data_in_sheet, args=(f'{notrip}-YES', 'K', index+2)).start()
            elif note != '':
                Thread(target=self.sheet.edit_data_in_sheet, args=(f'{notrip}-YES', 'K', index+2)).start()
            elif statusVehicle != '':
                Thread(target=self.sheet.edit_data_in_sheet, args=(f'{notrip}-YES', 'K', index+2)).start()

        print("DONE")

        
track = Tracking()
track.trackingDataNotrip()
