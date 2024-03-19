# by Hirai

import os
import serial
import serial.tools.list_ports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import re
import win32serviceutil
import win32service
import win32event
#import socket

#import csv

# for DEBUG
DEBUG = False
DEBUG_DATA = "0359 -54dBm 0.322mV T:25.33 H:25.81 LUX:295.84 X:-0.004 Y:-0.008 Z:0.976 SW:1001"

SPREADSHEET_KEY_FILE = "../secrets/spreadsheet-id.txt"
# WORKSHEET_NAME = "Sheet2"
JSON_FILE = "../secrets/project-watanabe-iot-b839ad323930.json"
LOG_DIR = "../log/"
DATA_LENGTH_NO_HF = 9
DATA_LENGTH_WITH_HF = 10
FREQUENCY = 500 #for test

class MySvc (win32serviceutil.ServiceFramework):
    _svc_name_ = "iot-service"              # TODO: 名前決める。
    _svc_display_name_ = "iot service"      # TODO: 名前決める。
    _svc_description_ = "descriptiooooooon" # TODO: 説明考える。

    def __init__(self,args):
        # 実行中のData_logger.pyが存在するディレクトリに移動する。
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        
        win32serviceutil.ServiceFramework.__init__(self, args) 
        self._stop_event = win32event.CreateEvent(None, 0, 0, None)

        # Googleスプレッドシートにアクセスするための認証情報を読み込む
        self.scope = ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
                ]
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, self.scope)
        self.client = gspread.authorize(self.creds)        

    def SvcStop(self):        
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self._stop_event)
    
    # サービス開始
    def SvcDoRun(self):
        # TODO: 実装
        self.run = True
        while self.run:
            self.main_loop()

    # スプレッドシートにデータを書き込む関数を定義する
    def write_to_sheet(self, ss_key, ser_num, data):
        # 送信機のシリアルナンバーに対応するシートを開く。
        # print(data)
        if len(data[0]) == DATA_LENGTH_NO_HF:
            sheetname = ser_num
        else:
            sheetname = ser_num + "_HF"
        worksheet = self.client.open_by_key(ss_key).worksheet(sheetname)

        #データを追加
        worksheet.append_rows(data)
        
        # 現在の日時を取得
        dt_now = datetime.datetime.now()
        now = dt_now.strftime('%Y年%m月%d日%H:%M:%S')

        print("Data have been written in Spreadsheet:",now, sheetname)
        return "ok"

    def load_ss_key(self):
        with open(SPREADSHEET_KEY_FILE) as f:
            key = f.readline()
        return key

    def shape_data(self, raw):
        """
        受信したデータの中から、数値のみを抽出してリスト化する。
        Args:
        - data <type:str>
        Return:
        - LoRa_num <type:str> 送信機のシリアルナンバー
        - shaped_data <type:list> 数値のみを抽出したデータ
        """
        # data = re.split(r'\s|:|([A-Za-z]+):', raw)   # 空白、コロン、[]を区切り文字としてリスト化
        data = re.split(r'\s*[^-\d.]+\s*', raw) # by NISHIMURA

        ser_num = data[0]   # 送信機のserial numberを取得
        data = data[1:]     # serial numberを取り除く
        shaped_data = []
        for d in data:
            try:
                shaped_data.append(float(d))
            except:
                shaped_data.append(d)
        return ser_num, shaped_data

    def main_loop(self):


        # スプレッドシートの名前とシートの名前を設定する
        # sheet_name = 'Sheet1'
        # worksheet_name = 'Data'
        # worksheet_name = "Sheet1"

        # シリアルポートを設定する

        ports = serial.tools.list_ports.comports()
        portnames = [p.name for p in ports]
        if len(portnames) == 0:
            print("使用可能なポートがありません。")
            exit()
        elif len(portnames) == 1:
            port = portnames[0]
        else:
            print("使用可能なシリアルポート:",*portnames)
            print("使用するシリアルポート名を入力してください。")
            name = input()
            if name in portnames:
                port = name
            else:
                print("シリアルポート名が正しくありません。")
                exit()

        baud_rate = 115200                      # ボーレート
        ser = serial.Serial(port, baud_rate)    # シリアル通信開始

        print("#################################")
        print("# Data recording is in progress #")
        print("#################################")

        # スクリプトを常時実行する
        cnt = 1
        while True:
            if ser.in_waiting:
                # シリアルポートからデータを読み取る

                # データの文字列を取得
                raw_data = ser.readline().decode().strip()
                # print(raw_data)

                # 受信機のシリアルナンバーと数値データを取得
                ser_num, data = shape_data(raw_data)

                # 現在の日時を取得
                dt_now = datetime.datetime.now()
                now = dt_now.strftime('%Y年%m月%d日%H:%M:%S')

                # 現在の日時と整形後のデータを連結
                now_data = [now] + data

                # ログを一時的に保存するためのファイルを開く。
                if not os.path.isfile(LOG_DIR + "temp_" + ser_num + ".csv"):
                    with open(LOG_DIR + "temp_" + ser_num + ".csv", mode='w') as f:
                        print(*now_data, file = f)
                else:
                    with open(LOG_DIR + "temp_" + ser_num + ".csv", mode='a') as f:
                        print(*now_data, file = f)
                
                # 現在一時ファイルに蓄積されているデータ数をカウント
                with open(LOG_DIR + "temp_" + ser_num + ".csv", mode='r') as f:
                    num_acc_data = len(f.readlines())
                
                # スプレッドシートへの書き込みを行う。
                if num_acc_data >= FREQUENCY:

                    # 対応するSpreadsheetのキーを取得
                    ss_key = self.load_ss_key()
                    
                    # 一時ファイルの中身を読み込む
                    temp_data = []
                    with open(LOG_DIR + "temp_" + ser_num + ".csv", mode='r') as f:
                        temp_data = f.readlines()
                    
                    # temp_dataの中身を整形
                    # temp_dataは2次元リストになる
                    temp_data_shaped = []
                    for td in temp_data:
                        td_shaped = list(td.split())                                   # 空白区切りしたものをリスト化
                        td_shaped = [td_shaped[0]] + list(map(float, td_shaped[1:]))   # 先頭以外をfloatに変換
                        temp_data_shaped.append(td_shaped)

                    # Spreadsheetにデータを書き込み
                    result = self.write_to_sheet(ss_key, ser_num, temp_data_shaped)

                    if result == "ok":
                        # 書き込みが正常に行われた場合は一時ファイルの中身を削除
                        with open(LOG_DIR + "/temp_" + ser_num + ".csv", mode='w') as f:
                            pass
                    else:
                        # 書き込みに失敗したら終了
                        with open("error.log", "w"):
                            print("Fail to write data...")
                        exit(1)
                else:
                    #print("Data received:", now, " #:",  cnt)
                    pass

def shape_data(raw):
    """
    受信したデータの中から、数値のみを抽出してリスト化する。
    Args:
    - data <type:str>
    Return:
    - LoRa_num <type:str> 送信機のシリアルナンバー
    - shaped_data <type:list> 数値のみを抽出したデータ
    """
    # data = re.split(r'\s|:|([A-Za-z]+):', raw)   # 空白、コロン、[]を区切り文字としてリスト化
    data = re.split(r'\s*[^-\d.]+\s*', raw) # by NISHIMURA

    ser_num = data[0]   # 送信機のserial numberを取得
    data = data[1:]     # serial numberを取り除く
    shaped_data = []
    for d in data:
        try:
            shaped_data.append(float(d))
        except:
            shaped_data.append(d)
    return ser_num, shaped_data

def debug():
    print("#################################")
    print("#           DEBUG               #")
    print("#################################")
    debug_data = DEBUG_DATA
    
    # データ整形
    ser_num, data = shape_data(debug_data)

    # 現在の日時を取得
    dt_now = datetime.datetime.now()           
    now = dt_now.strftime('%Y年%m月%d日 %H:%M:%S')
    
    # 現在の日時と整形後のデータを連結
    now_data = [now] + data
    print("ser_num, now_data", ser_num, now_data)
    print(now_data)

if __name__ == "__main__":
    if (not DEBUG):
        win32serviceutil.HandleCommandLine(MySvc)
    else:
        debug()