import serial
import serial.tools.list_ports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import re

# for DEBUG
DEBUG = False
DEBUG_DATA = "0359 -54dBm 0.322mV T:25.33 H:25.81 LUX:295.84 X:-0.004 Y:-0.008 Z:0.976 SW:1001"

SPREADSHEET_KEY_FILE = "../secret/spreadsheet-id.txt"
# WORKSHEET_NAME = "Sheet2"
JSON_FILE = "../secret/project-watanabe-iot-b839ad323930.json"
DATA_LENGTH_NO_HF = 9
DATA_LENGTH_WITH_HF = 10
FREQUENCY = 10

# スプレッドシートにデータを書き込む関数を定義する
def write_to_sheet(ss_key, ser_num, data):
    # 送信機のシリアルナンバーに対応するシートを開く。
    # print(data)
    if len(data) == DATA_LENGTH_NO_HF:
        worksheet = client.open_by_key(ss_key).worksheet(ser_num)
    else:
        worksheet = client.open_by_key(ss_key).worksheet(ser_num + "_HF")

    # 最終行にデータを追加
    worksheet.append_row(data)

def load_ss_key():
    with open(SPREADSHEET_KEY_FILE) as f:
        key = f.readline()
    return key

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

if __name__ == "__main__":
    if not DEBUG:
        # Googleスプレッドシートにアクセスするための認証情報を読み込む
        scope = ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
                ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
        client = gspread.authorize(creds)

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
        try:
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
                    now = dt_now.strftime('%Y年%m月%d日 %H:%M:%S')

                    # 現在の日時と整形後のデータを連結
                    now_data = [now] + data
                    
                    # print(ser_num, now_data)
                    # スプレッドシートへの書き込みを行う。
                    if cnt == FREQUENCY:
                        ss_key = load_ss_key()
                        print("Data received:", now, "(output to SpreadSheet)")
                        write_to_sheet(ss_key, ser_num, now_data)
                        cnt = 1
                    else:
                        print("Data received:", now)
                        cnt += 1

        except KeyboardInterrupt:
            print("stop")
            # スプレッドシートにデータを書き込む
    else:
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



