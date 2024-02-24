import os
import subprocess
import pandas as pd
import numpy as np
import datetime


### 1. 抽出区間の開始日と終了日を取得する
# スペースで分割して、文字列で取得
start_day_str_list = input("開始日を入力下さい。スペース区切り。 例) 2022 11 1 \n").split()

# リストの各要素の文字列型を整数型へ変換
start_day = list(map(int, start_day_str_list))
print('start day: ', start_day)

# 現時点を終了日として自動取得する
now = datetime.datetime.now()
end_day_str_list = now.strftime("%Y %m %d").split()
end_day = list(map(int, end_day_str_list))
print('end_day: ', end_day)

# ファイルの保存先ディレクトリ
save_dir = os.path.join(os.environ["USERPROFILE"], "Documents\Python_Projects\pc_log_on_off_311\log_output")
print('log output: ', save_dir)

# ログファイルを保存するパス（一時ファイルとして使用）
logtxt_path = os.path.join(save_dir, 'log.txt')
print(logtxt_path)

# ログファイルを整形してcsvファイルを作成するパス（一時ファイル）
logcsv_path = os.path.join(save_dir, 'log.csv')

# 目的のログオン，オフ日時情報のcsvファイルを保存するパス
save_csv_path = os.path.join(save_dir, 'logon_off.csv')



### 2. Windowsコマンド「wevtutil」により、イベントログを取得する
cmd = f'wevtutil qe system /f:text /rd:true /q:"*[*[EventID=7001 or EventID=7002]]" > ./log_output/log.txt'
print(cmd)
subprocess.call(cmd, shell = True)



### 3. ログファイルを整形してcsvファイルを作成する
with open(logtxt_path, 'r') as f1:
    with open(logcsv_path, 'w') as f2:
        f2.write('row_No,Date,Time,Event ID,Description\n') # 文字列で書き込む場合は writeメソッドを使う
        flag_Description = False
        for i, row in enumerate(f1, start = 1):
            if 'Date' in row:
                buf_date = row.strip()[6:] # stripメソッドで改行コードを削除（右端の文字列を消す）
                _date = buf_date[:10]
                _time = buf_date[11:19]
                f2.writelines([str(i), ',', _date, ',', _time, ',']) # リストで書き込む場合は writelinesメソッドを使う
            elif 'Event ID' in row:
                f2.writelines([row.strip()[10:], ','])
            elif 'Description' in row:
                flag_Description = True
            elif flag_Description:
                f2.write(row.strip() + '\n') # 文字列の最後に改行コードを付与
                flag_Description = False

# 作成したcsvファイルをpandasデータフレームで読み込む
df0 = pd.read_csv(logcsv_path, encoding = 'cp932')
print(df0)

# Descriptionの列に対して、ログオンとログオフの文字列がある行のみを残す
target_list = [
'カスタマー エクスペリエンス向上プログラムのユーザー ログオン通知',
'カスタマー エクスペリエンス向上プログラムのユーザー ログオフ通知',
]
df = df0.query('Description in @target_list')
print(df)

# csvファイルに上書き保存する
df.to_csv(logcsv_path, index = False, encoding = 'cp932')



### 4. 指定区間を抽出して表示。また、csvファイルへ保存する
# 型変換（Date列を日付型へ変換）
df['Date'] = pd.to_datetime(df['Date'])

# 型確認
print(df.dtypes)

# 日付で区間を絞り込み
df = df[df['Date'] >= datetime.datetime(*start_day)]  # 先頭のアスタリスクはlist形式のアンパックをしている
df = df[df['Date'] <= datetime.datetime(*end_day)]
print(df)

# 日付のグループを取得（重複を省く）
date_list = list(df.groupby('Date').groups.keys())
print('\n', date_list, '\n')

# ログオン，ログオフの日時をcsvへ出力する
with open(save_csv_path, 'w') as f:
    print('date', 'logon', 'logoff', 'operating_time')
    col_list = ['date', 'logon', 'logoff', 'operating_time']
    
    # リストの要素間にカンマを挿入して文字列にする
    col_str = ','.join(map(str, col_list))
    f.write(col_str + '\n')
    
    # 各日に対して、最も早い時刻をログオン、最も遅い時刻をログオフとする
    # 日中にPCを一時的にシャットダウンすることもあるかもしれない場合のため
    for _date in date_list:
        
        _df = df.groupby("Date").get_group(_date)
        #print('\n', _df)
        
        # ログオン時刻から、その日の最初の時刻を抽出
        try:
            _logon_df = _df[_df['Event ID'] == 7001]
            _t_min = _logon_df['Time'].min()
        except:
            _t_min = np.nan
        #print('_t_min', _t_min)
        
        # ログオフ時刻から、その日の最後の時刻を抽出
        try:        
            _logoff_df = _df[_df['Event ID'] == 7002]
            _t_max = _logoff_df['Time'].max()
        except:
            _t_max = np.nan       
        #print('_t_max', _t_max)

        # 稼働時間の計算
        try:
            _t_time = datetime.datetime.strptime(_t_max, '%H:%M:%S') \
                      - datetime.datetime.strptime(_t_min, '%H:%M:%S') \
                      + datetime.timedelta(hours = -1)  # 1時間を引く（昼休み）
            _date_list = [_date.date().strftime('%Y/%m/%d'), _t_min, _t_max, _t_time]
        except:
            _date_list = [_date.date().strftime('%Y/%m/%d'), _t_min, _t_max, np.nan]
        
        # 型を文字列へ変換
        _date_list = list(map(str, _date_list))
        _date_str = ','.join(map(str, _date_list))
        
        # csvファイルへ書き込む
        f.write(_date_str + '\n')
        print(*_date_list)



### 5. 一時ファイルを削除する
os.remove(logtxt_path)
os.remove(logcsv_path)
