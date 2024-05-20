# script4iot
- IoT班用Pythonスクリプト

## 使用方法
### 想定環境
- Windows11 Command Prompt
### 実行
- PythonをWindowsサービスに登録して使用する。

1. `pywin32`をインストールする。
``` bash
$ pip install pywin32
```

2. 以下のPATHを「システム環境変数」に追加する。
```
/path/to/your/Python/Python3XX
/path/to/your/Python/Python3XX/Scripts
/path/to/your/Python/Python3XX/Lib/site-packages/pywin32_system32
/path/to/your/Python/Python3XX/Lib/site-packages/win32
```

3. Command Promptを管理者権限で起動する。
4. `Data_logger.py`のパスに移動する。
```bash
$ cd /path/to/your/Data_logger.py
```
5. `Data_logger.py`をサービスに登録する。
```bash
$ python Data_logeer.py install
```
6. `Data_logger.py`をサービスとして実行する。
```bash
$ python Data_logger.py start
```

### 停止
1. `Data_logger.py`のパスに移動する。
```bash
$ cd /path/to/your/Data_logger.py
```

2. サービスを停止する。
```bash
$ python Data_logger.py stop
```

> **Caution:**
> `「指定されたサービスは削除の対象としてマークされています」`等のメッセージが表示され、うまくサービスを削除できない場合は、PCを再起動する。

### 参考
- https://qiita.com/Bashi50/items/1d98f80ccaa8746bff38
- 