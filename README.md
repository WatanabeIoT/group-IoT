# script4iot
- IoT班用スクリプト

## 使用方法
### 実行
- Windows Powershellに以下を入力して実行
```
Start-Process python -ArgumentList ".\Data_logger.py" -NoNewWindow -RedirectStandardOutput "../log/output.log" -RedirectStandardError "../log/error.log"
```
### プロセスの確認
- 以下を入力してプロセスを確認
```
Get-Process
```