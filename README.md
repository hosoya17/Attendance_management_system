# Attendance_management_system
[![Open in Visual Studio Code](https://img.shields.io/static/v1?logo=visualstudiocode&label=&message=Open%20in%20Visual%20Studio%20Code&labelColor=2c2c32&color=007acc&logoColor=007acc)](https://open.vscode.dev/hosoya17/Attendance_management_system)
## 開発の概要
windows、mac共に動作する勤怠管理システムです。
## システムの概要
・社員一覧.xlsxに表記している、社員IDとパスワードをテキストボックスに入力します。<br>
・出勤ボタンをクリックしたら勤怠管理.xlsxに社員ID、社員名、出勤日、出勤時間が追加されます。<br>
・退勤ボタンをクリックしたら勤怠管理.xlsxに社員ID、社員名、退勤日、退勤時間が追加されます。
### 開発環境
開発環境：Visual Studio Code<br>
開発言語：python3<br>
ライブラリ:tkinter, time, openpyxl<br>
[![My Skills](https://skillicons.dev/icons?i=vscode,py)](https://skillicons.dev)
### 注意事項
#### セキュリティ
社員ID、社員名、パスワードは社員一覧.xlsxで管理されています。<br>
実際にはセキュリティ上の観点から、パスワードを平文のまま保存することは望ましくありません。<br>
もし実務でご使用になられる場合は、パスワードの管理は十分にセキュリティ対策を行ってください。<br>
また、社員一覧.xlsxのデータはChatGPTでランダムに生成したものであり、実在する個人や企業とは関係ありませんので、ご理解のほどよろしくお願いいたします。<br>
#### 環境構築
このプログラムはMicrosoftのExcelがインストールされていない場合、xlsxファイルが開けない為、使用することができません。<br>
<br>
事前にopenpyxlをインストールする必要があります。インストール方法は以下の通りです。<br>

```Shell
pip install openpyxl
```
<br>
また、openpyxlはバージョンによって文法が異なります。<br>
念のため以下の方法でアップグレードしてください。

```Shell
pip install openpyxl --upgrade
```
<br>
Attendance_management_system.pyの48, 53, 78, 97, 102, 126行目の''の中はExcelフォルダの勤怠管理または、社員一覧ファイルのパスを指定してください。<br>
以下に記述例を示します。<br>
<br>
48行目は社員一覧.xlsxのパスを指定してください。<br>

```python
wb = load_workbook('C:\\python\\Excel\\社員一覧.xlsx')
```
<br>
53行目は勤怠管理.xlsxを指定してください。<br>

```python
wb_attendance = load_workbook('C:\\python\\Excel\\勤怠管理.xlsx')
```
<br>
78行目は勤怠管理.xlsxを指定してください。<br>

```python
wb_attendance.save('C:\\python\\Excel\\勤怠管理.xlsx')
```
<br>
97行目は社員一覧.xlsxを指定してください。<br>

```python
wb = load_workbook('C:\\python\\Excel\\社員一覧.xlsx')
```
<br>
102行目は勤怠管理.xlsxを指定してください。<br>

```python
wb_attendance = load_workbook('C:\\python\\Excel\\勤怠管理.xlsx')
```
<br>
126行目は勤怠管理.xlsxを指定してください。<br>

```python
wb_attendance.save('C:\\python\\Excel\\勤怠管理.xlsx')
```
