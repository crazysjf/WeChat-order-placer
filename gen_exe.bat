# 调试的时候需要去掉-w参数，才能有错误信息输出
pyinstaller.exe --onefile --hidden-import PySide2.QtXml --add-data "MainUI\form.ui;MainUI\" -w .\main_ui.py