

@chcp 65001




@set year=%date:~3,4%
@set mon=%date:~8,2%
@set day=%date:~11,2%
@set m=%time:~0,2%
@set h=%time:~3,2%
@set s=%time:~6,2%




@set today=%year%%mon%%day%







@echo 開始彙總底稿附件

@rem python執行
@python %~dp0main.py

@rem 開啟Excel，並傳入參數
@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" "WorkBook_技術部_%today%.xlsm" /batOpen

@rem pause