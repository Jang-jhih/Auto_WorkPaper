
from autoWP.autoWP import *



#%%
path = r'\\192.168.248.10\IA_Evidence\RD\調閱清單111.7.11_技術部'


#%%
# 取出紀錄表
All_AuditItem = PrintFilePath(path)

All_File_TXT = All_AuditItem[All_AuditItem['files'].str.contains('txt')]
 
#%% WorkPaper SHEET


All_content = []
for TXTcontent in All_File_TXT['原始路徑'].tolist():
    content = open(TXTcontent,'r', encoding='UTF-8').read()
    All_content.append(content)
    
All_content=pd.DataFrame({'原始路徑':All_File_TXT['原始路徑'].tolist(),
              '文件內容':All_content})

All_AuditItem = pd.merge(All_AuditItem,All_content,on = '原始路徑' ,how ='outer')

workbook =openpyxl.load_workbook(filename='VBA.xlsm', read_only=False, keep_vba=True)


worksheet = workbook.create_sheet("WorkPaper",1)





#%% PNG SHEET

All_png = All_AuditItem[All_AuditItem['files'].str.contains('png')]






        

#%% 塞入df

#確定要輸出的儲存格及順序
ExportColumns = ['主項目', '子項目1', '子項目2',  'files', '原始路徑','文件內容']

#建立sheet
wspng = workbook.create_sheet("PNG",2)



All_png = All_png[['主項目', '子項目1', '子項目2',  'files'      , '原始路徑']]
All_png.columns = ['主項目', '子項目1', '子項目2',  '檔案名稱'   , '原始路徑']
All_AuditItem =   All_AuditItem[['主項目', '子項目1', '子項目2',  'files'    , '原始路徑','文件內容']]
All_AuditItemcolumns =          ['主項目', '子項目1', '子項目2',  '檔案名稱' , '原始路徑','文件內容']
#塞圖片
inputPNG(sheet=wspng,PngPathList=All_png['原始路徑'].tolist())

#清理前三個欄位有副檔名及分段文字段落
All_png =ProcessCell(All_png)
All_AuditItem =ProcessCell(All_AuditItem)





#塞資料到sheet
inputDF(worksheet=wspng,df=All_png,header=True,index = False)
inputDF(worksheet=worksheet,df=All_AuditItem)




#微調png sheet
wspng.insert_cols(1,1)
wspng.cell(row=1, column=1).value = '圖片'
# wspng['A1']='png'

#%%存檔

from datetime import datetime

today = datetime.now()
today = datetime.strftime(today, "%Y%m%d")



department = path.split('\\')[-1].split('_')[-1]
workbook.save(f'WorkBook_{department}_{today}.xlsm')


#%%




