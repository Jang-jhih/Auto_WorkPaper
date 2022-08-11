from autoWP.autoWP import *

#%%
# path = r'\\192.168.248.10\IA_Evidence\RD\調閱清單111.7.11_技術部'
path = r'\\192.168.248.10\ia\全支付各單位辦法'

#%%
# 取出紀錄表
All_AuditItem = PrintFilePath(path)

All_File_TXT = All_AuditItem[All_AuditItem['files'].str.contains('txt')]
 

#%%提取TXT內容
All_content = CreatTxtContent(txtlist=All_File_TXT['原始路徑'].tolist())

#%% WorkPaper SHEET
    
All_content=pd.DataFrame({'原始路徑':All_File_TXT['原始路徑'].tolist(),
              '文件內容':All_content})

All_AuditItem = pd.merge(All_AuditItem,All_content,on = '原始路徑' ,how ='outer')


#%% PNG SHEET

All_png = All_AuditItem[All_AuditItem['files'].str.contains('png', na=False)]

All_png =ProcessCell(All_png)
All_AuditItem =ProcessCell(All_AuditItem)


HYPERLINK_png = CreatHyperlink(FilesList = All_png['files'].tolist(),RootList = All_png['root'].tolist())

All_png['HYPERLINK'] = HYPERLINK_png

HYPERLINK_workpaper = CreatHyperlink(FilesList = All_AuditItem['files'].tolist(),RootList = All_AuditItem['root'].tolist())
All_AuditItem['HYPERLINK'] = HYPERLINK_workpaper


workbook =openpyxl.load_workbook(filename='VBA.xlsm', read_only=False, keep_vba=True)

#建立sheet
worksheet = workbook.create_sheet("附件彙總",1)


wspng = workbook.create_sheet("PNG檔彙總",2)

#塞圖片
inputPNG(sheet=wspng,PngPathList=All_png['原始路徑'].tolist())


All_png = All_png[['主項目', '子項目1', '子項目2',  'HYPERLINK']]
All_png.columns = ['主項目', '子項目1', '子項目2',  '檔案名稱' ]

All_AuditItem = All_AuditItem[['主項目', '子項目1', '子項目2',  'HYPERLINK','ReviseTime','文件內容']]
All_AuditItem.columns = ['主項目', '子項目1', '子項目2',  '檔案名稱','ReviseTime','文件內容']

All_AuditItem['查核人員意見(x : 無)'] = ''
All_AuditItem['備註'] = ''




#%%
#塞資料到sheet
All_AuditItem = All_AuditItem.fillna('')


All_AuditItem.columns

for Serise in ['主項目', '子項目1', '子項目2',  '文件內容', '查核人員意見(x : 無)',
       '備註']:
    All_AuditItem[Serise] = clean(All_AuditItem[Serise])



inputDF(worksheet=wspng,df=All_png,header=True,index = False)
inputDF(worksheet=worksheet,df=All_AuditItem)


#%%


#微調png sheet
wspng.insert_cols(1,1)
wspng.cell(row=1, column=1).value = '圖片'

#%% 調整樣式
for worksheet,wspng in zip(worksheet['D'],wspng['E']):
    worksheet.style = "Hyperlink"
    wspng.style = "Hyperlink"



#%%存檔

from datetime import datetime

today = datetime.now()
today = datetime.strftime(today, "%Y%m%d")



department = path.split('\\')[-1].split('_')[-1]
# workbook.save(f'WorkBook_{department}_{today}.xlsm')
workbook.save(f'WorkBook.xlsm')


