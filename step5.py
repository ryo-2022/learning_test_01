!pip install pywin32
import win32com.client

#STEP5 Start
#Excelの起動
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

#ブックを開く
book = excel.Workbooks.Open("C:\\bbt\\チャレンジ課題_単票.xlsx")

#シートを選択する
sheet = book.WorkSheets("単票")
sheet.Select()

#検索キーを入力
myVal = input("検索キーを入力：")

if myVal == "all" or myVal == "ALL" :
    #シートを選択
    sheet_one = book.WorkSheets("港区_区役所一覧")
    sheet_one.Select()
    #最終行を取得
    xlUp = -4162
    lastrow = sheet_one.Cells(sheet.Rows.Count, 1).End(xlUp).Row
    #全ての行を印刷
    for i in range(lastrow - 1) :
        #検索キーを選択
        myKey = sheet_one.Cells(i+2, 1).Value
        sheet.Range("F2").Value = myKey
        #PDFで保存する
        sheet.ExportAsFixedFormat(Type=0, Filename="C:\\bbt\\" + myKey + ".pdf")
else:       
    #検索キーをセット
    sheet.Range("F2").Value = myVal
    #PDFで保存する
    sheet.ExportAsFixedFormat(Type=0, Filename="C:\\bbt\\" + myVal + ".pdf")

#Excelの終了
excel.Workbooks(1).Close(SaveChanges=0)
excel.Application.Quit() 
