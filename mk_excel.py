import xlwt #エクセルに出力するのに必要なモジュール

wb=xlwt.Workbook() #ワークブックの作成
ws=wb.add_sheet("test_sheet_1") #ワークブックにシートを追加

ws.write(0,0,"My name is:")
ws.write(0,1,"takashiba") #ワークシートに指定されたセルに値を書き込む

wb.save("test_excel_work.xls") #ワークブック名を指定してエクセルファイルとして保存する。