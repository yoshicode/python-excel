#Init
```
$ python3.8 -m venv excel-python-test
$ ls excel-python-test/
$ source excel-python-test/bin/activate
(excel-python-test) yoshi-2:py-ex yoshi$
```

#Fin
```
$ deactivate
```


##How to manage Excel by python
https://datumstudio.jp/blog/1722
```
#データフレームの取扱、Excelファイルのデータフレームへの読み込み用
import pandas as pd

#正規表現
import re

#input file name
input_file_csv = 'stat2015_02-04-02.xls'

#xls book Open
input_book = pd.ExcelFile(input_file_csv)

#xlsファイルのシートの名前をリストとして取得
input_sheet_name = input_book.sheet_names

# シートの総数
num_sheet = len(input_sheet_name)

output_df = pd.DataFrame(index=[], columns=[])
for i_sheet in range(0, num_sheet) :

    #シートを格納しているリストから読み込み
    input_sheet_df = input_book.parse(input_sheet_name[i_sheet],
                                      skiprows = 5,
                                      skip_footer = 2,
                                      parse_cols = "B:H,J:O",
                                      names =  range(0,13))

    #Change (%)やShare (%)の行を取り除く
    input_sheet_df = input_sheet_df[input_sheet_df[0] != 'Change (%)']
    input_sheet_df = input_sheet_df[input_sheet_df[0] != 'Share (%)']
    input_sheet_df = input_sheet_df[input_sheet_df[0] != 'Change(%)']
    input_sheet_df = input_sheet_df[input_sheet_df[0] != 'Share(%)']

    #データフレームの列名を振り直す
    input_sheet_df = input_sheet_df.rename(columns={0: 'year'})

    ####縦持ち横持ち変換####
    # 変換用に、空のデータフレームを生成
    input_trans_df = pd.DataFrame(index=[], columns=[])
    for i_m in range(1, len(input_sheet_df.columns)) : #データフレームの列について 1列目(1月)-12列目(12月)までループを回す     

        #year列（0列目）とi_m月（i_m列目）の列を抜き出す
        input_visitor_df = input_sheet_df.iloc[0:,[0, i_m]]

        #新たにmonthの列を追加し、その月の値 i_m をその列の成分に入力する
        input_visitor_df["month"] = i_m

        #人数の列をvisitorにrename
        input_visitor_df = input_visitor_df.rename(columns = {i_m:"visitor"})

        #以上で成形した各月のデータフレーム(input_df)を縦方向に結合していき、一つのデータフレームにしていく
        input_trans_df = pd.concat([input_trans_df,input_visitor_df])

    ####シート名の国名を表す列countryを追加####
    #シートの名前は、'韓国（総数）'みたいに余分な（総数）という部分がついているので正規表現で（総数）で取り除く
    input_sheet_name_resub = re.sub('（総数）', '', input_sheet_name[i_sheet])

    #新たにcountryの列を追加し、値を入力
    input_trans_df["country"] = input_sheet_name_resub

    #以上でシートごとに処理したデータフレームを一つのデータフレームに結合
    output_df = pd.concat([output_df,input_trans_df])


#旅行者(visitor)を整数型に変換
output_df = output_df.astype({'visitor':int})

#行のラベルを振り直す
output_df = output_df.reset_index(drop=True)

#列の順を year, month, visitorの順にする
output_df = output_df.ix[:,['year','month', "country",'visitor']]

#csvファイルとして出力してみる

#outpu file name
outputfile = 'visitor_by_month.csv'
print("Output file name is " + "\"" + outputfile + "\"")

#csv output
#output はutf-8 (windows Excelでは文字化けするが、utf-8の方が都合がいいことも多いので)
output_df.to_csv(outputfile,  sep=",", header=True, index =False,
                encoding="utf-8", line_terminator="\n" )

```

### Other method
```
#NaNを落とす
#dropna()を使う
input_sheet_df = input_sheet_df.dropna()

#NaN補完
#fillna(補完したい数とか)で可能
#空白を0人としてカウントしたいとかそういうとき
input_sheet_df = input_sheet_df.fillna(0)

#ある特定の要素だけの変更
#replaceを使う
#今回のデータだったら、国名を正式名に変えたいとかそういう場合
#(例) 以下では"米国"となっている要素だけを"アメリカ合衆国"に入れ替える
output_df = output_df.replace('米国', 'アメリカ合衆国')
```
