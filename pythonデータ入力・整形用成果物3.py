import pandas as pd

#整形用ファイルの読み込み
df = pd.read_excel('実務Excel練習3.xlsx')

#電話番号の「-」を削除
df['電話番号'] = df['電話番号'].str.replace('-','')

#金額の「,」を削除
df['金額'] = df['金額'].astype(str).replace(',','')

#金額を数値変換
df['金額'] = df['金額'].astype(int)

#金額の合計を算出
total = df['金額'].sum(numeric_only=True)

#合計金額をExcelのA4セルへ設定
df.loc['合計金額'] = ['合計金額',total,'','']

#整形したデータをExcelへ出力
df.to_excel('整形済みデータ3.xlsx',sheet_name='結果',index=False)

print(df)
print(total)