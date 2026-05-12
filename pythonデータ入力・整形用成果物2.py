import pandas as pd

#整形用ファイルの読み込み
df = pd.read_excel('実務Excel練習2.xlsx')

#合計列追加
df['合計'] = df['単価'] * df['数量']

#合計降順並び替え
df['合計'].sort_values(ascending=False)

#上位2件抽出
df = df.head(2)

#整形データ確認用
print(df)

#整形したデータをExcelファイルへ出力
df.to_excel('整形済みデータ2.xlsx',sheet_name='結果',index=False)