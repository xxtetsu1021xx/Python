import pandas as pd

#整形用ファイル読み込み
df = pd.read_excel('実務Excel練習1.xlsx')

#空白年齢を0に変換
df['年齢'] = df['年齢'].fillna(0)

#重複削除
df = df.drop_duplicates(subset='日付')

#日付順に並び替え
df = df.sort_values('日付',ascending=True)

#Excel出力内容事前確認用
print(df)

#整形したデータをExcelファイルへ出力
df.to_excel('整形済みデータ1.xlsx',sheet_name='結果',index=False)

