import pandas as pd

#整形用のデータ読み込み
df = pd.read_excel('実務Excel練習4.xlsx')

#年齢空白を「0」に変換
df['年齢'] = df['年齢'].fillna(0)

#同じ行のデータを削除
df = df.drop_duplicates()

#金額のカンマ削除
df['金額'] = df['金額'].astype(str).replace(',','')

#金額の数値変換
df['金額'] = df['金額'].astype(int)

#電話のハイフンを削除
df['電話'] = df['電話'].str.replace('-','')

#日付を日付型へ変換
df['日付'] = pd.to_datetime(df['日付'])

#日付の昇順へ並び替え
df.sort_values('日付' ,ascending=True)

#整形データ確認用
print(df)

#整形後のデータをExcelに出力
df.to_excel('整形済みデータ4.xlsx',sheet_name='結果',index=False)