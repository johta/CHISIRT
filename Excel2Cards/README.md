# Excel2Card

## これはなに
Excel2Cardです。Excelのリストからカードゲーム用のカードを自動生成します。

こんなかんじ↓↓↓
![figure](figure.png)


## 事前準備
- 必要なライブラリのインストール

  `pip install openpyxl　Pillow　textwrap3 `

- ディレクトリの作成

  `mkdir front/ rear/`


## 使い方
1. Excelのファイル名・シート名を[cardsheet.xlsx][sheet1]に変更します。
2. 不要な行・列を削除してください。
3. excel2card.pyを動かします。

  `python excel2card.py`

3. /front および /rear ディレクトリ配下にカードが生成されます

## 動作確認環境
- Python：3.7.3
- OS:macOS Catalina

## 注意
- ExcelファイルのフォーマットはCHISIRTカードゲームプロジェクトで使っているものに準拠しています。（持っていることが前提です。秘密のため上げません。）
- カードフォーマットは[調整中]です。
- anaconda環境だとpipとケンカしてうまく動かないかもしれません。
