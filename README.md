# excel-extructor-with-all-objects

Microsoft Excel ブックの差分を見る為に、テキストファイルに書き出します。  
図形や画像も含めて比較できるツールが少ないため、作成しました。  

## Use caution when export

画像を保存するため、クリップボードを使用しています。  
抽出中はクリップボードを使わないで下さい。  

## Loadmap

- v0.0.1 テキストの抽出
  - セルのテキストを抽出する
  - 図形のテキストを抽出する
- v0.0.2 修飾の抽出
  - セルのテキストの修飾を抽出する(色・太字・斜体・取り消し線・下線)
  - 図形のテキストの修飾を抽出する(色・太字・斜体・取り消し線・下線)
  - コメントとスレッドを抽出する
- v0.0.3 画像をファイルに抽出する
- v1.0.0 正式版
  - グループ化した図形を抽出する
  - 経過表示を追加

- 他に出力したい物があれば、[issue](https://github.com/suzukimitsuru/excel-extructor-with-all-objects/issues) でリクエストして下さい。 

## Excel for Mac Setup

Excel for Mac で excel-extructor-mac.xlsm を実行する場合、以下の AppleScript が必要です。  
コピー先のフォルダにコピーして下さい。  

- コピー元: `src/excel-extructor.applescript`
- コピー先: `~/Library/Application Scripts/com.microsoft.Excel/excel-extructor.applescript/`

## Similar tools

- [データを使用してブックを比較スプレッドシート検査 - Microsoft](https://support.microsoft.com/ja-jp/office/データを使用してブックを比較スプレッドシート検査-ebaf3d62-2af5-4cb1-af7d-e958cc5fad42)
- [WinMergeでExcelの差分を比較しよう](https://tech.robotpayment.co.jp/entry/2023/03/23/070000)
