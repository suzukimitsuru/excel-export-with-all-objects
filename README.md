# excel-extructor-with-all-objects

Microsoft Excel ブックの差分を見る為に、テキストファイルに書き出します。  
図形や画像も含めて比較できるツールが少ないため、作成しました。  

こちらのツールでは、比較もできます。

- [データを使用してブックを比較スプレッドシート検査 - Microsoft](https://support.microsoft.com/ja-jp/office/データを使用してブックを比較スプレッドシート検査-ebaf3d62-2af5-4cb1-af7d-e958cc5fad42) セルの比較ができます。
- [WinMergeでExcelの差分を比較しよう](https://tech.robotpayment.co.jp/entry/2023/03/23/070000) シートを画像にして比較できます。

## How to Use

ファイル名を完全パスで入力して、 Export ボタンを押します。  
excel-extructor は、フォルダを作成して、抽出結果を出力します。  

- sample
  - export.txt
  - Sheet3!Diagram_9.png
  - Sheet3!Picture_8.png

export.txt には、以下の様に出力します。

```
--- Sheet2:文字飾り セル ---
Sheet2!A1 "<色:0x808080>飾り無し</色>"
Sheet2!A2 "<色:0x808080><取り消し線>取り消し線</取り消し線></色>"
Sheet2!A3 "<色:0x808080><斜体>斜体</斜体></色>"
Sheet2!A4 "<色:0x808080><太字>太字</太字></色>"
Sheet2!A5 "<色:0x808080><太字><斜体>斜体で太字</斜体></太字></色>"

--- Sheet3:画像 図形 ---
Sheet3!"Rectangular Callout 3"(吹き出しの代替えテキスト) "<色:0xFFFFFF>吹き出しです。</色>"
Sheet3!"TextBox 4"(文字飾りの代替テキスト) "<太い二重下線>sharph</下線>Text"
Sheet3!"Right Arrow 6"(矢印の代替えテキスト) "<色:0xFFFFFF>右向き矢印</色>"
Sheet3!"TextBox 7"(横書きの
代替テキスト) "横書き"
Sheet3!"Picture 8"
Sheet3!"Diagram 9"(ピラミッド)
Sheet3!"Group 11"
    "Oval 12" "<色:0xFFFFFF>丸</色>"
    "Isosceles Triangle 13" "<色:0xFFFFFF>三角</色>"
    "Rectangle 14" "<色:0xFFFFFF>四角</色>"
```

## Use caution when export

画像を保存するため、クリップボードを使用しています。  
抽出中はクリップボードを使わないで下さい。  

## Excel for Mac Setup

Excel for Mac で excel-extructor-mac.xlsm を実行する場合、以下の AppleScript が必要です。  
コピー先のフォルダにコピーして下さい。  

- コピー元: `src/excel-extructor.applescript`
- コピー先: `~/Library/Application Scripts/com.microsoft.Excel/excel-extructor.applescript/`

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
