# ExcelSheetSorter
This Python program sort Excel sheet by your list.

### This program require the Libraly. Please, insatall.
openpyxl

## 説明 Explanation
Excelのシートを自分の用意した並び順のリストに従って並び替える。
辞書順や、番号順でない順番で並べる（例えば都道府県を北から並べる）ときに使える。

ソートリストは以下のような表形式のExcelファイルで用意してください。
このプログラムと同じディレクトリにリストを置くと動作します。

|index|sort_order|
---|---
|1|北海道|
|2|青森|
|3|秋田|
|4|岩手|
|…|…|
