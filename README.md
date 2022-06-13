# excelstl
頂点座標と三角形のインデックスリストを書いたExcelファイルを読み込んで、STLファイルを出力するアプリケーション。
An application that reads an Excel file with vertex and triangle arrays and outputs an STL file.

# 使い方

## 実行ファイルの配布
https://github.com/sasekazu/excelstl/releases

## Excelファイルの準備
Sheet1に頂点座標（頂点番号は1行目が1）を、Sheet2にインデックスリストを書く。
インデックスリストは頂点の順番が三角形の表から見て反時計回りになるようにする（右ネジの方向が法線ベクトルと一致するように）。
例えば3辺が10mmの四面体では、以下のようにする。

Sheet1

| x | y | z |
| ---- | ---- | ---- | 
|  0 |  0 |  0 |
| 10 |  0 |  0 |
|  0 | 10 |  0 |
|  0 |  0 | 10 |

※ヘッダーのx, y, zは書かない。書くとエラーになる。

Sheet2

| v1 | v2 | v3 |
| ---- | ---- | ---- | 
| 1 | 3 | 2 |
| 1 | 4 | 3 |
| 1 | 2 | 4 |
| 2 | 3 | 4 |

※ヘッダーのv1, v2, v3は書かない。書くとエラーになる。

[サンプルExcelファイル](https://github.com/sasekazu/excelstl/blob/main/tetra-10mm.xlsx)

## STLファイルの生成
excelstlを実行し、ウィンドウにExcelファイルをドラッグアンドドロップする。
名前を付けて保存ダイアログが出るので、保存場所とファイル名を入力し保存する。

[出力例](https://github.com/sasekazu/excelstl/blob/main/tetra-10mm.stl)
