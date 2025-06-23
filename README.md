# ExcelSTL
頂点座標と三角形のインデックスリストを書いたExcelファイルを読み込んで、STLファイルを出力するアプリケーション。
An application that reads an Excel file with vertex and triangle arrays and outputs an STL file.

# 使い方

## 実行ファイルの配布
https://github.com/sasekazu/excelstl/releases

## Excelファイルの準備
Sheet1に頂点座標（頂点番号は1行目が1）を、Sheet2にインデックスリストを書く。
2Dモード（押し出しモード）と3Dモードがある。

### 2Dモード
2次元空間の平面形状を入力し、押し出し処理によって3Dモデルを作成するモード。
例えば1辺1cmの正方形を入力し、押し出しによって直方体にしたい場合は、以下のようにする。

Sheet1（頂点座標）

| x | y |
| ---- | ---- |
| 0 | 0 |
| 1 | 0 | 
| 1 | 1 |
| 0 | 1 |

※ヘッダーのx, yは書かない。書くとエラーになる。

Sheet2（三角形インデックスリスト）

| v1 | v2 | v3 |
| ---- | ---- | ---- | 
| 1 | 2 | 3 | 
| 1 | 3 | 4 | 

※ヘッダーのv1, v2, v3は書かない。書くとエラーになる。
※インデックスリストは頂点の順番が三角形の表から見て反時計回りに書かれていることが望ましいが、2DモードではExcelSTLが自動的に修正する。

[サンプルExcelファイル](https://github.com/sasekazu/excelstl/blob/main/sample/square-1cm.xlsx)

### 3Dモード
三次元空間で閉じた形状を直接入力するモード。
3Dモードではインデックスリストは頂点の順番が三角形の表から見て反時計回りになるようにする（右ネジの方向が法線ベクトルと一致するように）。
例えば3辺が1cmの四面体では、以下のようにする。

Sheet1（頂点座標）

| x | y | z |
| ---- | ---- | ---- | 
| 0 | 0 | 0 |
| 1 | 0 | 0 |
| 0 | 1 | 0 |
| 0 | 0 | 1 |

※ヘッダーのx, y, zは書かない。書くとエラーになる。

Sheet2（三角形インデックスリスト）

| v1 | v2 | v3 |
| ---- | ---- | ---- | 
| 1 | 3 | 2 |
| 1 | 4 | 3 |
| 1 | 2 | 4 |
| 2 | 3 | 4 |

※ヘッダーのv1, v2, v3は書かない。書くとエラーになる。

[サンプルExcelファイル](https://github.com/sasekazu/excelstl/blob/main/sample/tetra-1cm.xlsx)

## STLファイルの生成
ExcelSTLを実行し、ウィンドウにExcelファイルをドラッグアンドドロップする。
頂点座標が2列の場合は2Dモード、3列の場合は3Dモードが自動的に選択され実行される。
名前を付けて保存ダイアログが出るので、保存場所とファイル名を入力し保存する。

[2Dモード出力例](https://github.com/sasekazu/excelstl/blob/main/sample/square-1cm.stl)
[3Dモード出力例](https://github.com/sasekazu/excelstl/blob/main/sample/tetra-1cm.stl)
