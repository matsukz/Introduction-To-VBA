# はじめに<br>

紙媒体なら試験時に参照可ということでMarkdownの練習がてらに作成しているカンペです

### 目次
- [はじめに](#はじめに)
    - [目次](#目次)
- [VBAの始め方](#vbaの始め方)
  - [開発の有効化](#開発の有効化)
  - [Visual Basic Editorの表示](#visual-basic-editorの表示)
- [セル関連](#セル関連)
  - [セルの選択（絶対参照）](#セルの選択絶対参照)
  - [セル選択（相対参照）](#セル選択相対参照)
  - [セル内の変更する](#セル内の変更する)
- [シート関連](#シート関連)
  - [シートの選択（移動）](#シートの選択移動)
- [計算](#計算)
  - [基本](#基本)
  - [シートをまたいだ計算](#シートをまたいだ計算)
- [変数](#変数)
  - [変数の型について](#変数の型について)
  - [変数の使い方](#変数の使い方)
- [条件分岐](#条件分岐)
  - [基本構文](#基本構文)
  - [条件式の書き方](#条件式の書き方)
    - [比較演算子について](#比較演算子について)
- [Forによる繰り返し](#forによる繰り返し)
  - [基本構文](#基本構文-1)
- [その他便利機能](#その他便利機能)
  - [オートフィル](#オートフィル)
  - [メッセージボックス](#メッセージボックス)
- [参考](#参考)
- [随時更新！](#随時更新)

<div style="page-break-before:always"></div>

# VBAの始め方
**はじめてVBAを作成するときにはExcelで設定を行う必要があります。**<br>
## 開発の有効化
1. Excelを起動する
2. 左下の`オプション`をクリック
3. `リボンのユーザー設定`を選択
4. `リボンのユーザー設定`内にある`開発`にチェックを入れる
   
## Visual Basic Editorの表示
1. 開発タブをクリック
2. `Visual Basic`をクリック
3. `挿入`から`標準モジュール`を選択
2. プロシージャの作成
    ```VBA:プロシージャ作成
    Sub マクロ名 ()

    End Sub
    ```
   と入力します<br>
* マクロ名で使用できる文字について<br>
        [Visual Basic の名前付け規則](https://learn.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/visual-basic-naming-rules) （閲覧日：2022/11/30）

<div style="page-break-before:always"></div>

# セル関連
## セルの選択（絶対参照）

* セル`A1`選択するだけ<br>
    ```VBA:A1の選択
    Range("A1")
    ```

* セル`A1からC3`と`D5`を選択する
    ```VBA:選択応用
    Range(A1:C3,D5)
    ```
    Rangeでの選択は複数のセル選択に利用できます。<br>
    後述の文字入力や色変更にも使用できます。
## セル選択（相対参照）

数値を用いてセルを選択します。[Forによる繰り返し](#forによる繰り返し)に向いています。基準は`A1`です。

```VBA:Cells
Cells(上下,左右)
```
* セル`A1`を選択する
    ```VBA:Cells_1
    Cells(1,1)
    ```
* セル`C5`を選択する
    ```VBA:Cells_2
    Cells(5,3)
    ```
* セル`F8`から下に変数`X`、右に変数`Y`移動したセルを選択する
    ```VBA:Cells_3
    Cells(8+X,6+Y)
    ```
    移動させる値がマイナスになれば反対方向に動きます。

<div style="page-break-before:always"></div>

## セル内の変更する

* セル`A1`に数値を入力する
    ```VBA:A1に数値を入力
    Range("A1")=100
    ```
    ```VBA:A1に数値を入力_2
    Cells(1,1)=100
    ```

* セル`B3`に文字列`あほばかまぬけ`と入力する
    ```VBA:B3に文字を入力
    Range("B3")="あほばかまぬけ"
    ```
    ```VBA:B3に文字を入力_2
    Cells(3,2)="あほばかまぬけ"
    ```

* セル`C4`の色を変更する
    ```VBA:セル色変更
    Range("C4").Interior.ColorIndex=色コード
    ```
    ```VBA:セル色変更_2
    Cells(4,3).Interior.ColorIndex=色コード
    ```

* セル`D5`の文字の色を変更する
    ```VBA:文字色変更
    Range("D5").Font.ColorIndex=色コード
    ```
    ```VBA:文字色変更_2
    Cells(5,4).Font.ColorIndex=色コード
    ```
    色コードについて<br>
    |  赤   | 明るい緑 |  青   |  黄   |  緑   |
    | :---: | :------: | :---: | :---: | :---: |
    |   3   |    4     |   5   |   6   |  10   |

    Rangeによるセルの複数選択に対応しています。

<div style="page-break-before:always"></div>

# シート関連
 **初期シート以外を選択する際は予めシートを追加しておく必要があります。**<br>
 **マクロが実行されるのは現在表示されるシート上となります。**<br>

* シート内セルの全クリア
    ```VBA:クリア
    Cells.delete
    ```

## シートの選択（移動）
* `sheet2`へ移動する。
    ```
    Worksheets("sheet2").select
    ```
* `sheet2`のセル`A1`を選択する
    ```
    Worksheets("sheet2").Range("A1")
    ```

    上記の色変更などが利用できます。

<div style="page-break-before:always"></div>

# 計算
## 基本
* セル`A1`に計算結果を表示する
    ```VBA:四則計算
    Range("A1")=1+2
    Range("A1")=3-4
    Range("A1")=5*6
    Range("A1")=7/8
    ```
* セル`A1`に`B2`と`C4`の計算結果を表示する
    ```
    Range("A1")=Range("B2")+Range("C4")
    ```
    セル同士計算でも四則計算同様の演算子を利用します。
* その他の演算子
    | 商（整数部） | 商（余り） | べき乗 |
    | :----------: | :--------: | :----: |
    |      \       |    Mod     |   ^    |
## シートをまたいだ計算
* 現在のシート`Sheet1`の`A1`に`B1`と`Sheet2`の`C1`を計算した結果を表示する
    ```VBA:シート演算
    Range("A1")=Range("B1")+Worksheets("sheet2").Range("C1")
    ```

<div style="page-break-before:always"></div>

# 変数
## 変数の型について

どのような内容を変数に代入するか**変数を使用する前**に指定する必要があります。<br>
|                値の形                 |   型    | 使用RAM |  備考  |
| :-----------------------------------: | :-----: | :-----: | :----: |
|            0から255の整数             |  Byte   |   2B    |
|        -32,768から32,767の整数        | Integer |   2B    |
| -2,147,483,648から2,147,483,647の整数 |  Long   |   4B    | 非推奨 |
|           ±3.4×10^38の少数            | Single  |   4B    |        |
|            約10×10^7の文字            | String  |   2B    |        |
|              すべての型               | Variant |   16B   | 非推奨 |
* 整数を代入する変数の作成
    ```VBA:変数（整数）
    Dim 変数名 As Integer
    ```
* 少数に対応した変数の作成
    ```VBA:変数（少数）
    Dim 変数名 As Single
    ```
* 文字列に対応した変数の作成
    ```VBA:変数（文字列）
    Dim 変数名 As String
    ```
    * 変数に使える文字について<br>
        [Visual Basic の名前付け規則](https://learn.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/visual-basic-naming-rules)（閲覧日：2022/11/30）

<div style="page-break-before:always"></div>

## 変数の使い方
* 数値`100`を代入する
    ```VBA:数値代入
    変数名 = 100
    ```
* 文字列`あほばかまぬけ`を代入する
    ```VBA:文字列代入
     変数名 = "あほばかまぬけ"
     ```
* セル`A1`の値を代入する
    ```VBA:セル代入
    変数名 = Range("A1")
    ```
* セル`A1`と`B1`の値を足した値を代入する
    ```VBA:計算代入
    変数名 = Range("A1") + Range("B1")
    ```
* 変数の値をセル`A1`に表示する
    ```VBA:変数表示
    Range("A1") = 変数名
    ```

<div style="page-break-before:always"></div>

# 条件分岐
## 基本構文
```VBA:IF基本
IF 条件式1 Then
    '条件式1に合致したときの処理
ElseIF 条件式2 Then
     '条件式1は合致しないが条件式2に合致したときの処理
Else
    'どの条件にも合致しないときの処理
End IF
```
ElseIF部は **必須ではありません。** 必要に応じて消したり加えたりしてください。<br>
[プロシージャ作成](#visual-basic-editorの表示)のEnd Subは自動入力されますが、**End IFは自動入力されません。**

## 条件式の書き方
* セル`A1`が`100`と等しいかを判断
    ```VBA:IF1
    IF Range("A1") = 100 Then
    ```
   * 利用できる算術演算子<br>
       [計算](#計算)と同じです。

### 比較演算子について

条件Aと条件Bが存在するとき

| 演算子 | 利用例  |               意味               |
| :----: | :-----: | :------------------------------: |
|   =    |  A = B  |           AとBは等しい           |
|   <    |  A < B  |          AはBより小さい          |
|   <=   | A <= B  |       AはBと等しいか小さい       |
|   >    |  A > B  |         AはBよりも大きい         |
|   >=   | A >= B  |       AはBと等しいか大きい       |
|  AND   | A AND B |   AとB両方の条件が合致している   |
|   OR   | A OR B  | AとBどちらかの条件に合致している |
|  NOT   |  NOT A  |       条件に合致しないとき       |
<br>
<br>
<br>

* セル`A1`の値が`90`以上か判断する
    ```VBA:IF2
    IF 90 <= Range("A1") Then
    ``` 
* セル`A1`の値が`90`以上`100`未満かを判断する<br>
  悪い例
    ```VBA:IF3
    IF 90 <= Range("A1") < 100 Then
    ```
  正しい式
  ```VBA:IF4
  IF 90 <= Range("A1") And Range("A1") < 100 Then
  ```

<div style="page-break-before:always"></div>

# Forによる繰り返し

特定の操作を任意の回数繰り返すことができます。<br>

## 基本構文

**数値に対応した変数を設定しておく必要があります。**
```VBA:For基本
Dim i As Byte
For i = A to B Step C
    '繰り返す操作
Next i
```
|      i       |   A    |   B    |             C              |
| :----------: | :----: | :----: | :------------------------: |
| 繰り返す回数 | 開始値 | 終了値 | 1回のループで増える`i`の数 |

セル選択には[セル選択（相対参照）](#セル選択相対参照)が有効です。

* セル`A1`から`A100`まで下に１ずつセルを選択する
    ```VBA:for_1
    Dim i As Byte
    For i = 0 To 100
        Cells(1+i,1)
    Next i
    ```
    IF同様`Next i`は自動入力されません。<br>
    IFによる条件分岐を繰り返す際は文末の**End IFを忘れないように**しましょう。

<div style="page-break-before:always"></div>

# その他便利機能
## オートフィル
1. 数値のオートフィル<br>
**連続したセルに数値が入力されている必要があります**

* `A1`に100、`A2`に200と手で入力してら`A10`まで100・・200・・300と連続した数値を表示する
    ```VBA:数値オートフィル
    Range("A1:A2").AutoFill Destination:=Range("A1:A10")
    ```

2. 文字列のオートフィル
* `A1`に`月`と入力し`A1`から`A7`に一週間分の曜日を表示させる
    ```VBA:文字列オートフィル
    Range("A1").AutoFill Destination:=Range("A1:A7")
    ```
    [#セル内の変更](#セル内の変更する)による方法でセルに値を入力してもOK

<div style="page-break-before:always"></div>

## メッセージボックス
* メッセージダイアログで文字列`あほばかまぬけ`を表示する
    ```VBA:Msg_1
    MsgBox "あほばかまぬけ"
    ```
    数値を表示させる際`""`で囲む必要があります。

* メッセージボックスでセル`A1`の内容を表示する
    ```VBA:Msg_2
    MsgBox Range("A1")
    ```
* メッセージボックスで変数`TEXT`の内容を表示させる
    ```VBA:Msg_3
    MsgBox TEXT
    ```
* （応用）セル`A1`の内容を表示するメッセージボックスのタイトル`A1の中身は？`を設定する
    ```VBA:_4
    MsgBox Range("C15"), vbOKOnly, "A1は？"
    ```
    * メッセージボックスについて<br>
    [第23回.メッセージボックス(MsgBox関数)](https://excel-ubara.com/excelvba1/EXCELVBA323.html)（閲覧日：2022/12/01）

<div style="page-break-before:always"></div>

# 参考
[Visual Studio Code拡張 Markdown All in One](https://qiita.com/kamorits/items/6f342da395ad57468ae3)（閲覧日：2022/11/30）

[Qiita マークダウン記法 一覧表・チートシート](https://zenn.dev/ctrlkeykoyubi/articles/vscode-markdown-all-in-one)（閲覧日：2022/11/30）

[エクセルの真髄 第23回.メッセージボックス(MsgBox関数)](https://excel-ubara.com/excelvba1/EXCELVBA323.html)（閲覧日：2022/12/01）

[VScodeのMarkdownからPDF変換時に改ページを挿入](https://qiita.com/0xmks/items/4fec4116bb42120f5180)（閲覧日2022/12/01）


# 随時更新！
* 作成：
    **Ma2kzzzz**
* 最終更新日：
    **2022年12月01日**
* 皆伝レベル：
    **１**
* 初音ミクのキャラクターランク：
    **59**
