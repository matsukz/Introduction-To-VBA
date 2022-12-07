# はじめり<br>

紙媒体なら試験時に参照可ということでMarkdownの練習がてら作成しているカンペです。

# 目次
- [はじめり](#はじめり)
- [目次](#目次)
- [VBAの始め方](#vbaの始め方)
  - [開発の有効化](#開発の有効化)
  - [Visual Basic Editorの表示](#visual-basic-editorの表示)
  - [プロシージャの作成](#プロシージャの作成)
- [セル関連](#セル関連)
  - [セルの選択（絶対参照）](#セルの選択絶対参照)
  - [セル選択（相対参照）](#セル選択相対参照)
  - [セル内の変更する](#セル内の変更する)
- [シート関連](#シート関連)
  - [シートの選択（移動）](#シートの選択移動)
  - [よくあるエラー（シート関連）](#よくあるエラーシート関連)
- [計算](#計算)
  - [基本](#基本)
  - [シートをまたいだ計算](#シートをまたいだ計算)
- [変数](#変数)
  - [変数の型について](#変数の型について)
  - [変数の使い方](#変数の使い方)
  - [よくあるエラー（変数）](#よくあるエラー変数)
- [条件分岐](#条件分岐)
  - [基本構文](#基本構文)
  - [比較演算子について](#比較演算子について)
  - [よくあるエラー（条件分岐）](#よくあるエラー条件分岐)
  - [条件式の書き方](#条件式の書き方)
- [Forによる繰り返し](#forによる繰り返し)
  - [基本構文](#基本構文-1)
  - [よくあるエラー（繰り返しと条件分岐の組み合わせ）](#よくあるエラー繰り返しと条件分岐の組み合わせ)
- [その他便利機能](#その他便利機能)
  - [オートフィル](#オートフィル)
  - [メッセージボックス](#メッセージボックス)
  - [罫線](#罫線)
- [参考](#参考)
- [このノートについて](#このノートについて)

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
## プロシージャの作成

```VBA:プロシージャ作成
Sub マクロ名 ()

End Sub
```
`Sub`と`End Sub`の間に処理を記述します。<br>
実行するには`開発`タブ内の`マクロ`を選択します。

* マクロ名で使用できる文字について<br>
        [Visual Basic の名前付け規則](https://learn.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/visual-basic-naming-rules) （閲覧日：2022/11/30）

<div style="page-break-before:always"></div>

# セル関連
## セルの選択（絶対参照）

* セル`A1`選択するだけ<br>
    ```VBA:A1の選択
    Range("A1").Select
    ```

* セル`A1からC3`と`D5`を選択する
    ```VBA:選択応用
    Range(A1:C3,D5).Select
    ```
    `Range`での選択は複数のセル選択が可能です。<br>
    [次ページ](#セル内の変更する)の文字入力や色変更にも使用できます。
## セル選択（相対参照）

数値を用いてセルを選択します。[Forによる繰り返し](#forによる繰り返し)に向いています。基準は`A1`です。

```VBA:Cells
Cells(上下,左右)
```
* セル`A1`を選択する
    ```VBA:Cells_1
    Cells(1,1).Select
    ```
* セル`C5`を選択する
    ```VBA:Cells_2
    Cells(5,3).Select
    ```
* セル`F8`から下に変数`X`、右に変数`Y`移動したセルを選択する
    ```VBA:Cells_3
    Cells(8+X,6+Y).Select
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

* セル`B3`に文字列`テキスト`と入力する
    ```VBA:B3に文字を入力
    Range("B3")="テキスト"
    ```
    ```VBA:B3に文字を入力_2
    Cells(3,2)="テキスト"
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
    |  黒   |  赤   | 明るい緑 |  青   |  黄   |  緑   |
    | :---: | :---: | :------: | :---: | :---: | :---: |
    |   1   |   3   |    4     |   5   |   6   |  10   |

    `Range`によるセルの複数選択に対応しています。

<div style="page-break-before:always"></div>

# シート関連
 **初期シート以外を選択する際は事前にシートを追加しておく必要があります。**<br>
 **マクロは指定がない限りExcelで現在表示されるシート上での実行となります。**<br>

* シート内セルの全クリア
    ```VBA:クリア
    Cells.Delete
    ```

## シートの選択（移動）
* `sheet2`へ移動する
    ```
    Worksheets("sheet2").Select
    ```
* `sheet2`のセル`A1`を選択する
    ```
    Worksheets("sheet2").Range("A1").Select
    ```

    [前述](#セル内の変更する)の色変更などが利用できます。

## よくあるエラー（シート関連）
```VBA:sheet_ERROR
実行エラー'9':
インデックスが有効な範囲にありません。
 ```
  →存在しないワークシートを選択しようとしているので、シートの数や名前を確認してください。

<div style="page-break-before:always"></div>

# 計算
## 基本

計算結果を表示する場所の指定を忘れないようにしましょう。

* セル`A1`に計算結果を表示する
    ```VBA:計算1
    Range("A1")=1+2
    ```
    ```VBA:計算2
    Range("A1")=3-4
    ```
    ```VBA:計算3
    Range("A1")=5*6
    ```
    ```VBA:計算4
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
|          0から255の正の整数           |  Byte   |   2B    |
|        -32,768から32,767の整数        | Integer |   2B    |
| -2,147,483,648から2,147,483,647の整数 |  Long   |   4B    | 非推奨 |
|           ±3.4×10^38の少数            | Single  |   4B    |        |
|            約10×10^7の文字            | String  |   2B    |        |
|              すべての型               | Variant |   16B   | 非推奨 |

参考：[変数の型](https://katakago.sakura.ne.jp/pgm/vba/pgm01/var-type.html)（閲覧日：2022/12/06）
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

プロシージャ外で変数の宣言をすると、異なるプロシージャで同じ変数を利用できます。
```VBA:Dim_Ex1
Dim 変数名 As 型
変数名 = 値
Sub Hoge()

End Sub
```
<div style="page-break-before:always"></div>

## 変数の使い方
* 数値`100`を代入する
    ```VBA:数値代入
    変数名 = 100
    ```
* 文字列`テキスト`を代入する
    ```VBA:文字列代入
     変数名 = "テキスト"
     ```
* セル`A1`の値を代入する
    ```VBA:セル代入
    変数名 = Range("A1")
    ```
* セル`A1`と`B1`の値を合計した結果を代入する
    ```VBA:計算代入
    変数名 = Range("A1") + Range("B1")
    ```
* 変数の値をセル`A1`に表示する
    ```VBA:変数表示
    Range("A1") = 変数名
    ```

## よくあるエラー（変数）
* オーバーフロー
    ```VBA:DimERROR_1
    実行エラー'6':
    オーバーフローしました。
    ```
    →代入できる値の範囲を超えています。より大きな方で宣言し直してください。
* 不一致
    ```VBA:DimERROR_2
    実行エラー'13':
    型が一致しません。
    ```
    →数値の型に文字を代入しようとしていませんか？

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
* 記述時の注意
    * 処理を記述しないと **なにもしない**という処理になります。<br>
    * インデントは視認性のためにTabキーにて挿入しています。
    * ElseIF部は **必須ではありません。** 必要に応じて消したり加えたりしてください。<br>
    * [プロシージャ作成](#visual-basic-editorの表示)のEnd Subは自動入力されますが、**End IFは自動入力されません。**

## 比較演算子について

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
|  NOT   |  NOT A  |      条件Aに合致しないとき       |

利用できる算術演算子は[計算](#計算)と同じです。

## よくあるエラー（条件分岐）
```VBA:IF_ERROR
コンパイルエラー:
修正候補:Then または GoTo
```
→条件式末尾のThenが抜けています。
<div style="page-break-before:always"></div>

## 条件式の書き方

条件には数値や文字列、`Range`、`Cells`、変数が利用できます。

* セル`A1`が数値`100`と等しいかを判断
    ```VBA:IF1
    IF Range("A1") = 100 Then
    ```

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

* セル`A1`の値が**偶数ではない**ことの判断（NOTの利用）
  ```VBA:IF5
  IF NOT Range("A1") Mod 2 = 0
  ```
  
<div style="page-break-before:always"></div>

# Forによる繰り返し

特定の操作を任意の回数繰り返すことができます。
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
| 利用する変数 | 開始値 | 終了値 | 1回のループで増える`i`の数 |

セル選択には[セル選択（相対参照）](#セル選択相対参照)が便利です。

* セル`A1`から`A100`まで下に１ずつセルを選択する
    ```VBA:for_1
    Dim i As Byte
    For i = 0 To 99
        Cells(1+i,1).Select
    Next i
    ```

* （応用）`A1`から`A10`に入力されている数が偶数か奇数かを判断する
    ```VBA:IF3
    Dim i As Byte
        For i = 0 To 99 Step 1
            IF Cells(1+i,1) Mod 2 = 0 Then
                '偶数のときの処理
            Else
                '奇数のときの処理
            End IF
        Next i
    ```
## よくあるエラー（繰り返しと条件分岐の組み合わせ）
```VBA:For_ERROR
コンパイルエラー:
Nextに対応するForがありません。
```
→End IFを忘れていませんか？

<div style="page-break-before:always"></div>

# その他便利機能
## オートフィル
1. 数値のオートフィル<br>
**連続したセルに数値が入力されている必要があります**

* `A1`に100、`A2`に200と手で入力し、`A10`まで100・・200・・300と連続した数値を表示する
    ```VBA:数値オートフィル
    Range("A1:A2").AutoFill Destination:=Range("A1:A10")
    ```

1. 文字列のオートフィル
* `A1`に`月`と入力し`A1`から`A7`に一週間分の曜日を表示させる
    ```VBA:文字列オートフィル
    Range("A1").AutoFill Destination:=Range("A1:A7")
    ```
    手入力や`Range`・`Cells`による方法でセルに値を入力してもOK

<div style="page-break-before:always"></div>

## メッセージボックス
* メッセージダイアログで文字列`テキスト`を表示する
    ```VBA:Msg_1
    MsgBox "テキスト"
    ```
    数値を表示させる際でも`""`で囲む必要があります。

* メッセージボックスでセル`A1`の内容を表示する
    ```VBA:Msg_2
    MsgBox Range("A1")
    ```
* メッセージボックスで変数`TEXT`の内容を表示させる
    ```VBA:Msg_3
    MsgBox TEXT
    ```
* （応用）メッセージボックスに`A1の値は（A1の値）です`と表示する
    ```VBA_4
    MsgBox "A1の値は" & Range("A1") & "です"
    ```
    `&`を挟むことで結合できます。
* （応用）セル`A1`の内容を表示するメッセージボックスのタイトル`A1の中身は？`を設定する
    ```VBA:_5
    MsgBox Range("A1"), vbOKOnly, "A1の中身は？"
    ```
    * メッセージボックスについて<br>
    [第23回.メッセージボックス(MsgBox関数)](https://excel-ubara.com/excelvba1/EXCELVBA323.html)（閲覧日：2022/12/01）

<div style="page-break-before:always"></div>

## 罫線
```VBA:Border
罫線を引く範囲.BorderAround 線のスタイル定数, 線の太さ定数, 色コード
```
* 枠線を引く範囲には[3ページ](#セル関連)の`Range`や`Cells`が利用できます。
* 色コードには[4ページ](#セル内の変更する)の`ColorIndex`が利用できます。

* 線のスタイルについて
    | 罫線の種類 |      定数       |
    | :--------: | :-------------: |
    |   線なし   | xlLineStyleNone |
    |   一重線   |  xlContinuous   |
    |   二重線   |    xlDouble     |
    |    破線    |     xlDash      |
    |  一点鎖線  |    xlDashDot    |
    |  二点鎖線  |  xlDashDotDot   |
    |    点線    |      xlDot      |
    |   斜破線   | xlSlantDashDot  |
* 線の太さについて
    |  太さ  |    定数    |
    | :----: | :--------: |
    | 極細線 | xlHairline |
    |  細線  |   xlThin   |
    | 中太線 |  xlMedium  |
    |  太線  |  xlThick   |

* セル`C1`の周りに赤色の細い破線を引く
```VBA:Border_1
Range("C1").BorderAround xlDash, xlThin, 3
```
* セル`C1`から`G5`の周りに青色の太い一重線を引く
```VBA:Border_2
Range("C1:G5").BorderAround xlContinuous, xlThick, 5
```

<div style="page-break-before:always"></div>
<ノート>
<div style="page-break-before:always"></div>

# 参考
FOM出版　よくわかるMicrosoft Excel 2019/2016/2013 マクロ/VBA

[Visual Studio Code拡張 Markdown All in One](https://zenn.dev/ctrlkeykoyubi/articles/vscode-markdown-all-in-one)（閲覧日：2022/11/30）

[Qiita マークダウン記法 一覧表・チートシート](https://qiita.com/kamorits/items/6f342da395ad57468ae3)（閲覧日：2022/11/30）

[Visual Basic の名前付け規則](https://learn.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/visual-basic-naming-rules)（閲覧日：2022/11/30）

[変数の型](https://katakago.sakura.ne.jp/pgm/vba/pgm01/var-type.html)（閲覧日：2022/12/06）

[エクセルの真髄 第23回.メッセージボックス(MsgBox関数)](https://excel-ubara.com/excelvba1/EXCELVBA323.html)（閲覧日：2022/12/01）

[VScodeのMarkdownからPDF変換時に改ページを挿入](https://qiita.com/0xmks/items/4fec4116bb42120f5180)（閲覧日2022/12/01）


# このノートについて
* 作成：
    **matsukz**
* レポジトリURL：
    https://github.com/matsukz/3-1ClassNote
* 最終更新日：
    **2022年12月07日**
* 皆伝レベル：
    **１**
* 初音ミクのキャラクターランク：
    **59**
* 彼女：
    **なし**
