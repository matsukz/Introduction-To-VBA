# テスト用まとめ<br>

紙媒体なら試験時に参照可ということでMarkdownの練習がてらに作成しているカンペです

### 目次
- [テスト用まとめ](#テスト用まとめ)
    - [目次](#目次)
  - [VBAの始め方](#vbaの始め方)
  - [セル関連](#セル関連)
  - [シート関連](#シート関連)
  - [計算](#計算)
  - [変数](#変数)
  - [条件分岐（IF)](#条件分岐if)
  - [その他便利機能](#その他便利機能)

## VBAの始め方
**はじめてVBAを作成するときにはExcelで設定を行う必要があります。**<br>
* 開発の有効化
    1. Excelを起動する
    2. 左下の`オプション`をクリック
    3. `リボンのユーザー設定`を選択
    4. `リボンのユーザー設定`内にある`開発`にチェックを入れる
   
1. Visual Basic Editorの表示
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

## セル関連

* セルの選択

    * セル`A1`選択するだけ<br>
        ```VBA:A1の選択
        Range("A1")
        ```

    * セル`A1からC3・D5`を選択する
        ```VBA:選択応用
        Range(A1:C3,D5)
        ```
        後述の文字入力や色変更にも使用できます。

* セル内の変更する

    * セル`A1`に数値を入力する
        ```VBA:A1に数値を入力
        Range("A1")=100
        ```
    * セル`A1`に文字を入力する
        ```VBA:A1に、文字を入力
        Range("A1")="あほばかまぬけ"
        ```
    * セル`A1`の色を変更する
        ```VBA:セル色変更
        Range("A1").Interior.ColorIndex=色コード
        ```
    * セル`A1`の文字の色を変更する
        ```
        Range("A1").Font.ColorIndex=色コード
        ```
        色コードについて<br>
        |  赤   |  青   |  黄   |  緑   |
        | :---: | :---: | :---: | :---: |
        |   3   |   5   |   6   |  10   |

    セルの複数選択に対応しています。

    * シート内セルの全クリア
        ```VBA:クリア
        Cells.delete
        ```
## シート関連
 **初期シート以外の選択は予めシートを追加しておく必要があります。**<br>
 **マクロが実行されるのは現在表示されるシート上となります。**<br>
* シートの選択（移動）
    * `sheet2`へ移動する。
        ```
        Worksheets("sheet2").select
        ```
    * `sheet2`のセル`A1`を選択する
        ```
        Worksheets("sheet2").Range("A1")
        ```
        上記の色変更などが利用できます。

## 計算
* 基本
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
        |      /       |    Mod     |   ^    |
* シートをまたいだ計算
    * 現在のシート`Sheet1`の`A1`に`B1`と`Sheet2`の`C1`を計算した結果を表示する
        ```VBA:シート演算
        Range("A1")=Range("B1")+Worksheets("sheet2").Range("C1")
        ```
## 変数
* 変数の設定
    * 変数の型について<br>
        どのような内容を変数に代入するか**変数を使用する前**に指定する必要があります。<br>
        | 値の形         | 型   | 使用RAM |備考|
        | :--------------: | :----: | :-------: |:----:|
        | 0から255の整数 | Byte | 2B      |
        |-32,768から32,767の整数|Integer|2B|
        |-2,147,483,648から2,147,483,647の整数|Long|4B|非推奨|
        |±3.4×10^38の少数|Single|4B||
        |約10×10^7の文字|String|2B||
        |すべての型|Variant|16B|非推奨|
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
* 変数の使い方
    * 数値`100`を代入する
        ```VBA:数値代入
        変数名 = 100
        ```
    * 文字列`あほばかまぬけ`を代入する
        ```VBA:文字列代入
        変数名 = "あほばかまぬけ"
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
## 条件分岐（IF)

* 基本構文
    ```VBA:IF基本
    IF 条件式1 Then
        '条件式1に合致したときの処理
    ElseIF 条件式2 Then
        '条件式1は合致しないが条件式2に合致したときの処理
    Else
        'どの条件にも合致しないときの処理
    End IF
    ```
    ElseIF部は必須ではありません。必要に応じて消したり加えたりしてください。

* 条件式の書き方
    * セル`A1`が`100`と等しいかを判断
        ```VBA:IF1
        IF Range("A1") = 100 Then
        ```
       * 利用できる算術演算子<br>
           [計算](#計算)と同じです。
    * セル`A1`の値が`90`以上か判断する
        ```VBA:IF2
        IF 90 <= Range("A1") Then
        ``` 
    * セル`A1`の値が`90`以上`100`未満かを判断する
        ```VBA:IF3
        IF 90 <= Range("A1") And Range("A1") < 100 Then
        ```

## その他便利機能
* オートフィル
    1. 数値のオートフィル<br>
    **連続したセルに数値が入力されている必要があります**
    * `A1`に100、`A2`に200と手で入力しから`A10`まで100・・200・・300と連続した数値を表示する
        ```VBA:数値オートフィル
        Range("A1:A2").AutoFill Destination:=Range("A1:A10")
        ```

    2. 文字列のオートフィル
    * `A1`に`月`と入力し`A1`から`A7`に一週間分の曜日を表示させる
        ```VBA:文字列オートフィル
        Range("A1").AutoFill Destination:=Range("A1:A7")
        ```
    Rangeでセルに値を入力してもOK
