# 基礎 Visual Studio Code まとめ
---
## Part1 はじめての VisualBasic プログラミング

略

---
## Part2 Visual Basicの基礎を身に付ける
### <u>Chapter3 数値や文字列を扱う</u>
<br>

- キーワード

```
And As Boolean Byte Call Case Catch Class Const Continue Date Default Dim
Do Double Each Else Elseif End EndIf Error Event Exit False For Function
Get Handles If Imports In Integer Is Long Loop Me Mod New Next Not Nothing
Object Of On Option Or Private Property Public Return Select Short Single
Static Step Stop String Sub Then Throw To True Try Wend When While With
```

- n 進数

```
'&B	2進数  
'&O	8進数  
'&H	16進数 
```

- 変数の宣言

```
Dim (変数名) As (データ型)
Dim i As Integer
Dim weight As Double
Dim ClientName As String
```
<br>


### <u>Chapter4 条件によって処理を変える</u>

- AndAlso 演算子と OrElse 演算子
  AldAlso 演算では、前に指定した式が False なら後ろの式を評価せずに、False を返す.
  OrElse 演算子では　前に指定した式が True なら後ろの式を評価せずに、結果としてそのまま True を返す.
  AndAlso 演算や OrElse 演算子は、後ろの式が前の式に依存するような場合に便利.

```
If obj Is reOption Ans obj.Checked Then
    MessageBox.Show("ラジオボタンがチェックされています")
End If
'objがCheckedプロパティのないコントロールを参照している場合、Andの後ろの
'obj.Checkedという部分にエラーがでてしまう.
```
<br>

### <u>Chapter5 処理を繰り返す</u>

- Until型とWhile型のDo...Loopステートメント

Until型では、条件が満たされていない間、繰り返しを実行し、While型では、じょうけんが満たされている間、繰り返しを実行する.
  - 後判断-Until型
  ```
  Dim i As Integer = 0
  'iが5以上になるまで繰り返す.
  Do
    Debug.WriteLine("Hello World")
    i += 1
  Loop Until i > 5
  ```
  - 後判断-While型

  ```
  Do i As integer = 0
  'iが5未満の内は繰り返す.
  Do
    Debug.WriteLine("Hello World")
    i += 1
  Loop While i < 5
  ```

  - 前判断-Until型
   ```
  Dim i As Integer = 0
  'iが5以上になるまで繰り返す.
  Do Until i > 5
    Debug.WriteLine("Hello World")
    i += 1
  Loop 
  ```

  - 前判断-While型  
  ```
  Do i As integer = 0
  'iが5未満の内は繰り返す.
  Do While i < 5
    Debug.WriteLine("Hello World")
    i += 1
  Loop 
  ```
- For...Nextステートメント

For...Nextステートメントはある変数の初期値、終了値、増分値で決まるような繰り返しが簡潔に書ける.

```
'For 変数名 = 初期値 To 終了値 Step 増分値(1の際は省略可能)
'	・・・
'Next 変数名

'ex)
Dim i As Integer
For i = 0 To 4 Step 1
  Debug.WriteLine("Hello World")
Next i
```

- For Each...Nextステートメント

複数のデータを含むようなデータ（コレクション）の個々の要素を順に取り扱う事ができる.
インデックスを使用するとコントロールに項目を追加したり削除したりする事で、正しくインデックで表せない事が出てくるが、ForEach...Nextであれば表せることもある.
```
'For Each 変数名 As データ型 In コレクション
'	・・・
'Next 変数名

'ex)
For Each anItem As String In ListBox1.Items
	Debug.WriteLine(anItem)
Next
```
<br>

### <u>Chapter6 配列を利用する</u>

- 配列の宣言
```
' Dim 変数名(インデックスの最大値) As データ型

'ex)
Dim Score(9) 						As Integer
Dim DiscountRate(2) 		As Double
Dim Grade(3) 						As String 
Dim DataMatrix(2, 3)		As Integer
```

- 配列のサイズの変更

配列のサイズを変更するにはReDimステートメントを用いる.
ReDimで配列のサイズを変更した場合は、元の要素の値が失われる.
元の要素を保持するには、Preserveキーワードを付ける.
```
Dim x() As Double = {0.5, 1.0, 1.5}
Dim y() As Double = {0.5, 1.0, 1.5} 

ReDim x(5)
ReDim Preserve y(5)

Debug.Write(x(1).ToString("F1")) '0.0と出力される
Debug.Write(y(1).ToString("F1")) '1.0と出力される
```
<br>

### <u>Chapter7 プロシージャを使ってコードをまとめる</u>

- SubプロシージャとFunctionプロシージャ

返し値を持たない(Void型)プロシージャの場合はSubプロシージャを、
返し値を持つプロシージャを作成する場合はFunctionプロシージャを用いる.

- Subプロシージャの作成

Private/Publicでアクセス範囲を指定（Functionプロシージャも同様）
```
'Private(同フォーム内からのみアクセス)
Private Sub SetBackColor (Red As Integer, Green As Integer)

End Sub

'Public(他のフォームからもアクセス可能)
Public Sub SetBackColor (Red As Integer, Green As Integer)

End Sub
```

- Functionプロシージャの作成

Functionプロシージャでは戻り値のデータ型も定義する
```
'戻り値がstring型のFunctionプロシージャ
Private Function GetDayOfWeek(Dow As Integer) As String

End Function
```

- 値渡しと参照渡し

プロシージャの引数の定義時にByVal/ByRefを付ける事で値渡しと参照渡しを指定

```
'arg1,arg2は値がコピーされ、呼び出し側の引数を変える事はないが、
'ansを変更することで、呼び出し元のansも値を変更できる
Sub Multiple (ByVal arg1, ByVal arg2, ByRef ans)

End Sub
```
配列を参照する変数やString型の変数等、参照型の変数を引数として定義する時、
何も指定しないかByValを指定すると、参照がそのまま渡される.
ByRefを指定すると参照型の変数の参照が渡される.

- プロシージャの多重定義



---
## Part3 本格的なプログラミングにチャレンジする
