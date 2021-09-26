# 基礎 Visual Studio Code まとめ


<!-- vscode-markdown-toc -->
* 1. [Part1 はじめての VisualBasic プログラミング](#Part1VisualBasic)
* 2. [Part2 Visual Basicの基礎を身に付ける](#Part2VisualBasic)
	* 2.1. [<u>Chapter3 数値や文字列を扱う</u>](#uChapter3u)
	* 2.2. [<u>Chapter4 条件によって処理を変える</u>](#uChapter4u)
	* 2.3. [<u>Chapter5 処理を繰り返す</u>](#uChapter5u)
	* 2.4. [<u>Chapter6 配列を利用する</u>](#uChapter6u)
	* 2.5. [<u>Chapter7 プロシージャを使ってコードをまとめる</u>](#uChapter7u)
	* 2.6. [<u>Chapter8 クラスを利用する</u>](#uChapter8u)
* 3. [Part3 本格的なプログラミングにチャレンジする](#Part3)


---
##  1. <a name='Part1VisualBasic'></a>Part1 はじめての VisualBasic プログラミング

略

---
##  2. <a name='Part2VisualBasic'></a>Part2 Visual Basicの基礎を身に付ける
###  2.1. <a name='uChapter3u'></a><u>Chapter3 数値や文字列を扱う</u>
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


###  2.2. <a name='uChapter4u'></a><u>Chapter4 条件によって処理を変える</u>

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

###  2.3. <a name='uChapter5u'></a><u>Chapter5 処理を繰り返す</u>

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

###  2.4. <a name='uChapter6u'></a><u>Chapter6 配列を利用する</u>

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

###  2.5. <a name='uChapter7u'></a><u>Chapter7 プロシージャを使ってコードをまとめる</u>

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
'arg1,arg2は値がコピーされ、呼び出し側の引数の値を変える事はないが、
'ansを変更することで、呼び出し元のansも値を変更できる
Sub Multiple (ByVal arg1, ByVal arg2, ByRef ans)

End Sub
```
配列を参照する変数やString型の変数等、参照型の変数を引数として定義する時、
何も指定しないかByValを指定すると、参照がそのまま渡される.
ByRefを指定すると参照型の変数の参照が渡される.

- プロシージャの多重定義

引数の数や引数の型が異なれば、同名の関数を定義することができる.
```
Sub hoge(huga As String)
End Sub

Sub hoge(huga As Integer)
End Sub

Sub hoge(huga As Integer, piyo As Integer)
End Sub

Dim str As String
Dim num1, num2 As Integer

'1つめがコール
hoege(str)
'2つめがコール
hoege(num1)
'3つめがコール
hoege(num1, num2)
```

- 省略可能な引数を定義する

Optionalというキーワードを用いて省略可能な引数を指定する事ができる.
ただし、Optionalを付けて定義した引数には規定値を指定しておく必要があり、
Optionalを指定した引数移行の引数は全て省略可能である必要がある.
```
'piyoは省略可能で、規定値は0.
Sub hoge(hoge As Integer, Optional piyo As Integer = 0)
End Sub
```

- 名前付き引数を指定して呼び出す

省略できる引数がたくさんある場合、「,」で見づらくなることがあるが、
指定したい引数のみを「引数名:=値」の形式で書く事ができる

```
Sub hoge(huga As Integer, Optional piyo1 As Integer, piyo2 As Integer, 
        piyo3 AS Integer)
End Sub

'指定したいOptionalの引数のみを指定して呼び出し
hoge (huga := 3, piyo2 :=5)
```

- タプルを利用して複数の値を返す

タプルという機能を利用するFunctionプロシージャの戻り値を複数にする事ができる.
タプルとは複数の値をひとまとめにしたもので以下の様に使用する.
```
Dim member As (Integer, String) 'タプルの宣言
member = (3, "sato")            'タプルへの代入
member.item1 = 5                '要素ごとの代入
member.item2 = "tanaka"         '要素ごとの代入

Dim member As (number As Integer, name As String) '要素に名前を付けてタプルを宣言
member.number = 8
member.name ="suzuki"
```
上のタプルを用いてFunctionプロシージャの返し値を複数にする.

```
'返し値をタプルで宣言する
Function Hoge(huga As Integer) As (piyo1 As Integer, piyo2 As Integer)
  
  '処理

  Return (piyo1,piyo2)
End Function
```

<br>

###  2.6. <a name='uChapter8u'></a><u>Chapter8 クラスを利用する</u>

- オブジェクトをクラスから作成する

オブジェクトの作成にはNewキーワードを用いる.
Newキーワードを使ってオブジェクトを作成すると、作成されたオブジェクトの参照が返される.その参照を変数に代入する事で、オブジェクトを利用できるになる.
```
'Dim r As Random  -> オブジェクトを参照する変数の宣言.
'r = New Random() -> 新しいオブジェクトを作成して、参照を変数に代入.
'下では、上記二つの処理を宣言時に実行している. 
Dim r As Random = New Random ()
```

- オブジェクトのメンバを利用する

オブジェクトのメンバ（プロパティ・メソッド）は、オブジェクトを参照する変数の後に
「.」をつけ書く.

```
'Randomクラスのオブジェクトを作成し、Nextメソッドを使って乱数を作成
Dim i As Integer
Dim r As Random = New Random()
i = r.Next(10)
```
ちなみにフォーム上に配置するコントロールも様々なクラスのオブジェクトであり、
コード上で作成可能.
```
Dim btnOption As Button = New Button() 'Buttonクラスのオブジェクトを作成し、参照

btnOption.Text = "オプション(&O)"      'Textプロパティを設定する
btnOption.Parent = Me                  '親コントロールをMe(現在のフォーム)とする
```

- クラスの定義

クラスは以下のように定義する.
```
'必要に応じてPublicまたはPrivateを定義する
Public Class Person
End Class
```

- プロパティの定義

プロパティの実際の値を入れておくための変数はPrivateで宣言し、Get/SetプロシージャをPublicで定義する事で、外部より値を取得/変更する際に、不正な操作かをチェックする事ができる.
(下記のコードではその様な操作は特に記述していない.)
また、PropertyプロシージャにReadOnly/WriteOnlyを指定する事で読み出し/書き込み専用のプロパティとする事ができる.

```
Public Class Person
  'プロパティの値を記憶しておく為の変数の宣言
  Private mHeight As Double
  Private mHWeight as Double
  Private mAddress As String
  Private mLastUpdate As String

  'Propertyプロシージャを定義する
  Public property Height() As Double
    
    'プロパティの値を返す
    Get
      Return mHeight
    End Get

    'プロパティの値を設定する
    Set(Value As Double)
      mHeight = value
    End Set

  End property

  '読み出し専用のプロシージャ
  Public ReadOnly Property Address() As String
    
    'プロパティの値を返す
    Get
      Return mAddress
    End Get

    'Setステートメントは記述不可

  End property
  
  '書き込み専用のプロシージャ
  Public WriteOnly property UpdateDate() As String

    'Getステートメントは記述不可

    'プロパティの値を設定する
    Set(Date As String)
      mLastUpdate = Date
    End Set

End Class
```

- メソッドを定義する

メソッドはクラス内のSubプロシージャ、Functionプロシージャを指す.
```
Public Class Person
  'プロパティの値を記憶しておく為の変数の宣言
  Private mHeight As Double
  Private mHWeight as Double

  '・・・・・・・・・・・・・
  '・・・・・・・・・・・・・

    'プロパティの値を設定する
    Set(Date As String)
      mLastUpdate = Date
    End Set
  
  'メソッドを定義する
  Public Function GetBmi() As  Double
    Return mWeight / (mHeight / 100) ^ 2
  End Function

End Class
```
因みに上記の例では値を単純に返しているだけなので、ReadOnlyのプロパティとして
定義しても可.

- 定義したクラスの利用

上で記述したPersonクラスのオブジェクト生成.
```
Sub ShowBmi()
  Dim aPerson As Person = New Person()

  aPerson.Height = 170
  aPerson.Weight = 64
  MessageBox.show(aPerson.GetBmi().ToString("F2"))

End Sub
```

- 共有メンバーの定義

共有メソッドとするにはSharedというキーワードをメソッドの定義の最初に指定する.
```
Public Class Person
  ・・・
  ・・・
  Public Shared Function GetBmi(Height As Double, Weight As Double) As Double
    Return Weight / (Height / 100) ^ 2
  End Function
```
上記のように定地する事でオブジェクトを作成することなく、メソッドを使用する事ができる.
因みに、共有されていないメンバーをインスタンスメンバーと呼ぶ.

- クラスを継承して派生クラスを定義

作成済のPersonクラスを継承してMaleクラスを定義する.
Maleクラスのオブジェクトをさくせいする事で基本クラスのPersonのプロパティ、メソッドが使用できる.

```
Public Class Male
  Inherits Person
End Class
```
ただし、基本クラスでPrivateで宣言された変数やプロパティ、メソッドは派生クラスで使用できない.
また、基本クラスと派生クラスのみから利用できるメンバを定義したい場合は、基本クラスでの宣言時にProtectedを指定する.
```
Public Class Person
  Protected mHeight As Double
  ・・・
  ・・・
```

- メソッドのオーバーライド

子クラスでは基本クラスで定義済のメソッドの処理を変更したい場合、メソッドの定義しなおしが可能.
これをオーバーライドという.
オーバーライドしたい場合同じ名前のメソッドを派生クラスに書き、メソッド定義にOverridesというキーワードを付けておく.
ただし、基本クラスに書かれている同じな目のメソッドにはOverridableというキーワードを付けておく必要がある.
```
Public Class Child
  Inherits Person
  Public Overrides Function GerBmi() As Double
    ・・・
    ・・・
```

- インターフェースを利用する

複数のクラスで共通に使用されるような機能を実装する場合はインターフェイスを定義する.
インターフェイスを定義するにはInterfaceキーワードを使う.
インターフェイスでは、メソッドの定義のみを書き、メソッドの中身は書かず、具体的な処理はインターフェースを利用するクラスの方で書く.
```
Public Interface ISing
  Sub Sing(SongName As String)
End Interface
```
インターフェースを利用するにはクラス定義の後に「Implements インターフェース名」と書き、インターフェイスで定義されているメソッドの内容を書く.
メソッド名の後にもImplementsキーワードと「インターフェイス名.メソッド名」を書く必要がある.
```
Public Class Male
  Inherits Person
  Implements ISing
  ・・・
  ・・・
  Public Sub Sing(value As String) Implements ISing.Sing
    ・・・
  End Sub
```

---
##  3. <a name='Part3'></a>Part3 本格的なプログラミングにチャレンジする
