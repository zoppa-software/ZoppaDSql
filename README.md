# ZoppaDSql 
SQL文の内部に制御文を埋め込む方式で動的SQLを行うライブラリです。 

## 説明
動的SQLはコードからSQL文を自動生成する方法が一般的と思っています。  
しかし、コードから自動生成されるSQL文は細かな調整が難しいというイメージがあり、また、データベース操作が埋め込まれるので処理の検索がちょっと難しいと思います。  
 
そのため、MybatisのXMLの動的SQLを参考にライブラリを作成しました。  
以下の例をご覧ください。 
``` vb
Dim answer = "" &
"select * from employees 
{trim}
where
    {if empNo}emp_no < 20000{end if} and
    ({if first_name}first_name like 'A%'{end if} or 
     {if gender}gender = 'F'{end if})
{end trim}
limit 10".Compile(New With {.empNo = False, .first_name = False, .gender = False})
```

SQLを記述した文字列内に `{}` で囲まれた文が制御文になります。  
拡張メソッド `Compile` にパラメータとなるクラスを引き渡して実行するとクラスのプロパティを読み込み、SQLを動的に構築します。  
パラメータは全て `False` なので、実行結果は以下のようになります。 
``` vb
select * from employees 
limit 10
```

## 比較、または特徴
* 素に近いSQL文を動的に変更するため、コードから自動生成されるSQL文より直観的です  
* SQLが文字列であるため、プログラム言語による差がなくなります  
* select、insert、update、delete を文字列検索すれはデータベース処理を検索できます  

## 依存関係
ライブラリは .NET Standard 2.0 で記述しています。そのため、.net framework 4.5以降、.net core 4以降で使用できます。  
その他のライブラリへの依存関係はありません。

## 使い方
### SQL文に置き換え式、制御式を埋め込む
SQL文を部分的に置き換える（置き換え式）、また、部分的に除外するや繰り返すなど（制御式）を埋め込みます。  
埋め込みは `#{` *参照するプロパティ* `}`、`{` *制御式* `}` の形式で記述します。  
  
#### 埋め込み式  
`#{` *参照するプロパティ* `}` を使用すると、`Compile`で引き渡したオブジェクトのプロパティを参照して置き換えます。  
以下は文字列プロパティを参照しています。 `'`で囲まれて出力していることに注目してください。   
``` vb
Dim ansstr1 = "where parson_name = #{name}".Compile(New With {.name = "zouta takshi"})
Assert.Equal(ansstr1, "where parson_name = 'zouta takshi'")
```
次に数値プロパティを参照します。  
``` vb
Dim ansnum1 = "where age >= #{age}".Compile(New With {.age = 12})
Assert.Equal(ansnum1, "where age >= 12")
```
次にnullを参照します。  
``` vb
Dim ansnull = "set col1 = #{null}".Compile(New With {.null = Nothing})
Assert.Equal(ansnull, "set col1 = null")
```
  
埋め込みたい文字列にはデーブル名など `'` で囲みたくない場面があります。この場合、`!{}` (または `${}`)を使用します。  
``` vb
Dim ansstr2 = "select * from !{table}".Compile(New With {.table = "sample_table"})
Assert.Equal(ansstr2, "select * from sample_table")
```

#### 制御式  
SQL文を部分的に除外、または繰り返すなど制御を行います。  
  
* **if文**  
条件が真であるならば、その部分を出力します。  
`{if 条件式}`、`{else if 条件式}`、`{else}`、`{end if}`で囲まれた部分を判定します。  
以下の例を参照してください。  
``` vb
Dim query = "" &
"select * from table1
where
  {if num = 1}col1 = #{num}
  {else if num = 2}col2 = #{num}
  {else}col3 = #{num}
  {end if}"

' num = 1 ならば {if num = 1}の部分を出力
Dim ans1 = query.Compile(New With {.num = 1})
Assert.Equal(ans1,
"select * from table1
where
  col1 = 1")

' num = 2 ならば {else if num = 2}の部分を出力
Dim ans2 = query.Compile(New With {.num = 2})
Assert.Equal(ans2,
"select * from table1
where
  col2 = 2")

' num = 5 ならば {else}の部分を出力
Dim ans3 = query.Compile(New With {.num = 5})
Assert.Equal(ans3,
"select * from table1
where
  col3 = 5")
```
  
* **foreach文**  
パラメータの要素分、出力を繰り返します。  
`{foreach 一時変数 in パラメータ(配列など)}`、`{end for}`で囲まれた範囲を繰り返します。一時変数に要素が格納されるので foreachの範囲内で置き換え式を使用して出力してください。    
``` vb
"SELECT
    *
FROM
    customers 
WHERE
    FirstName in ({trim}{foreach nm in names}#{nm}, {end for}{end trim})
".Compile(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
```
上記の例では foreachで囲まれた範囲の文、`#{nm}, `がパラメータ`names`の要素数分("Helena", "Dan", "Aaron")を繰り返して出力します。  
(お気づきのとおり、最後の要素では`,`が不要に出力されます、この`,`を取り除くのが **trim**文です)  
出力は以下のとおりです。  
``` vb
SELECT
    *
FROM
    customers 
WHERE
    FirstName in ('Helena', 'Dan', 'Aaron')
```
ただし、このような場合、通常の置き換え式で配列など繰り返し要素を与えれば `,` で結合して置き換えるように実装しています。  
以下の例は上記と同じ結果を出力します。  
``` vb
SELECT
    *
FROM
    customers 
WHERE
    FirstName in (#{names})
".Compile(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
```
  
* **trim文**  
**注意！、trim処理は仕様が複雑なので期待している結果が得られないかもしれません。**


### SQLクエリを実行し、簡単なマッパー機能を使用してインスタンスを生成する
#### 基本的な使い方
動的に生成したSQL文をDapperやEntity Frameworkで利用することができます。  
また、簡単に利用するために [IDbConnection](https://learn.microsoft.com/ja-jp/dotnet/api/system.data.idbconnection)の拡張メソッド、および、シンプルなマッピング処理を用意しています。  
  
以下の例は、パラメータの`seachId`が`NULL`以外ならば、`ArtistId`が`seachId`と等しいことという動的SQLを `query`変数に格納しました。
``` vb
Dim query = "" &
"select
  albumid, title, name
from
  albums
inner join artists on
  albums.ArtistId = artists.ArtistId
where
  {if seachId <> NULL }albums.ArtistId = @seachId{end if}"
```
次に、SQLiteの`IDbConnection`の実装である`SQLiteConnection`をOpenした後、`ExecuteRecordsSync`拡張メソッドを実行するとSQLの実行結果が`AlbumInfo`クラスのリストで取得できます。  
``` vb
Dim answer As List(Of AlbumInfo)
Using sqlite As New SQLiteConnection("Data Source=chinook.db")
    sqlite.Open()

    answer = Await sqlite.ExecuteRecordsSync(Of AlbumInfo)(query, New With {.seachId = 11})
    ' answerにSQLの実行結果が格納されます
End Using
```
`AlbumInfo`クラスの実装は以下のとおりです。  
マッピングは一般的にはプロパティ、フィールドをマッピングしますが、ZoppaDSqlはSQLの実行結果の各カラムの型と一致する**コンストラクタ**を検索してインスタンスを生成します。  
``` vb
Public Class AlbumInfo
    Public ReadOnly Property AlbumId As Integer
    Public ReadOnly Property AlbumTitle As String
    Public ReadOnly Property ArtistName As String

    Public Sub New(id As Long, title As String, nm As String)
        Me.AlbumId = id
        Me.AlbumTitle = title
        Me.ArtistName = nm
    End Sub
End Class
```
#### SQL実行設定
トランザクション、SQLタイムアウトの設定も拡張メソッドで行います。  
### パラメータにCSVファイルを与えてSQLクエリを実行します
### ログファイル出力機能を有効にします
動的SQLを使用したとき、生成されたSQL文を確認したい時があります。そのためにZoppaDSqlではログファイル出力機能があります。  
デフォルトのログファイル出力機能を有効にするには以下のコードを実行します。引数にログファイルパスを指定するとログファイルを変更することができます。初期値はカレントディレクトリの`zoppa_dsql.txt`ファイルに出力されます。  
``` vb
ZoppaDSqlManager.UseDefaultLogger()
```  
ログの出力は別スレッドで行っているため、アプリケーション終了前などに出力完了を待機してください。
``` vb
ZoppaDSqlManager.LogWaitFinish()
```  
別のログ出力機能を使用する場合は、`ZoppaDSql.ILogWriter`インターフェイスを実装したクラスを定義して、`ZoppaDSql.SetCustomLogger`で設定してください。

### 付属機能
#### 簡単な式を評価し、結果を得ることができます
#### カンマ区切りで文字列を分割できます
#### CSVファイルを読み込みます
#### CSVファイルを読み込み、簡単なマッパー機能を使用してインスタンスを生成します

## インストール
ソースをビルドして `ZoppaDSql.dll` ファイルを生成して参照してください。  
Nugetにライブラリを公開しています。[ZoppaDSql](https://www.nuget.org/packages/ZoppaDSql/#readme-body-tab)を参照してください。

## 作成情報
* 造田　崇（zoppa software）
* ミウラ第3システムカンパニー 
* takashi.zouta@kkmiuta.jp

## ライセンス
[apache 2.0](https://www.apache.org/licenses/LICENSE-2.0.html)
