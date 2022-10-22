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
ライブラリは .NET Standard 2.0 で記述しています。そのため、.net framework 4.6.1以降、.net core 2.0以降で使用できます。  
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
#### 通常のtrimの使い方
`{trim}`～`{end trim}`で囲まれた範囲は前後の空白をトリムします。  
``` vb
Dim ans1 = "{trim}   a = #{12 * 13}   {end trim}".Compile()
Assert.Equal(ans1, "a = 156")
```
それに加えて、`{trim `文字列`}`と指定すると、末尾が指定した文字列ならば削除します。  
``` vb
Dim ans2 = "{trim trush}   a = #{'11' + '29'}trush{end trim}".Compile()
Assert.Equal(ans2, "a = '1129'")
```
#### foreach文と組み合わせたtrimの使い方
`{trim}{foreach}～{end for}{end trim}`と記述している場合、末尾の`,`は自動的にトリムされます。  
以下が例です。'なにぬねの'の末尾に表示されるはずの`,`はトリムされており出力されません。  
``` vb
Dim strs As New List(Of String)()
strs.Add("あいうえお")
strs.Add("かきくけこ")
strs.Add("さしすせそ")
strs.Add("たちつてと")
strs.Add("なにぬねの")

Dim ans3 = "{trim}
{foreach str in strs}
    #{str},
{end for}
{end trim}".Compile(New With {.strs = strs})
Assert.Equal(ans3, "'あいうえお',
    'かきくけこ',
    'さしすせそ',
    'たちつてと',
    'なにぬねの'")
```
#### where句と組み合わせたtrimの使い方
`{trim}where {if}～{end if}{end trim}`と記述している場合、if文が空になったとき`where`をトリムします。  
また、`and`、`or`の関係演算子を構文解析して不要な場合、削除します。  
``` vb
Dim query3 = "" &
"select * from tb1
{trim}
where
    ({if af}a = 1{end if} or {if bf}b = 2{end if}) and ({if cf}c = 3{end if} or {if df}d = 4{end if})
{end trim}
"
' こちらは where句から全てトリムされます
Dim ans6 = query3.Compile(New With {.af = False, .bf = False, .cf = False, .df = False})
Assert.Equal(ans6.Trim(), "select * from tb1")

' b = 2、c = 3 が非出力なので or がトリムされます
Dim ans7 = query3.Compile(New With {.af = True, .bf = False, .cf = False, .df = True})
Assert.Equal(ans7.Trim(),
"select * from tb1
where
    (a = 1 ) and ( d = 4)")
```
  
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
以下はトランザクションの例です。  
``` vb
Dim zodiacs = New Zodiac() {
    New Zodiac("Aries", "牡羊座", New Date(2022, 3, 21), New Date(2022, 4, 19)),
    New Zodiac("Taurus", "牡牛座", New Date(2022, 4, 20), New Date(2022, 5, 20)),
    New Zodiac("Gemini", "双子座", New Date(2022, 5, 21), New Date(2022, 6, 21)),
    New Zodiac("Cancer", "蟹座", New Date(2022, 6, 22), New Date(2022, 7, 22)),
    New Zodiac("Leo", "獅子座", New Date(2022, 7, 23), New Date(2022, 8, 22)),
    New Zodiac("Virgo", "乙女座", New Date(2022, 8, 23), New Date(2022, 9, 22)),
    New Zodiac("Libra", "天秤座", New Date(2022, 9, 23), New Date(2022, 10, 23)),
    New Zodiac("Scorpio", "蠍座", New Date(2022, 10, 24), New Date(2022, 11, 22)),
    New Zodiac("Sagittarius", "射手座", New Date(2022, 11, 23), New Date(2022, 12, 21)),
    New Zodiac("Capricom", "山羊座", New Date(2022, 12, 22), New Date(2023, 1, 19)),
    New Zodiac("Aquuarius", "水瓶座", New Date(2023, 1, 20), New Date(2023, 2, 18)),
    New Zodiac("Pisces", "魚座", New Date(2023, 2, 19), New Date(2023, 3, 20))
}

Dim tran = Me.mSQLite.BeginTransaction()
Try
    Me.mSQLite.SetTransaction(tran).ExecuteQuery(
        "INSERT INTO Zodiac (name, jp_name, from_date, to_date) 
        VALUES (@Name, @JpName, @FromDate, @ToDate)", Nothing, zodiacs)
    tran.Commit()
Catch ex As Exception
    tran.Rollback()
End Try
```
トランザクションは`IDbConnection`から適切に取得してください。  
`SetTransaction`という拡張メソッドを用意しているのでトランザクションを与えます、その後はコミット、ロールバックを実行してください。  
拡張メソッドは以下のものがあります。  
| メソッド | 内容 |  
| ---- | ---- | 
| SetTransaction | トランザクションを設定します。 | 
| SetTimeoutSecond | SQLタイムアウトを設定します（秒数）、デフォルト値は30秒です。</br>デフォルト値は `ZoppaDSqlSetting.DefaultTimeoutSecond` で変更してください。 | 
| SetParameterPrepix | SQLパラメータの接頭辞を設定します、デフォルトは `@` です。</br>デフォルト値は `ZoppaDSqlSetting.DefaultParameterPrefix` で変更してください。 | 
| SetCommandType | SQLのコマンドタイプを設定します、デフォルト値は `CommandType.Text` です。 |
| SetParameterChecker | SQLパラメータチェック式を設定します。デフォルト値は `Null` です。</br>デフォルトs式は `ZoppaDSqlSetting.DefaultSqlParameterCheck` で変更してください。 |
  
### インスタンス生成をカスタマイズします
検索結果が 多対1 など一つのインスタンスで表現できない場合、インスタンスの生成をカスタマイズする必要があります。  
ZoppaDSqlではインスタンスを生成する式を引数で与えることで対応します。  
以下の例では、`Person`テーブルと`Zodiac`テーブルが多対1の関係で、そのまま`Person`クラスと`Zodiac`クラスに展開します。SQLの実行結果 1レコードで 1つの`Person`クラスを生成し、リレーションキーで`Zodiac`クラスをコレクションに保持して、関連を表現する`Persons`プロパティに追加します。
``` vb
Dim ansZodiacs As New PrimaryKeyList(Of Zodiac)()
Dim ansPersons = Me.mSQLite.ExecuteRecords(Of Person)(
    "select " &
    "  Person.Name, Person.birth_day, Zodiac.name, Zodiac.jp_name, Zodiac.from_date, Zodiac.to_date " &
    "from Person " &
    "left outer join Zodiac on " &
    "  Person.zodiac = Zodiac.name",
    Function(prm As Object()) As Person
        ' 一つのPersonを生成
        Dim pson = New Person(prm(0).ToString(), prm(2).ToString(), CDate(prm(1)))

        ' リレーションキーでZodiacを保持し、Personsプロパティに関連を追加
        Dim zdic As Zodiac = Nothing
        Dim zdicKey = prm(2).ToString()
        If Not ansZodiacs.TrySearchValue(zdic, zdicKey) Then
            zdic = New Zodiac(zdicKey, prm(3).ToString(), CDate(prm(4)), CDate(prm(5)))
            ansZodiacs.Regist(zdic, zdicKey)
        End If
        zdic.Persons.Add(pson)

        Return pson
    End Function
)
```
  
### パラメータにCSVファイルを与えてSQLクエリを実行します
単体テストなど大量データのインサート用にCSVファイルから直接パラメータ値を取得する仕組みを用意しました。  
以下の例を参照してください。  
``` vb
Using sr As New CsvReaderStream("Sample.csv")
    Using tran = sqlite.BeginTransaction()
        ' 実行するSQL文
        Dim query = "insert into SampleDB (indexno, name) values (@indexno, @name)"

        ' CSVファイルの各列の型を保持するbuilderを生成
        Dim builder As New CsvParameterBuilder()
        builder.Add("indexno", CsvType.CsvInteger)
        builder.Add("name", CsvType.CsvString)

        ' CSVストリームとbuilderをパラメータに与えて実行
        ' ※ 非同期実行するとファイルが先にCloseされるため同期実行します
        sqlite.SetTransaction(tran).ExecuteQuery(query, sr, builder)

        tran.Commit()
    End Using
End Using
```
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
``` vb
' 数式
Dim ans1 = "(28 - 3) / (2 + 3)".Executes().Contents
Assert.Equal(ans1, 5)

' 比較式
Dim ans2 = "0.1 * 5 <= 0.4".Executes().Contents
Assert.Equal(ans2, False)
```
#### カンマ区切りで文字列を分割できます
`"`のエスケープを考慮して文字列を区切ります。分割した文字列は`CsvItem`構造体に格納されます。`Text`プロパティではエスケープは解除されませんが、`UnEscape`メソッドではエスケープを解除します。  
``` vb
Dim csv = CsvSpliter.CreateSpliter("あ,い,う,え,""お,を""").Split()
Assert.Equal(csv(0).UnEscape(), "あ")
Assert.Equal(csv(1).UnEscape(), "い")
Assert.Equal(csv(2).UnEscape(), "う")
Assert.Equal(csv(3).UnEscape(), "え")
Assert.Equal(csv(4).UnEscape(), "お,を")
Assert.Equal(csv(4).Text, """お,を""")
```
#### CSVファイルを読み込みます
* ストリーム、イテレータを使用して読み込みます
基本的な使い方はストリームを用意し、イテレータを使用して読み込みます。
``` vb
Using sr As New CsvReaderStream("CsvFiles\Sample3.csv", Encoding.GetEncoding("shift_jis"))
    For Each pointer In sr
        Console.Out.WriteLine($"{pointer.Items(0).UnEscape()}, {pointer.Items(1).UnEscape()}, …")
    Next
End Using
```
ストリームクラスは`IEnumerable`インターフェイスを実装しているため一行の情報を`CsvReaderStream.Pointer`構造体で取得できます。  
読み込んだ行番号(`Row`プロパティ)と各項目の情報(`Items`プロパティ、`CsvItem`構造体のリスト)を持っているため参照します。  

* 条件を指定して読み込みます
`WhereCsv`メソッドを使用すると二つの式を与えてCSVの情報をインスタンスに変換できます。  
    * 一つ目の式は対象の行を決定するための式です  
    * 二つ目の式はインスタンスを生成して返します  

以下の例では、一つ目の式にヘッダ行を除くため2行目以降（`row >= 1`）を指定し、二つ目の式は`Sample1Csv`のインスタンスを生成します。  
``` vb
Dim ans As New List(Of Sample1Csv)()
Using sr As New CsvReaderStream("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
ans = sr.WhereCsv(Of Sample1Csv)(
    Function(row, item) row >= 1,
    Function(row, item) New Sample1Csv(item(0).UnEscape(), item(1).UnEscape(), item(2).UnEscape())
).ToList()
End Using
```

``` vb
Class Sample1Csv
    Public ReadOnly Property Item1 As String
    Public ReadOnly Property Item2 As String
    Public ReadOnly Property Item3 As String

    Public Sub New(s1 As String, s2 As String, s3 As String)
        Me.Item1 = s1
        Me.Item2 = s2
        Me.Item3 = s3
    End Sub
End Class
```
`ICsvType`インターフェイスを実装したインスタンスを与えることで、インスタンスを生成するコンストラクタを指定できます。  
以下の例では`String`型の引数を3つ持つコンストラクタを利用してインスタンスを生成します。
``` vb
Dim ans As New List(Of Sample1Csv)()
Using sr As New CsvReaderStream("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
ans = sr.WhereCsv(Of Sample1Csv)(
    Function(row, item) row >= 1,
    CsvType.CsvString, CsvType.CsvString, CsvType.CsvString
).ToList()
End Using
```

## 注意
Oracle DB を対象にマッパー機能でSQLパラメーターを与えるとき、ODP.NET, Managed Driver の OracleParameter の以下の仕様から日本語の検索が正しくできないと思われます。  
> OracleParameter.DbType に DbType.String を設定した場合、OracleDbType には OracleDbType.Varchar2 が設定されます  
> *OracleDbType.NVarchar2 ではありません*  

上記の仕様が問題になる可能性があるので、`IDbDataParameter` を使用前にチェックする式を設定する機能を追加しました。  
以下の例をご覧ください。  
``` vb
Using ora As New OracleConnection()
    ora.ConnectionString = "接続文字列"
    ora.Open()

    Dim tbl = ora.
        SetParameterPrepix(PrefixType.Colon).
        SetParameterChecker(
            Sub(chk)
                Dim prm = TryCast(chk, Oracle.ManagedDataAccess.Client.OracleParameter)
                If prm?.OracleDbType = OracleDbType.Varchar2 Then
                    prm.OracleDbType = OracleDbType.NVarchar2
                End If
            End Sub).
        ExecuteRecords(Of RFLVGROUP)(
            "select * from GROUP where SYAIN_NO = :SyNo ",
            New With {.SyNo = CType("105055", DbString)}
        )
End Using
```
拡張メソッド `SetParameterChecker` では生成した `IDbDataParameter` を順次引き渡すので、`OracleDbType` を式内で変更しています。
全てのSQLに適用したい場合はデフォルトの式を変更します。  
``` vb
ZoppaDSqlSetting.DefaultSqlParameterCheck =
    Sub(chk)
        Dim prm = TryCast(chk, Oracle.ManagedDataAccess.Client.OracleParameter)
        If prm?.OracleDbType = OracleDbType.Varchar2 Then
            prm.OracleDbType = OracleDbType.NVarchar2
        End If
    End Sub
```
  
## インストール
ソースをビルドして `ZoppaDSql.dll` ファイルを生成して参照してください。  
Nugetにライブラリを公開しています。[ZoppaDSql](https://www.nuget.org/packages/ZoppaDSql/)を参照してください。

## 作成情報
* 造田　崇（zoppa software）
* ミウラ第3システムカンパニー 
* takashi.zouta@kkmiuta.jp

## ライセンス
[apache 2.0](https://www.apache.org/licenses/LICENSE-2.0.html)
