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

以上、簡単な説明となります。**ライブラリの詳細は[Githubのページ](https://github.com/zoppa-software/ZoppaDSql)を参照してください。**