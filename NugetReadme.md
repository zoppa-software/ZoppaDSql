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

## 更新について
* 1.1.1 大量ログ出力に対応するためログ出力機能を修正
* 1.1.0 `Select`クエリの結果を`Object`配列リストで返す`ExecuteArrays`メソッドを追加  
        一列の結果を指定の型で取得する`ExecuteDatas`を追加
* 1.0.9 PrimaryKeyListにSearchValueを追加して、インスタンスの関連付け処理コードを読みやすく修正  
        Microsoft Accessのパラメータが位置指定のため、`SetOrderName`でパラメータ位置を設定できるようにしました
* 1.0.8 実行結果をDataTable(`ZoppaDSqlManager.ExecuteTable`)、DynamicObject(`ZoppaDSqlManager.ExecuteObject`)で取得するメソッドを追加、および、Mapper機能でDBNullをnullに変更する機能を追加
* 1.0.7 マッパー機能について、Oracle ODP.NET, Managed Driver の仕様から期待通りの動作がされない恐れがあるため `ZoppaDSqlSetting.SetParameterChecker` を追加
