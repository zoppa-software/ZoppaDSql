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

  
* **trim文**  
**注意！、trim処理は仕様が複雑なので期待している結果が得られないかもしれません。**


### SQL文を動的にコンパイルする
### SQLクエリを実行し、簡単なマッパー機能を使用してインスタンスを生成する
### パラメータにCSVファイルを与えてSQLクエリを実行します
### 付属機能
### ログファイル出力機能を有効にします
#### 簡単な式を評価し、結果を得ることができます
#### カンマ区切りで文字列を分割できます
#### CSVファイルを読み込みます
#### CSVファイルを読み込み、簡単なマッパー機能を使用してインスタンスを生成します

## インストール
ソースをビルドして `ZoppaDSql.dll` ファイルを生成して参照してください。  
Nugetにライブラリを公開しています。[ZoppaDSql](https://www.nuget.org/packages/ZoppaDSql/#readme-body-tab)を参照してください。

## 注意点

## 作成情報
* 造田　崇（zoppa software）
* ミウラ第3システムカンパニー 
* takashi.zouta@kkmiuta.jp

## ライセンス
[apache 2.0](https://www.apache.org/licenses/LICENSE-2.0.html)