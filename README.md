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
