Option Strict On
Option Explicit On

Imports System.Data
Imports System.Dynamic
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Text
Imports ZoppaDSql.Analysis
Imports ZoppaDSql.Analysis.Tokens
Imports ZoppaDSql.Csv

''' <summary>DSql APIモジュール。</summary>
Public Module ZoppaDSqlManager

    ''' <summary>ログ出力機能。</summary>
    Private mLogger As ILogWriter

    ''' <summary>パラメータをログ出力します。</summary>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub LoggingParameter(parameter As Object)
        If parameter IsNot Nothing Then
            LoggingDebug("Parameters")
            Dim props = parameter.GetType().GetProperties()
            For Each prop In props
                Dim v = prop.GetValue(parameter)
                LoggingDebug($"・{prop.Name}={GetValueString(v)} ({prop.PropertyType})")
            Next
        End If
    End Sub

    ''' <summary>オブジェクトの文字列表現を取得します。</summary>
    ''' <param name="value">オブジェクト。</param>
    ''' <returns>文字列表現。</returns>
    Private Function GetValueString(value As Object) As String
        If TypeOf value Is IEnumerable Then
            Dim buf As New StringBuilder()
            For Each v In CType(value, IEnumerable)
                If buf.Length > 0 Then buf.Append(", ")
                buf.Append(v)
            Next
            Return buf.ToString()
        Else
            Return If(value?.ToString(), "[null]")
        End If
    End Function

    ''' <summary>SQLパラメータ定義を設定します。</summary>
    ''' <param name="command">SQLコマンド。</param>
    ''' <param name="parameter">パラメータ。</param>
    ''' <param name="varFormat">変数フォーマット。</param>
    ''' <param name="prmChecker">SQLパラメータチェック。</param>
    ''' <param name="propNames">プロパティ名リスト。</param>
    ''' <returns>プロパティインフォ。</returns>
    Private Function SetSqlParameterDefine(command As IDbCommand,
                                           parameter As Object(),
                                           varFormat As String,
                                           prmChecker As Action(Of IDbDataParameter),
                                           propNames As String()) As PropertyInfo()
        Dim props = New List(Of PropertyInfo)()
        Dim params = parameter.Where(Function(v) v IsNot Nothing)
        If params.Any() Then
            LoggingDebug("params class define")

            ' プロパティインフォを取得
            props = params.First().GetType().GetProperties().ToList()
            If propNames?.Length > 0 Then
                Dim dic = props.ToDictionary(Of String, PropertyInfo)(Function(v) v.Name.ToLower(), Function(v) v)
                props.Clear()
                For Each nm In propNames
                    Dim p As PropertyInfo = Nothing
                    If dic.TryGetValue(nm.ToLower(), p) Then
                        props.Add(p)
                    End If
                Next
            End If

            Dim prms As New List(Of IDbDataParameter)()
            For Each prop In props
                ' SQLパラメータを作成
                Dim prm = command.CreateParameter()

                ' パラメータの名前、方向を設定
                prm.ParameterName = String.Format(varFormat, prop.Name)
                If GetDbType(prop.PropertyType, prm.DbType) Then
                    If prop.CanRead AndAlso prop.CanWrite Then
                        prm.Direction = ParameterDirection.Input
                    Else
                        prm.Direction = If(prop.CanWrite, ParameterDirection.Output, ParameterDirection.Input)
                    End If
                    LoggingDebug($"・Name = {prm.ParameterName} Direction = {[Enum].GetName(GetType(ParameterDirection), prm.Direction)}")

                    prms.Add(prm)
                End If
            Next

            command.Parameters.Clear()
            For Each prm In prms
                command.Parameters.Add(prm)
            Next

            ' パラメータのチェックをする
            If prmChecker IsNot Nothing Then
                For Each prm As IDbDataParameter In command.Parameters
                    prmChecker(prm)
                Next
            End If
        End If
        Return props.ToArray()
    End Function

    ''' <summary>SQLパラメータに値を設定します。</summary>
    ''' <param name="command">SQLコマンド。</param>
    ''' <param name="parameter">パラメータ。</param>
    ''' <param name="properties">プロパティインフォリスト。</param>
    ''' <param name="varFormat">変数フォーマット。</param>
    Private Sub SetParameter(command As IDbCommand, parameter As Object, properties As PropertyInfo(), varFormat As String)
        For Each prop In properties
            ' 名前と値を取得
            Dim propName = String.Format(varFormat, prop.Name)
            Dim propVal = If(prop.GetValue(parameter), DBNull.Value)

            ' 変数に設定
            If command.Parameters.Contains(propName) Then
                LoggingDebug($"・{propName}={propVal}")
                CType(command.Parameters(propName), IDbDataParameter).Value = propVal
            End If
        Next
    End Sub

    ''' <summary>SQLパラメータの書式を取得します。</summary>
    ''' <param name="varPrefix">SQLパラメータの接頭辞。</param>
    ''' <returns>SQLパラメータの書式。</returns>
    Private Function GetVariantFormat(varPrefix As PrefixType) As String
        Dim ans As String = ""
        Select Case varPrefix
            Case PrefixType.AtMark
                ans = "@"
            Case PrefixType.Colon
                ans = ":"
        End Select
        LoggingDebug($"Variant format : '{ans}'")

        Return ans & "{0}"
    End Function

    ''' <summary>動的SQLをコンパイルします。</summary>
    ''' <param name="sqlQuery">動的SQL。</param>
    ''' <param name="parameter">動的SQL、クエリパラメータ用の情報。</param>
    ''' <returns>コンパイル結果。</returns>
    <Extension()>
    Public Function Compile(sqlQuery As String, Optional parameter As Object = Nothing) As String
        Try
            LoggingDebug($"Compile SQL : {sqlQuery}")
            LoggingParameter(parameter)
            Dim ans = ParserAnalysis.Replase(sqlQuery, parameter)
            LoggingDebug($"Answer SQL : {ans}")
            Return ans
        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>引数の文字列を評価して値を取得します。</summary>
    ''' <param name="expression">評価する文字列。</param>
    ''' <param name="parameter">環境値。</param>
    ''' <returns>評価結果。</returns>
    <Extension()>
    Public Function Executes(expression As String, Optional parameter As Object = Nothing) As IToken
        Try
            LoggingDebug($"Executes Command : {expression}")
            LoggingParameter(parameter)
            Dim ans = ParserAnalysis.Executes(expression, parameter)
            LoggingDebug($"Answer Command : {ans}")
            Return ans
        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>引数で指定した型をDBのデータ型に変換します。</summary>
    ''' <param name="propType">データ型。</param>
    ''' <param name="dbType">DBのデータ型(戻り値)</param>
    ''' <returns>変換できたら真。</returns>
    Private Function GetDbType(propType As Type, ByRef dbType As DbType) As Boolean
        Dim res = True
        Select Case propType
            Case GetType(String)
                dbType = DbType.String

            Case GetType(DbString)
                dbType = DbType.String

            Case GetType(DBAnsiString)
                dbType = DbType.AnsiString

            Case GetType(Integer), GetType(Integer?)
                dbType = DbType.Int32

            Case GetType(Long), GetType(Long?)
                dbType = DbType.Int64

            Case GetType(Short), GetType(Short?)
                dbType = DbType.Int16

            Case GetType(Single), GetType(Single?)
                dbType = DbType.Single

            Case GetType(Double), GetType(Double?)
                dbType = DbType.Double

            Case GetType(Decimal), GetType(Decimal?)
                dbType = DbType.Decimal

            Case GetType(Date), GetType(Date?)
                dbType = DbType.Date

            Case GetType(TimeSpan), GetType(TimeSpan?)
                dbType = DbType.Time

            Case GetType(Object)
                dbType = DbType.Object

            Case GetType(Boolean), GetType(Boolean?)
                dbType = DbType.Boolean

            Case GetType(Byte), GetType(Byte?)
                dbType = DbType.Byte

            Case GetType(Byte())
                dbType = DbType.Binary

            Case GetType(Char())
                dbType = DbType.String

            Case Else
                res = False
        End Select
        Return res
    End Function

    ''' <summary>DbType.String型の変数を作成する。</summary>
    ''' <param name="str">文字列。</param>
    ''' <returns>DbType.String型</returns>
    <Extension()>
    Public Function DbStr(str As String) As DbString
        Return New DbString(str)
    End Function

    ''' <summary>DbType.AnsiString型の変数を作成する。</summary>
    ''' <param name="str">文字列。</param>
    ''' <returns>DbType.AnsiString型</returns>
    <Extension()>
    Public Function DbAnsi(str As String) As DBAnsiString
        Return New DBAnsiString(str)
    End Function

    ''' <summary>マッピングするコンストラクタを取得します。</summary>
    ''' <typeparam name="T">対象の型。</typeparam>
    ''' <param name="reader">データリーダー。</param>
    ''' <param name="allowNul">Null許容の列リスト。</param>
    ''' <returns>コンストラクタインフォ。</returns>
    Private Function CreateConstructorInfo(Of T)(reader As IDataReader, allowNul As List(Of Integer)) As ConstructorInfo
        ' 各列のデータ型を取得
        Dim columNames As New List(Of String)()
        Dim columTypes As New List(Of Type)()
        Dim columAllow As New List(Of Boolean)()
        allowNul.Clear()

        Dim tbl = reader.GetSchemaTable()
        Dim idx As Integer = 0
        For Each r As DataRow In tbl.Rows
            columNames.Add(r("ColumnName").ToString())

            columTypes.Add(CType(r("DataType"), Type))

            Dim allow = CBool(r("AllowDBNull"))
            columAllow.Add(allow)
            If allow Then
                allowNul.Add(idx)
            End If
            idx += 1
        Next

        ' コンストラクタを取得
        Dim constructor = GetType(T).GetConstructor(columTypes.ToArray())
        If constructor IsNot Nothing Then
            Return constructor
        Else
            Dim info As New StringBuilder()
            For i As Integer = 0 To columTypes.Count - 1
                info.AppendFormat(
                    "{0}:{1}{2}{3}",
                    columNames(i),
                    columTypes(i).Name,
                    If(columAllow(i) AndAlso columTypes(i).IsValueType, "?", ""),
                    If(i < columTypes.Count - 1, ",", "")
                )
            Next
            Throw New DSqlException($"取得データに一致するコンストラクタがありません:{info.ToString()}")
        End If
    End Function

    ''' <summary>DBNullを見つけたら nullに変更します。</summary>
    ''' <param name="fields">チェック対象の配列。</param>
    ''' <param name="allowNul">DBNullを許容している列リスト。</param>
    Private Sub ChangeDBNull(fields As Object(), allowNul As List(Of Integer))
        For Each i In allowNul
            If TypeOf fields(i) Is DBNull Then
                fields(i) = Nothing
            End If
        Next
    End Sub

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteRecords
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         dynamicParameter As Object,
                                         ParamArray sqlParameter As Object()) As List(Of T)
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim recoreds As New List(Of T)()
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat, setting.ParameterChecker, setting.PropertyNames)

                Dim constructor As ConstructorInfo = Nothing
                Dim allowNull As New List(Of Integer)()
                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    If prm IsNot Nothing Then
                        SetParameter(command, prm, props, varFormat)
                    End If

                    Using reader = command.ExecuteReader()
                        ' マッピングコンストラクタを設定
                        If constructor Is Nothing Then
                            constructor = CreateConstructorInfo(Of T)(reader, allowNull)
                        End If

                        ' 一行取得してインスタンスを生成
                        Dim fields = New Object(reader.FieldCount - 1) {}
                        Do While reader.Read()
                            If reader.GetValues(fields) >= reader.FieldCount Then
                                ChangeDBNull(fields, allowNull)
                                recoreds.Add(CType(constructor.Invoke(fields), T))
                            End If
                        Loop
                    End Using
                Next
            End Using
            Return recoreds

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(setting As Settings,
                                                   query As String,
                                                   dynamicParameter As Object,
                                                   ParamArray sqlParameter As Object()) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(setting, query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         dynamicParameter As Object,
                                         ParamArray sqlParameter As Object()) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, sqlParameter)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(connect As IDbConnection,
                                                   query As String,
                                                   dynamicParameter As Object,
                                                   ParamArray sqlParameter As Object()) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String) As List(Of T)
        Return ExecuteRecords(Of T)(setting, query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(setting As Settings,
                                                   query As String) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(setting, query, Nothing, New Object() {})
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(connect As IDbConnection,
                                                   query As String) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(New Settings(connect), query, Nothing, New Object() {})
            End Function
        )
    End Function

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteRecords(CSV)
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         dynamicParameter As Object,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder) As List(Of T)
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim recoreds As New List(Of T)()
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' パラメータの定義を設定
                SetSqlParameterDefine(command, csvParameter, varFormat)

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                Dim constructor As ConstructorInfo = Nothing
                Dim allowNull As New List(Of Integer)()
                For Each prm In csvStream
                    ' パラメータ変数に値を設定
                    Dim i As Integer = 0
                    For Each pair In csvParameter
                        ' 名前と値を取得
                        Dim propName = String.Format(varFormat, pair.Name)
                        Dim propVal = If(pair.CsvColumnType.Convert(prm.Items(i)), DBNull.Value)
                        i += 1

                        ' 変数に設定
                        LoggingDebug($"・{propName}={propVal}")
                        CType(command.Parameters(propName), IDbDataParameter).Value = propVal
                    Next

                    ' 一行取得してインスタンスを生成
                    Using reader = command.ExecuteReader()
                        ' マッピングコンストラクタを設定
                        If constructor Is Nothing Then
                            constructor = CreateConstructorInfo(Of T)(reader, allowNull)
                        End If

                        ' 一行取得してインスタンスを生成
                        Dim fields = New Object(reader.FieldCount - 1) {}
                        Do While reader.Read()
                            If reader.GetValues(fields) >= reader.FieldCount Then
                                ChangeDBNull(fields, allowNull)
                                recoreds.Add(CType(constructor.Invoke(fields), T))
                            End If
                        Loop
                    End Using
                Next
            End Using
            Return recoreds

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         dynamicParameter As Object,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, csvStream, csvParameter)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder) As List(Of T)
        Return ExecuteRecords(Of T)(setting, query, Nothing, csvStream, csvParameter)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, Nothing, csvStream, csvParameter)
    End Function

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteRecords(createrMethod)
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         dynamicParameter As Object,
                                         sqlParameter As Object(),
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim recoreds As New List(Of T)()
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat, setting.ParameterChecker, setting.PropertyNames)

                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    If prm IsNot Nothing Then
                        SetParameter(command, prm, props, varFormat)
                    End If

                    ' 一行取得してインスタンスを生成
                    Using reader = command.ExecuteReader()
                        Dim fields = New Object(reader.FieldCount - 1) {}
                        Do While reader.Read()
                            If reader.GetValues(fields) >= reader.FieldCount Then
                                recoreds.Add(createrMethod(fields))
                            End If
                        Loop
                    End Using
                Next
            End Using
            Return recoreds

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(setting As Settings,
                                                   query As String,
                                                   dynamicParameter As Object,
                                                   sqlParameter As Object(),
                                                   createrMethod As Func(Of Object(), T)) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(setting, query, dynamicParameter, sqlParameter, createrMethod)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         dynamicParameter As Object,
                                         sqlParameter As Object(),
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, sqlParameter, createrMethod)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(connect As IDbConnection,
                                                   query As String,
                                                   dynamicParameter As Object,
                                                   sqlParameter As Object(),
                                                   createrMethod As Func(Of Object(), T)) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, sqlParameter, createrMethod)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         dynamicParameter As Object,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(setting, query, dynamicParameter, New Object() {}, createrMethod)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(setting As Settings,
                                                   query As String,
                                                   dynamicParameter As Object,
                                                   createrMethod As Func(Of Object(), T)) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(setting, query, dynamicParameter, New Object() {}, createrMethod)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         dynamicParameter As Object,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, New Object() {}, createrMethod)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(connect As IDbConnection,
                                                   query As String,
                                                   dynamicParameter As Object,
                                                   createrMethod As Func(Of Object(), T)) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, New Object() {}, createrMethod)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(setting, query, Nothing, New Object() {}, createrMethod)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(setting As Settings,
                                                   query As String,
                                                   createrMethod As Func(Of Object(), T)) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(setting, query, Nothing, New Object() {}, createrMethod)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, Nothing, New Object() {}, createrMethod)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteRecordsSync(Of T)(connect As IDbConnection,
                                                   query As String,
                                                   createrMethod As Func(Of Object(), T)) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteRecords(Of T)(New Settings(connect), query, Nothing, New Object() {}, createrMethod)
            End Function
        )
    End Function

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteRecords(createrMethod, CSV)
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         dynamicParameter As Object,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim recoreds As New List(Of T)()
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' パラメータの定義を設定
                SetSqlParameterDefine(command, csvParameter, varFormat)

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                For Each prm In csvStream
                    ' パラメータ変数に値を設定
                    Dim i As Integer = 0
                    For Each pair In csvParameter
                        ' 名前と値を取得
                        Dim propName = String.Format(varFormat, pair.Name)
                        Dim propVal = If(pair.CsvColumnType.Convert(prm.Items(i)), DBNull.Value)
                        i += 1

                        ' 変数に設定
                        LoggingDebug($"・{propName}={propVal}")
                        CType(command.Parameters(propName), IDbDataParameter).Value = propVal
                    Next

                    ' 一行取得してインスタンスを生成
                    Using reader = command.ExecuteReader()
                        Dim fields = New Object(reader.FieldCount - 1) {}
                        Do While reader.Read()
                            If reader.GetValues(fields) >= reader.FieldCount Then
                                recoreds.Add(createrMethod(fields))
                            End If
                        Loop
                    End Using
                Next
            End Using
            Return recoreds

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         dynamicParameter As Object,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, csvStream, csvParameter, createrMethod)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(setting, query, Nothing, csvStream, csvParameter, createrMethod)
    End Function

    ''' <summary>SQLクエリを実行し、指定の型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <param name="createrMethod">インスタンス生成式。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, Nothing, csvStream, csvParameter, createrMethod)
    End Function

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteTable
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行し、データテーブルを取得します。</summary>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteTable(setting As Settings,
                                 query As String,
                                 dynamicParameter As Object,
                                 ParamArray sqlParameter As Object()) As DataTable
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim res As New DataTable(query)
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat, setting.ParameterChecker, setting.PropertyNames)

                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    If prm IsNot Nothing Then
                        SetParameter(command, prm, props, varFormat)
                    End If

                    Using reader = command.ExecuteReader()
                        Dim tbl = reader.GetSchemaTable()
                        For Each r As DataRow In tbl.Rows
                            Dim clm = New DataColumn(r("ColumnName").ToString(), CType(r("DataType"), Type))
                            res.Columns.Add(clm)
                        Next

                        ' 一行取得して行を追加
                        Do While reader.Read()
                            Dim r = res.NewRow()
                            Dim fields = New Object(reader.FieldCount - 1) {}
                            reader.GetValues(fields)
                            r.ItemArray = fields
                            res.Rows.Add(r)
                        Loop
                    End Using
                Next
            End Using
            Return res

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行し、データテーブルを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteTableSync(setting As Settings,
                                           query As String,
                                           dynamicParameter As Object,
                                           ParamArray sqlParameter As Object()) As Task(Of DataTable)
        Return Await Task.Run(
            Function()
                Return ExecuteTable(setting, query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、データテーブルを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteTable(connect As IDbConnection,
                                 query As String,
                                 dynamicParameter As Object,
                                 ParamArray sqlParameter As Object()) As DataTable
        Return ExecuteTable(New Settings(connect), query, dynamicParameter, sqlParameter)
    End Function

    ''' <summary>SQLクエリを実行し、データテーブルを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteTableSync(connect As IDbConnection,
                                           query As String,
                                           dynamicParameter As Object,
                                           ParamArray sqlParameter As Object()) As Task(Of DataTable)
        Return Await Task.Run(
            Function()
                Return ExecuteTable(New Settings(connect), query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、データテーブルを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteTable(setting As Settings,
                                 query As String) As DataTable
        Return ExecuteTable(setting, query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、データテーブルを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteTableSync(setting As Settings,
                                           query As String) As Task(Of DataTable)
        Return Await Task.Run(
            Function()
                Return ExecuteTable(setting, query, Nothing, New Object() {})
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、データテーブルを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteTable(connect As IDbConnection,
                                 query As String) As DataTable
        Return ExecuteTable(New Settings(connect), query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、データテーブルを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteTableSync(connect As IDbConnection,
                                           query As String) As Task(Of DataTable)
        Return Await Task.Run(
            Function()
                Return ExecuteTable(New Settings(connect), query, Nothing, New Object() {})
            End Function
        )
    End Function

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteObject
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行し、ダイナミックオブジェクトを取得します。</summary>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteObject(setting As Settings,
                                 query As String,
                                 dynamicParameter As Object,
                                 ParamArray sqlParameter As Object()) As List(Of DynamicObject)
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim res As New List(Of DynamicObject)()
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat, setting.ParameterChecker, setting.PropertyNames)

                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    If prm IsNot Nothing Then
                        SetParameter(command, prm, props, varFormat)
                    End If

                    Using reader = command.ExecuteReader()
                        Dim tbl = reader.GetSchemaTable()
                        Dim clms As New List(Of String)()
                        For Each r As DataRow In tbl.Rows
                            clms.Add(r("ColumnName").ToString())
                        Next

                        ' 一行取得して行を追加
                        Dim fields = New Object(reader.FieldCount - 1) {}
                        Do While reader.Read()
                            reader.GetValues(fields)
                            res.Add(New DynamicRecord(clms, fields))
                        Loop
                    End Using
                Next
            End Using
            Return res

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行し、ダイナミックオブジェクトを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteObjectSync(setting As Settings,
                                            query As String,
                                            dynamicParameter As Object,
                                            ParamArray sqlParameter As Object()) As Task(Of List(Of DynamicObject))
        Return Await Task.Run(
            Function()
                Return ExecuteObject(setting, query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、ダイナミックオブジェクトを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteObject(connect As IDbConnection,
                                  query As String,
                                  dynamicParameter As Object,
                                  ParamArray sqlParameter As Object()) As List(Of DynamicObject)
        Return ExecuteObject(New Settings(connect), query, dynamicParameter, sqlParameter)
    End Function

    ''' <summary>SQLクエリを実行し、ダイナミックオブジェクトを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteObjectSync(connect As IDbConnection,
                                            query As String,
                                            dynamicParameter As Object,
                                            ParamArray sqlParameter As Object()) As Task(Of List(Of DynamicObject))
        Return Await Task.Run(
            Function()
                Return ExecuteObject(New Settings(connect), query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、ダイナミックオブジェクトを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteObject(setting As Settings,
                                  query As String) As List(Of DynamicObject)
        Return ExecuteObject(setting, query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、ダイナミックオブジェクトを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteObjectSync(setting As Settings,
                                            query As String) As Task(Of List(Of DynamicObject))
        Return Await Task.Run(
            Function()
                Return ExecuteObject(setting, query, Nothing, New Object() {})
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、ダイナミックオブジェクトを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteObject(connect As IDbConnection,
                                  query As String) As List(Of DynamicObject)
        Return ExecuteObject(New Settings(connect), query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、ダイナミックオブジェクトを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteObjectSync(connect As IDbConnection,
                                            query As String) As Task(Of List(Of DynamicObject))
        Return Await Task.Run(
            Function()
                Return ExecuteObject(New Settings(connect), query, Nothing, New Object() {})
            End Function
        )
    End Function

    ''' <summary>動的レコードです。</summary>
    Private Class DynamicRecord
        Inherits DynamicObject

        ' データリスト
        Private ReadOnly mData As New Dictionary(Of String, Object)()

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="clms">列名。</param>
        ''' <param name="fields">値。</param>
        Public Sub New(clms As List(Of String), fields As Object())
            For i As Integer = 0 To clms.Count - 1
                Me.mData.Add(clms(i).ToLower(), fields(i))
            Next
        End Sub

        ''' <summary>動的メンバーを取得します。</summary>
        ''' <param name="binder">バインダー。</param>
        ''' <param name="result">取得値。</param>
        ''' <returns>取得出来たら真。</returns>
        Public Overrides Function TryGetMember(binder As GetMemberBinder, ByRef result As Object) As Boolean
            Return Me.mData.TryGetValue(binder.Name.ToLower(), result)
        End Function

    End Class

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteQuery
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行します。</summary>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Function ExecuteQuery(setting As Settings,
                                 query As String,
                                 dynamicParameter As Object,
                                 ParamArray sqlParameter As Object()) As Integer
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim ans As Integer = 0
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat, setting.ParameterChecker, setting.PropertyNames)

                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    If prm IsNot Nothing Then
                        SetParameter(command, prm, props, varFormat)
                    End If

                    ' SQLを実行
                    ans += command.ExecuteNonQuery()
                Next
            End Using
            Return ans

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行します（非同期）</summary>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Async Function ExecuteQuerySync(setting As Settings,
                                           query As String,
                                           dynamicParameter As Object,
                                           ParamArray sqlParameter As Object()) As Task(Of Integer)
        Return Await Task.Run(
            Function()
                Return ExecuteQuery(setting, query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行します。</summary>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Function ExecuteQuery(connect As IDbConnection,
                                 query As String,
                                 dynamicParameter As Object,
                                 ParamArray sqlParameter As Object()) As Integer
        Return ExecuteQuery(New Settings(connect), query, dynamicParameter, sqlParameter)
    End Function

    ''' <summary>SQLクエリを実行します（非同期）</summary>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Async Function ExecuteQuerySync(connect As IDbConnection,
                                           query As String,
                                           dynamicParameter As Object,
                                           ParamArray sqlParameter As Object()) As Task(Of Integer)
        Return Await Task.Run(
            Function()
                Return ExecuteQuery(New Settings(connect), query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行します。</summary>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Function ExecuteQuery(setting As Settings,
                                 query As String) As Integer
        Return ExecuteQuery(setting, query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行します（非同期）</summary>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Async Function ExecuteQuerySync(setting As Settings,
                                           query As String) As Task(Of Integer)
        Return Await Task.Run(
            Function()
                Return ExecuteQuery(setting, query, Nothing, New Object() {})
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行します。</summary>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Function ExecuteQuery(connect As IDbConnection,
                                 query As String) As Integer
        Return ExecuteQuery(New Settings(connect), query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行します（非同期）</summary>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Async Function ExecuteQuerySync(connect As IDbConnection,
                                           query As String) As Task(Of Integer)
        Return Await Task.Run(
            Function()
                Return ExecuteQuery(New Settings(connect), query, Nothing, New Object() {})
            End Function
        )
    End Function

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteQuery(CSV)
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLパラメータの定義を設定します。</summary>
    ''' <param name="command">SQLコマンド。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <param name="varFormat">パラメータ変数書式。</param>
    Private Sub SetSqlParameterDefine(command As IDbCommand, csvParameter As CsvParameterBuilder, varFormat As String)
        Dim props As New Dictionary(Of String, ICsvType)()
        If csvParameter.Count > 0 Then
            LoggingDebug("Parameter class define")

            command.Parameters.Clear()
            For Each pair In csvParameter
                ' SQLパラメータを作成
                Dim prm = command.CreateParameter()

                ' パラメータの名前、方向を設定
                prm.ParameterName = String.Format(varFormat, pair.Name)
                GetDbType(pair.CsvColumnType.ColumnType, prm.DbType)
                prm.Direction = ParameterDirection.Input
                LoggingDebug($"・Name = {prm.ParameterName} Direction = {[Enum].GetName(GetType(ParameterDirection), prm.Direction)}")

                command.Parameters.Add(prm)
            Next
        End If
    End Sub

    ''' <summary>SQLクエリを実行します。</summary>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Function ExecuteQuery(setting As Settings,
                                 query As String,
                                 dynamicParameter As Object,
                                 csvStream As CsvReaderStream,
                                 csvParameter As CsvParameterBuilder) As Integer
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim ans As Integer = 0
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' パラメータの定義を設定
                SetSqlParameterDefine(command, csvParameter, varFormat)

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                For Each prm In csvStream
                    ' パラメータ変数に値を設定
                    Dim i As Integer = 0
                    For Each pair In csvParameter
                        ' 名前と値を取得
                        Dim propName = String.Format(varFormat, pair.Name)
                        Dim propVal = If(pair.CsvColumnType.Convert(prm.Items(i)), DBNull.Value)
                        i += 1

                        ' 変数に設定
                        LoggingDebug($"・{propName}={propVal}")
                        CType(command.Parameters(propName), IDbDataParameter).Value = propVal
                    Next

                    ' SQLを実行
                    ans += command.ExecuteNonQuery()
                Next
            End Using
            Return ans

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行します。</summary>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Function ExecuteQuery(connect As IDbConnection,
                                 query As String,
                                 dynamicParameter As Object,
                                 csvStream As CsvReaderStream,
                                 csvParameter As CsvParameterBuilder) As Integer
        Return ExecuteQuery(New Settings(connect), query, dynamicParameter, csvStream, csvParameter)
    End Function

    ''' <summary>SQLクエリを実行します。</summary>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Function ExecuteQuery(setting As Settings,
                                 query As String,
                                 csvStream As CsvReaderStream,
                                 csvParameter As CsvParameterBuilder) As Integer
        Return ExecuteQuery(setting, query, Nothing, csvStream, csvParameter)
    End Function

    ''' <summary>SQLクエリを実行します。</summary>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="csvStream">CSVストリーム。</param>
    ''' <param name="csvParameter">CSV変換パラメータ。</param>
    ''' <returns>影響行数。</returns>
    <Extension()>
    Public Function ExecuteQuery(connect As IDbConnection,
                                 query As String,
                                 csvStream As CsvReaderStream,
                                 csvParameter As CsvParameterBuilder) As Integer
        Return ExecuteQuery(New Settings(connect), query, Nothing, csvStream, csvParameter)
    End Function

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteArrays
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行し、オブジェクト配列のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteArrays(setting As Settings,
                                  query As String,
                                  dynamicParameter As Object,
                                  ParamArray sqlParameter As Object()) As List(Of Object())
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim recoreds As New List(Of Object())()
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat, setting.ParameterChecker, setting.PropertyNames)

                Dim allowNull As New List(Of Integer)()
                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    If prm IsNot Nothing Then
                        SetParameter(command, prm, props, varFormat)
                    End If

                    Using reader = command.ExecuteReader()
                        ' 一行取得してインスタンスを生成
                        Do While reader.Read()
                            Dim fields = New Object(reader.FieldCount - 1) {}
                            If reader.GetValues(fields) >= reader.FieldCount Then
                                ChangeDBNull(fields, allowNull)
                                recoreds.Add(fields)
                            End If
                        Loop
                    End Using
                Next
            End Using
            Return recoreds

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行し、オブジェクト配列のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteArraysSync(setting As Settings,
                                            query As String,
                                            dynamicParameter As Object,
                                            ParamArray sqlParameter As Object()) As Task(Of List(Of Object()))
        Return Await Task.Run(
            Function()
                Return ExecuteArrays(setting, query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、オブジェクト配列のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteArrays(connect As IDbConnection,
                                  query As String,
                                  dynamicParameter As Object,
                                  ParamArray sqlParameter As Object()) As List(Of Object())
        Return ExecuteArrays(New Settings(connect), query, dynamicParameter, sqlParameter)
    End Function

    ''' <summary>SQLクエリを実行し、オブジェクト配列のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteArraysSync(connect As IDbConnection,
                                            query As String,
                                            dynamicParameter As Object,
                                            ParamArray sqlParameter As Object()) As Task(Of List(Of Object()))
        Return Await Task.Run(
            Function()
                Return ExecuteArrays(New Settings(connect), query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、オブジェクト配列のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteArrays(setting As Settings,
                                  query As String) As List(Of Object())
        Return ExecuteArrays(setting, query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、オブジェクト配列のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteArraysSync(setting As Settings,
                                            query As String) As Task(Of List(Of Object()))
        Return Await Task.Run(
            Function()
                Return ExecuteArrays(setting, query, Nothing, New Object() {})
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、オブジェクト配列のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteArrays(connect As IDbConnection,
                                  query As String) As List(Of Object())
        Return ExecuteArrays(New Settings(connect), query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、オブジェクト配列のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteArraysSync(connect As IDbConnection,
                                            query As String) As Task(Of List(Of Object()))
        Return Await Task.Run(
            Function()
                Return ExecuteArrays(New Settings(connect), query, Nothing, New Object() {})
            End Function
        )
    End Function

    '-----------------------------------------------------------------------------------------------------------------
    ' ExecuteDatas
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>SQLクエリを実行し、指定データ型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteDatas(Of T)(setting As Settings,
                                       query As String,
                                       dynamicParameter As Object,
                                       ParamArray sqlParameter As Object()) As List(Of T)
        Try
            LoggingInformation($"Execute SQL : {query}")
            LoggingInformation($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingInformation($"Use command type : {setting.CommandType}")
            LoggingInformation($"Timeout seconds : {setting.TimeOutSecond}")
            Dim varFormat = GetVariantFormat(setting.ParameterPrefix)

            Dim values As New List(Of T)()
            Using command = setting.DbConnection.CreateCommand()
                ' タイムアウト秒を設定
                command.CommandTimeout = setting.TimeOutSecond

                ' トランザクションを設定
                If setting.Transaction IsNot Nothing Then
                    command.Transaction = setting.Transaction
                End If

                ' SQLクエリを設定
                command.CommandText = ParserAnalysis.Replase(query, dynamicParameter)
                LoggingInformation($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' SQLコマンドタイプを設定
                command.CommandType = setting.CommandType

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat, setting.ParameterChecker, setting.PropertyNames)

                Dim allowNull As New List(Of Integer)()
                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    If prm IsNot Nothing Then
                        SetParameter(command, prm, props, varFormat)
                    End If

                    Using reader = command.ExecuteReader()
                        ' 一行取得してインスタンスを生成
                        Dim fields = New Object(reader.FieldCount - 1) {}
                        Do While reader.Read()
                            If reader.GetValues(fields) >= reader.FieldCount Then
                                ChangeDBNull(fields, allowNull)
                                If fields.Length > 0 AndAlso TypeOf fields(0) Is T Then
                                    values.Add(CType(fields(0), T))
                                Else
                                    values.Add(Nothing)
                                End If
                            End If
                        Loop
                    End Using
                Next
            End Using
            Return values

        Catch ex As Exception
            LoggingError(ex.Message)
            LoggingError(ex.StackTrace)
            Throw
        End Try
    End Function

    ''' <summary>SQLクエリを実行し、指定データ型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteDatasSync(Of T)(setting As Settings,
                                                 query As String,
                                                 dynamicParameter As Object,
                                                 ParamArray sqlParameter As Object()) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteDatas(Of T)(setting, query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定データ型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteDatas(Of T)(connect As IDbConnection,
                                       query As String,
                                       dynamicParameter As Object,
                                       ParamArray sqlParameter As Object()) As List(Of T)
        Return ExecuteDatas(Of T)(New Settings(connect), query, dynamicParameter, sqlParameter)
    End Function

    ''' <summary>SQLクエリを実行し、指定データ型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <param name="dynamicParameter">動的SQLパラメータ。</param>
    ''' <param name="sqlParameter">SQLパラメータオブジェクト。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteDatasSync(Of T)(connect As IDbConnection,
                                                 query As String,
                                                 dynamicParameter As Object,
                                                 ParamArray sqlParameter As Object()) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteDatas(Of T)(New Settings(connect), query, dynamicParameter, sqlParameter)
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定データ型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteDatas(Of T)(setting As Settings,
                                       query As String) As List(Of T)
        Return ExecuteDatas(Of T)(setting, query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、指定データ型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="setting">実行パラメータ設定。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteDatasSync(Of T)(setting As Settings,
                                                 query As String) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteDatas(Of T)(setting, query, Nothing, New Object() {})
            End Function
        )
    End Function

    ''' <summary>SQLクエリを実行し、指定データ型のリストを取得します。</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Function ExecuteDatas(Of T)(connect As IDbConnection,
                                        query As String) As List(Of T)
        Return ExecuteDatas(Of T)(New Settings(connect), query, Nothing, New Object() {})
    End Function

    ''' <summary>SQLクエリを実行し、指定データ型のリストを取得します（非同期）</summary>
    ''' <typeparam name="T">戻り値の型。</typeparam>
    ''' <param name="connect">DBコネクション。</param>
    ''' <param name="query">SQLクエリ。</param>
    ''' <returns>実行結果。</returns>
    <Extension()>
    Public Async Function ExecuteDatasSync(Of T)(connect As IDbConnection,
                                                  query As String) As Task(Of List(Of T))
        Return Await Task.Run(
            Function()
                Return ExecuteDatas(Of T)(New Settings(connect), query, Nothing, New Object() {})
            End Function
        )
    End Function

End Module
