Option Strict On
Option Explicit On

Imports System.Data
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Text
Imports ZoppaDSql.Analysis
Imports ZoppaDSql.Analysis.Tokens
Imports ZoppaDSql.Csv
Imports ZoppaDSql.ZoppaDSqlSetting

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
    ''' <returns>プロパティインフォ。</returns>
    Private Function SetSqlParameterDefine(command As IDbCommand, parameter As Object(), varFormat As String) As PropertyInfo()
        Dim props = New PropertyInfo(-1) {}
        If parameter.Length > 0 Then
            LoggingDebug("Parameter class define")

            ' プロパティインフォを取得
            props = parameter(0).GetType().GetProperties()

            command.Parameters.Clear()
            For Each prop In props
                ' SQLパラメータを作成
                Dim prm = command.CreateParameter()

                ' パラメータの名前、方向を設定
                prm.ParameterName = String.Format(varFormat, prop.Name)
                prm.DbType = GetDbType(prop.PropertyType)
                If prop.CanRead AndAlso prop.CanWrite Then
                    prm.Direction = ParameterDirection.Input
                Else
                    prm.Direction = If(prop.CanWrite, ParameterDirection.Output, ParameterDirection.Input)
                End If
                LoggingDebug($"・Name = {prm.ParameterName} Direction = {[Enum].GetName(GetType(ParameterDirection), prm.Direction)}")

                command.Parameters.Add(prm)
            Next
        End If
        Return props
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
            LoggingDebug($"・{propName}={propVal}")
            CType(command.Parameters(propName), IDbDataParameter).Value = propVal
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
    ''' <returns>DBのデータ型。</returns>
    Private Function GetDbType(propType As Type) As DbType
        Select Case propType
            Case GetType(String)
                Return DbType.String

            Case GetType(Integer), GetType(Integer?)
                Return DbType.Int32

            Case GetType(Long), GetType(Long?)
                Return DbType.Int64

            Case GetType(Short), GetType(Short?)
                Return DbType.Int16

            Case GetType(Single), GetType(Single?)
                Return DbType.Single

            Case GetType(Double), GetType(Double?)
                Return DbType.Double

            Case GetType(Decimal), GetType(Decimal?)
                Return DbType.Decimal

            Case GetType(Date), GetType(Date?)
                Return DbType.Date

            Case GetType(TimeSpan), GetType(TimeSpan?)
                Return DbType.Time

            Case GetType(Object)
                Return DbType.Object

            Case GetType(Boolean), GetType(Boolean?)
                Return DbType.Boolean

            Case GetType(Byte), GetType(Byte?)
                Return DbType.Byte

            Case GetType(Byte())
                Return DbType.Binary

            Case GetType(Char())
                Return DbType.String

            Case Else
                Throw New DSqlException("SQLパラメータに使用できないプロパティを使用しています")
        End Select
    End Function

    ''' <summary>マッピングするコンストラクタを取得します。</summary>
    ''' <typeparam name="T">対象の型。</typeparam>
    ''' <param name="reader">データリーダー。</param>
    ''' <returns>コンストラクタインフォ。</returns>
    Private Function CreateConstructorInfo(Of T)(reader As IDataReader) As ConstructorInfo
        ' 各列のデータ型を取得
        Dim columns As New List(Of Type)()
        For i As Integer = 0 To reader.FieldCount - 1
            'Dim dbType = GetDbType(reader.GetFieldType(i))
            columns.Add(reader.GetFieldType(i))
        Next

        ' コンストラクタを取得
        Dim constructor = GetType(T).GetConstructor(columns.ToArray())
        If constructor IsNot Nothing Then
            Return constructor
        Else
            Dim info = String.Join(",", columns.Select(Function(c) c.Name).ToArray())
            Throw New DSqlException($"取得データに一致するコンストラクタがありません:{info}")
        End If
    End Function

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
            LoggingDebug($"Execute SQL : {query}")
            LoggingDebug($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingDebug($"Timeout seconds : {setting.TimeOutSecond}")
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
                LoggingDebug($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If dynamicParameter IsNot Nothing AndAlso sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat)

                Dim constructor As ConstructorInfo = Nothing
                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    SetParameter(command, prm, props, varFormat)

                    Using reader = command.ExecuteReader()
                        ' マッピングコンストラクタを設定
                        If constructor Is Nothing Then
                            constructor = CreateConstructorInfo(Of T)(reader)
                        End If

                        ' 一行取得してインスタンスを生成
                        Dim fields = New Object(reader.FieldCount - 1) {}
                        Do While reader.Read()
                            If reader.GetValues(fields) >= reader.FieldCount Then
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
    ' ExecuteQuery(CSV)
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="setting"></param>
    ''' <param name="query"></param>
    ''' <param name="dynamicParameter"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         dynamicParameter As Object,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder) As List(Of T)
        Try
            LoggingDebug($"Execute SQL : {query}")
            LoggingDebug($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingDebug($"Timeout seconds : {setting.TimeOutSecond}")
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
                LoggingDebug($"Answer SQL : {command.CommandText}")

                ' パラメータの定義を設定
                SetSqlParameterDefine(command, csvParameter, varFormat)

                Dim constructor As ConstructorInfo = Nothing
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
                            constructor = CreateConstructorInfo(Of T)(reader)
                        End If

                        ' 一行取得してインスタンスを生成
                        Dim fields = New Object(reader.FieldCount - 1) {}
                        Do While reader.Read()
                            If reader.GetValues(fields) >= reader.FieldCount Then
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="connect"></param>
    ''' <param name="query"></param>
    ''' <param name="dynamicParameter"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         dynamicParameter As Object,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, csvStream, csvParameter)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="setting"></param>
    ''' <param name="query"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder) As List(Of T)
        Return ExecuteRecords(Of T)(setting, query, Nothing, csvStream, csvParameter)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="connect"></param>
    ''' <param name="query"></param>
    ''' <param name="dynamicParameter"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <returns></returns>
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
            LoggingDebug($"Execute SQL : {query}")
            LoggingDebug($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingDebug($"Timeout seconds : {setting.TimeOutSecond}")
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
                LoggingDebug($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If dynamicParameter IsNot Nothing AndAlso sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat)

                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    SetParameter(command, prm, props, varFormat)

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
    ' ExecuteQuery(createrMethod, CSV)
    '-----------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="setting"></param>
    ''' <param name="query"></param>
    ''' <param name="dynamicParameter"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <param name="createrMethod"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         dynamicParameter As Object,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Try
            LoggingDebug($"Execute SQL : {query}")
            LoggingDebug($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingDebug($"Timeout seconds : {setting.TimeOutSecond}")
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
                LoggingDebug($"Answer SQL : {command.CommandText}")

                ' パラメータの定義を設定
                SetSqlParameterDefine(command, csvParameter, varFormat)

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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="connect"></param>
    ''' <param name="query"></param>
    ''' <param name="dynamicParameter"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <param name="createrMethod"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         dynamicParameter As Object,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, dynamicParameter, csvStream, csvParameter, createrMethod)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="setting"></param>
    ''' <param name="query"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <param name="createrMethod"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(setting As Settings,
                                         query As String,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(setting, query, Nothing, csvStream, csvParameter, createrMethod)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="connect"></param>
    ''' <param name="query"></param>
    ''' <param name="dynamicParameter"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <param name="createrMethod"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteRecords(Of T)(connect As IDbConnection,
                                         query As String,
                                         csvStream As CsvReaderStream,
                                         csvParameter As CsvParameterBuilder,
                                         createrMethod As Func(Of Object(), T)) As List(Of T)
        Return ExecuteRecords(Of T)(New Settings(connect), query, Nothing, csvStream, csvParameter, createrMethod)
    End Function

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
            LoggingDebug($"Execute SQL : {query}")
            LoggingDebug($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingDebug($"Timeout seconds : {setting.TimeOutSecond}")
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
                LoggingDebug($"Answer SQL : {command.CommandText}")

                ' SQLパラメータが空なら動的パラメータを展開
                If dynamicParameter IsNot Nothing AndAlso sqlParameter.Length = 0 Then
                    sqlParameter = New Object() {dynamicParameter}
                End If

                ' パラメータの定義を設定
                Dim props = SetSqlParameterDefine(command, sqlParameter, varFormat)

                For Each prm In sqlParameter
                    ' パラメータ変数に値を設定
                    SetParameter(command, prm, props, varFormat)

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
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="command"></param>
    ''' <param name="csvParameter"></param>
    ''' <param name="varFormat"></param>
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
                prm.DbType = GetDbType(pair.CsvColumnType.ColumnType)
                prm.Direction = ParameterDirection.Input
                LoggingDebug($"・Name = {prm.ParameterName} Direction = {[Enum].GetName(GetType(ParameterDirection), prm.Direction)}")

                command.Parameters.Add(prm)
            Next
        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="setting"></param>
    ''' <param name="query"></param>
    ''' <param name="dynamicParameter"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteQuery(setting As Settings,
                                 query As String,
                                 dynamicParameter As Object,
                                 csvStream As CsvReaderStream,
                                 csvParameter As CsvParameterBuilder) As Integer
        Try
            LoggingDebug($"Execute SQL : {query}")
            LoggingDebug($"Use Transaction : {setting.Transaction IsNot Nothing}")
            LoggingDebug($"Timeout seconds : {setting.TimeOutSecond}")
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
                LoggingDebug($"Answer SQL : {command.CommandText}")

                ' パラメータの定義を設定
                SetSqlParameterDefine(command, csvParameter, varFormat)

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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="connect"></param>
    ''' <param name="query"></param>
    ''' <param name="dynamicParameter"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteQuery(connect As IDbConnection,
                                 query As String,
                                 dynamicParameter As Object,
                                 csvStream As CsvReaderStream,
                                 csvParameter As CsvParameterBuilder) As Integer
        Return ExecuteQuery(New Settings(connect), query, dynamicParameter, csvStream, csvParameter)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="setting"></param>
    ''' <param name="query"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteQuery(setting As Settings,
                                 query As String,
                                 csvStream As CsvReaderStream,
                                 csvParameter As CsvParameterBuilder) As Integer
        Return ExecuteQuery(setting, query, Nothing, csvStream, csvParameter)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="connect"></param>
    ''' <param name="query"></param>
    ''' <param name="csvStream"></param>
    ''' <param name="csvParameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ExecuteQuery(connect As IDbConnection,
                                 query As String,
                                 csvStream As CsvReaderStream,
                                 csvParameter As CsvParameterBuilder) As Integer
        Return ExecuteQuery(New Settings(connect), query, Nothing, csvStream, csvParameter)
    End Function

End Module
