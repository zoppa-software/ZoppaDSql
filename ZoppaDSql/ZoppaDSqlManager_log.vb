Option Strict On
Option Explicit On

Imports System.IO
Imports System.IO.Compression

''' <summary>DSql APIモジュール。</summary>
Partial Module ZoppaDSqlManager

    ''' <summary>案内ログを出力します。</summary>
    ''' <param name="message">出力メッセージ。</param>
    Friend Sub LoggingInformation(message As String)
        Try
            mLogger?.LoggingInformation(message)
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>デバッグログを出力します。</summary>
    ''' <param name="message">出力メッセージ。</param>
    Friend Sub LoggingDebug(message As String)
        Try
            mLogger?.LoggingDebug(message)
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>エラーログを出力します。</summary>
    ''' <param name="message">出力メッセージ。</param>
    Friend Sub LoggingError(message As String)
        Try
            mLogger?.LoggingError(message)
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>デフォルトログを使用します。</summary>
    ''' <param name="logFilePath">出力ファイルパス。</param>
    ''' <param name="encode">出力エンコード。</param>
    ''' <param name="maxLogSize">最大ログファイルサイズ。</param>
    ''' <param name="logGeneration">ログ世代数。</param>
    ''' <param name="logLevel">ログレベル。</param>
    ''' <param name="dateChange">日付の変更でログを切り替えるかの設定。</param>
    Public Sub UseDefaultLogger(Optional logFilePath As String = "zoppa_dsql.txt",
                                Optional encode As Text.Encoding = Nothing,
                                Optional maxLogSize As Integer = 30 * 1024 * 1024,
                                Optional logGeneration As Integer = 10,
                                Optional logLevel As LogLevel = LogLevel.DebugLevel,
                                Optional dateChange As Boolean = False)
        Dim fi As New FileInfo(logFilePath)
        If Not fi.Directory.Exists Then
            fi.Directory.Create()
        End If
        mLogger = New Logger(logFilePath, If(encode, Text.Encoding.Default), maxLogSize, logGeneration, logLevel, dateChange)
    End Sub

    ''' <summary>カスタムログを設定します。</summary>
    ''' <param name="logger">カスタムログ。</param>
    Public Sub SetCustomLogger(logger As ILogWriter)
        mLogger = logger
    End Sub

    ''' <summary>書き込み中のログがあれば、書き込みが完了するまで待機します。</summary>
    Public Sub LogWaitFinish()
        mLogger?.WaitFinish()
    End Sub

    ''' <summary>ログ出力レベル。</summary>
    Public Enum LogLevel

        ''' <summary>エラーレベル。</summary>
        ErrorLevel = 0

        ''' <summary>案内レベル。</summary>
        InfomationLevel = 1

        ''' <summary>デバッグレベル。</summary>
        DebugLevel = 2

    End Enum

    ''' <summary>ログデータ。</summary>
    Private Structure LogData

        ''' <summary>書き込み日時。</summary>
        Public ReadOnly WriteTime As Date

        ''' <summary>ログレベル。</summary>
        Public ReadOnly LogLevel As LogLevel

        ''' <summary>ログメッセージ。。</summary>
        Public ReadOnly LogMessage As String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="wtm">書き込み日時。</param>
        ''' <param name="lv">ログレベル。</param>
        ''' <param name="msg">ログメッセージ。</param>
        Public Sub New(wtm As Date, lv As LogLevel, msg As String)
            Me.WriteTime = wtm
            Me.LogLevel = lv
            Me.LogMessage = msg
        End Sub

        ''' <summary>文字列表現を取得します。</summary>
        ''' <returns>ログ。</returns>
        Public Overrides Function ToString() As String
            Dim lv = "ERROR"
            Select Case Me.LogLevel
                Case LogLevel.InfomationLevel
                    lv = "INFO"
                Case LogLevel.DebugLevel
                    lv = "DEBUG"
            End Select
            Return $"[{WriteTime.ToString("yyyy/MM/dd HH:mm:ss")} {lv}] {Me.LogMessage}"
        End Function

    End Structure

    ''' <summary>ログ出力機能。</summary>
    Private NotInheritable Class Logger
        Implements ILogWriter

        ' 対象ファイル
        Private ReadOnly mLogFile As FileInfo

        ' 出力エンコード
        Private ReadOnly mEncode As Text.Encoding

        ' 最大ログサイズ
        Private ReadOnly mMaxLogSize As Integer

        ' 最大ログ世代数
        Private ReadOnly mLogGen As Integer

        ' ログ出力レベル
        Private ReadOnly mLogLevel As LogLevel

        ' 日付が変わったら切り替えるかのフラグ
        Private ReadOnly mDateChange As Boolean

        ' 書込みバッファ
        Private mQueue As New Queue(Of LogData)()

        ' 前回書込み完了日時
        Private mPrevWriteDate As Date

        ' 書込み中フラグ
        Private mWriting As Boolean

        ''' <summary>ログ設定を行う。</summary>
        ''' <param name="logFilePath">出力ファイル名。</param>
        ''' <param name="encode">出力エンコード。</param>
        ''' <param name="maxLogSize">最大ログファイルサイズ。</param>
        ''' <param name="logGeneration">ログ世代数。</param>
        ''' <param name="logLevel">ログ出力レベル。</param>
        ''' <param name="dateChange">日付の変更でログを切り替えるかの設定。</param>
        Public Sub New(logFilePath As String,
                       encode As Text.Encoding,
                       maxLogSize As Integer,
                       logGeneration As Integer,
                       logLevel As LogLevel,
                       dateChange As Boolean)
            Me.mLogFile = New FileInfo(logFilePath)
            Me.mEncode = encode
            Me.mMaxLogSize = maxLogSize
            Me.mLogGen = logGeneration
            Me.mWriting = False
            Me.mLogLevel = logLevel
            Me.mDateChange = dateChange
            Me.mPrevWriteDate = Date.MaxValue
        End Sub

        ''' <summary>ログをファイルに出力します。</summary>
        ''' <param name="message">出力するログ。</param>
        Public Sub Write(message As LogData)
            ' 書き出す情報をため込む
            Dim wrt As Boolean
            SyncLock Me
                Me.mQueue.Enqueue(message)
                wrt = Me.mWriting
            End SyncLock

            ' 別スレッドでファイルに出力
            If Not wrt Then
                Me.mWriting = True
                Task.Run(Sub() Me.Write())
            End If
        End Sub

        ''' <summary>ログをファイルに出力する。</summary>
        Private Sub Write()
            Me.mLogFile.Refresh()

            If Me.mLogFile.Exists AndAlso
               (Me.mLogFile.Length > Me.mMaxLogSize OrElse Me.ChangeOfDate) Then
                Try
                    ' 以前のファイルをリネーム
                    Dim ext = Path.GetExtension(Me.mLogFile.Name)
                    Dim nm = Me.mLogFile.Name.Substring(0, Me.mLogFile.Name.Length - ext.Length)
                    Dim tn = Date.Now.ToString("yyyyMMddHHmmssfff")

                    File.Move(Me.mLogFile.FullName, $"{mLogFile.Directory.FullName}\{nm}_{tn}\{nm}{ext}")
                    ZipFile.CreateFromDirectory(
                        $"{mLogFile.Directory.FullName}\{nm}_{tn}",
                        $"{mLogFile.Directory.FullName}\{nm}_{tn}.zip"
                    )
                    Directory.Delete($"{mLogFile.Directory.FullName}\{nm}_{tn}", True)

                    ' 過去ファイルを整理
                    Dim oldfiles = Directory.GetFiles(Me.mLogFile.Directory.FullName, $"{nm}*.zip").ToList()
                    oldfiles.Sort()
                    Do While oldfiles.Count > Me.mLogGen
                        File.Delete(oldfiles.First())
                        oldfiles.RemoveAt(0)
                    Loop
                Catch ex As Exception
                    SyncLock Me
                        Me.mWriting = False
                    End SyncLock
                    Return
                End Try
            End If

            Try
                Using sw As New StreamWriter(Me.mLogFile.FullName, True, Me.mEncode)
                    Dim writed As Boolean
                    Do
                        Threading.Thread.Sleep(10)
                                
                        ' キュー内の文字列を取得
                        writed = False
                        Dim ln As LogData? = Nothing
                        SyncLock Me
                            If Me.mQueue.Count > 0 Then
                                ln = Me.mQueue.Dequeue()
                            Else
                                Me.mWriting = False
                            End If
                        End SyncLock

                        ' ファイルに書き出す
                        If ln IsNot Nothing Then
                            sw.WriteLine(ln)
                            writed = True
                        End If
                    Loop While writed
                End Using
            Catch ex As Exception
                SyncLock Me
                    Me.mWriting = False
                End SyncLock
            Finally
                Me.mPrevWriteDate = Date.Now
            End Try
        End Sub

        ''' <summary>案内レベルログを出力します。</summary>
        ''' <param name="message">ログ。</param>
        Public Sub LoggingInformation(message As String) Implements ILogWriter.LoggingInformation
            If Me.mLogLevel >= LogLevel.InfomationLevel Then
                Me.Write(New LogData(Date.Now, LogLevel.InfomationLevel, message))
            End If
        End Sub

        ''' <summary>デバッグレベルログを出力します。</summary>
        ''' <param name="message">ログ。</param>
        Public Sub LoggingDebug(message As String) Implements ILogWriter.LoggingDebug
            If Me.mLogLevel >= LogLevel.DebugLevel Then
                Me.Write(New LogData(Date.Now, LogLevel.DebugLevel, message))
            End If
        End Sub

        ''' <summary>エラーレベルログを出力します。</summary>
        ''' <param name="message">ログ。</param>
        Public Sub LoggingError(message As String) Implements ILogWriter.LoggingError
            Me.Write(New LogData(Date.Now, LogLevel.ErrorLevel, message))
        End Sub

        ''' <summary>ログ出力終了を待機します。</summary>
        Public Sub WaitFinish() Implements ILogWriter.WaitFinish
            For i As Integer = 0 To 99 ' 事情があって書き込めないとき無限ループするためループ回数制限する
                If Me.IsWriting Then
                    Threading.Thread.Sleep(1000)
                Else
                    Exit For
                End If
            Next
        End Sub

        ''' <summary>書き込み中状態を取得します。</summary>
        ''' <returns>書き込み中状態。</returns>
        Public ReadOnly Property IsWriting() As Boolean
            Get
                SyncLock Me
                    Return Me.mWriting
                End SyncLock
            End Get
        End Property

        ''' <summary>日付の変更でログを切り替えるならば真を返します。</summary>
        ''' <returns>切り替えるならば真。</returns>
        Private ReadOnly Property ChangeOfDate() As Boolean
            Get
                Return Me.mDateChange AndAlso
                    Me.mPrevWriteDate.Date < Date.Now.Date
            End Get
        End Property

    End Class

End Module
