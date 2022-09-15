Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.Common
Imports System.Runtime.CompilerServices
Imports System.Transactions

''' <summary>SQL実行用パラメータ APIです。</summary>
Public Module ZoppaDSqlSetting

    ' タイムアウト設定
    Private mDefTimeOut As Integer = 30

    ''' <summary>デフォルトのタイムアウト値を設定、取得します。</summary>
    ''' <returns>タイムアウト値。</returns>
    Public Property DefaultTimeoutSecond As Integer
        Get
            Return mDefTimeOut
        End Get
        Set(value As Integer)
            If value >= -1 Then
                mDefTimeOut = value
            Else
                Throw New DSqlException("タイムアウト値に不正な値を指定しています。")
            End If
        End Set
    End Property

    ''' <summary>デフォルトパラメータ接頭辞を設定、取得します。</summary>
    ''' <returns>パラメータ接頭辞。</returns>
    Public Property DefaultParameterPrefix As PrefixType = PrefixType.AtMark

    ''' <summary>タイムアウト値を設定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="timeout">タイムアウト値。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetTimeoutSecond(dbConnection As IDbConnection, timeout As Integer) As Settings
        Return New Settings(dbConnection, timeout, Nothing, DefaultParameterPrefix)
    End Function

    ''' <summary>タイムアウト値を設定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="timeout">タイムアウト値。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetTimeoutSecond(prevSetting As Settings, timeout As Integer) As Settings
        With prevSetting
            Return New Settings(.DbConnection, timeout, .Transaction, .ParameterPrefix)
        End With
    End Function

    ''' <summary>トランザクションを設定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="transaction">トランザクション。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetTransaction(dbConnection As IDbConnection, transaction As IDbTransaction) As Settings
        Return New Settings(dbConnection, DefaultTimeoutSecond, transaction, DefaultParameterPrefix)
    End Function

    ''' <summary>トランザクションを設定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="transaction">トランザクション。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetTransaction(prevSetting As Settings, transaction As IDbTransaction) As Settings
        With prevSetting
            Return New Settings(.DbConnection, .TimeOutSecond, transaction, .ParameterPrefix)
        End With
    End Function

    ''' <summary>使用するパラメータ接頭辞を指定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="prefix">パラメータ接頭辞。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetParameterPrepix(dbConnection As IDbConnection, prefix As PrefixType) As Settings
        Return New Settings(dbConnection, DefaultTimeoutSecond, Nothing, prefix)
    End Function

    ''' <summary>使用するパラメータ接頭辞を指定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="prefix">パラメータ接頭辞。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetParameterPrepix(prevSetting As Settings, prefix As PrefixType) As Settings
        With prevSetting
            Return New Settings(.DbConnection, .TimeOutSecond, .Transaction, prefix)
        End With
    End Function

    ''' <summary>SQL実行用パラメータ定義です。</summary>
    Public Structure Settings

        ''' <summary>DB接続情報です。</summary>
        Friend ReadOnly DbConnection As IDbConnection

        ' タイムアウト設定
        Private mTimeOut As Integer

        ' トランザクション設定
        Private mTran As IDbTransaction

        ' 変数接頭辞設定
        Private mVarPrefix As PrefixType

        ''' <summary>タイムアウト設定を取得します。</summary>
        ''' <returns>タイムアウト時間（秒数）</returns>
        Public ReadOnly Property TimeOutSecond As Integer
            Get
                Return Me.mTimeOut
            End Get
        End Property

        ''' <summary>トランザクションを取得します。</summary>
        ''' <returns>トランザクションオブジェクト。</returns>
        Public ReadOnly Property Transaction As IDbTransaction
            Get
                Return Me.mTran
            End Get
        End Property

        ''' <summary>パラメータの接頭辞を取得します。</summary>
        ''' <returns>接頭辞の種類。</returns>
        Public ReadOnly Property ParameterPrefix As PrefixType
            Get
                Return Me.mVarPrefix
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="dbCon">DBコネクション。</param>
        Friend Sub New(dbCon As IDbConnection)
            Me.DbConnection = dbCon
            Me.mTimeOut = DefaultTimeoutSecond
            Me.mTran = Nothing
            Me.mVarPrefix = DefaultParameterPrefix
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="dbCon">DBコネクション。</param>
        ''' <param name="timeOut">タイムアウト値。</param>
        ''' <param name="tran">トランザクション。</param>
        ''' <param name="varPrex">パラメータ接頭辞。</param>
        Friend Sub New(dbCon As IDbConnection,
                       timeOut As Integer,
                       tran As IDbTransaction,
                       varPrex As PrefixType)
            Me.DbConnection = dbCon
            Me.mTimeOut = timeOut
            Me.mTran = tran
            Me.mVarPrefix = varPrex
        End Sub

    End Structure

End Module

