﻿Option Strict On
Option Explicit On

Namespace Analysis.Tokens

    ''' <summary>トークンインターフェイス。</summary>
    Public Interface IToken

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        ReadOnly Property Contents() As Object

        ''' <summary>トークン名を取得する。</summary>
        ''' <returns>トークン名。</returns>
        ReadOnly Property TokenName As String

    End Interface

    ''' <summary>命令付トークンインターフェイス。</summary>
    Friend Interface ICommandToken

        ''' <summary>命令トークンリストを取得します。</summary>
        ''' <returns>命令トークンリスト。</returns>
        ReadOnly Property CommandTokens As List(Of IToken)

    End Interface

End Namespace

