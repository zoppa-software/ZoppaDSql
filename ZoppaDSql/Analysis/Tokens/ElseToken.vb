﻿Option Strict On
Option Explicit On

Namespace Analysis.Tokens

    ''' <summary>Elseトークン。</summary>
    Public NotInheritable Class ElseToken
        Implements IToken

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyInstance() As New Lazy(Of ElseToken)(Function() New ElseToken())

        ''' <summary>唯一のインスタンスを返します。</summary>
        Public Shared ReadOnly Property Value() As ElseToken
            Get
                Return LazyInstance.Value
            End Get
        End Property

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Throw New NotImplementedException("使用できません")
            End Get
        End Property

        ''' <summary>トークン名を取得する。</summary>
        ''' <returns>トークン名。</returns>
        Public ReadOnly Property TokenName As String Implements IToken.TokenName
            Get
                Return NameOf(ElseToken)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        Private Sub New()

        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "Else"
        End Function

    End Class

End Namespace
