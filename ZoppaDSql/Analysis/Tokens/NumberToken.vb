Option Strict On
Option Explicit On

Namespace Analysis.Tokens

    ''' <summary>数値トークン。</summary>
    Public NotInheritable Class NumberToken
        Implements IToken

        ' 数値
        Private ReadOnly mValue As Double

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Return Me.mValue
            End Get
        End Property

        ''' <summary>トークン名を取得する。</summary>
        ''' <returns>トークン名。</returns>
        Public ReadOnly Property TokenName As String Implements IToken.TokenName
            Get
                Return NameOf(NumberToken)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">数値の文字列。</param>
        Public Sub New(value As String)
            Me.mValue = Double.Parse(value)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">数値。</param>
        Public Sub New(value As Double)
            Me.mValue = value
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me.mValue.ToString()
        End Function

    End Class

End Namespace
