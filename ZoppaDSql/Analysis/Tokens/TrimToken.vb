Option Strict On
Option Explicit On

Namespace Analysis.Tokens

    ''' <summary>Trimトークン。</summary>
    Public NotInheritable Class TrimToken
        Implements IToken

        ''' <summary>末尾をトリムするならば真を返します。</summary>
        Public ReadOnly Property IsTrimRight As Boolean

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
                Return NameOf(TrimToken)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        Private Sub New()
            Me.IsTrimRight = True
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="isRight">末尾をトリムするならば真。</param>
        Public Sub New(isRight As Boolean)
            Me.IsTrimRight = isRight
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "Trim"
        End Function

    End Class

End Namespace
