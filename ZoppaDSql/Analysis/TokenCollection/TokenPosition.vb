Option Strict On
Option Explicit On

Imports ZoppaDSql.Analysis.Tokens

Namespace TokenCollection

    ''' <summary>トークン位置情報です。</summary>
    Public Structure TokenPosition

        ' 参照しているトークン
        Private ReadOnly mToken As IToken

        ' トークン位置
        Public ReadOnly Position As Integer

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object
            Get
                Return Me.mToken.Contents
            End Get
        End Property

        ''' <summary>トークン名を取得する。</summary>
        ''' <returns>トークン名。</returns>
        Public ReadOnly Property TokenName As String
            Get
                Return Me.mToken.TokenName
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="src">参照するトークン。</param>
        ''' <param name="pos">トークン位置。</param>
        Public Sub New(src As IToken, pos As Integer)
            Me.mToken = src
            Me.Position = pos
        End Sub

        ''' <summary>型を指定してトークンを取得します。</summary>
        ''' <typeparam name="T">取得する型。</typeparam>
        ''' <returns>トークン情報。</returns>
        Friend Function GetToken(Of T As {Class, IToken})() As T
            Return TryCast(Me.mToken, T)
        End Function

        ''' <summary>コマンドトークンを取得します。</summary>
        ''' <returns>コマンドトークン。</returns>
        Friend Function GetCommandToken() As ICommandToken
            Return TryCast(Me.mToken, ICommandToken)
        End Function

        ''' <summary>文字列表現を取得する。</summary>
        ''' <returns>文字列。</returns>
        Public Overrides Function ToString() As String
            Return Me.mToken.ToString()
        End Function

    End Structure

End Namespace
