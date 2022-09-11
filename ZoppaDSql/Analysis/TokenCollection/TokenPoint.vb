Option Strict On
Option Explicit On

Imports ZoppaDSql.Analysis.Tokens

Namespace Analysis

    Public Structure TokenPoint
        Implements IToken

        Private ReadOnly Token As IToken

        Public ReadOnly Position As Integer

        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Return Me.Token.Contents
            End Get
        End Property

        Public ReadOnly Property TokenName As String Implements IToken.TokenName
            Get
                Return Me.Token.TokenName
            End Get
        End Property

        Public Sub New(src As IToken, pos As Integer)
            Me.Token = src
            Me.Position = pos
        End Sub

        Friend Function GetToken(Of T As {Class, IToken})() As T
            Return TryCast(Me.Token, T)
        End Function

        Friend Function GetCommandToken() As ICommandToken
            Return TryCast(Me.Token, ICommandToken)
        End Function

        Public Overrides Function ToString() As String
            Return Me.Token.ToString()
        End Function

    End Structure

End Namespace
