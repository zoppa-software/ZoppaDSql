﻿Option Strict On
Option Explicit On

Imports ZoppaDSql.Analysis.Tokens

Namespace Analysis.Express

    ''' <summary>減算式。</summary>
    Public NotInheritable Class MinusExpress
        Implements IExpression

        ' 左辺式
        Private ReadOnly mTml As IExpression

        ' 右辺式
        Private ReadOnly mTmr As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tml">左辺式。</param>
        ''' <param name="tmr">右辺式。</param>
        Public Sub New(tml As IExpression, tmr As IExpression)
            If tml IsNot Nothing AndAlso tmr IsNot Nothing Then
                Me.mTml = tml
                Me.mTmr = tmr
            Else
                Throw New DSqlAnalysisException("減算式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As EnvironmentValue) As IToken Implements IExpression.Executes
            Dim tml = Me.mTml?.Executes(env)
            Dim tmr = Me.mTmr?.Executes(env)

            Try
                Return New NumberToken(Convert.ToDouble(tml.Contents) - Convert.ToDouble(tmr.Contents))
            Catch ex As Exception
                Throw New DSqlAnalysisException($"減算ができません。{tml.Contents} - {tmr.Contents}", ex)
            End Try
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:-"
        End Function

    End Class

End Namespace