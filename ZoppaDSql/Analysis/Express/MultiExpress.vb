﻿Option Strict On
Option Explicit On

Imports ZoppaDSql.Analysis.Environments
Imports ZoppaDSql.Analysis.Tokens

Namespace Analysis.Express

    ''' <summary>乗算式。</summary>
    Public NotInheritable Class MultiExpress
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
                Throw New DSqlAnalysisException("乗算式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Dim tml = Me.mTml?.Executes(env)
            Dim tmr = Me.mTmr?.Executes(env)

            Dim nml = TryCast(tml, NumberToken)
            Dim nmr = TryCast(tmr, NumberToken)
            If nml IsNot Nothing AndAlso nmr IsNot Nothing Then
                Return nml.MultiComputation(nmr)
            Else
                Try
                    Dim lf = If(nml, NumberToken.Create(If(tml?.Contents.ToString(), "null")))
                    Dim rt = If(nmr, NumberToken.Create(If(tmr?.Contents.ToString(), "null")))
                    Return nml.MultiComputation(nmr)
                Catch ex As Exception
                    Throw New DSqlAnalysisException($"乗算ができません。{tml.Contents} * {tmr.Contents}", ex)
                End Try
            End If
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:*"
        End Function

    End Class

End Namespace

