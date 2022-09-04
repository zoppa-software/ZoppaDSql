﻿Option Strict On
Option Explicit On

Imports ZoppaDSql.Analysis.Tokens

Namespace Analysis.Express

    ''' <summary>等号式。</summary>
    Public NotInheritable Class EqualExpress
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
                Throw New DSqlAnalysisException("等号式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As EnvironmentValue) As IToken Implements IExpression.Executes
            Dim tml = Me.mTml?.Executes(env)
            Dim tmr = Me.mTmr?.Executes(env)

            Dim bval As Boolean = False
            If TypeOf tml?.Contents Is Double AndAlso TypeOf tmr?.Contents Is Double Then
                ' 数値比較
                bval = EqualDouble(tml, tmr)
            ElseIf tml?.Contents Is Nothing AndAlso tmr?.Contents Is Nothing Then
                ' 両方 Null比較
                bval = True
            Else
                bval = If(tml?.Contents?.Equals(tmr?.Contents), False)
            End If
            Return If(bval, CType(TrueToken.Value, IToken), FalseToken.Value)
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:="
        End Function

    End Class

End Namespace