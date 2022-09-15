Option Strict On
Option Explicit On

Imports ZoppaDSql.Analysis.Environments
Imports ZoppaDSql.Analysis.Tokens

Namespace Analysis.Express

    ''' <summary>式インターフェイス。</summary>
    Public Interface IExpression

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Function Executes(env As IEnvironmentValue) As IToken

    End Interface

    ''' <summary>式のヘルパーモジュール。</summary>
    Module ExpressionModule

        ''' <summary>数値トークンの値を比較し、等しいか判定する。</summary>
        ''' <param name="tml">左辺トークン。</param>
        ''' <param name="tmr">右辺トークン。</param>
        ''' <returns>等しければ真を返す。</returns>
        Public Function EqualDouble(tml As IToken, tmr As IToken) As Boolean
            Dim lv = Convert.ToDouble(tml.Contents)
            Dim rv = Convert.ToDouble(tmr.Contents)
            Return (Math.Abs(lv - rv) < 0.00000000000001)
        End Function

    End Module

End Namespace