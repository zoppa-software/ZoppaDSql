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

End Namespace