Option Strict On
Option Explicit On

Namespace Csv

    ''' <summary>CSV操作例外。</summary>
    Public Class CsvException
        Inherits Exception

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="message">例外メッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

    End Class

End Namespace
