Option Strict On
Option Explicit On

''' <summary>ログ出力インターフェイス</summary>
Public Interface ILogWriter

    ''' <summary>案内レベルログを出力します。</summary>
    ''' <param name="message">ログ。</param>
    Sub LoggingInformation(message As String)

    ''' <summary>デバッグレベルログを出力します。</summary>
    ''' <param name="message">ログ。</param>
    Sub LoggingDebug(message As String)

    ''' <summary>エラーレベルログを出力します。</summary>
    ''' <param name="message">ログ。</param>
    Sub LoggingError(message As String)

    ''' <summary>ログ出力終了を待機します。</summary>
    Sub WaitFinish()

End Interface
