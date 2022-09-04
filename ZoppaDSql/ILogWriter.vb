Option Strict On
Option Explicit On

Public Interface ILogWriter

    Sub LoggingInformation(message As String)

    Sub LoggingDebug(message As String)

    Sub LoggingError(message As String)

    Sub WaitFinish()

End Interface
