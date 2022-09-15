Option Strict On
Option Explicit On

Namespace Csv

    ''' <summary>CSV項目情報。</summary>
    Public Structure CsvItem

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="items">CSV項目スライス。</param>
        Public Sub New(items As ReadOnlyMemory)
            Me.Raw = items
        End Sub

        ''' <summary>エスケープを解除して返す。</summary>
        ''' <returns>CSV項目文字列。</returns>
        Public Function UnEscape() As String
            Dim str = Me.Text
            If str.Length >= 2 AndAlso str(0) = """"c AndAlso str(str.Length - 1) = """"c Then
                ' "のエスケープを解除して返す
                Dim slice = Me.Raw.Slice(1, Me.Raw.Length - 2)
                Return slice.ToString().Replace("""""", """")
            Else
                ' 前後トリムして返す
                Return str.Trim()
            End If
        End Function

        ''' <summary>生値を取得する。</summary>
        ''' <returns>生値。</returns>
        Public ReadOnly Property Raw() As ReadOnlyMemory

        ''' <summary>値の文字列を取得する。</summary>
        ''' <returns>CSV項目文字列。</returns>
        Public ReadOnly Property Text() As String
            Get
                Return Me.Raw.ToString()
            End Get
        End Property

        ''' <summary>インスタンスを表現する文字列を取得する。</summary>
        ''' <returns>文字列。</returns>
        Public Overrides Function ToString() As String
            Return Me.UnEscape()
        End Function

    End Structure

End Namespace