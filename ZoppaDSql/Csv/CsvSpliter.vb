Option Strict On
Option Explicit On

Imports System.IO

Namespace Csv

    ''' <summary>カンマ区切りで文字列を分割する機能です。</summary>
    Public NotInheritable Class CsvSpliter

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputStream">入力ストリーム。</param>
        Private Sub New(inputStream As StreamReader)
            Me.InnerStream = inputStream
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputText">入力文字列。</param>
        Private Sub New(inputText As String)
            Me.InnerStream = New StringReader(inputText)
        End Sub

        ''' <summary>カンマ区切り分割機能を生成します。</summary>
        ''' <param name="inputStream">入力ストリーム。</param>
        ''' <returns>カンマ区切り分割機能。</returns>
        Public Shared Function CreateSpliter(inputStream As StreamReader) As CsvSpliter
            Return New CsvSpliter(inputStream)
        End Function

        ''' <summary>カンマ区切り分割機能を生成します。</summary>
        ''' <param name="inputText">分割する文字列。</param>
        ''' <returns>カンマ区切り分割機能。</returns>
        Public Shared Function CreateSpliter(inputText As String) As CsvSpliter
            Return New CsvSpliter(inputText)
        End Function

        ''' <summary>カンマ区切り分割を行います。</summary>
        ''' <returns>カンマ区切り分割結果。</returns>
        Friend Function ReadLine() As ReadResult
            Dim rchars As New List(Of Char)(4096)
            Dim index As Integer = 0
            Dim spoint As New List(Of Integer)(256)
            Dim esc As Boolean = False

            spoint.Add(0)
            Do While Me.InnerStream.Peek() <> -1
                Dim c = Convert.ToChar(Me.InnerStream.Read())

                Select Case c
                    Case ChrW(13)
                        rchars.Add(c)
                        index += 1
                        If Not esc AndAlso Me.InnerStream.Peek() = 10 Then
                            Me.InnerStream.Read()
                            rchars.Add(ChrW(10))
                            index -= 1
                            Exit Do
                        End If

                    Case ChrW(10)
                        rchars.Add(c)
                        index += 1
                        If Not esc Then
                            index -= 1
                            Exit Do
                        End If

                    Case """"c
                        rchars.Add(c)
                        index += 1

                        If esc Then
                            If Me.InnerStream.Peek() = AscW(""""c) Then
                                rchars.Add(""""c)
                                index += 1
                                Me.InnerStream.Read()
                            Else
                                esc = False
                            End If
                        Else
                            esc = True
                        End If

                    Case ","c
                        rchars.Add(c)
                        index += 1
                        If Not esc Then
                            spoint.Add(index)
                        End If

                    Case Else
                        rchars.Add(c)
                        index += 1
                End Select
            Loop
            spoint.Add(index + 1)

            Return New ReadResult(rchars.ToArray(), spoint)
        End Function

        ''' <summary>内部より一行を読み込み、分割して返します。</summary>
        ''' <returns>分割した項目の配列。</returns>
        Public Function Split() As List(Of CsvItem)
            Dim ans = Me.ReadLine()
            Dim res As New List(Of CsvItem)(ans.Items.Count)
            For Each i In ans.Items
                res.Add(New CsvItem(i))
            Next
            Return res
        End Function

        ''' <summary>内部ストリームを取得します。</summary>
        Private ReadOnly Property InnerStream() As TextReader

        ''' <summary>分割結果を保持する。</summary>
        Friend Structure ReadResult

            ''' <summary>分割位置リスト。</summary>
            Private ReadOnly mSPoint As List(Of Integer)

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="chars">読み込んだ文字配列。</param>
            ''' <param name="spoint">分割位置。</param>
            Public Sub New(chars As Char(), spoint As List(Of Integer))
                Me.Chars = chars
                Me.mSPoint = spoint
            End Sub

            ''' <summary>読み込んだ文字配列を取得します。</summary>
            ''' <returns>文字配列。</returns>
            Public ReadOnly Property Chars() As Char()

            ''' <summary>項目があれば真を返します。</summary>
            ''' <returns>項目があれば真。</returns>
            Public ReadOnly Property HasCsv() As Boolean
                Get
                    Return (Me.Chars.Length > 0)
                End Get
            End Property

            ''' <summary>項目のリストを返します。</summary>
            ''' <returns>項目リスト。</returns>
            Public ReadOnly Property Items As List(Of ReadOnlyMemory)
                Get
                    Dim src As New ReadOnlyMemory(Me.Chars)
                    Dim split As New List(Of ReadOnlyMemory)(Me.mSPoint.Count - 1)
                    For i As Integer = 0 To Me.mSPoint.Count - 2
                        split.Add(src.Slice(Me.mSPoint(i), (Me.mSPoint(i + 1) - 1) - Me.mSPoint(i)))
                    Next
                    Return split
                End Get
            End Property

        End Structure

    End Class

End Namespace
