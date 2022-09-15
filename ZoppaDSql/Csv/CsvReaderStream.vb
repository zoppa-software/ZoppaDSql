Option Strict On
Option Explicit On

Imports System.Reflection

Namespace Csv

    ''' <summary>カンマ区切りファイル読み込みストリームです。</summary>
    Public Class CsvReaderStream
        Inherits IO.StreamReader
        Implements IEnumerable(Of Pointer)

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        Public Sub New(stream As IO.Stream)
            MyBase.New(stream)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        Public Sub New(path As String)
            MyBase.New(path)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        Public Sub New(stream As IO.Stream, detectEncodingFromByteOrderMarks As Boolean)
            MyBase.New(stream, detectEncodingFromByteOrderMarks)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding)
            MyBase.New(stream, encoding)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        Public Sub New(path As String, detectEncodingFromByteOrderMarks As Boolean)
            MyBase.New(path, detectEncodingFromByteOrderMarks)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        Public Sub New(path As String, encoding As Text.Encoding)
            MyBase.New(path, encoding)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        Public Sub New(path As String, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="path">入力ファイルパス。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        Public Sub New(path As String, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, bufferSize)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="stream">元となるストリーム。</param>
        ''' <param name="encoding">テキストエンコード。</param>
        ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
        ''' <param name="bufferSize">バッファサイズ。</param>
        ''' <param name="leaveOpen">ストリームを開いたままにするならば真。</param>
        Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, leaveOpen As Boolean)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize, leaveOpen)
        End Sub

        ''' <summary>列挙子を取得する。</summary>
        ''' <returns>列挙子。</returns>
        Public Iterator Function GetEnumerator() As IEnumerator(Of Pointer) Implements IEnumerable(Of Pointer).GetEnumerator
            Dim spliter = CsvSpliter.CreateSpliter(Me)
            Dim ans = spliter.ReadLine()
            Dim index As Integer = 0
            Do While ans.HasCsv
                Yield New Pointer(index, ans.Items)
                index += 1
                ans = spliter.ReadLine()
            Loop
        End Function

        ''' <summary>列挙子を取得する。</summary>
        ''' <returns>列挙子。</returns>
        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return Me.GetEnumerator()
        End Function

        ''' <summary>指定した型へ変換して取得するクエリ。</summary>
        ''' <typeparam name="TResult">変換後の型。</typeparam>
        ''' <param name="func">変換するためのラムダ式。</param>
        ''' <returns>変換後型の列挙子。</returns>
        Public Iterator Function SelectCsv(Of TResult)(func As Func(Of Integer, List(Of CsvItem), TResult)) As IEnumerable(Of TResult)
            For Each item In Me
                Yield func(item.Row, item.Items)
            Next
        End Function

        Public Iterator Function SelectCsv(Of T)(ParamArray columTypes As ICsvType()) As IEnumerable(Of T)
            Dim clmTps = columTypes.Select(Function(v) v.ColumnType).ToArray()

            Dim constructor = GetConstructor(Of T)(clmTps)

            Dim fields As New ArrayList()
            For Each item In Me
                SetFields(columTypes, item, fields)
                Yield CType(constructor.Invoke(fields.ToArray()), T)
            Next
        End Function

        ''' <summary>条件に一致した行を指定した型へ変換して取得するクエリ。</summary>
        ''' <typeparam name="TResult">変換後の型。</typeparam>
        ''' <param name="condition">条件判定するラムダ式。</param>
        ''' <param name="func">変換するためのラムダ式。</param>
        ''' <returns>変換後型の列挙子。</returns>
        Public Iterator Function WhereCsv(Of TResult)(condition As Func(Of Integer, List(Of CsvItem), Boolean), func As Func(Of Integer, List(Of CsvItem), TResult)) As IEnumerable(Of TResult)
            For Each item In Me
                If condition(item.Row, item.Items) Then
                    Yield func(item.Row, item.Items)
                End If
            Next
        End Function

        Public Iterator Function WhereCsv(Of T)(condition As Func(Of Integer, List(Of CsvItem), Boolean),
                                                ParamArray columTypes As ICsvType()) As IEnumerable(Of T)
            Dim clmTps = columTypes.Select(Function(v) v.ColumnType).ToArray()

            Dim constructor = GetConstructor(Of T)(clmTps)

            Dim fields As New ArrayList()
            For Each item In Me
                If condition(item.Row, item.Items) Then
                    SetFields(columTypes, item, fields)
                    Yield CType(constructor.Invoke(fields.ToArray()), T)
                End If
            Next
        End Function

        Private Shared Function GetConstructor(Of T)(clmTps() As Type) As ConstructorInfo
            Dim constructor = GetType(T).GetConstructor(clmTps)
            If constructor Is Nothing Then
                constructor = GetType(T).GetConstructor(New Type() {GetType(Object())})
            End If
            If constructor Is Nothing Then
                Dim info = String.Join(",", clmTps.Select(Function(c) c.Name).ToArray())
                Throw New CsvException($"取得データに一致するコンストラクタがありません:{info}")
            End If
            Return constructor
        End Function

        Private Shared Sub SetFields(columTypes As ICsvType(), item As Pointer, fields As ArrayList)
            fields.Clear()
            For i As Integer = 0 To Math.Min(columTypes.Length, item.Items.Count) - 1
                Try
                    fields.Add(columTypes(i).Convert(item.Items(i)))
                Catch ex As Exception
                    Throw New CsvException($"変換に失敗しました:{i},{item.Row} {item.Items(i).Raw} -> {columTypes(i).ColumnType.Name}")
                End Try
            Next
        End Sub

        ''' <summary>項目のポインタ。</summary>
        Public Structure Pointer

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="row">行位置。</param>
            ''' <param name="items">分割項目。</param>
            Public Sub New(row As Integer, items As List(Of ReadOnlyMemory))
                Me.Row = row
                Me.Items = New List(Of CsvItem)(items.Count)
                For Each i In items
                    Me.Items.Add(New CsvItem(i))
                Next
            End Sub

            ''' <summary>行位置を取得する。</summary>
            ''' <returns>行位置。</returns>
            Public ReadOnly Property Row() As Integer

            ''' <summary>分割項目の配列を取得する。</summary>
            ''' <returns>分割項目配列。</returns>
            Public ReadOnly Property Items() As List(Of CsvItem)

            ''' <summary>文字表現を取得する。</summary>
            ''' <returns>文字列。</returns>
            Public Overrides Function ToString() As String
                Return $"row index:{Me.Row} colum count:{Me.Items.Count}"
            End Function

        End Structure

    End Class

End Namespace