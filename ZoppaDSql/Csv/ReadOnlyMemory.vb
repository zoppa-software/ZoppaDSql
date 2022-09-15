Option Strict On
Option Explicit On

Namespace Csv

    ''' <summary>読み込み専用オブジェクトです。</summary>
    Public Structure ReadOnlyMemory

        ' 文字配列
        Private ReadOnly mItems As Char()

        ' 分割開始位置
        Private ReadOnly mStart As Integer

        ' 文字数
        Private ReadOnly mLength As Integer

        ''' <summary>文字数を取得します。</summary>
        ''' <returns>文字数。</returns>
        Public ReadOnly Property Length() As Integer
            Get
                Return Me.mLength
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="item">分割元文字配列。</param>
        Public Sub New(item As Char())
            Me.mItems = item
            Me.mStart = 0
            Me.mLength = item.Length
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="source">分割元文字配列。</param>
        ''' <param name="start">分割開始位置。</param>
        ''' <param name="length">文字数。</param>
        Public Sub New(source As ReadOnlyMemory, start As Integer, length As Integer)
            Me.mItems = source.mItems
            Me.mStart = start
            Me.mLength = length
        End Sub

        ''' <summary>部分を抽出します。</summary>
        ''' <param name="start">抽出開始位置。</param>
        ''' <param name="length">抽出文字数。</param>
        ''' <returns>抽出文字列。</returns>
        Public Function Slice(start As Integer, length As Integer) As ReadOnlyMemory
            Return New ReadOnlyMemory(Me, Me.mStart + start, length)
        End Function

        ''' <summary>文字列を取得します。</summary>
        ''' <returns>文字列。</returns>
        Public Overrides Function ToString() As String
            Return New String(Me.mItems, Me.mStart, Me.mLength)
        End Function

    End Structure

End Namespace
