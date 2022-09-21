Option Strict On
Option Explicit On

Imports ZoppaDSql.Csv

''' <summary>CSVパラメータ構築機能です。</summary>
Public NotInheritable Class CsvParameterBuilder
    Implements IEnumerable(Of Pair)

    ' CSV列名リスト
    Private ReadOnly mNames As New List(Of String)()

    ' CSV列名、列型コレクション
    Private mMap As New SortedDictionary(Of String, ICsvType)()

    ''' <summary>登録した情報数を取得します。</summary>
    ''' <returns>数。</returns>
    Public ReadOnly Property Count() As Integer
        Get
            Return Me.mMap.Count
        End Get
    End Property

    ''' <summary>列名を指定して、列型情報を取得します。</summary>
    ''' <param name="name">列名。</param>
    ''' <returns>列型。</returns>
    Default Property Item(name As String) As ICsvType
        Get
            Dim v As ICsvType = Nothing
            If Me.mMap.TryGetValue(name, v) Then
                Return v
            Else
                Throw New CsvException("列型情報が取得できません")
            End If
        End Get
        Set(value As ICsvType)
            Me.Add(name, value)
        End Set
    End Property

    ''' <summary>位置を指定して、列名と列型情報を取得します。</summary>
    ''' <param name="index">位置。</param>
    ''' <returns>列名、列型のペア情報。</returns>
    Default ReadOnly Property Item(index As Integer) As Pair
        Get
            If index >= 0 AndAlso index < Me.mNames.Count Then
                Dim mn = Me.mNames(index)
                Return New Pair(mn, Me.mMap(mn))
            Else
                Throw New CsvException("列名、列型情報が取得できません")
            End If
        End Get
    End Property

    ''' <summary>登録している情報を消去します。</summary>
    Public Sub Clear()
        Me.mNames.Clear()
        Me.mMap.Clear()
    End Sub

    ''' <summary>列名、列型情報を追加します。</summary>
    ''' <param name="name">列名。</param>
    ''' <param name="csvType">列型。</param>
    Public Sub Add(name As String, csvType As ICsvType)
        If Me.mMap.ContainsKey(name) Then
            Me.mMap(name) = csvType
        Else
            Me.mMap.Add(name, csvType)
            Me.mNames.Add(name)
        End If
    End Sub

    ''' <summary>列挙子を取得します。</summary>
    ''' <returns>列挙子。</returns>
    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    ''' <summary>列挙子を取得します。</summary>
    ''' <returns>列挙子。</returns>
    Public Iterator Function GetEnumerator() As IEnumerator(Of Pair) Implements IEnumerable(Of Pair).GetEnumerator
        For Each pair In Me.mMap
            Yield New Pair(pair.Key, pair.Value)
        Next
    End Function

    ''' <summary>名前と CSV列型を保持したペア構造体。</summary>
    Public Structure Pair

        ''' <summary>CSV列名。</summary>
        Public ReadOnly Name As String

        ''' <summary>CSV列型情報。</summary>
        Public ReadOnly CsvColumnType As ICsvType

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="nm">CSV列名。</param>
        ''' <param name="csvTp">CSV列型情報。</param>
        Public Sub New(nm As String, csvTp As ICsvType)
            Me.Name = nm
            Me.CsvColumnType = csvTp
        End Sub

    End Structure

End Class