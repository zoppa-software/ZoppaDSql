Option Strict On
Option Explicit On

Public Class PrimaryKeyList(Of T)
    Implements IEnumerable(Of T)

    Private mInnerList As New SortedDictionary(Of UniqueKey, T)

    Default Public ReadOnly Property Item(index As Integer) As T
        Get
            If index >= 0 AndAlso index < Me.mInnerList.Keys.Count Then
                Dim key = Me.mInnerList.Keys(index)
                Return Me.mInnerList(key)
            Else
                Throw New IndexOutOfRangeException("インデックスが要素の範囲外です")
            End If
        End Get
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Return Me.mInnerList.Count
        End Get
    End Property

    Public Sub RemoveAt(index As Integer)
        If index >= 0 AndAlso index < Me.mInnerList.Keys.Count Then
            Dim key = Me.mInnerList.Keys(index)
            Me.mInnerList.Remove(key)
        Else
            Throw New IndexOutOfRangeException("インデックスが要素の範囲外です")
        End If
    End Sub

    Public Sub Regist(item As T, ParamArray primaryKey As Object())
        Dim key = UniqueKey.Create(primaryKey)
        Me.mInnerList.Add(key, item)
    End Sub

    Public Sub Clear()
        Me.mInnerList.Clear()
    End Sub

    Public Function IndexOf(item As T) As Integer
        Dim i As Integer = 0
        For Each pair In Me.mInnerList
            If pair.Value.Equals(item) Then
                Return i
            End If
            i += 1
        Next
        Return -1
    End Function

    Public Function Contains(item As T) As Boolean
        For Each pair In Me.mInnerList
            If pair.Value.Equals(item) Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function ContainsKey(ParamArray keys As Object()) As Boolean
        Dim key = UniqueKey.Create(keys)
        Return Me.mInnerList.ContainsKey(key)
    End Function

    Public Function GetValue(ParamArray keys As Object()) As T
        Dim key = UniqueKey.Create(keys)
        Return Me.mInnerList(key)
    End Function

    Public Function TrySearchValue(ByRef v As T, ParamArray keys As Object()) As Boolean
        Dim key = UniqueKey.Create(keys)
        Return Me.mInnerList.TryGetValue(key, v)
    End Function

    Public Function Remove(item As T) As Boolean
        Dim key As UniqueKey = Nothing
        For Each pair In Me.mInnerList
            If pair.Value.Equals(item) Then
                key = pair.Key
                Exit For
            End If
        Next
        If key IsNot Nothing Then
            Me.mInnerList.Remove(key)
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
        Return Me.mInnerList.Values.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

End Class
