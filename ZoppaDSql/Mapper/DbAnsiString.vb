Option Strict On
Option Explicit On

Public NotInheritable Class DBAnsiString
    Implements IComparable(Of DBAnsiString)

    Private ReadOnly mStr As String

    Public Sub New(str As String)
        Me.mStr = str
    End Sub

    Public Shared Widening Operator CType(ByVal v As DBAnsiString) As String
        Return v.mStr
    End Operator

    Public Shared Widening Operator CType(ByVal v As String) As DBAnsiString
        Return New DBAnsiString(v)
    End Operator

    Public Overrides Function ToString() As String
        Return Me.mStr
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Dim other = TryCast(obj, DBAnsiString)
        If other IsNot Nothing Then
            Return (Me.mStr = other.mStr)
        End If
        Return False
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return Me.mStr.GetHashCode()
    End Function

    Public Function CompareTo(other As DBAnsiString) As Integer Implements IComparable(Of DBAnsiString).CompareTo
        Return Me.mStr.CompareTo(other.mStr)
    End Function

    Public Shared Operator =(left As DBAnsiString, right As DBAnsiString) As Boolean
        Return Not left.Equals(right)
    End Operator

    Public Shared Operator <>(left As DBAnsiString, right As DBAnsiString) As Boolean
        Return left.Equals(right)
    End Operator

End Class
