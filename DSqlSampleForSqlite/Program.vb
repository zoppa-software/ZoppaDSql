Imports System
Imports System.Data.SQLite
Imports System.Diagnostics.CodeAnalysis
Imports ZoppaDSql

Module Program

    Sub Main(args As String())
        Dim ans = Manage().Result
    End Sub

    Private Async Function Manage() As Task(Of Integer)
        Using sqlite As New SQLiteConnection("Data Source=chinook.db")
            sqlite.Open()

            Dim query1 = "" &
"select
  albumid, title, name
from
  albums
inner join artists on
  albums.ArtistId = artists.ArtistId
{trim}
where
  {if seachId <> NULL }albums.ArtistId = @seachId{end if}
{end trim}
"
            Dim ans1 = Await sqlite.ExecuteRecordsSync(Of AlbumInfo)(query1, New With {.seachId = Nothing})
            For Each v In ans1
                Console.WriteLine("AlbumId={0}, AlbumTitle={1}, ArtistName={2}", v.AlbumId, v.AlbumTitle, v.ArtistName)
            Next

            Dim ans2 = Await sqlite.ExecuteRecordsSync(Of AlbumInfo)(query1, New With {.seachId = 11})
            For Each v In ans2
                Console.WriteLine("AlbumId={0}, AlbumTitle={1}, ArtistName={2}", v.AlbumId, v.AlbumTitle, v.ArtistName)
            Next
        End Using
        Return 0
    End Function

    Public Class AlbumInfo

        Public ReadOnly Property AlbumId As Integer

        Public ReadOnly Property AlbumTitle As String

        Public ReadOnly Property ArtistName As String

        Public Sub New(id As Long, title As String, nm As String)
            Me.AlbumId = id
            Me.AlbumTitle = title
            Me.ArtistName = nm
        End Sub

    End Class

End Module
