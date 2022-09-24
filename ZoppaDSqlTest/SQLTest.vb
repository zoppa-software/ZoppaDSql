Imports System
Imports ZoppaDSql
Imports Xunit
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis
Imports System.Text
Imports ZoppaDSql.Csv
Imports System.Data.SQLite

Namespace ZoppaDSqlTest

    Public Class SQLTest
        Implements IDisposable

        Private mSQLite As SQLiteConnection

        Public Sub New()
            Me.mSQLite = New SQLiteConnection("Data Source=sample.db")
            Me.mSQLite.Open()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Me.mSQLite?.Dispose()
        End Sub

        <Fact>
        Public Sub CUIDTest()
            Dim tran1 = Me.mSQLite.BeginTransaction()
            Try
                Me.mSQLite.SetTransaction(tran1).ExecuteQuery(
"CREATE TABLE Zodiac (
name TEXT,
jp_name TEXT,
from_date DATETIME NOT NULL,
to_date DATETIME NOT NULL,
PRIMARY KEY(name)
)")
                tran1.Commit()
            Catch ex As Exception
                tran1.Rollback()
            End Try

            Dim zodiacs = New Zodiac() {
                New Zodiac("Aries", "牡羊座", New Date(2022, 3, 21), New Date(2022, 4, 19)),
                New Zodiac("Taurus", "牡牛座", New Date(2022, 4, 20), New Date(2022, 5, 20)),
                New Zodiac("Gemini", "双子座", New Date(2022, 5, 21), New Date(2022, 6, 21)),
                New Zodiac("Cancer", "蟹座", New Date(2022, 6, 22), New Date(2022, 7, 22)),
                New Zodiac("Leo", "獅子座", New Date(2022, 7, 23), New Date(2022, 8, 22)),
                New Zodiac("Virgo", "乙女座", New Date(2022, 8, 23), New Date(2022, 9, 22)),
                New Zodiac("Libra", "天秤座", New Date(2022, 9, 23), New Date(2022, 10, 23)),
                New Zodiac("Scorpio", "蠍座", New Date(2022, 10, 24), New Date(2022, 11, 22)),
                New Zodiac("Sagittarius", "射手座", New Date(2022, 11, 23), New Date(2022, 12, 21)),
                New Zodiac("Capricom", "山羊座", New Date(2022, 12, 22), New Date(2023, 1, 19)),
                New Zodiac("Aquuarius", "水瓶座", New Date(2023, 1, 20), New Date(2023, 2, 18)),
                New Zodiac("Pisces", "魚座", New Date(2023, 2, 19), New Date(2023, 3, 20))
            }

            Dim tran = Me.mSQLite.BeginTransaction()
            Try
                Me.mSQLite.SetTransaction(tran).ExecuteQuery(
                    "INSERT INTO Zodiac (name, jp_name, from_date, to_date) 
                    VALUES (@Name, @JpName, @FromDate, @ToDate)", Nothing, zodiacs)
                tran.Commit()
            Catch ex As Exception
                tran.Rollback()
            End Try

            Dim ansZodiacs = Me.mSQLite.ExecuteRecords(Of Zodiac)(
                "select name, jp_name, from_date, to_date from Zodiac"
            )
            Assert.Equal(zodiacs, ansZodiacs.ToArray())

            Dim tran3 = Me.mSQLite.BeginTransaction()
            Try
                Me.mSQLite.SetTransaction(tran3).ExecuteQuery(
"DROP TABLE Zodiac")
                tran3.Commit()
            Catch ex As Exception
                tran3.Rollback()
            End Try
        End Sub

    End Class

    Public Class Zodiac
        Public Property Name As String
        Public Property JpName As String
        Public Property FromDate As Date
        Public Property ToDate As Date
        Public Sub New(nm As String, jp As String, frm As Date, tod As Date)
            Me.Name = nm
            Me.JpName = jp
            Me.FromDate = frm
            Me.ToDate = tod
        End Sub
        Public Overrides Function Equals(obj As Object) As Boolean
            With TryCast(obj, Zodiac)
                Return Me.Name = .Name AndAlso
                       Me.JpName = .JpName AndAlso
                       Me.FromDate = .FromDate AndAlso
                       Me.ToDate = .ToDate
            End With
        End Function
    End Class

End Namespace

