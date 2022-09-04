Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports ZoppaDSql

Module MainModule

    Sub Main()
        Dim ans = Manage().Result
    End Sub

    Private Async Function Manage() As Task(Of Integer)
        Using sqlsvr As New SqlConnection("Server=tcp:zoppa-dsql-sample.database.windows.net,1433;Initial Catalog=DSqlSample;Persist Security Info=False;User ID={};Password={};MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;")
            sqlsvr.Open()

            Dim query1 = "" &
"select
  [City], [CountryRegion], [AddressLine1]
from
  [SalesLT].[Address]
{trim}
where
  {if searchCuntry <> NULL}[SalesLT].[Address].[CountryRegion] = #{searchCuntry} and{end if}
  {if searchCity <> NULL}[SalesLT].[Address].[City] = @searchCity{end if}
{end trim}
"
            'Dim ans1 = Await sqlsvr.ExecuteRecordsSync(Of AAddressInfo)(query1, New With {.searchCuntry = "United States", .searchCity = "Gilbert"})
            'For Each v In ans1
            '    Console.WriteLine("都市={0}, 国={1}, Address={2}", v.City, v.Country, v.Address)
            'Next

            Dim ans2 = Await sqlsvr.ExecuteRecordsSync(Of AAddressInfo)(query1, New With {.searchCuntry = "United States", .searchCity = Nothing})
            For Each v In ans2
                Console.WriteLine("都市={0}, 国={1}, Address={2}", v.City, v.Country, v.Address)
            Next
        End Using
        Return 0
    End Function

    Public Class AAddressInfo

        Public ReadOnly Property City As String

        Public ReadOnly Property Country As String

        Public ReadOnly Property Address As String

        Public Sub New(cty As String, cnty As String, addr As String)
            Me.City = cty
            Me.Country = cnty
            Me.Address = addr
        End Sub

    End Class

End Module
