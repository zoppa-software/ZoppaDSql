Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports ZoppaDSql

Module MainModule

    Sub Main()
        Dim ans = Manage().Result
    End Sub

    Private Async Function Manage() As Task(Of Integer)
        ZoppaDSqlManager.UseDefaultLogger()

        Dim query0 = "" &
"select
  [City], [CountryRegion], [AddressLine1]
from
  [SalesLT].[Address]
{trim}
where
  ({if searchCuntry <> NULL}[SalesLT].[Address].[CountryRegion] = #{searchCuntry}{end if} and
  {if searchCity <> NULL}[SalesLT].[Address].[City] = @searchCity{end if}) or
  {if searchCity2 <> NULL}[SalesLT].[Address].[City] = @searchCity2{end if}
{end trim}
"
        Dim k = ZoppaDSqlManager.Compile(query0, New With {.searchCuntry = Nothing, .searchCity = Nothing, .searchCity2 = "2"})


            sqlsvr.Open()

            Dim query1 = "" &
"select
  [City], [CountryRegion], [AddressLine1]
from
  [SalesLT].[Address]
{trim}
where
  {if searchCuntry <> NULL}[SalesLT].[Address].[CountryRegion] = #{searchCuntry}{end if} and
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

        ZoppaDSqlManager.LogWaitFinish()

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
