Imports System
Imports ZoppaDSql
Imports Xunit
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis
Imports System.Text
Imports ZoppaDSql.Csv

Namespace ZoppaDSqlTest

    Public Class CsvTest

        Class Sample1Csv

            Public ReadOnly Property Item1 As String

            Public ReadOnly Property Item2 As String

            Public ReadOnly Property Item3 As String

            Public Sub New(s1 As String, s2 As String, s3 As String)
                Me.Item1 = s1
                Me.Item2 = s2
                Me.Item3 = s3
            End Sub

            Public Overrides Function Equals(obj As Object) As Boolean
                Dim other = TryCast(obj, Sample1Csv)
                Return Me.Item1 = other.Item1 AndAlso
                       Me.Item2 = other.Item2 AndAlso
                       Me.Item3 = other.Item3
            End Function

        End Class

        Public Sub New()
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        End Sub

        <Fact>
        Public Sub ReadTest()
            Dim ans1 As New List(Of Sample1Csv)()
            Using sr As New CsvReaderStream("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
                ans1 = sr.SelectCsv(Of Sample1Csv)(CsvType.CsvString, CsvType.CsvString, CsvType.CsvString).ToList()
            End Using
            Assert.Equal(5, ans1.Count)
            Assert.Equal(New Sample1Csv("項目1", "項目2", "項目3"), ans1(0))
            Assert.Equal(New Sample1Csv("1", "2", "3"), ans1(1))
            Assert.Equal(New Sample1Csv("1,1", $"2{vbCrLf}2", " 3.""3 "), ans1(2))
            Assert.Equal(New Sample1Csv("a", "b", "c"), ans1(3))
            Assert.Equal(New Sample1Csv("1", "2", "3"), ans1(4))

            Dim ans2 As New List(Of Sample1Csv)()
            Using sr As New CsvReaderStream("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
                ans2 = sr.WhereCsv(Of Sample1Csv)(
                    Function(row, item) row >= 1,
                    CsvType.CsvString, CsvType.CsvString, CsvType.CsvString
                ).ToList()
            End Using
            Assert.Equal(4, ans2.Count)
            Assert.Equal(New Sample1Csv("1", "2", "3"), ans2(0))
            Assert.Equal(New Sample1Csv("1,1", $"2{vbCrLf}2", " 3.""3 "), ans2(1))
            Assert.Equal(New Sample1Csv("a", "b", "c"), ans2(2))
            Assert.Equal(New Sample1Csv("1", "2", "3"), ans2(3))
        End Sub

        Class Sample2Csv
            Public ReadOnly Property ItemBytes As Byte()
            Public ReadOnly Property ItemByte As Byte
            Public ReadOnly Property ItemBool As Boolean
            Public ReadOnly Property ItemDate As Date
            Public ReadOnly Property ItemDec As Decimal
            Public ReadOnly Property ItemDbl As Double
            Public ReadOnly Property ItemSht As Short
            Public ReadOnly Property ItemInt As Integer
            Public ReadOnly Property ItemLng As Long
            Public ReadOnly Property ItemSng As Single
            Public ReadOnly Property ItemStr As String
            Public ReadOnly Property ItemTime As TimeSpan
            Public Sub New(pBytes As Byte(),
                           pByte As Byte?,
                           pBool As Boolean?,
                           pDate As Date?,
                           pDec As Decimal?,
                           pDbl As Double?,
                           pSht As Short?,
                           pInt As Integer?,
                           pLng As Long?,
                           pSng As Single?,
                           pStr As String,
                           pTime As TimeSpan?)
                Me.ItemBytes = pBytes
                Me.ItemByte = pByte
                Me.ItemBool = pBool
                Me.ItemDate = pDate
                Me.ItemDec = pDec
                Me.ItemDbl = pDbl
                Me.ItemSht = pSht
                Me.ItemInt = pInt
                Me.ItemLng = pLng
                Me.ItemSng = pSng
                Me.ItemStr = pStr
                Me.ItemTime = pTime
            End Sub
        End Class

        <Fact>
        Public Sub ConvertTest()
            Dim ans1 As New List(Of Sample2Csv)()
            Using sr As New CsvReaderStream("CsvFiles\Sample2.csv", Encoding.GetEncoding("shift_jis"))
                ans1 = sr.SelectCsv(Of Sample2Csv)(
                    CsvType.CsvBinary,
                    CsvType.CsvByte,
                    CsvType.CsvBoolean,
                    CsvType.CsvDateTime,
                    CsvType.CsvDecimal,
                    CsvType.CsvDouble,
                    CsvType.CsvShort,
                    CsvType.CsvInteger,
                    CsvType.CsvLong,
                    CsvType.CsvSingle,
                    CsvType.CsvString,
                    CsvType.CsvTime
                ).ToList()
            End Using
            Assert.Equal(1, ans1.Count)
            Assert.Equal(New Byte() {&H1, &H6, &HC, &H11, &H16, &H1C, &HFF}, ans1(0).ItemBytes)
            Assert.Equal(10, ans1(0).ItemByte)
            Assert.Equal(True, ans1(0).ItemBool)
            Assert.Equal(New Date(2022, 9, 14), ans1(0).ItemDate)
            Assert.Equal(100, ans1(0).ItemDec)
            Assert.Equal(10.5, ans1(0).ItemDbl)
            Assert.Equal(1000, ans1(0).ItemSht)
            Assert.Equal(20, ans1(0).ItemInt)
            Assert.Equal(200, ans1(0).ItemLng)
            Assert.Equal(20.8, ans1(0).ItemSng)
            Assert.Equal("ABC", ans1(0).ItemStr)
            Assert.Equal(New TimeSpan(3, 21, 0), ans1(0).ItemTime)
        End Sub

        <Fact>
        Public Sub SplitTest()
            Dim csv = CsvSpliter.CreateSpliter("あ,い,う,え,""お,を""").Split()
            Assert.Equal(csv(0).UnEscape(), "あ")
            Assert.Equal(csv(1).UnEscape(), "い")
            Assert.Equal(csv(2).UnEscape(), "う")
            Assert.Equal(csv(3).UnEscape(), "え")
            Assert.Equal(csv(4).UnEscape(), "お,を")
            Assert.Equal(csv(4).Text, """お,を""")

            Using sr As New CsvReaderStream("CsvFiles\Sample3.csv", Encoding.GetEncoding("shift_jis"))
                For Each pointer In sr
                    Console.Out.WriteLine($"{pointer.Items(0).UnEscape()}, {pointer.Items(1).UnEscape()}, …")
                Next
            End Using

            Dim ans As New List(Of Sample1Csv)()
            Using sr As New CsvReaderStream("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
                ans = sr.WhereCsv(Of Sample1Csv)(
                    Function(row, item) row >= 1,
                    Function(row, item) New Sample1Csv(item(0).UnEscape(), item(1).UnEscape(), item(2).UnEscape())
                ).ToList()
            End Using
            Assert.Equal(4, ans.Count)
            Assert.Equal(New Sample1Csv("1", "2", "3"), ans(0))
            Assert.Equal(New Sample1Csv("1,1", $"2{vbCrLf}2", " 3.""3 "), ans(1))
            Assert.Equal(New Sample1Csv("a", "b", "c"), ans(2))
            Assert.Equal(New Sample1Csv("1", "2", "3"), ans(3))
        End Sub

    End Class

End Namespace
