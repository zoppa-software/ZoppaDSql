Imports System
Imports ZoppaDSql
Imports Xunit
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis
Imports System.Runtime.InteropServices

Namespace ZoppaDSqlTest

    Public Class CompileTest

        <Fact>
        Public Sub WhereTrimTest()
#If False Then
            Dim query1 = "" &
"select * from employees 
{trim}
where
    {if empNo}emp_no < 20000{end if} and
    ({if first_name}first_name like 'A%'{end if} or 
     {if gender}gender = 'F'{end if})
{end trim}
limit 10"
            Dim ans1 = query1.Compile(New With {.empNo = False, .first_name = False, .gender = False})
            Assert.Equal(ans1.Trim(),
"select * from employees 
limit 10")

            Dim ans2 = query1.Compile(New With {.empNo = True, .first_name = False, .gender = False})
            Assert.Equal(ans2.Trim(),
"select * from employees 
where
    emp_no < 20000
limit 10")

            Dim ans3 = query1.Compile(New With {.empNo = False, .first_name = True, .gender = False})
            Assert.Equal(ans3.Trim(),
"select * from employees 
where
    (first_name like 'A%')
limit 10")
#End If

            Dim query2 = "" &
"select * from employees 
{trim}
where
    ({if empNo}emp_no = sysdate(){end if})
{end trim}
limit 10"
            Dim ans4 = query2.Compile(New With {.empNo = False})
            Assert.Equal(ans4.Trim(),
"select * from employees 
limit 10")

            Dim ans5 = query2.Compile(New With {.empNo = True})
            Assert.Equal(ans5.Trim(),
"select * from employees 
where
    (emp_no = sysdate())
limit 10")
        End Sub

        <Fact>
        Public Sub ConnmaTrimTest()
            Dim query1 = "" &
"SELECT
    *
FROM
    customers 
WHERE
    FirstName in ({trim}{foreach nm in names}#{nm}{}, {end for}{end trim})
"
            Dim ans1 = query1.Compile(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
            Assert.Equal(ans1.Trim(),
"SELECT
    *
FROM
    customers 
WHERE
    FirstName in ('Helena', 'Dan', 'Aaron')")
        End Sub

        <Fact>
        Public Sub CompileErrorTest()
            Dim query1 = "" &
"SELECT
    *
FROM
    customers 
WHERE
    FirstName in ({trim}
        {foreach nm in names}#{nm}{}, 
        {end foe}
    {end trim})
"
            Assert.Throws(Of DSqlAnalysisException)(
                Sub()
                    query1.Compile(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
                End Sub
            )
        End Sub

        <Fact>
        Public Sub ReplaseTest()
            Dim ansstr1 = "where parson_name = #{name}".Compile(New With {.name = "zouta takshi"})
            Assert.Equal(ansstr1, "where parson_name = 'zouta takshi'")

            Dim ansnum1 = "where age >= #{age}".Compile(New With {.age = 12})
            Assert.Equal(ansnum1, "where age >= 12")

            Dim ansnull = "set col1 = #{null}".Compile(New With {.null = Nothing})
            Assert.Equal(ansnull, "set col1 = null")

            Dim ansstr2 = "select * from !{table}".Compile(New With {.table = "sample_table"})
            Assert.Equal(ansstr2, "select * from sample_table")

            Dim ansstr3 = "select * from sample_${table}".Compile(New With {.table = "table"})
            Assert.Equal(ansstr3, "select * from sample_table")

            Dim ans1 = "a = #{'123'}".Compile()
            Assert.Equal(ans1, "a = '123'")

            Dim ans2 = "b = #{9 * 9}".Compile()
            Assert.Equal(ans2, "b = 81")

            Dim ans3 = "b = #{9 * num}".Compile(New With {.num = 3})
            Assert.Equal(ans3, "b = 27")
        End Sub

    End Class

End Namespace
