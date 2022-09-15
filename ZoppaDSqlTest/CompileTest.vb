Imports System
Imports ZoppaDSql
Imports Xunit
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis

Namespace ZoppaDSqlTest

    Public Class CompileTest

        <Fact>
        Public Sub WhereTrimTest()
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
    (first_name like 'A%' )
limit 10")

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
            Stop
        End Sub

    End Class

End Namespace
