﻿Imports System
Imports ZoppaDSql
Imports Xunit
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis
Imports System.Runtime.InteropServices

Namespace ZoppaDSqlTest

    Public Class CompileTest

        <Fact>
        Public Sub WhereTrimTest()
            Dim query3 = "" &
"select * from tb1
{trim}
where
    ({if af}a = 1{end if} or {if bf}b = 2{end if}) and ({if cf}c = 3{end if} or {if df}d = 4{end if})
{end trim}
"
            Dim ans7 = query3.Compile(New With {.af = True, .bf = False, .cf = False, .df = True})
            Assert.Equal(ans7.Trim(),
"select * from tb1
where
    (a = 1) and (d = 4)")

            Dim ans6 = query3.Compile(New With {.af = False, .bf = False, .cf = False, .df = False})
            Assert.Equal(ans6.Trim(), "select * from tb1")

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

            Dim query2 = "" &
"SELECT
    *
FROM
    customers 
WHERE
    FirstName in (#{names})
"
            Dim ans2 = query2.Compile(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
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

        <Fact>
        Public Sub IfTest()
            Dim query = "" &
"select * from table1
where
  {if num = 1}col1 = #{num}
  {else if num = 2}col2 = #{num}
  {else}col3 = #{num}
  {end if}"
            Dim ans1 = query.Compile(New With {.num = 1})
            Assert.Equal(ans1,
"select * from table1
where
  col1 = 1")
            Dim ans2 = query.Compile(New With {.num = 2})
            Assert.Equal(ans2,
"select * from table1
where
  col2 = 2")
            Dim ans3 = query.Compile(New With {.num = 5})
            Assert.Equal(ans3,
"select * from table1
where
  col3 = 5")
        End Sub

        <Fact>
        Public Sub TrimTest()
            Dim ans1 = "{trim}   a = #{12 * 13}   {end trim}".Compile()
            Assert.Equal(ans1, "a = 156")

            Dim ans2 = "{trim trush}   a = #{'11' + '29'}trush{end trim}".Compile()
            Assert.Equal(ans2, "a = '1129'")

            Dim strs As New List(Of String)()
            strs.Add("あいうえお")
            strs.Add("かきくけこ")
            strs.Add("さしすせそ")
            strs.Add("たちつてと")
            strs.Add("なにぬねの")
            Dim ans3 = "{trim}
{foreach str in strs}
    #{str},
{end for}
{end trim}".Compile(New With {.strs = strs})
            Assert.Equal(ans3, "'あいうえお',
    'かきくけこ',
    'さしすせそ',
    'たちつてと',
    'なにぬねの'")
        End Sub

        <Fact>
        Public Sub TrimTest2()
            Dim query = "SELECT
{if sel}
	T_摘要M.摘要CD,T_摘要M.摘要名
{else}
	count(*) as cnt
{end if}
FROM T_摘要M
{trim}
WHERE
    {if txZyCd <> ''}T_摘要M.担当者CD like '%!{txZyCd}%'{end if} AND
    {if txZyNm <> ''}T_摘要M.摘要名 like '%!{txZyNm}%'{end if}
{end trim}
{if sel}
ORDER BY T_摘要M.摘要CD
{end if}	
"
            Dim ans1 = query.Compile(New With {.sel = True, .txZyCd = "A", .txZyNm = ""})
            Assert.Equal(ans1, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名
FROM T_摘要M
WHERE
    T_摘要M.担当者CD like '%A%'
ORDER BY T_摘要M.摘要CD")

            Dim ans2 = query.Compile(New With {.sel = True, .txZyCd = "", .txZyNm = ""})
            Assert.Equal(ans2, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名
FROM T_摘要M
ORDER BY T_摘要M.摘要CD")

            Dim ans3 = query.Compile(New With {.sel = True, .txZyCd = "", .txZyNm = "B"})
            Assert.Equal(ans3, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名
FROM T_摘要M
WHERE
    T_摘要M.摘要名 like '%B%'
ORDER BY T_摘要M.摘要CD")

            Dim ans4 = query.Compile(New With {.sel = True, .txZyCd = "A", .txZyNm = "B"})
            Assert.Equal(ans4, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名
FROM T_摘要M
WHERE
    T_摘要M.担当者CD like '%A%' AND
    T_摘要M.摘要名 like '%B%'
ORDER BY T_摘要M.摘要CD")
        End Sub

        <Fact>
        Public Sub TrimTest3()
            Dim query = "SELECT
{if sel}
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
{else}
	count(*) as cnt
{end if}
FROM T_摘要M
{trim}
WHERE
    {if txZyCd <> ''}T_摘要M.担当者CD like '%!{txZyCd}%'{end if} AND
    {if txZyNm <> ''}T_摘要M.摘要名 like '%!{txZyNm}%'{end if} AND
    {if txZyNo <> ''}T_摘要M.摘要No like '%!{txZyNo}%'{end if}
{end trim}
{if sel}
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No
{end if}	
"
            Dim ans1 = query.Compile(New With {.sel = True, .txZyCd = "A", .txZyNm = "", .txZyNo = ""})
            Assert.Equal(ans1, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
FROM T_摘要M
WHERE
    T_摘要M.担当者CD like '%A%'
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No")

            Dim ans2 = query.Compile(New With {.sel = True, .txZyCd = "", .txZyNm = "B", .txZyNo = ""})
            Assert.Equal(ans2, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
FROM T_摘要M
WHERE
    T_摘要M.摘要名 like '%B%'
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No")

            Dim ans3 = query.Compile(New With {.sel = True, .txZyCd = "", .txZyNm = "", .txZyNo = "C"})
            Assert.Equal(ans3, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
FROM T_摘要M
WHERE
    T_摘要M.摘要No like '%C%'
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No")

            Dim ans4 = query.Compile(New With {.sel = True, .txZyCd = "A", .txZyNm = "B", .txZyNo = ""})
            Assert.Equal(ans4, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
FROM T_摘要M
WHERE
    T_摘要M.担当者CD like '%A%' AND
    T_摘要M.摘要名 like '%B%'
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No")

            Dim ans5 = query.Compile(New With {.sel = True, .txZyCd = "A", .txZyNm = "", .txZyNo = "C"})
            Assert.Equal(ans5, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
FROM T_摘要M
WHERE
    T_摘要M.担当者CD like '%A%' AND
    T_摘要M.摘要No like '%C%'
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No")

            Dim ans6 = query.Compile(New With {.sel = True, .txZyCd = "", .txZyNm = "B", .txZyNo = "C"})
            Assert.Equal(ans6, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
FROM T_摘要M
WHERE
    T_摘要M.摘要名 like '%B%' AND
    T_摘要M.摘要No like '%C%'
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No")

            Dim ans7 = query.Compile(New With {.sel = True, .txZyCd = "A", .txZyNm = "B", .txZyNo = "C"})
            Assert.Equal(ans7, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
FROM T_摘要M
WHERE
    T_摘要M.担当者CD like '%A%' AND
    T_摘要M.摘要名 like '%B%' AND
    T_摘要M.摘要No like '%C%'
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No")

            Dim ans8 = query.Compile(New With {.sel = True, .txZyCd = "", .txZyNm = "", .txZyNo = ""})
            Assert.Equal(ans8, "SELECT
	T_摘要M.摘要CD,T_摘要M.摘要名,T_摘要M.摘要No
FROM T_摘要M
ORDER BY T_摘要M.摘要CD,T_摘要M.摘要No")
        End Sub

    End Class

End Namespace
