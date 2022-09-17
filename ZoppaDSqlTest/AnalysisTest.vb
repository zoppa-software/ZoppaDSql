Imports System
Imports ZoppaDSql
Imports Xunit
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis
Imports ZoppaDSql.Analysis.Environments

Namespace ZoppaDSqlTest
    Public Class AnalysisTest

        <Fact>
        Sub ParentTest()
            Dim ans1 = "(28 - 3) / (2 + 3)".Executes().Contents
            Assert.Equal(ans1, 5)

            Dim ans2 = "(28 - 3)".Executes().Contents
            Assert.Equal(ans2, 25)
        End Sub

        <Fact>
        Sub ExpressionTest()
            Dim a301 = "100 = 100".Executes().Contents
            Assert.Equal(a301, True)
            Dim a302 = "100 = 200".Executes().Contents
            Assert.Equal(a302, False)
            Dim a303 = "'abc' = 'abc'".Executes().Contents
            Assert.Equal(a303, True)
            Dim a304 = "'abc' = 'def'".Executes().Contents
            Assert.Equal(a304, False)

            Dim a401 = "'abd' >= 'abc'".Executes().Contents
            Assert.Equal(a401, True)
            Dim a402 = "'abc' >= 'abd'".Executes().Contents
            Assert.Equal(a402, False)
            Dim a403 = "'abc' >= 'abc'".Executes().Contents
            Assert.Equal(a403, True)
            Dim a404 = "0.1+0.1+0.1 >= 0.4".Executes().Contents
            Assert.Equal(a404, False)
            Dim a405 = "0.1*5 >= 0.4".Executes().Contents
            Assert.Equal(a405, True)
            Dim a406 = "0.1*3 >= 0.3".Executes().Contents
            Assert.Equal(a406, True)

            Dim a501 = "'abd' > 'abc'".Executes().Contents
            Assert.Equal(a501, True)
            Dim a502 = "'abc' > 'abd'".Executes().Contents
            Assert.Equal(a502, False)
            Dim a503 = "'abc' > 'abc'".Executes().Contents
            Assert.Equal(a503, False)
            Dim a504 = "0.1+0.1+0.1 > 0.4".Executes().Contents
            Assert.Equal(a504, False)
            Dim a505 = "0.1*5 > 0.4".Executes().Contents
            Assert.Equal(a505, True)
            Dim a506 = "0.1*3 > 0.3".Executes().Contents
            Assert.Equal(a506, False)

            Dim a601 = "'abd' <= 'abc'".Executes().Contents
            Assert.Equal(a601, False)
            Dim a602 = "'abc' <= 'abd'".Executes().Contents
            Assert.Equal(a602, True)
            Dim a603 = "'abc' <= 'abc'".Executes().Contents
            Assert.Equal(a603, True)
            Dim a604 = "0.1+0.1+0.1 <= 0.4".Executes().Contents
            Assert.Equal(a604, True)
            Dim a605 = "0.1*5 <= 0.4".Executes().Contents
            Assert.Equal(a605, False)
            Dim a606 = "0.1*3 <= 0.3".Executes().Contents
            Assert.Equal(a606, True)

            Dim a701 = "'abd' < 'abc'".Executes().Contents
            Assert.Equal(a701, False)
            Dim a702 = "'abc' < 'abd'".Executes().Contents
            Assert.Equal(a702, True)
            Dim a703 = "'abc' < 'abc'".Executes().Contents
            Assert.Equal(a703, False)
            Dim a704 = "0.1+0.1+0.1 < 0.4".Executes().Contents
            Assert.Equal(a704, True)
            Dim a705 = "0.1*5 < 0.4".Executes().Contents
            Assert.Equal(a705, False)
            Dim a706 = "0.1*3 < 0.3".Executes().Contents
            Assert.Equal(a706, False)

            Dim a801 = "100 <> 100".Executes().Contents
            Assert.Equal(a801, False)
            Dim a802 = "100 <> 200".Executes().Contents
            Assert.Equal(a802, True)
            Dim a803 = "'abc' <> 'abc'".Executes().Contents
            Assert.Equal(a803, False)
            Dim a804 = "'abc' <> 'def'".Executes().Contents
            Assert.Equal(a804, True)
        End Sub

        <Fact>
        Sub Expression2Test()
            Dim a101 = "100 - 99.9".Executes().Contents
            Assert.True(Math.Abs(a101 - 0.1) < 0.0000000000001)
            Dim a102 = "46 - 12".Executes().Contents
            Assert.Equal(a102, 34)
            Dim a103 = "-2 - -2".Executes().Contents
            Assert.Equal(a103, 0)

            Dim a201 = "0.1 * 6".Executes().Contents
            Assert.True(Math.Abs(a201 - 0.6) < 0.0000000000001)
            Dim a202 = "26 * 2".Executes().Contents
            Assert.Equal(a202, 52)
            Dim a203 = "-2 * -2".Executes().Contents
            Assert.Equal(a203, 4)

            Dim a301 = "100 / 25".Executes().Contents
            Assert.Equal(a301, 4)

            Dim a401 = "'桃' + '太郎'".Executes().Contents
            Assert.Equal(a401, "桃太郎")
        End Sub

        <Fact>
        Sub BinaryTest()
            Dim a101 = "true and true".Executes().Contents
            Assert.Equal(a101, True)
            Dim a102 = "true and false".Executes().Contents
            Assert.Equal(a102, False)
            Dim a103 = "false and true".Executes().Contents
            Assert.Equal(a103, False)
            Dim a104 = "false and false".Executes().Contents
            Assert.Equal(a104, False)

            Dim a201 = "true or true".Executes().Contents
            Assert.Equal(a201, True)
            Dim a202 = "true or false".Executes().Contents
            Assert.Equal(a202, True)
            Dim a203 = "false or true".Executes().Contents
            Assert.Equal(a203, True)
            Dim a204 = "false or false".Executes().Contents
            Assert.Equal(a204, False)
        End Sub

        <Fact>
        Sub ExceptionTest()
            Try
                Dim a0 = "1 / 0".Executes()
                Assert.True(False, "0割例外が発生しない")
            Catch ex As DivideByZeroException

            End Try

            Try
                Dim a1 = "123 and '123'".Executes()
                Assert.True(False, "論理積エラーが発生しない")
            Catch ex As DSqlAnalysisException

            End Try

            Try
                Dim a1 = "123 or '123'".Executes()
                Assert.True(False, "論理和エラーが発生しない")
            Catch ex As DSqlAnalysisException

            End Try

            Try
                Dim a3 = "'abc' - 99".Executes()
                Assert.True(False, "減算エラーが発生しない")
            Catch ex As DSqlAnalysisException

            End Try
        End Sub

        <Fact>
        Sub EnvironmentTest()
            Dim env As New EnvironmentObjectValue(New With {.prm1 = 100, .prm2 = "abc"})
            Assert.Equal(env.GetValue("prm1"), 100)
            Assert.Equal(env.GetValue("prm2"), "abc")

            Assert.Equal(env.GetValue("prm2"), "abc")

            env.AddVariant("prm3", "def")
            Assert.Equal(env.GetValue("prm3"), "def")

            Try
                env.LocalVarClear()
                Assert.Equal(env.GetValue("prm3"), "def")
                Assert.True(False, "論理積エラーが発生しない")
            Catch ex As DSqlAnalysisException

            End Try
        End Sub

    End Class
End Namespace

