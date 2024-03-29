﻿Imports System
Imports ZoppaDSql
Imports Xunit
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis
Imports System.Runtime.InteropServices
Imports ZoppaDSql.Analysis.Tokens

Namespace ZoppaDSqlTest

    Public Class CalcTest

        <Fact>
        Public Sub SignChangeTest()
            Dim op1 As New NumberToken.NumberIntegerValueToken(12)
            Dim oh_op1 = op1.SignChange()
            Assert.Equal(oh_op1.Contents, CInt(-12))

            Dim op2 As New NumberToken.NumberLongValueToken(34)
            Dim oh_op2 = op2.SignChange()
            Assert.Equal(oh_op2.Contents, CLng(-34))

            Dim op3 As New NumberToken.NumberDecimalValueToken(56.7)
            Dim oh_op3 = op3.SignChange()
            Assert.Equal(oh_op3.Contents, New Decimal(-56.7))

            Dim op4 As New NumberToken.NumberDoubleValueToken(8.9)
            Dim oh_op4 = op4.SignChange()
            Assert.Equal(oh_op4.Contents, -8.9)
        End Sub

        <Fact>
        Public Sub PlusComputationTest()
            Dim op1 As New NumberToken.NumberIntegerValueToken(12)
            Dim op1_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op1_2 As New NumberToken.NumberLongValueToken(2)
            Dim op1_3 As New NumberToken.NumberDecimalValueToken(3.45)
            Dim op1_4 As New NumberToken.NumberDoubleValueToken(6.78)
            Assert.True(op1.PlusComputation(op1_1).Contents = 13)
            Assert.True(op1.PlusComputation(op1_2).Contents = 14)
            Assert.True(op1.PlusComputation(op1_3).Contents = 15.45)
            Assert.True(op1.PlusComputation(op1_4).Contents = 18.78)

            Dim op2 As New NumberToken.NumberLongValueToken(12)
            Dim op2_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op2_2 As New NumberToken.NumberLongValueToken(2)
            Dim op2_3 As New NumberToken.NumberDecimalValueToken(3.45)
            Dim op2_4 As New NumberToken.NumberDoubleValueToken(6.78)
            Assert.True(op2.PlusComputation(op2_1).Contents = 13)
            Assert.True(op2.PlusComputation(op2_2).Contents = 14)
            Assert.True(op2.PlusComputation(op2_3).Contents = 15.45)
            Assert.True(op2.PlusComputation(op2_4).Contents = 18.78)

            Dim op3 As New NumberToken.NumberDecimalValueToken(12)
            Dim op3_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op3_2 As New NumberToken.NumberLongValueToken(2)
            Dim op3_3 As New NumberToken.NumberDecimalValueToken(3.45)
            Dim op3_4 As New NumberToken.NumberDoubleValueToken(6.78)
            Assert.True(op3.PlusComputation(op3_1).Contents = 13)
            Assert.True(op3.PlusComputation(op3_2).Contents = 14)
            Assert.True(op3.PlusComputation(op3_3).Contents = 15.45)
            Assert.True(op3.PlusComputation(op3_4).Contents = 18.78)

            Dim op4 As New NumberToken.NumberDoubleValueToken(12)
            Dim op4_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op4_2 As New NumberToken.NumberLongValueToken(2)
            Dim op4_3 As New NumberToken.NumberDecimalValueToken(3.45)
            Dim op4_4 As New NumberToken.NumberDoubleValueToken(6.78)
            Assert.True(op4.PlusComputation(op4_1).Contents = 13)
            Assert.True(op4.PlusComputation(op4_2).Contents = 14)
            Assert.True(op4.PlusComputation(op4_3).Contents = 15.45)
            Assert.True(op4.PlusComputation(op4_4).Contents = 18.78)
        End Sub

        <Fact>
        Public Sub MinusComputationTest()
            Dim op1 As New NumberToken.NumberIntegerValueToken(10)
            Dim op1_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op1_2 As New NumberToken.NumberLongValueToken(2)
            Dim op1_3 As New NumberToken.NumberDecimalValueToken(3.4)
            Dim op1_4 As New NumberToken.NumberDoubleValueToken(5.6)
            Assert.Equal(op1.MinusComputation(op1_1).Contents, 9)
            Assert.Equal(op1.MinusComputation(op1_2).Contents, 8)
            Assert.Equal(op1.MinusComputation(op1_3).Contents, 6.6)
            Assert.Equal(op1.MinusComputation(op1_4).Contents, 4.4)

            Dim op2 As New NumberToken.NumberLongValueToken(10)
            Dim op2_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op2_2 As New NumberToken.NumberLongValueToken(2)
            Dim op2_3 As New NumberToken.NumberDecimalValueToken(3.4)
            Dim op2_4 As New NumberToken.NumberDoubleValueToken(5.6)
            Assert.Equal(op2.MinusComputation(op2_1).Contents, 9)
            Assert.Equal(op2.MinusComputation(op2_2).Contents, 8)
            Assert.Equal(op2.MinusComputation(op2_3).Contents, 6.6)
            Assert.Equal(op2.MinusComputation(op2_4).Contents, 4.4)

            Dim op3 As New NumberToken.NumberDecimalValueToken(10)
            Dim op3_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op3_2 As New NumberToken.NumberLongValueToken(2)
            Dim op3_3 As New NumberToken.NumberDecimalValueToken(3.4)
            Dim op3_4 As New NumberToken.NumberDoubleValueToken(5.6)
            Assert.Equal(op3.MinusComputation(op3_1).Contents, 9)
            Assert.Equal(op3.MinusComputation(op3_2).Contents, 8)
            Assert.Equal(op3.MinusComputation(op3_3).Contents, 6.6)
            Assert.Equal(op3.MinusComputation(op3_4).Contents, 4.4)

            Dim op4 As New NumberToken.NumberDoubleValueToken(10)
            Dim op4_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op4_2 As New NumberToken.NumberLongValueToken(2)
            Dim op4_3 As New NumberToken.NumberDecimalValueToken(3.4)
            Dim op4_4 As New NumberToken.NumberDoubleValueToken(5.6)
            Assert.Equal(op4.MinusComputation(op4_1).Contents, 9)
            Assert.Equal(op4.MinusComputation(op4_2).Contents, 8)
            Assert.Equal(op4.MinusComputation(op4_3).Contents, 6.6)
            Assert.Equal(op4.MinusComputation(op4_4).Contents, 4.4)
        End Sub

        <Fact>
        Public Sub MultiComputationTest()
            Dim op1 As New NumberToken.NumberIntegerValueToken(10)
            Dim op1_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op1_2 As New NumberToken.NumberLongValueToken(2)
            Dim op1_3 As New NumberToken.NumberDecimalValueToken(3.4)
            Dim op1_4 As New NumberToken.NumberDoubleValueToken(5.6)
            Assert.Equal(op1.MultiComputation(op1_1).Contents, 10)
            Assert.Equal(op1.MultiComputation(op1_2).Contents, 20)
            Assert.True(op1.MultiComputation(op1_3).Contents = 34.0)
            Assert.Equal(op1.MultiComputation(op1_4).Contents, 56.0)

            Dim op2 As New NumberToken.NumberLongValueToken(10)
            Dim op2_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op2_2 As New NumberToken.NumberLongValueToken(2)
            Dim op2_3 As New NumberToken.NumberDecimalValueToken(3.4)
            Dim op2_4 As New NumberToken.NumberDoubleValueToken(5.6)
            Assert.Equal(op2.MultiComputation(op2_1).Contents, 10)
            Assert.Equal(op2.MultiComputation(op2_2).Contents, 20)
            Assert.True(op2.MultiComputation(op2_3).Contents = 34.0)
            Assert.Equal(op2.MultiComputation(op2_4).Contents, 56.0)

            Dim op3 As New NumberToken.NumberDecimalValueToken(10)
            Dim op3_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op3_2 As New NumberToken.NumberLongValueToken(2)
            Dim op3_3 As New NumberToken.NumberDecimalValueToken(3.4)
            Dim op3_4 As New NumberToken.NumberDoubleValueToken(5.6)
            Assert.True(op3.MultiComputation(op3_1).Contents = 10)
            Assert.True(op3.MultiComputation(op3_2).Contents = 20)
            Assert.True(op3.MultiComputation(op3_3).Contents = 34.0)
            Assert.True(op3.MultiComputation(op3_4).Contents = 56.0)

            Dim op4 As New NumberToken.NumberDoubleValueToken(10)
            Dim op4_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op4_2 As New NumberToken.NumberLongValueToken(2)
            Dim op4_3 As New NumberToken.NumberDecimalValueToken(3.4)
            Dim op4_4 As New NumberToken.NumberDoubleValueToken(5.6)
            Assert.Equal(op4.MultiComputation(op4_1).Contents, 10)
            Assert.Equal(op4.MultiComputation(op4_2).Contents, 20)
            Assert.True(op4.MultiComputation(op4_3).Contents = 34.0)
            Assert.Equal(op4.MultiComputation(op4_4).Contents, 56.0)
        End Sub

        <Fact>
        Public Sub DivComputationTest()
            Dim op1 As New NumberToken.NumberIntegerValueToken(64)
            Dim op1_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op1_2 As New NumberToken.NumberLongValueToken(2)
            Dim op1_3 As New NumberToken.NumberDecimalValueToken(3.2)
            Dim op1_4 As New NumberToken.NumberDoubleValueToken(1.6)
            Assert.Equal(op1.DivComputation(op1_1).Contents, 64)
            Assert.Equal(op1.DivComputation(op1_2).Contents, 32)
            Assert.True(op1.DivComputation(op1_3).Contents = 20)
            Assert.Equal(op1.DivComputation(op1_4).Contents, 40)

            Dim op2 As New NumberToken.NumberLongValueToken(64)
            Dim op2_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op2_2 As New NumberToken.NumberLongValueToken(2)
            Dim op2_3 As New NumberToken.NumberDecimalValueToken(3.2)
            Dim op2_4 As New NumberToken.NumberDoubleValueToken(1.6)
            Assert.Equal(op2.DivComputation(op2_1).Contents, 64)
            Assert.Equal(op2.DivComputation(op2_2).Contents, 32)
            Assert.True(op2.DivComputation(op2_3).Contents = 20)
            Assert.Equal(op2.DivComputation(op2_4).Contents, 40)

            Dim op3 As New NumberToken.NumberDecimalValueToken(64)
            Dim op3_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op3_2 As New NumberToken.NumberLongValueToken(2)
            Dim op3_3 As New NumberToken.NumberDecimalValueToken(3.2)
            Dim op3_4 As New NumberToken.NumberDoubleValueToken(1.6)
            Assert.True(op3.DivComputation(op3_1).Contents = 64)
            Assert.True(op3.DivComputation(op3_2).Contents = 32)
            Assert.True(op3.DivComputation(op3_3).Contents = 20)
            Assert.True(op3.DivComputation(op3_4).Contents = 40)

            Dim op4 As New NumberToken.NumberDoubleValueToken(64)
            Dim op4_1 As New NumberToken.NumberIntegerValueToken(1)
            Dim op4_2 As New NumberToken.NumberLongValueToken(2)
            Dim op4_3 As New NumberToken.NumberDecimalValueToken(3.2)
            Dim op4_4 As New NumberToken.NumberDoubleValueToken(1.6)
            Assert.Equal(op4.DivComputation(op4_1).Contents, 64)
            Assert.Equal(op4.DivComputation(op4_2).Contents, 32)
            Assert.True(op4.DivComputation(op4_3).Contents = 20)
            Assert.Equal(op4.DivComputation(op4_4).Contents, 40)
        End Sub

    End Class

End Namespace
