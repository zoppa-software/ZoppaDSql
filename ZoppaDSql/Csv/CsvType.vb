Option Strict On
Option Explicit On

Namespace Csv

    Public Interface ICsvType

        ReadOnly Property ColumnType As Type

        Function Convert(input As CsvItem) As Object

    End Interface

    Public Module CsvType

        Public ReadOnly CsvBinary As ICsvType = New BinaryConverter()

        Public ReadOnly CsvByte As ICsvType = New ByteConverter()

        Public ReadOnly CsvBoolean As ICsvType = New BooleanConverter()

        Public ReadOnly CsvDateTime As ICsvType = New DateConverter()

        Public ReadOnly CsvDecimal As ICsvType = New DecimalConverter()

        Public ReadOnly CsvDouble As ICsvType = New DoubleConverter()

        Public ReadOnly CsvShort As ICsvType = New ShortConverter()

        Public ReadOnly CsvInteger As ICsvType = New IntegerConverter()

        Public ReadOnly CsvLong As ICsvType = New LongConverter()

        Public ReadOnly CsvSingle As ICsvType = New SingleConverter()

        Public ReadOnly CsvString As ICsvType = New StringConverter()

        Public ReadOnly CsvTime As ICsvType = New TimeConverter()

        Private NotInheritable Class BinaryConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Byte())
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim inp = input.UnEscape
                Dim ans As New List(Of Byte)()

                Dim buf = New Char(1) {}
                Using sr As New IO.StringReader(If(inp.Length Mod 2 = 0, inp, "0" & inp))
                    Do While sr.Peek() <> -1
                        sr.Read(buf, 0, 2)

                        Dim b = 0
                        For i As Integer = 0 To 1
                            Select Case buf(i)
                                Case "0"c To "9"c
                                    b = b * 16 + (AscW(buf(i)) - AscW("0"c))
                                Case "a"c To "f"c
                                    b = b * 16 + 10 + (AscW(buf(i)) - AscW("a"c))
                                Case "A"c To "F"c
                                    b = b * 16 + 10 + (AscW(buf(i)) - AscW("A"c))
                            End Select
                        Next
                        ans.Add(CByte(b))
                    Loop
                End Using

                Return ans.ToArray()
            End Function
        End Class

        Private NotInheritable Class ByteConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Byte?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToByte(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class BooleanConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Boolean?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToBoolean(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class DateConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Date?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToDateTime(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class DecimalConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Decimal?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToDecimal(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class DoubleConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Double?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToDouble(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class ShortConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Short?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToInt16(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class IntegerConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Integer?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToInt32(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class LongConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Long?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToInt64(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class SingleConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(Single?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", System.Convert.ToSingle(s), Nothing)
            End Function
        End Class

        Private NotInheritable Class StringConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(String)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Return input.UnEscape
            End Function
        End Class

        Private NotInheritable Class TimeConverter
            Implements ICsvType

            Public ReadOnly Property ColumnType As Type Implements ICsvType.ColumnType
                Get
                    Return GetType(TimeSpan?)
                End Get
            End Property

            Public Function Convert(input As CsvItem) As Object Implements ICsvType.Convert
                Dim s = input.UnEscape
                Return If(s <> "", TimeSpan.Parse(s), Nothing)
            End Function
        End Class

    End Module

End Namespace
