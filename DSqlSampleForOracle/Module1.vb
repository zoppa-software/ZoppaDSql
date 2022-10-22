Imports Oracle.ManagedDataAccess.Client
Imports ZoppaDSql

Module Module1

    Sub Main()
        'ZoppaDSqlSetting.DefaultSqlParameterCheck = Sub(chk)
        '                                                Dim prm = TryCast(chk, Oracle.ManagedDataAccess.Client.OracleParameter)
        '                                                If prm?.OracleDbType = OracleDbType.Varchar2 Then
        '                                                    prm.OracleDbType = OracleDbType.NVarchar2
        '                                                End If
        '                                            End Sub

        Using ora As New OracleConnection()
            ora.ConnectionString = "接続文字列"
            ora.Open()

            Dim tbl = ora.
                SetParameterPrepix(PrefixType.Colon).
                SetParameterChecker(
                    Sub(chk)
                        Dim prm = TryCast(chk, Oracle.ManagedDataAccess.Client.OracleParameter)
                        If prm?.OracleDbType = OracleDbType.Varchar2 Then
                            prm.OracleDbType = OracleDbType.NVarchar2
                        End If
                    End Sub).
                ExecuteRecords(Of RFLVGROUP)(
                    "select * from GROUP where SYAIN_NO = :SyNo ",
                    New With {.SyNo = CType("105055", DbString)}
                )
        End Using
    End Sub

    Public Class RFLVGROUP

        Public ReadOnly v1 As String

        Public ReadOnly v2 As String

        Public ReadOnly v3 As String

        Public Sub New(p1 As String, p2 As String, p3 As String)
            Me.v1 = p1
            Me.v2 = p2
            Me.v3 = p3
        End Sub

    End Class

End Module
