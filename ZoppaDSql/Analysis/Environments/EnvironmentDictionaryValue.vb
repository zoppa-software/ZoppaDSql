Option Strict On
Option Explicit On

Namespace Analysis.Environments

    ''' <summary>環境値情報（ディクショナリ）。</summary>
    Public NotInheritable Class EnvironmentDictionaryValue
        Implements IEnvironmentValue

        ' ローカル変数
        Private ReadOnly mVariants As New Dictionary(Of String, Object)

        ''' <summary>ローカル変数を消去します。</summary>
        Public Sub LocalVarClear() Implements IEnvironmentValue.LocalVarClear
            Me.mVariants.Clear()
        End Sub

        ''' <summary>ローカル変数を追加する。</summary>
        ''' <param name="name">変数名。</param>
        ''' <param name="value">変数値。</param>
        Public Sub AddVariant(name As String, value As Object) Implements IEnvironmentValue.AddVariant
            If Me.mVariants.ContainsKey(name) Then
                Me.mVariants(name) = value
            Else
                Me.mVariants.Add(name, value)
            End If
        End Sub

        ''' <summary>指定した名称の変数値を取得します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <returns>値。</returns>
        Public Function GetValue(name As String) As Object Implements IEnvironmentValue.GetValue
            If Me.mVariants.ContainsKey(name) Then
                Return Me.mVariants(name)
            Else
                Throw New DSqlAnalysisException($"指定した変数が定義されていません:{name}")
            End If
        End Function
    End Class

End Namespace
