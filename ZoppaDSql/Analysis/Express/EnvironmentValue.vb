﻿Option Strict On
Option Explicit On

Imports System.Reflection
Imports ZoppaDSql.Analysis.Tokens

Namespace Analysis.Express

    ''' <summary>環境値情報。</summary>
    Public NotInheritable Class EnvironmentValue

        ' パラメータの型
        Private ReadOnly mType As Type

        ' パラメータ実体
        Private ReadOnly mTarget As Object

        ' プロパティディクショナリ
        Private ReadOnly mPropDic As New Dictionary(Of String, PropertyInfo)

        ' ローカル変数
        Private ReadOnly mVariants As New Dictionary(Of String, Object)

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="target">パラメータ。</param>
        Public Sub New(target As Object)
            Me.mType = target?.GetType()
            Me.mTarget = target
        End Sub

        ''' <summary>ローカル変数を消去します。</summary>
        Public Sub LocalVarClear()
            Me.mVariants.Clear()
        End Sub

        ''' <summary>ローカル変数を追加する。</summary>
        ''' <param name="name">変数名。</param>
        ''' <param name="value">変数値。</param>
        Public Sub AddVariant(name As String, value As Object)
            If Me.mVariants.ContainsKey(name) Then
                Me.mVariants(name) = value
            Else
                Me.mVariants.Add(name, value)
            End If
        End Sub

        ''' <summary>指定した名称のプロパティから値を取得します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <returns>値。</returns>
        Public Function GetValue(name As String) As Object
            If Me.mVariants.ContainsKey(name) Then
                ' ローカル変数に存在しているため、それを返す
                Return Me.mVariants(name)
            Else
                If Not Me.mPropDic.ContainsKey(name) Then
                    ' プロパティを取得して値を返す
                    Dim prop = Me.mType?.GetProperty(name)
                    If prop IsNot Nothing Then
                        Me.mPropDic.Add(name, prop)
                        Return prop.GetValue(Me.mTarget)
                    Else
                        Throw New DSqlAnalysisException($"指定したプロパティが定義されていません:{name}")
                    End If
                Else
                    ' 保持しているプロパティ参照から値を返す
                    Return Me.mPropDic(name).GetValue(Me.mTarget)
                End If
            End If
        End Function

    End Class

End Namespace
