Option Strict On
Option Explicit On

Imports System.Data.Common
Imports System.IO
Imports System.Linq.Expressions
Imports System.Text
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis.Tokens

Namespace Analysis

    Partial Module ParserAnalysis

        ''' <summary>Trimを評価します。</summary>
        ''' <param name="trimToken">Trimブロック。</param>
        ''' <param name="buffer">出力先バッファ。</param>
        ''' <param name="prm">環境値情報。</param>
        Private Sub EvaluationTrim(trimToken As TokenLink, buffer As StringBuilder, prm As EnvironmentValue)
            Dim answer As New StringBuilder()

            ' Trim内のトークンを評価
            Dim tmpbuf As New StringBuilder()
            Evaluation(trimToken.Children, tmpbuf, prm)

            ' 評価結果を再度トークン化
            Dim tokens = LexicalAnalysis.SplitTrimToken(tmpbuf.ToString())
            Dim tknPtr = New TokenPtr(tokens)

            '
            Do While tknPtr.HasNext
                If tknPtr.Current.TokenName = NameOf(SpaceToken) Then
                    answer.Append(tknPtr.Current.Contents)
                    tknPtr.Move(1)
                Else
                    Exit Do
                End If
            Loop

            Dim keyToken = tknPtr.Current
            Dim keyStr = keyToken.Contents.ToString()
            If keyStr.ToLower() = "where" Then
                tknPtr.Move(1)

                Dim logicalChecker As New LogicalChecker()
                Dim facChecker As New FactorChecker()
                Dim parenChecker As New ParenChecker()
                logicalChecker.NextChecker = facChecker
                facChecker.NextChecker = parenChecker
                parenChecker.NextChecker = logicalChecker

                If logicalChecker.Check(tknPtr, answer) Then
                    answer.Insert(0, keyStr)
                End If
            ElseIf keyStr.ToLower() = "set" Then
                tknPtr.Move(1)
                Dim cmaChecker As New CommaChecker()
                Dim facChecker As New FactorChecker()
                cmaChecker.NextChecker = facChecker

                If cmaChecker.Check(tknPtr, answer) Then
                    answer.Insert(0, keyStr)
                End If
            Else
                Dim logicalChecker As New LogicalChecker()
                Dim facChecker As New FactorChecker()
                Dim parenChecker As New ParenChecker()
                logicalChecker.NextChecker = facChecker
                facChecker.NextChecker = parenChecker
                parenChecker.NextChecker = logicalChecker
                logicalChecker.Check(tknPtr, answer)
            End If

            buffer.Append(answer.ToString())
        End Sub

        ''' <summary>括弧内部式をチェックします（トリム用）</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <param name="leftToken">(トークン。</param>
        ''' <param name="nxtChecker">次のチェッカー。</param>
        ''' <param name="answer">トリム結果</param>
        ''' <returns>括弧内部式。</returns>
        Private Function CheckTrimParen(reader As TokenPtr, leftToken As TokenPoint, nxtChecker As ITrimChecker, answer As StringBuilder) As Boolean
            Dim tmp As New List(Of TokenPoint)()
            Dim lv As Integer = 0
            Dim rightToken As TokenPoint
            Do While reader.HasNext
                Dim tkn = reader.Current
                reader.Move(1)

                Select Case tkn.TokenName
                    Case NameOf(RParenToken)
                        If lv > 0 Then
                            lv -= 1
                        Else
                            rightToken = tkn
                            Exit Do
                        End If
                    Case Else
                        tmp.Add(tkn)
                End Select
            Loop

            Dim buf As New StringBuilder()
            If nxtChecker.Check(New TokenPtr(tmp), buf) Then
                answer.Append("(")
                answer.Append(buf.ToString())
                answer.Append(")")
                Return True
            Else
                Return (leftToken.Position = rightToken.Position - 1)
            End If
        End Function

        ''' <summary>Trimチェックインターフェイス。</summary>
        Private Interface ITrimChecker

            ''' <summary>チェックする。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <param name="answer">トリム結果。</param>
            ''' <returns>チェック結果。</returns>
            Function Check(reader As TokenPtr, answer As StringBuilder) As Boolean

        End Interface

        ''' <summary>括弧チェック。</summary>
        Private NotInheritable Class ParenChecker
            Implements ITrimChecker

            ''' <summary>次のチェッカーを設定、取得する。</summary>
            Friend Property NextChecker() As ITrimChecker

            ''' <summary>チェックする。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <param name="answer">トリム結果。</param>
            ''' <returns>チェック結果。</returns>
            Public Function Check(reader As TokenPtr, answer As StringBuilder) As Boolean Implements ITrimChecker.Check
                Dim tkn = reader.Current
                If tkn.TokenName = NameOf(LParenToken) Then
                    reader.Move(1)
                    Return CheckTrimParen(reader, tkn, Me.NextChecker, answer)
                Else
                    Return Me.NextChecker.Check(reader, answer)
                End If
            End Function

        End Class

        ''' <summary>論理チェック。</summary>
        Private NotInheritable Class LogicalChecker
            Implements ITrimChecker

            ''' <summary>次のチェッカーを設定、取得する。</summary>
            Friend Property NextChecker() As ITrimChecker

            ''' <summary>チェックする。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <param name="answer">トリム結果。</param>
            ''' <returns>チェック結果。</returns>
            Public Function Check(reader As TokenPtr, answer As StringBuilder) As Boolean Implements ITrimChecker.Check
                Dim bufleft As New StringBuilder()
                Dim extleft = Me.NextChecker.Check(reader, bufleft)

                Do While reader.HasNext
                    Dim ope = reader.Current
                    Select Case ope.TokenName
                        Case NameOf(AndOrToken)
                            reader.Move(1)

                            Dim bufright As New StringBuilder()
                            Dim extright = Me.NextChecker.Check(reader, bufright)
                            If extleft AndAlso extright Then
                                bufleft.Append(ope.Contents)
                            End If
                            bufleft.Append(bufright.ToString())
                            extleft = True

                        Case Else
                            Exit Do
                    End Select
                Loop

                answer.Append(bufleft.ToString())
                Return (bufleft.Length > 0)
            End Function

        End Class

        ''' <summary>カンマチェック。</summary>
        Private NotInheritable Class CommaChecker
            Implements ITrimChecker

            ''' <summary>次のチェッカーを設定、取得する。</summary>
            Friend Property NextChecker() As ITrimChecker

            ''' <summary>チェックする。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <param name="answer">トリム結果。</param>
            ''' <returns>チェック結果。</returns>
            Public Function Check(reader As TokenPtr, answer As StringBuilder) As Boolean Implements ITrimChecker.Check
                Dim bufleft As New StringBuilder()
                Dim extleft = Me.NextChecker.Check(reader, bufleft)

                Do While reader.HasNext
                    Dim ope = reader.Current
                    Select Case ope.TokenName
                        Case NameOf(CommaToken)
                            reader.Move(1)

                            Dim bufright As New StringBuilder()
                            Dim extright = Me.NextChecker.Check(reader, bufright)
                            If extleft AndAlso extright Then
                                bufleft.Append(",")
                            End If
                            bufleft.Append(bufright.ToString())
                            extleft = True

                        Case Else
                            Exit Do
                    End Select
                Loop

                answer.Append(bufleft.ToString())
                Return (bufleft.Length > 0)
            End Function

        End Class

        ''' <summary>要素チェック。</summary>
        Private NotInheritable Class FactorChecker
            Implements ITrimChecker

            ''' <summary>次のチェッカーを設定、取得する。</summary>
            Friend Property NextChecker() As ITrimChecker

            ''' <summary>チェックする。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>チェック結果。</returns>
            Public Function Check(reader As TokenPtr, answer As StringBuilder) As Boolean Implements ITrimChecker.Check
                Dim buf As New StringBuilder()

                If reader.HasNext Then
                    Dim res = False

                    Do While reader.HasNext
                        Dim ope = reader.Current
                        Select Case ope.TokenName
                            Case NameOf(LParenToken)
                                reader.Move(1)
                                If Me.NextChecker IsNot Nothing Then
                                    res = res Or CheckTrimParen(reader, ope, Me.NextChecker, buf)
                                Else
                                    buf.Append("(")
                                    res = True
                                End If

                            Case NameOf(RParenToken)
                                reader.Move(1)
                                buf.Append(")")
                                res = True

                            Case NameOf(AndOrToken)
                                Exit Do

                            Case NameOf(SpaceToken)
                                buf.Append(ope.Contents)
                                reader.Move(1)

                            Case Else
                                buf.Append(ope.Contents)
                                reader.Move(1)
                                res = True
                        End Select
                    Loop
                    If res Then answer.Append(buf.ToString())
                    Return res
                Else
                    Return False
                End If
            End Function
        End Class

    End Module

End Namespace