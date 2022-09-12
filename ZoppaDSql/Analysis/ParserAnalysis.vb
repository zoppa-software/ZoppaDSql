Option Strict On
Option Explicit On

Imports System.Data.Common
Imports System.IO
Imports System.Linq.Expressions
Imports System.Text
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis.Tokens

Namespace Analysis

    ''' <summary>トークンを解析する。</summary>
    Friend Module ParserAnalysis

        ''' <summary>動的SQLの置き換えを実行します。</summary>
        ''' <param name="sqlQuery">動的SQL。</param>
        ''' <param name="parameter">パラメータ。</param>
        ''' <returns>置き換え結果。</returns>
        Public Function Replase(sqlQuery As String, Optional parameter As Object = Nothing) As String
            ' トークン配列に変換
            Dim tokens = LexicalAnalysis.Compile(sqlQuery)
            LoggingDebug("tokens")
            For i As Integer = 0 To tokens.Count - 1
                LoggingDebug($"{i + 1} : {tokens(i)}")
            Next

            ' 置き換え式を階層化する
            Dim hierarchy = CreateHierarchy(tokens)

            ' 置き換え式を解析して出力文字列を取得します。
            Dim buffer As New StringBuilder()
            Dim values As New EnvironmentValue(parameter)
            Evaluation(hierarchy.Children, buffer, values)

            ' 空白行を取り除いて返す
            Dim ans As New StringBuilder(buffer.Length)
            Using sr As New StringReader(buffer.ToString())
                Do While sr.Peek() <> -1
                    Dim ln = sr.ReadLine()
                    If ln.Trim() <> "" Then ans.AppendLine(ln)
                Loop
            End Using
            Return ans.ToString().Trim()
        End Function

        ''' <summary>置き換え式の階層を作成します。</summary>
        ''' <param name="src">トークンリスト。</param>
        ''' <returns>階層リンク。</returns>
        Public Function CreateHierarchy(src As List(Of TokenPoint)) As TokenLink
            Dim root As New TokenLink(Nothing)
            If src?.Count > 0 Then
                Dim tmp As New List(Of TokenPoint)(src)
                tmp.Reverse()

                CreateHierarchy(root, tmp)
            End If
            Return root
        End Function

        ''' <summary>置き換え式の階層を作成します。</summary>
        ''' <param name="node">階層ノード。</param>
        ''' <param name="tokens">対象トークンリスト。</param>
        ''' <param name="partitionTokens"></param>
        Private Sub CreateHierarchy(node As TokenLink,
                                           tokens As List(Of TokenPoint),
                                           ParamArray partitionTokens As String())
            Dim limit As New HashSet(Of String)()
            For Each tkn In partitionTokens
                limit.Add(tkn)
            Next

            Do While tokens.Count > 0
                Dim tkn = tokens(tokens.Count - 1)
                If limit.Contains(tkn.TokenName) Then
                    ' 末端トークンなら階層終了
                    Return
                Else
                    Dim cnode = node.AddChild(tkn)
                    Select Case tkn.TokenName
                        Case NameOf(IfToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy(cnode, tokens, NameOf(ElseIfToken), NameOf(ElseToken), NameOf(EndIfToken))

                        Case NameOf(ElseIfToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy(cnode, tokens, NameOf(ElseToken), NameOf(EndIfToken))

                        Case NameOf(ElseToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy(cnode, tokens, NameOf(EndIfToken))

                        Case NameOf(ForEachToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy(cnode, tokens, NameOf(EndForToken))

                        Case NameOf(TrimToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy(cnode, tokens, NameOf(EndTrimToken))

                        Case Else
                            tokens.RemoveAt(tokens.Count - 1)
                    End Select
                End If
            Loop
        End Sub

        ''' <summary>トークンリストを評価します。</summary>
        ''' <param name="children">階層情報。</param>
        ''' <param name="buffer">出力先バッファ。</param>
        ''' <param name="prm">環境値情報。</param>
        Private Sub Evaluation(children As List(Of TokenLink), buffer As StringBuilder, prm As EnvironmentValue)
            Dim tmp As New List(Of TokenLink)(children)
            tmp.Reverse()

            Do While tmp.Count > 0
                Dim chd = tmp(tmp.Count - 1)

                Select Case chd.TargetToken.TokenName
                    Case NameOf(IfToken)
                        EvaluationIf(tmp, buffer, prm)

                    Case NameOf(ForEachToken)
                        EvaluationFor(chd, buffer, prm)
                        tmp.RemoveAt(tmp.Count - 1)

                    Case NameOf(TrimToken)
                        EvaluationTrim(chd, buffer, prm)
                        tmp.RemoveAt(tmp.Count - 1)

                    Case NameOf(ReplaseToken)
                        Dim rtoken = chd.TargetToken.GetToken(Of ReplaseToken)()
                        If rtoken IsNot Nothing Then
                            Dim ans = prm.GetValue(If(rtoken.Contents?.ToString(), ""))
                            If TypeOf ans Is String AndAlso rtoken.IsEscape Then
                                Dim s = ans.ToString()
                                s = s.Replace("'"c, "''")
                                s = s.Replace("\"c, "\\")
                                ans = $"'{s}'"
                            End If
                            buffer.Append(ans)
                        End If
                        tmp.RemoveAt(tmp.Count - 1)

                    Case NameOf(EndIfToken), NameOf(EndForToken), NameOf(EndTrimToken)
                        tmp.RemoveAt(tmp.Count - 1)

                    Case Else
                        buffer.Append(chd.TargetToken.Contents)
                        tmp.RemoveAt(tmp.Count - 1)
                End Select
            Loop
        End Sub

        ''' <summary>Ifを評価します。</summary>
        ''' <param name="tmp">階層情報。</param>
        ''' <param name="buffer">出力先バッファ。</param>
        ''' <param name="prm">環境値情報。</param>
        Private Sub EvaluationIf(tmp As List(Of TokenLink), buffer As StringBuilder, prm As EnvironmentValue)
            ' If、ElseIf、Elseブロックを集める
            Dim blocks As New List(Of TokenLink)()
            For i As Integer = tmp.Count - 1 To 0 Step -1
                Dim iftkn = tmp(i)
                tmp.RemoveAt(tmp.Count - 1)
                If iftkn.TargetToken.TokenName <> NameOf(EndIfToken) Then
                    blocks.Add(iftkn)
                Else
                    Exit For
                End If
            Next

            ' If、ElseIf、Elseブロックを評価
            For Each iftkn In blocks
                Select Case iftkn.TargetToken.TokenName
                    Case NameOf(IfToken), NameOf(ElseIfToken)
                        ' 条件を評価して真ならば、ブロックを出力
                        Dim ifans = Executes(iftkn.TargetToken.GetCommandToken().CommandTokens, prm)
                        If TypeOf ifans.Contents Is Boolean AndAlso CBool(ifans.Contents) Then
                            Evaluation(iftkn.Children, buffer, prm)
                            Return
                        End If

                    Case NameOf(ElseToken)
                        Evaluation(iftkn.Children, buffer, prm)
                        Return
                End Select
            Next
            buffer.Append(" ")
        End Sub

        ''' <summary>Foreachを評価します。</summary>
        ''' <param name="fortoken">Foreachブロック。</param>
        ''' <param name="buffer">出力先バッファ。</param>
        ''' <param name="prm">環境値情報。</param>
        Private Sub EvaluationFor(forToken As TokenLink, buffer As StringBuilder, prm As EnvironmentValue)
            Dim forTkn = forToken.TargetToken.GetToken(Of ForEachToken)()

            ' カウンタ変数
            Dim valkey As String = ""

            ' ループ元コレクション
            Dim collection As IEnumerable = Nothing

            ' 構文を解析して変数、ループ元コレクションを取得
            With forTkn
                If .CommandTokens.Count = 3 AndAlso
                   .CommandTokens(0).TokenName = NameOf(IdentToken) AndAlso
                   .CommandTokens(1).TokenName = NameOf(InToken) AndAlso
                   .CommandTokens(2).TokenName = NameOf(IdentToken) Then
                    Dim colln = prm.GetValue(If(.CommandTokens(2).Contents?.ToString(), ""))
                    If TypeOf colln Is IEnumerable Then
                        valkey = If(.CommandTokens(0).Contents?.ToString(), "")
                        collection = CType(colln, IEnumerable)
                    End If
                End If
            End With

            ' Foreachして出力
            For Each v In collection
                prm.AddVariant(valkey, v)
                Evaluation(forToken.Children, buffer, prm)
            Next
            prm.LocalVarClear()
        End Sub

        ''' <summary>式を解析して結果を取得します。</summary>
        ''' <param name="expression">式文字列。</param>
        ''' <param name="parameter">パラメータ。</param>
        ''' <returns>解析結果。</returns>
        Public Function Executes(expression As String, Optional parameter As Object = Nothing) As IToken
            ' トークン配列に変換
            Dim tokens = LexicalAnalysis.SimpleCompile(expression)

            ' 式木を実行
            Return Executes(tokens, New EnvironmentValue(parameter))
        End Function

        ''' <summary>式を解析して結果を取得します。</summary>
        ''' <param name="tokens">対象トークン。。</param>
        ''' <param name="parameter">パラメータ。</param>
        ''' <returns>解析結果。</returns>
        Public Function Executes(tokens As List(Of TokenPoint), parameter As EnvironmentValue) As IToken
            ' 式木を作成
            Dim logicalParser As New LogicalParser()
            Dim compParser As New ComparisonParser()
            Dim addOrSubParser As New AddOrSubParser()
            Dim multiOrDivParser As New MultiOrDivParser()
            Dim facParser As New FactorParser()
            Dim parenParser As New ParenParser()

            logicalParser.NextParser = compParser
            compParser.NextParser = addOrSubParser
            addOrSubParser.NextParser = multiOrDivParser
            multiOrDivParser.NextParser = facParser
            facParser.NextParser = parenParser
            parenParser.NextParser = logicalParser

            Dim tknPtr = New TokenPtr(tokens)
            Dim expr = logicalParser.Parser(tknPtr)

            ' 結果を取得する
            If Not tknPtr.HasNext Then
                Return expr.Executes(parameter)
            Else
                Throw New DSqlAnalysisException("未評価のトークンがあります")
            End If
        End Function

        ''' <summary>括弧内部式を取得します。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <param name="nxtParser">次のパーサー。</param>
        ''' <returns>括弧内部式。</returns>
        Private Function CreateParenExpress(reader As TokenPtr, nxtParser As IParser) As ParenExpress
            Dim tmp As New List(Of TokenPoint)()
            Dim lv As Integer = 0
            Do While reader.HasNext
                Dim tkn = reader.Current
                reader.Move(1)

                Select Case tkn.TokenName
                    Case NameOf(LParenToken)
                        tmp.Add(tkn)
                        lv += 1

                    Case NameOf(RParenToken)
                        If lv > 0 Then
                            tmp.Add(tkn)
                            lv -= 1
                        Else
                            Exit Do
                        End If
                    Case Else
                        tmp.Add(tkn)
                End Select
            Loop
            Return New ParenExpress(nxtParser.Parser(New TokenPtr(tmp)))
        End Function

        ''' <summary>解析インターフェイス。</summary>
        Private Interface IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Function Parser(reader As TokenPtr) As IExpression

        End Interface

        ''' <summary>括弧解析。</summary>
        Private NotInheritable Class ParenParser
            Implements IParser

            ''' <summary>次のパーサーを設定、取得する。</summary>
            Friend Property NextParser() As IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Public Function Parser(reader As TokenPtr) As IExpression Implements IParser.Parser
                Dim tkn = reader.Current
                If tkn.TokenName = NameOf(LParenToken) Then
                    reader.Move(1)
                    Return CreateParenExpress(reader, Me.NextParser)
                Else
                    Return Me.NextParser.Parser(reader)
                End If
            End Function

        End Class

        ''' <summary>論理解析。</summary>
        Private NotInheritable Class LogicalParser
            Implements IParser

            ''' <summary>次のパーサーを設定、取得する。</summary>
            Friend Property NextParser() As IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Public Function Parser(reader As TokenPtr) As IExpression Implements IParser.Parser
                Dim tml = Me.NextParser.Parser(reader)

                Do While reader.HasNext
                    Dim ope = reader.Current
                    Select Case ope.TokenName
                        Case NameOf(AndToken)
                            reader.Move(1)
                            tml = New AndExpress(tml, Me.NextParser.Parser(reader))

                        Case NameOf(OrToken)
                            reader.Move(1)
                            tml = New OrExpress(tml, Me.NextParser.Parser(reader))

                        Case Else
                            Exit Do
                    End Select
                Loop

                Return tml
            End Function

        End Class

        ''' <summary>比較解析。</summary>
        Private NotInheritable Class ComparisonParser
            Implements IParser

            ''' <summary>次のパーサーを設定、取得する。</summary>
            Friend Property NextParser() As IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Public Function Parser(reader As TokenPtr) As IExpression Implements IParser.Parser
                Dim tml = Me.NextParser.Parser(reader)

                If reader.HasNext Then
                    Dim ope = reader.Current
                    Select Case ope.TokenName
                        Case NameOf(EqualToken)
                            reader.Move(1)
                            tml = New EqualExpress(tml, Me.NextParser.Parser(reader))

                        Case NameOf(NotEqualToken)
                            reader.Move(1)
                            tml = New NotEqualExpress(tml, Me.NextParser.Parser(reader))

                        Case NameOf(GreaterToken)
                            reader.Move(1)
                            tml = New GreaterExpress(tml, Me.NextParser.Parser(reader))

                        Case NameOf(GreaterEqualToken)
                            reader.Move(1)
                            tml = New GreaterEqualExpress(tml, Me.NextParser.Parser(reader))

                        Case NameOf(LessToken)
                            reader.Move(1)
                            tml = New LessExpress(tml, Me.NextParser.Parser(reader))

                        Case NameOf(LessEqualToken)
                            reader.Move(1)
                            tml = New LessEqualExpress(tml, Me.NextParser.Parser(reader))
                    End Select
                End If

                Return tml
            End Function

        End Class

        ''' <summary>加算、減算解析。</summary>
        Private NotInheritable Class AddOrSubParser
            Implements IParser

            ''' <summary>次のパーサーを設定、取得する。</summary>
            Friend Property NextParser() As IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Public Function Parser(reader As TokenPtr) As IExpression Implements IParser.Parser
                Dim tml = Me.NextParser.Parser(reader)

                Do While reader.HasNext
                    Dim ope = reader.Current
                    Select Case ope.TokenName
                        Case NameOf(PlusToken)
                            reader.Move(1)
                            tml = New PlusExpress(tml, Me.NextParser.Parser(reader))

                        Case NameOf(MinusToken)
                            reader.Move(1)
                            tml = New MinusExpress(tml, Me.NextParser.Parser(reader))

                        Case Else
                            Exit Do
                    End Select
                Loop

                Return tml
            End Function

        End Class

        ''' <summary>乗算、除算解析。</summary>
        Private NotInheritable Class MultiOrDivParser
            Implements IParser

            ''' <summary>次のパーサーを設定、取得する。</summary>
            Friend Property NextParser() As IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Public Function Parser(reader As TokenPtr) As IExpression Implements IParser.Parser
                Dim tml = Me.NextParser.Parser(reader)

                Do While reader.HasNext
                    Dim ope = reader.Current
                    Select Case ope.TokenName
                        Case NameOf(MultiToken)
                            reader.Move(1)
                            tml = New MultiExpress(tml, Me.NextParser.Parser(reader))

                        Case NameOf(DivToken)
                            reader.Move(1)
                            tml = New DivExpress(tml, Me.NextParser.Parser(reader))

                        Case Else
                            Exit Do
                    End Select
                Loop

                Return tml
            End Function
        End Class

        ''' <summary>要素解析。</summary>
        Private NotInheritable Class FactorParser
            Implements IParser

            ''' <summary>次のパーサーを設定、取得する。</summary>
            Friend Property NextParser() As IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Public Function Parser(reader As TokenPtr) As IExpression Implements IParser.Parser
                Dim tkn = reader.Current

                Select Case tkn.TokenName
                    Case NameOf(IdentToken), NameOf(NumberToken), NameOf(StringToken),
                         NameOf(QueryToken), NameOf(ReplaseToken), NameOf(ObjectToken),
                         NameOf(TrueToken), NameOf(FalseToken), NameOf(NullToken)
                        reader.Move(1)
                        Return New ValueExpress(tkn.GetToken(Of IToken)())

                    Case NameOf(LParenToken)
                        reader.Move(1)
                        Return CreateParenExpress(reader, Me.NextParser)

                    Case NameOf(PlusToken), NameOf(MinusToken), NameOf(NotToken)
                        reader.Move(1)
                        Dim nxtExper = Me.Parser(reader)
                        If TypeOf nxtExper Is ValueExpress Then
                            Return New UnaryExpress(tkn.GetToken(Of IToken)(), nxtExper)
                        Else
                            Throw New DSqlAnalysisException($"前置き演算子{tkn}が値の前に配置していません")
                        End If

                    Case Else
                        Throw New DSqlAnalysisException("Factor要素の解析に失敗")
                End Select
            End Function

        End Class

    End Module

End Namespace
