Option Strict On
Option Explicit On

Imports System.Data.Common
Imports System.IO
Imports System.Linq.Expressions
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows
Imports ZoppaDSql.Analysis.Environments
Imports ZoppaDSql.Analysis.Express
Imports ZoppaDSql.Analysis.Tokens
Imports ZoppaDSql.TokenCollection

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
            Dim values = New EnvironmentObjectValue(parameter)
            Dim ansParts As New List(Of EvaParts)()
            Evaluation(hierarchy.Children, buffer, values, ansParts)

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
        Public Function CreateHierarchy(src As List(Of TokenPosition)) As TokenHierarchy
            Dim root As New TokenHierarchy(Nothing)
            If src?.Count > 0 Then
                Dim tmp As New List(Of TokenPosition)(src)
                tmp.Reverse()

                CreateHierarchy("", root, tmp)
            End If
            Return root
        End Function

        ''' <summary>置き換え式の階層を作成します。</summary>
        ''' <param name="errMsg">エラーメッセージ。</param>
        ''' <param name="node">階層ノード。</param>
        ''' <param name="tokens">対象トークンリスト。</param>
        ''' <param name="partitionTokens"></param>
        Private Sub CreateHierarchy(errMsg As String,
                                    node As TokenHierarchy,
                                    tokens As List(Of TokenPosition),
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
                            CreateHierarchy("ifブロックが閉じられていません", cnode, tokens, NameOf(ElseIfToken), NameOf(ElseToken), NameOf(EndIfToken))

                        Case NameOf(ElseIfToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy("ifブロックが閉じられていません", cnode, tokens, NameOf(ElseToken), NameOf(EndIfToken))

                        Case NameOf(ElseToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy("ifブロックが閉じられていません", cnode, tokens, NameOf(EndIfToken))

                        Case NameOf(ForEachToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy("foreachブロックが閉じられていません", cnode, tokens, NameOf(EndForToken))

                        Case NameOf(TrimToken)
                            tokens.RemoveAt(tokens.Count - 1)
                            CreateHierarchy("trimブロックが閉じられていません", cnode, tokens, NameOf(EndTrimToken))

                        Case Else
                            tokens.RemoveAt(tokens.Count - 1)
                    End Select
                End If
            Loop
            If partitionTokens.Count > 0 Then
                Throw New DSqlAnalysisException(errMsg)
            End If
        End Sub

        ''' <summary>トークンリストを評価します。</summary>
        ''' <param name="children">階層情報。</param>
        ''' <param name="buffer">出力先バッファ。</param>
        ''' <param name="prm">環境値情報。</param>
        ''' <param name="ansParts">評価結果リスト。</param>
        Private Sub Evaluation(children As List(Of TokenHierarchy),
                               buffer As StringBuilder,
                               prm As IEnvironmentValue,
                               ansParts As List(Of EvaParts))
            Dim tmp As New List(Of TokenHierarchy)(children)
            tmp.Reverse()

            Do While tmp.Count > 0
                Dim chd = tmp(tmp.Count - 1)

                Select Case chd.TargetToken.TokenName
                    Case NameOf(IfToken)
                        EvaluationIf(tmp, buffer, prm, ansParts)

                    Case NameOf(ForEachToken)
                        EvaluationFor(chd, buffer, prm, ansParts)
                        tmp.RemoveAt(tmp.Count - 1)

                    Case NameOf(TrimToken)
                        EvaluationTrim(chd, buffer, prm, ansParts)
                        tmp.RemoveAt(tmp.Count - 1)

                    Case NameOf(ReplaseToken)
                        Dim rtoken = chd.TargetToken.GetToken(Of ReplaseToken)()
                        If rtoken IsNot Nothing Then
                            Dim rval = prm.GetValue(If(rtoken.Contents?.ToString(), ""))
                            Dim ans = GetRefValue(rval, rtoken.IsEscape)
                            buffer.Append(ans)
                            ansParts.Add(New EvaParts(chd.TargetToken, ans.ToString()))
                        End If
                        tmp.RemoveAt(tmp.Count - 1)

                    Case NameOf(EndIfToken)
                        ' ※ EvaluationIfで処理するため不要です

                    Case NameOf(EndForToken), NameOf(EndTrimToken)
                        ansParts.Add(New EvaParts(chd.TargetToken, ""))
                        tmp.RemoveAt(tmp.Count - 1)

                    Case Else
                        buffer.Append(chd.TargetToken.Contents)
                        ansParts.Add(New EvaParts(chd.TargetToken, chd.TargetToken.Contents.ToString()))
                        tmp.RemoveAt(tmp.Count - 1)
                End Select
            Loop
        End Sub

        Private Function GetRefValue(ans As Object, isEscape As Boolean) As String
            If TypeOf ans Is String AndAlso isEscape Then
                Dim s = ans.ToString()
                s = s.Replace("'"c, "''")
                s = s.Replace("\"c, "\\")
                Return $"'{s}'"

            ElseIf TypeOf ans Is String Then
                Return ans.ToString()

            ElseIf TypeOf ans Is IEnumerable Then
                Dim buf As New StringBuilder()
                For Each itm In CType(ans, IEnumerable)
                    If buf.Length > 0 Then
                        buf.Append(", ")
                    End If
                    buf.Append(GetRefValue(itm, isEscape))
                Next
                Return buf.ToString()

            ElseIf ans Is Nothing Then
                Return "null"
            Else
                Return ans.ToString()
            End If
        End Function

        ''' <summary>Ifを評価します。</summary>
        ''' <param name="tmp">階層情報。</param>
        ''' <param name="buffer">出力先バッファ。</param>
        ''' <param name="prm">環境値情報。</param>
        ''' <param name="ansParts">評価結果リスト。</param>
        Private Sub EvaluationIf(tmp As List(Of TokenHierarchy),
                                 buffer As StringBuilder,
                                 prm As IEnvironmentValue,
                                 ansParts As List(Of EvaParts))
            ' If、ElseIf、Elseブロックを集める
            Dim blocks As New List(Of TokenHierarchy)()
            Dim endifTkn As TokenPosition = Nothing
            For i As Integer = tmp.Count - 1 To 0 Step -1
                Dim iftkn = tmp(i)
                endifTkn = iftkn.TargetToken
                tmp.RemoveAt(tmp.Count - 1)
                If iftkn.TargetToken.TokenName <> NameOf(EndIfToken) Then
                    blocks.Add(iftkn)
                Else
                    Exit For
                End If
            Next

            ' If、ElseIf、Elseブロックを評価
            Dim output = False
            For Each iftkn In blocks
                Select Case iftkn.TargetToken.TokenName
                    Case NameOf(IfToken), NameOf(ElseIfToken)
                        ' 条件を評価して真ならば、ブロックを出力
                        Dim ifans = Executes(iftkn.TargetToken.GetCommandToken().CommandTokens, prm)
                        If TypeOf ifans.Contents Is Boolean AndAlso CBool(ifans.Contents) Then
                            ansParts.Add(New EvaParts(iftkn.TargetToken, ""))
                            Evaluation(iftkn.Children, buffer, prm, ansParts)
                            output = True
                            Exit For
                        End If

                    Case NameOf(ElseToken)
                        ansParts.Add(New EvaParts(iftkn.TargetToken, ""))
                        Evaluation(iftkn.Children, buffer, prm, ansParts)
                        output = True
                        Exit For
                End Select
            Next

            If output Then
                ansParts.Add(New EvaParts(endifTkn, ""))
            Else
                For Each iftkn In blocks
                    ansParts.Add(New EvaParts(iftkn.TargetToken))
                Next
                ansParts.Add(New EvaParts(endifTkn))
            End If
        End Sub

        ''' <summary>Foreachを評価します。</summary>
        ''' <param name="fortoken">Foreachブロック。</param>
        ''' <param name="buffer">出力先バッファ。</param>
        ''' <param name="prm">環境値情報。</param>
        ''' <param name="ansParts">評価結果トークンリスト。</param>
        Private Sub EvaluationFor(forToken As TokenHierarchy,
                                  buffer As StringBuilder,
                                  prm As IEnvironmentValue,
                                  ansParts As List(Of EvaParts))
            Dim forTkn = forToken.TargetToken.GetToken(Of ForEachToken)()
            ansParts.Add(New EvaParts(forToken.TargetToken, ""))

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
                Evaluation(forToken.Children, buffer, prm, ansParts)
            Next

            prm.LocalVarClear()
        End Sub

        ''' <summary>Trimを評価します。</summary>
        ''' <param name="trimToken">Trimブロック。</param>
        ''' <param name="buffer">出力先バッファ。</param>
        ''' <param name="prm">環境値情報。</param>
        Private Sub EvaluationTrim(trimToken As TokenHierarchy,
                                   buffer As StringBuilder,
                                   prm As IEnvironmentValue,
                                   ansParts As List(Of EvaParts))
            Dim answer As New StringBuilder()

            ' Trim内のトークンを評価
            Dim tmpbuf As New StringBuilder()
            Dim subAnsParts As New List(Of EvaParts)()
            Evaluation(trimToken.Children, tmpbuf, prm, subAnsParts)

            ' 先頭の空白を取り除く
            Do While subAnsParts.Count > 0
                If subAnsParts(0).outString?.Trim() <> "" OrElse (subAnsParts(0).IsControlToken AndAlso subAnsParts(0).isOutpit) Then
                    Exit Do
                Else
                    subAnsParts.RemoveAt(0)
                End If
            Loop

            ' 末尾の空白を取り除く
            For i As Integer = subAnsParts.Count - 1 To 0 Step -1
                If subAnsParts(i).outString?.Trim() <> "" OrElse (subAnsParts(i).IsControlToken AndAlso subAnsParts(i).isOutpit) Then
                    Exit For
                Else
                    subAnsParts.RemoveAt(i)
                End If
            Next

            ' トリムルールに従ってトリムします
            If subAnsParts.First().token.TokenName = NameOf(ForEachToken) AndAlso
               subAnsParts.Last().token.TokenName = NameOf(EndForToken) Then
                RemoveTrimEndByForeach(trimToken.TargetToken.GetToken(Of TrimToken)(), subAnsParts, ",")
                For Each parts In subAnsParts
                    ansParts.Add(parts)
                    buffer.Append(parts.outString)
                Next

            ElseIf subAnsParts.First().token.TokenName = NameOf(QueryToken) AndAlso
                   subAnsParts.First().outString?.Trim().ToLower().StartsWith("where") Then
                buffer.Append(RemoveAndOrByWhere(subAnsParts))

            ElseIf subAnsParts.First().token.TokenName = NameOf(QueryToken) AndAlso
                   subAnsParts.First().outString?.Trim().ToLower().StartsWith("set") Then
                'RemoveAndOrByWhere(subAnsParts)
                'RemoveParentByWhere(subAnsParts)

                'Dim tmp As New StringBuilder()
                'For Each tkn In subAnsParts
                '    If tkn.isOutpit Then
                '        ansParts.Add(tkn)
                '        tmp.Append(tkn.outString)
                '    End If
                'Next

                'Dim tmpstr = tmp.ToString()
                'If tmpstr.Trim().ToLower() <> "where" Then
                '    buffer.Append(tmpstr)
                'End If

            Else
                RemoveTrimEndByForeach(trimToken.TargetToken.GetToken(Of TrimToken)(), subAnsParts, "")
                For Each tkn In subAnsParts
                    ansParts.Add(tkn)
                    buffer.Append(tkn.outString)
                Next
            End If
        End Sub

        Private Sub RemoveTrimEndByForeach(trimTkn As TrimToken, ansParts As List(Of EvaParts), defTrmStr As String)
            Dim trmstr = If(trimTkn.TrimString?.Trim() = "", defTrmStr, trimTkn.TrimString?.Trim())

            If trmstr <> "" Then
                For i As Integer = ansParts.Count - 1 To 0 Step -1
                    If ansParts(i).token.TokenName = NameOf(QueryToken) AndAlso Not ansParts(i).IsSpace Then
                        Dim str = ansParts(i).outString.TrimEnd()
                        If str.EndsWith(trmstr) Then
                            ansParts(i).outString = str.Substring(0, str.Length - trmstr.Length)
                        End If
                        Exit For
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ansParts"></param>
        Private Function RemoveAndOrByWhere(ansParts As List(Of EvaParts)) As String
            Dim copyParts As New List(Of EvaParts)()
            Dim partsLink As New Dictionary(Of EvaParts, List(Of EvaParts))()
            For Each ps In ansParts
                If ps.token.TokenName = NameOf(QueryToken) Then
                    Dim ch = ps.outString.ToArray()
                    Dim i As Integer = 0
                    Dim tokens As New List(Of (str As String, keywd As Boolean, space As Boolean))()
                    Do While i < ch.Length
                        If ch(i) = "("c Then
                            tokens.Add(("(", True, False))
                            i += 1
                        ElseIf ch(i) = ")"c Then
                            tokens.Add((")", True, False))
                            i += 1
                        ElseIf Char.IsWhiteSpace(ch(i)) Then
                            Dim buf As New StringBuilder()
                            For j As Integer = i To ch.Length - 1
                                If Char.IsWhiteSpace(ch(j)) Then
                                    buf.Append(ch(j))
                                    i = i + 1
                                Else
                                    Exit For
                                End If
                            Next
                            tokens.Add((buf.ToString(), False, True))
                        Else
                            Dim buf As New StringBuilder()
                            For j As Integer = i To ch.Length - 1
                                If Not Char.IsWhiteSpace(ch(j)) OrElse ch(j) = "("c OrElse ch(j) = ")"c Then
                                    buf.Append(ch(j))
                                    i = i + 1
                                Else
                                    Exit For
                                End If
                            Next
                            Dim str = buf.ToString()
                            tokens.Add((str, str.ToLower() = "and" OrElse str.ToLower() = "or", False))
                        End If
                    Loop

                    Dim cnt As Integer = 0
                    Dim def = New Integer(tokens.Count - 1) {}
                    For j As Integer = 0 To tokens.Count - 1
                        If tokens(j).keywd Then
                            If j > 0 AndAlso Not tokens(j - 1).keywd Then cnt += 1
                            def(j) = cnt
                            cnt += 1
                        ElseIf tokens(j).space AndAlso (j = 0 OrElse tokens(j - 1).keywd) AndAlso (j = tokens.Count - 1 OrElse tokens(j + 1).keywd) Then
                            def(j) = cnt
                            cnt += 1
                        Else
                            def(j) = cnt
                        End If
                    Next

                    Dim children As New List(Of EvaParts)()
                    i = 0
                    Do While i < tokens.Count
                        Dim buf As New StringBuilder()
                        buf.Append(tokens(i).str)
                        i += 1
                        For j As Integer = i To tokens.Count - 1
                            If def(j - 1) = def(j) Then
                                buf.Append(tokens(j).str)
                                i += 1
                            Else
                                Exit For
                            End If
                        Next
                        children.Add(New EvaParts(ps.token, buf.ToString()))
                    Loop
                    copyParts.AddRange(children)
                    partsLink.Add(ps, children)
                Else
                    copyParts.Add(ps)
                End If
            Next

            Dim parts = New EvaPartsPointer(copyParts)
            Do While parts.HasNext
                If parts.Current.IsSpace Then
                    parts.Move()
                ElseIf parts.Current.outString.Trim().ToLower().StartsWith("where") Then
                    parts.Move()
                    Exit Do
                Else
                    Exit Do
                End If
            Loop

            LogicalTrim(parts)

            Dim tmp As New StringBuilder()
            For Each tkn In parts.parts
                If tkn.isOutpit Then
                    ansParts.Add(tkn)
                    tmp.Append(tkn.outString)
                End If
            Next

            Dim tmpstr = tmp.ToString()
            If tmpstr.Trim().ToLower() <> "where" Then
                For Each parent In partsLink.Keys
                    Dim buf As New StringBuilder()
                    Dim view = False
                    For Each ch In partsLink(parent)
                        If ch.isOutpit Then
                            buf.Append(ch.outString)
                            view = True
                        End If
                    Next
                    parent.outString = buf.ToString()
                    parent.isOutpit = view
                Next
                Return tmpstr
            Else
                For Each ps In ansParts
                    ps.isOutpit = False
                Next
                Return ""
            End If
        End Function

        Private Function LogicalTrim(parts As EvaPartsPointer) As Boolean
            Dim tml = FactorTrim(parts)
            Do While parts.HasNext
                Dim ope = parts.Current
                If ope.outString.ToLower() = "and" OrElse
                   ope.outString.ToLower() = "or" Then
                    parts.Move()
                    Dim tmr = FactorTrim(parts)
                    If Not tml OrElse Not tmr Then
                        ope.isOutpit = False
                    End If
                    tml = tml Or tmr
                Else
                    Exit Do
                End If
            Loop
            Return tml
        End Function

        Private Function FactorTrim(parts As EvaPartsPointer) As Boolean
            Dim facts As New List(Of EvaParts)()
            Dim isOut As Boolean = True
            Do While parts.HasNext
                If parts.Current.outString = "(" Then
                    isOut = isOut And ParenTrim(parts)
                ElseIf parts.Current.outString = ")" Then
                    Exit Do
                ElseIf parts.Current.outString.ToLower() = "and" OrElse
                       parts.Current.outString.ToLower() = "or" Then
                    Exit Do
                Else
                    facts.Add(parts.Current)
                    isOut = isOut And parts.Current.isOutpit
                    parts.Move()
                End If
            Loop

            If isOut Then
                Return True
            Else
                For Each pt In facts
                    If Not pt.IsSpace AndAlso pt.isOutpit Then
                        Return True
                    End If
                Next
                For Each pt In facts
                    pt.isOutpit = False
                Next
                Return False
            End If
        End Function

        Private Function ParenTrim(parts As EvaPartsPointer) As Boolean
            Dim lParen = parts.Current
            parts.Move()

            Dim isOut = LogicalTrim(parts)

            Dim rParen = parts.Current
            parts.Move()

            If isOut Then
                Return True
            Else
                lParen.isOutpit = False
                rParen.isOutpit = False
                Return False
            End If
        End Function

        ''' <summary>式を解析して結果を取得します。</summary>
        ''' <param name="expression">式文字列。</param>
        ''' <param name="parameter">パラメータ。</param>
        ''' <returns>解析結果。</returns>
        Public Function Executes(expression As String, Optional parameter As Object = Nothing) As IToken
            ' トークン配列に変換
            Dim tokens = LexicalAnalysis.SimpleCompile(expression)

            ' 式木を実行
            Return Executes(tokens, New EnvironmentObjectValue(parameter))
        End Function

        ''' <summary>式を解析して結果を取得します。</summary>
        ''' <param name="tokens">対象トークン。。</param>
        ''' <param name="parameter">パラメータ。</param>
        ''' <returns>解析結果。</returns>
        Friend Function Executes(tokens As List(Of TokenPosition), parameter As IEnvironmentValue) As IToken
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

            Dim tknPtr = New TokenStream(tokens)
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
        Private Function CreateParenExpress(reader As TokenStream, nxtParser As IParser) As ParenExpress
            Dim tmp As New List(Of TokenPosition)()
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
            Return New ParenExpress(nxtParser.Parser(New TokenStream(tmp)))
        End Function

        Private Class EvaPartsPointer

            Public parts As List(Of EvaParts)

            Public index As Integer = 0

            Public ReadOnly Property Current As EvaParts
                Get
                    Return Me.parts(Me.index)
                End Get
            End Property

            Public ReadOnly Property HasNext As Boolean
                Get
                    Return Me.index < Me.parts.Count
                End Get
            End Property

            Public Sub New(srcParts As List(Of EvaParts))
                Me.parts = New List(Of EvaParts)(srcParts)
            End Sub

            Public Sub Move()
                Me.index += 1
            End Sub

        End Class

        Private NotInheritable Class EvaParts

            Public ReadOnly token As TokenPosition

            Public isOutpit As Boolean

            Public outString As String

            Public ReadOnly Property IsControlToken As Boolean
                Get
                    Return Me.token.TokenName = NameOf(IfToken) OrElse
                           Me.token.TokenName = NameOf(ForEachToken) OrElse
                           Me.token.TokenName = NameOf(TrimToken) OrElse
                           Me.token.TokenName = NameOf(EndIfToken) OrElse
                           Me.token.TokenName = NameOf(EndForToken) OrElse
                           Me.token.TokenName = NameOf(EndTrimToken)
                End Get
            End Property

            Public ReadOnly Property IsSpace As Boolean
                Get
                    Return (Me.outString?.Trim() = "")
                End Get
            End Property

            Public Sub New(tkn As TokenPosition)
                Me.token = tkn
                Me.isOutpit = False
                Me.outString = ""
            End Sub

            Public Sub New(tkn As TokenPosition, outStr As String)
                Me.token = tkn
                Me.isOutpit = True
                Me.outString = outStr
            End Sub

            Public Overrides Function ToString() As String
                Return $"{Me.outString} view:{Me.isOutpit} ctrl:{Me.IsControlToken}"
            End Function

        End Class

        ''' <summary>解析インターフェイス。</summary>
        Private Interface IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Function Parser(reader As TokenStream) As IExpression

        End Interface

        ''' <summary>括弧解析。</summary>
        Private NotInheritable Class ParenParser
            Implements IParser

            ''' <summary>次のパーサーを設定、取得する。</summary>
            Friend Property NextParser() As IParser

            ''' <summary>解析を実行する。</summary>
            ''' <param name="reader">入力トークンストリーム。</param>
            ''' <returns>解析結果。</returns>
            Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
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
            Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
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
            Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
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
            Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
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
            Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
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
            Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
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
