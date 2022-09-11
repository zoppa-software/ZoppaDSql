Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>トークン階層リンク。</summary>
    Friend NotInheritable Class TokenLink

        ' 対象トークン
        Private mToken As TokenPoint

        ' 子要素リスト
        Private mChildren As List(Of TokenLink)

        ''' <summary>対象トークンを取得します。</summary>
        Public ReadOnly Property TargetToken As TokenPoint
            Get
                Return Me.mToken
            End Get
        End Property

        ''' <summary>子要素リストを取得します。</summary>
        Public ReadOnly Property Children As List(Of TokenLink)
            Get
                Return Me.mChildren
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tkn">対象トークン。</param>
        Public Sub New(tkn As TokenPoint)
            Me.mToken = tkn
        End Sub

        ''' <summary>子要素を追加します。</summary>
        ''' <param name="token">対象トークン。</param>
        ''' <returns>追加した階層リンク。</returns>
        Friend Function AddChild(token As TokenPoint) As TokenLink
            If Me.mChildren Is Nothing Then
                Me.mChildren = New List(Of TokenLink)()
            End If
            Dim cnode = New TokenLink(token)
            Me.mChildren.Add(cnode)
            Return cnode
        End Function

    End Class

End Namespace
