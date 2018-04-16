Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax


Public Class FoldingSyntaxWalker
    Inherits VisualBasicSyntaxWalker

    Public Property Items As List(Of Tuple(Of Integer, Integer, String)) = Nothing

    Public Sub New()
        MyBase.New(SyntaxWalkerDepth.Node)

        Me.Items = New List(Of Tuple(Of Integer, Integer, String))

    End Sub

    Public Sub Parse(source As String)

        Dim tree = VisualBasicSyntaxTree.ParseText(source)
        Dim node = tree.GetRoot()

        Me.Visit(node)

    End Sub

    Public Overrides Sub Visit(node As SyntaxNode)
        MyBase.Visit(node)
    End Sub

    Public Overrides Sub VisitCompilationUnit(node As CompilationUnitSyntax)

        Me.AddRegionData(node)
        MyBase.VisitCompilationUnit(node)

    End Sub

    Public Overrides Sub VisitNamespaceBlock(node As NamespaceBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitNamespaceBlock(node)

    End Sub

    Public Overrides Sub VisitClassBlock(node As ClassBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitClassBlock(node)

    End Sub

    Public Overrides Sub VisitStructureBlock(node As StructureBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitStructureBlock(node)

    End Sub

    Public Overrides Sub VisitInterfaceBlock(node As InterfaceBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitInterfaceBlock(node)

    End Sub

    Public Overrides Sub VisitModuleBlock(node As ModuleBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitModuleBlock(node)

    End Sub

    Public Overrides Sub VisitEnumBlock(node As EnumBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitEnumBlock(node)

    End Sub

    Public Overrides Sub VisitConstructorBlock(node As ConstructorBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitConstructorBlock(node)

    End Sub

    Public Overrides Sub VisitOperatorBlock(node As OperatorBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitOperatorBlock(node)

    End Sub

    ' SubBlock, FunctionBlock が含まれているか？
    Public Overrides Sub VisitMethodBlock(node As MethodBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitMethodBlock(node)

    End Sub

    Public Overrides Sub VisitPropertyBlock(node As PropertyBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitPropertyBlock(node)

    End Sub

    ' 以下が含まれているか？
    ' Property(GetAccessorBlock, SetAccessorBlock)
    ' Custom Event(AddHandlerAccessorBlock, RemoveHandlerAccessorBlock, RaiseEventAccessorBlock)
    Public Overrides Sub VisitAccessorBlock(node As AccessorBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitAccessorBlock(node)

    End Sub

    Public Overrides Sub VisitEventBlock(node As EventBlockSyntax)

        Me.AddBlockData(node)
        MyBase.VisitEventBlock(node)

    End Sub

    Private Sub AddRegionData(node As VisualBasicSyntaxNode)

        ' どうやら、GetDirectives() が、階層関係なく再帰的に取得してくるみたい、かつ各ブロックで、ところどころ該当する同じ値を取得してくるので、
        ' トップレベルのノードだけで、全ての #Region を取得して返却する

        ' 宣言順にリスト登録されている。対応する開始 Region と終了 Region を取得するため、スタックを利用する
        ' （開始ノードが、終了分までの情報を持っていない）

        Dim regionStack = New Stack(Of DirectiveTriviaSyntax)
        For Each child In node.GetDirectives()

            Dim childKind = child.Kind()
            Select Case childKind

                Case SyntaxKind.RegionDirectiveTrivia

                    regionStack.Push(child)

                Case SyntaxKind.EndRegionDirectiveTrivia

                    Dim startSyntax = regionStack.Pop()

                    Dim startLength = startSyntax.Span.Start
                    Dim endLength = child.Span.End

                    ' #Region "aaa" のうち、冒頭の「#Region」と文字列を囲うダブルコーテーションを除去
                    Dim header = startSyntax.ToString()
                    header = header.Substring("#Region ".Length)
                    header = header.Substring(1)
                    header = header.Substring(0, header.Length - 1)

                    Me.Items.Add(Tuple.Create(startLength, endLength, header))

            End Select

        Next

    End Sub

    Private Sub AddBlockData(node As VisualBasicSyntaxNode)

        ' ブロック系は、開始ノードが、開始と終了の両方の文字列位置を持っているので、取得する

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim header = node.ToString()
        If header.Contains(Environment.NewLine) Then
            header = header.Substring(0, header.IndexOf(Environment.NewLine))
            header = $"{header} ..."
        End If

        Me.Items.Add(Tuple.Create(startLength, endLength, header))

    End Sub

End Class
