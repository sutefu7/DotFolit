Imports System.Data
Imports System.IO
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax


' 「親情報の引継ぎ渡しの仕組み」がうまく思いつかなかったため、
' VisualBasicSyntaxWalker クラスを継承した実装は、いったん保留
' 以前の自前処理を使う


Public Class MethodSyntaxWalker

    Public Property MethodModel As MethodTemplateModel = Nothing

    Public Sub New()

    End Sub

    Public Sub Parse(methodRange As String)

        Dim tree = VisualBasicSyntaxTree.ParseText(methodRange)
        Dim root = TryCast(tree.GetRoot(), VisualBasicSyntaxNode)
        Dim node = root.DescendantNodes().FirstOrDefault(Function(x)

                                                             Select Case x.Kind()
                                                                 Case SyntaxKind.SubNewStatement,
                                                                      SyntaxKind.SubStatement,
                                                                      SyntaxKind.FunctionStatement,
                                                                      SyntaxKind.OperatorStatement,
                                                                      SyntaxKind.DeclareSubStatement,
                                                                      SyntaxKind.DeclareFunctionStatement
                                                                     Return True
                                                                 Case Else
                                                                     Return False
                                                             End Select

                                                         End Function)

        Me.MethodModel = New MethodTemplateModel With {.Signature = node.ToString()}
        Me.ParseInternal(root, Me.MethodModel)

    End Sub

    Private Sub ParseInternal(node As VisualBasicSyntaxNode, parentModel As BaseTemplateModel)

        Select Case node.Kind()

            Case SyntaxKind.CompilationUnit,
                 SyntaxKind.ConstructorBlock,
                 SyntaxKind.OperatorBlock,
                 SyntaxKind.SubBlock, SyntaxKind.FunctionBlock

                ' 入れ子を再帰
                For Each child As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(child, parentModel)
                Next


            Case SyntaxKind.SingleLineIfStatement
                ' 単一行 If

                Dim containerModel = New ContainerIfTemplateModel
                parentModel.Children.Add(containerModel)

                Dim childModel = New IfTemplateModel
                containerModel.Children.Add(childModel)

                ' 入れ子を再帰
                For Each child As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(child, childModel)
                Next


            Case SyntaxKind.MultiLineIfBlock

                Dim containerModel = New ContainerIfTemplateModel
                parentModel.Children.Add(containerModel)

                Dim childNodes = node.ChildNodes().Where(Function(x) (TypeOf x Is IfStatementSyntax) OrElse (TypeOf x Is ElseIfBlockSyntax) OrElse (TypeOf x Is ElseBlockSyntax))
                For Each childNode In childNodes

                    Select Case True

                        Case TypeOf childNode Is IfStatementSyntax

                            Dim childModel = New IfTemplateModel
                            containerModel.Children.Add(childModel)

                            ' If の場合、Block ではなく Statement なので、この下階層（子リスト）には、ループや条件分岐が（ソース的にはあったとしても）含まれない
                            ' １階層上にさかのぼって、ElseIfBlock, または ElseBlock 以外で、ループや条件分岐がある場合、各自を再帰する
                            Dim innerBlocks = node.ChildNodes().Where(Function(x) (TypeOf x IsNot IfStatementSyntax) AndAlso (TypeOf x IsNot ElseIfBlockSyntax) AndAlso (TypeOf x IsNot ElseBlockSyntax))
                            For Each innerBlock As VisualBasicSyntaxNode In innerBlocks
                                Me.ParseInternal(innerBlock, childModel)
                            Next


                        Case TypeOf childNode Is ElseIfBlockSyntax

                            Dim childModel = New ElseIfTemplateModel
                            containerModel.Children.Add(childModel)

                            For Each grandChild As VisualBasicSyntaxNode In childNode.ChildNodes()
                                Me.ParseInternal(grandChild, childModel)
                            Next


                        Case TypeOf childNode Is ElseBlockSyntax

                            Dim childModel = New ElseTemplateModel
                            containerModel.Children.Add(childModel)

                            For Each grandChild As VisualBasicSyntaxNode In childNode.ChildNodes()
                                Me.ParseInternal(grandChild, childModel)
                            Next

                    End Select

                Next


            Case SyntaxKind.SelectBlock

                Dim containerModel = New ContainerSelectTemplateModel
                parentModel.Children.Add(containerModel)

                Dim childNode As SyntaxNode = Nothing
                Dim childNodes = node.ChildNodes().Where(Function(x) TypeOf x Is CaseBlockSyntax)

                For i As Integer = 0 To childNodes.Count() - 1

                    ' 念のためのチェック
                    childNode = childNodes(i)
                    If TypeOf childNode Is CaseBlockSyntax Then

                        If i = 0 Then

                            Dim childModel = New SelectTemplateModel
                            containerModel.Children.Add(childModel)

                            For Each grandChild As VisualBasicSyntaxNode In childNode.ChildNodes()
                                Me.ParseInternal(grandChild, childModel)
                            Next

                        ElseIf childNode.DescendantNodes().Any(Function(x) TypeOf x Is ElseCaseClauseSyntax) Then

                            Dim childModel = New CaseElseTemplateModel
                            containerModel.Children.Add(childModel)

                            For Each grandChild As VisualBasicSyntaxNode In childNode.ChildNodes()
                                Me.ParseInternal(grandChild, childModel)
                            Next

                        Else

                            Dim childModel = New CaseTemplateModel
                            containerModel.Children.Add(childModel)

                            For Each grandChild As VisualBasicSyntaxNode In childNode.ChildNodes()
                                Me.ParseInternal(grandChild, childModel)
                            Next

                        End If

                    End If

                Next


            Case SyntaxKind.ForBlock

                Dim childModel = New ForTemplateModel
                parentModel.Children.Add(childModel)

                For Each childNode As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(childNode, childModel)
                Next


            Case SyntaxKind.ForEachBlock

                Dim childModel = New ForEachTemplateModel
                parentModel.Children.Add(childModel)

                For Each childNode As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(childNode, childModel)
                Next


            Case SyntaxKind.WhileBlock

                Dim childModel = New WhileTemplateModel
                parentModel.Children.Add(childModel)

                For Each childNode As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(childNode, childModel)
                Next


            Case SyntaxKind.SimpleDoLoopBlock

                Dim childModel = New DoLoopTemplateModel
                parentModel.Children.Add(childModel)

                For Each childNode As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(childNode, childModel)
                Next


            Case SyntaxKind.DoLoopUntilBlock

                Dim childModel = New DoUntilTemplateModel
                parentModel.Children.Add(childModel)

                For Each childNode As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(childNode, childModel)
                Next


            Case SyntaxKind.DoLoopWhileBlock

                Dim childModel = New DoWhileTemplateModel
                parentModel.Children.Add(childModel)

                For Each childNode As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(childNode, childModel)
                Next


            Case SyntaxKind.UsingBlock

                Dim childModel = New UsingTemplateModel
                parentModel.Children.Add(childModel)

                For Each childNode As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(childNode, childModel)
                Next


            Case SyntaxKind.TryBlock

                ' Try
                Dim tryModel = New TryTemplateModel
                parentModel.Children.Add(tryModel)

                Dim innerBlocks1 = node.ChildNodes().Where(Function(x) (TypeOf x IsNot TryStatementSyntax) AndAlso (TypeOf x IsNot CatchBlockSyntax) AndAlso (TypeOf x IsNot FinallyBlockSyntax))
                For Each innerBlock As VisualBasicSyntaxNode In innerBlocks1
                    Me.ParseInternal(innerBlock, tryModel)
                Next

                ' Catch
                If node.ChildNodes().Any(Function(x) TypeOf x Is CatchBlockSyntax) Then

                    Dim catchModel = New CatchTemplateModel
                    parentModel.Children.Add(catchModel)

                    Dim innerBlocks2 = node.ChildNodes().Where(Function(x) TypeOf x Is CatchBlockSyntax)
                    For Each innerBlock As VisualBasicSyntaxNode In innerBlocks2
                        Me.ParseInternal(innerBlock, catchModel)
                    Next

                End If

                ' Finally
                If node.ChildNodes().Any(Function(x) TypeOf x Is FinallyBlockSyntax) Then

                    Dim finallyModel = New FinallyTemplateModel
                    parentModel.Children.Add(finallyModel)

                    Dim innerBlocks2 = node.ChildNodes().Where(Function(x) TypeOf x Is FinallyBlockSyntax)
                    For Each innerBlock As VisualBasicSyntaxNode In innerBlocks2
                        Me.ParseInternal(innerBlock, finallyModel)
                    Next

                End If

        End Select

    End Sub

End Class
