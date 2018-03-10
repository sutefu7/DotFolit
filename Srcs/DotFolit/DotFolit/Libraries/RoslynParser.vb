Imports System.Data
Imports System.IO
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax


' 実装がほぼ完了しているので今更直さないが、
' 本来は、VisualBasicSyntaxWalker クラスを継承して、各ノードやトークンなどを回りながら、やりたい事をする。というやり方みたい
' 以下は、それの自前実装みたいなことをしている


Public Class RoslynParser

    Public Sub New()

    End Sub

    Public Sub Parse(nsTable As DataTable, rootNamespace As String, sourceFile As String)

        Dim source = File.ReadAllText(sourceFile, EncodeResolver.GetEncoding(sourceFile))
        Dim tree = VisualBasicSyntaxTree.ParseText(source)
        Dim root = TryCast(tree.GetRoot(), VisualBasicSyntaxNode)

        Me.ParseInternal(nsTable, sourceFile, rootNamespace, "Namespace", root)

    End Sub

    Private Sub ParseInternal(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim kind = node.Kind()

        Select Case kind

            Case SyntaxKind.CompilationUnit

                For Each child As VisualBasicSyntaxNode In node.ChildNodes
                    Me.ParseInternal(nsTable, sourceFile, containerFullName, containerDefineKind, child)
                Next

                'For Each child In node.ChildTokens
                'Next

            Case SyntaxKind.NamespaceBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End
                Dim defineName = String.Empty
                Dim statementNode = node.ChildNodes().OfType(Of NamespaceStatementSyntax)().FirstOrDefault()

                If statementNode.ChildNodes().OfType(Of QualifiedNameSyntax)().Any() Then
                    Dim childNode = statementNode.ChildNodes().OfType(Of QualifiedNameSyntax)().FirstOrDefault()
                    defineName = childNode.ToString()

                ElseIf statementNode.ChildNodes().OfType(Of IdentifierNameSyntax)().Any() Then
                    Dim childNode = statementNode.ChildNodes().OfType(Of IdentifierNameSyntax)().FirstOrDefault()
                    defineName = childNode.ToString()

                End If

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Namespace"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = $"{containerFullName}.{defineName}"
                row("DefineType") = DBNull.Value
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = DBNull.Value
                row("IsPartial") = DBNull.Value
                row("IsShared") = DBNull.Value
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)

                For Each child As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(nsTable, sourceFile, $"{containerFullName}.{defineName}", "Namespace", child)
                Next


            Case SyntaxKind.ClassBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of ClassStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' ジェネリックを定義している場合、個数を追加
                If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                ' 内側のコンテナの場合、ドットつなぎではなくプラスつなぎにする
                Dim defineFullName = $"{containerFullName}.{defineName}"
                Dim containerKeywords = New List(Of String) From {"Class", "Structure", "Interface", "Module"}
                If containerKeywords.Contains(containerDefineKind) Then
                    defineFullName = $"{containerFullName}+{defineName}"
                End If

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Class"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = defineFullName
                row("DefineType") = DBNull.Value
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = DBNull.Value
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)

                For Each child As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(nsTable, sourceFile, defineFullName, "Class", child)
                Next


            Case SyntaxKind.StructureBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of StructureStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' ジェネリックを定義している場合、個数を追加
                If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                ' 内側のコンテナの場合、ドットつなぎではなくプラスつなぎにする
                Dim defineFullName = $"{containerFullName}.{defineName}"
                Dim containerKeywords = New List(Of String) From {"Class", "Structure", "Interface", "Module"}
                If containerKeywords.Contains(containerDefineKind) Then
                    defineFullName = $"{containerFullName}+{defineName}"
                End If

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Structure"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = defineFullName
                row("DefineType") = DBNull.Value
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = DBNull.Value
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)

                For Each child As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(nsTable, sourceFile, defineFullName, "Structure", child)
                Next


            Case SyntaxKind.InterfaceBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of InterfaceStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' ジェネリックを定義している場合、個数を追加
                If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                ' 内側のコンテナの場合、ドットつなぎではなくプラスつなぎにする
                Dim defineFullName = $"{containerFullName}.{defineName}"
                Dim containerKeywords = New List(Of String) From {"Class", "Structure", "Interface", "Module"}
                If containerKeywords.Contains(containerDefineKind) Then
                    defineFullName = $"{containerFullName}+{defineName}"
                End If

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Interface"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = defineFullName
                row("DefineType") = DBNull.Value
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = DBNull.Value
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)

                For Each child As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(nsTable, sourceFile, defineFullName, "Interface", child)
                Next


            Case SyntaxKind.ModuleBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of ModuleStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' (VB2015 時点)、モジュールは、ジェネリック定義はできない仕様

                ' 内側のコンテナの場合、ドットつなぎではなくプラスつなぎにする
                ' → モジュールは他のコンテナに含めることが出来ない仕様かも
                Dim defineFullName = $"{containerFullName}.{defineName}"
                Dim containerKeywords = New List(Of String) From {"Class", "Structure", "Interface", "Module"}
                If containerKeywords.Contains(containerDefineKind) Then
                    defineFullName = $"{containerFullName}+{defineName}"
                End If

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Module"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = defineFullName
                row("DefineType") = DBNull.Value
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = DBNull.Value
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)

                For Each child As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.ParseInternal(nsTable, sourceFile, defineFullName, "Module", child)
                Next


            Case SyntaxKind.EnumBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of EnumStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' 内側のコンテナの場合、ドットつなぎではなくプラスつなぎにする
                Dim defineFullName = $"{containerFullName}.{defineName}"
                Dim containerKeywords = New List(Of String) From {"Class", "Structure", "Interface", "Module"}
                If containerKeywords.Contains(containerDefineKind) Then
                    defineFullName = $"{containerFullName}+{defineName}"
                End If

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Enum"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = defineFullName
                row("DefineType") = DBNull.Value
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = DBNull.Value
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


            Case SyntaxKind.ConstructorBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of SubNewStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.NewKeyword).ToString()

                ' ジェネリックを定義している場合、個数を追加
                If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                Dim listNode = statementNode.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
                Dim methodArguments = Me.GetMethodArguments(listNode)

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Constructor"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = $"{containerFullName}.{defineName}"
                row("DefineType") = DBNull.Value
                row("ReturnType") = "Void"
                row("MethodArguments") = methodArguments
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


            Case SyntaxKind.OperatorBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of OperatorStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ToString()
                defineName = defineName.Substring(defineName.IndexOf("Operator ") + "Operator ".Length)
                defineName = defineName.Substring(0, defineName.IndexOf("("))

                ' ビルド後は、記号名から対応する文字列名称に変換される
                Dim langTable = MemoryDB.Instance.DB.Tables("LanguageConversion")
                For Each langRow As DataRow In langTable.Rows

                    Dim vbType = CStr(langRow("VBNet"))
                    Dim fxType = CStr(langRow("NETFramework"))

                    If defineName = vbType Then
                        defineName = fxType
                        Exit For
                    End If

                Next

                ' ジェネリックを定義している場合、個数を追加
                ' → オペレータはジェネリック定義できない仕様かも
                If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                Dim listNode = statementNode.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
                Dim methodArguments = Me.GetMethodArguments(listNode)

                Dim typeNode = statementNode.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
                Dim typeName = Me.GetTypeName(typeNode)

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Operator"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = $"{containerFullName}.{defineName}"
                row("DefineType") = DBNull.Value
                row("ReturnType") = typeName
                row("MethodArguments") = methodArguments
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


            Case SyntaxKind.SubBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of MethodStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' ジェネリックを定義している場合、個数を追加
                If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                Dim listNode = statementNode.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
                Dim methodArguments = Me.GetMethodArguments(listNode)

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Method"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = $"{containerFullName}.{defineName}"
                row("DefineType") = DBNull.Value
                row("ReturnType") = "Void"
                row("MethodArguments") = methodArguments
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


            Case SyntaxKind.FunctionBlock

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim statementNode = node.ChildNodes().OfType(Of MethodStatementSyntax)().FirstOrDefault()
                Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' ジェネリックを定義している場合、個数を追加
                If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                Dim listNode = statementNode.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
                Dim methodArguments = Me.GetMethodArguments(listNode)

                Dim typeNode = statementNode.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
                Dim typeName = Me.GetTypeName(typeNode)

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Method"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = $"{containerFullName}.{defineName}"
                row("DefineType") = DBNull.Value
                row("ReturnType") = typeName
                row("MethodArguments") = methodArguments
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


            Case SyntaxKind.FieldDeclaration

                Dim isPartial = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)

                Dim declaratorNodes = node.ChildNodes().OfType(Of VariableDeclaratorSyntax)()
                For Each declaratorNode In declaratorNodes

                    ' 先に型を取得してしまう
                    Dim typeNode = declaratorNode.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
                    Dim typeName = Me.GetTypeName(typeNode)

                    ' フィールド名の取得
                    Dim fieldNodes = declaratorNode.ChildNodes().OfType(Of ModifiedIdentifierSyntax)()
                    For Each fieldNode In fieldNodes

                        Dim startLength = fieldNode.Span.Start
                        Dim endLength = fieldNode.Span.End
                        Dim defineName = fieldNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
                        Dim bracket = String.Empty

                        ' 変数名にカッコが付いている場合、型に移す
                        If fieldNode.ChildNodes().OfType(Of ArrayRankSpecifierSyntax)().Any() Then

                            Dim arrayRankNodes = fieldNode.ChildNodes().OfType(Of ArrayRankSpecifierSyntax)()
                            For Each arrayRankNode In arrayRankNodes
                                bracket &= arrayRankNode.ToString()
                            Next

                        End If

                        Dim row = nsTable.NewRow()
                        row("DefineKind") = "Field"
                        row("ContainerFullName") = containerFullName
                        row("DefineFullName") = $"{containerFullName}.{defineName}"
                        row("DefineType") = $"{typeName}{bracket}"
                        row("ReturnType") = DBNull.Value
                        row("MethodArguments") = DBNull.Value
                        row("IsPartial") = isPartial
                        row("IsShared") = isShared
                        row("SourceFile") = sourceFile
                        row("StartLength") = startLength
                        row("EndLength") = endLength
                        row("StartLineNumber") = -1
                        row("EndLineNumber") = -1
                        nsTable.Rows.Add(row)

                    Next

                Next


            Case SyntaxKind.PropertyBlock, SyntaxKind.PropertyStatement
                ' PropertyBlock / 通常プロパティ用（setter, getter いづれかの記述あり）
                ' PropertyStatement / 自動実装プロパティ用（setter, getter 記述無し）

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                If kind = SyntaxKind.PropertyBlock Then
                    node = node.ChildNodes().OfType(Of PropertyStatementSyntax)().FirstOrDefault()
                End If

                Dim isPartial = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)

                Dim defineName = node.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
                Dim typeNode = node.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
                Dim typeName = Me.GetTypeName(typeNode)

                Dim listNode = node.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
                Dim methodArguments = Me.GetMethodArguments(listNode)

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Property"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = $"{containerFullName}.{defineName}"
                row("DefineType") = typeName
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = methodArguments
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


            Case SyntaxKind.DelegateSubStatement

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim isPartial = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = node.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' ジェネリックを定義している場合、個数を追加
                If node.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = node.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                ' 内側のコンテナの場合、ドットつなぎではなくプラスつなぎにする
                Dim defineFullName = $"{containerFullName}.{defineName}"
                Dim containerKeywords = New List(Of String) From {"Class", "Structure", "Interface", "Module"}
                If containerKeywords.Contains(containerDefineKind) Then
                    defineFullName = $"{containerFullName}+{defineName}"
                End If

                Dim listNode = node.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
                Dim methodArguments = Me.GetMethodArguments(listNode)

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Delegate"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = defineFullName
                row("DefineType") = DBNull.Value
                row("ReturnType") = "Void"
                row("MethodArguments") = methodArguments
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


            Case SyntaxKind.DelegateFunctionStatement

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim isPartial = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
                Dim defineName = node.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()

                ' ジェネリックを定義している場合、個数を追加
                If node.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

                    Dim genericRangeNode = node.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
                    Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
                    defineName = $"{defineName}`{genericNodes.Count()}"

                End If

                ' 内側のコンテナの場合、ドットつなぎではなくプラスつなぎにする
                Dim defineFullName = $"{containerFullName}.{defineName}"
                Dim containerKeywords = New List(Of String) From {"Class", "Structure", "Interface", "Module"}
                If containerKeywords.Contains(containerDefineKind) Then
                    defineFullName = $"{containerFullName}+{defineName}"
                End If

                Dim listNode = node.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
                Dim methodArguments = Me.GetMethodArguments(listNode)

                Dim typeNode = node.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
                Dim typeName = Me.GetTypeName(typeNode)

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Delegate"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = defineFullName
                row("DefineType") = DBNull.Value
                row("ReturnType") = typeName
                row("MethodArguments") = methodArguments
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


            Case SyntaxKind.EventBlock, SyntaxKind.EventStatement

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                If kind = SyntaxKind.EventBlock Then
                    node = node.ChildNodes().OfType(Of EventStatementSyntax)().FirstOrDefault()
                End If

                Dim isPartial = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
                Dim isShared = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)

                Dim defineName = node.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
                Dim typeNode = node.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
                Dim typeName = String.Empty

                If typeNode IsNot Nothing Then
                    typeName = Me.GetTypeName(typeNode)
                End If

                Dim listNode = node.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
                Dim methodArguments = String.Empty

                If listNode IsNot Nothing Then
                    methodArguments = Me.GetMethodArguments(listNode)
                End If

                Dim row = nsTable.NewRow()
                row("DefineKind") = "Event"
                row("ContainerFullName") = containerFullName
                row("DefineFullName") = $"{containerFullName}.{defineName}"
                row("DefineType") = typeName
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = methodArguments
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                row("StartLineNumber") = -1
                row("EndLineNumber") = -1
                nsTable.Rows.Add(row)


        End Select

    End Sub

    Private Function GetTypeName(node As SyntaxNode) As String

        Dim item = "Object"
        If node Is Nothing Then
            Return item
        End If

        If TypeOf node IsNot SimpleAsClauseSyntax Then
            Return item
        End If

        Select Case True

            Case node.ChildNodes().OfType(Of GenericNameSyntax)().Any()

                Dim childNode = node.ChildNodes().OfType(Of GenericNameSyntax)().FirstOrDefault()
                item = childNode.ToString()


            Case node.ChildNodes().OfType(Of ArrayTypeSyntax)().Any()

                Dim childNode = node.ChildNodes().OfType(Of ArrayTypeSyntax)().FirstOrDefault()
                item = childNode.ToString()


            Case node.ChildNodes().OfType(Of QualifiedNameSyntax)().Any()

                Dim childNode = node.ChildNodes().OfType(Of QualifiedNameSyntax)().FirstOrDefault()
                item = childNode.ToString()


            Case node.ChildNodes().OfType(Of PredefinedTypeSyntax)().Any()

                Dim childNode = node.ChildNodes().OfType(Of PredefinedTypeSyntax)().FirstOrDefault()
                item = childNode.ToString()


            Case node.ChildNodes().OfType(Of IdentifierNameSyntax)().Any()

                Dim childNode = node.ChildNodes().OfType(Of IdentifierNameSyntax)().FirstOrDefault()
                item = childNode.ToString()


        End Select

        Return item

    End Function

    Private Function GetMethodArguments(node As SyntaxNode) As String

        Dim item = "0()"
        If node Is Nothing Then
            Return item
        End If

        If TypeOf node IsNot ParameterListSyntax Then
            Return item
        End If

        item = String.Empty
        Dim count = node.ChildNodes().OfType(Of ParameterSyntax)().Count()
        For Each parameterNode In node.ChildNodes().OfType(Of ParameterSyntax)()

            If item <> String.Empty Then
                item &= ", "
            End If

            Dim simpleNode = parameterNode.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
            Dim tmpItem = Me.GetTypeName(simpleNode)
            item &= Me.RemoveNamespace(tmpItem)

        Next

        item = $"{count}({item})"
        Return item

    End Function

    Private Function RemoveNamespace(parameterType As String) As String

        If parameterType.Contains(".") Then
            parameterType = parameterType.Substring(parameterType.LastIndexOf(".") + 1)
        End If

        Return parameterType

    End Function

    Public Iterator Function GetFoldingItems(source As String) As IEnumerable(Of Tuple(Of Integer, Integer, String))

        Dim tree = VisualBasicSyntaxTree.ParseText(source)
        Dim root = TryCast(tree.GetRoot(), VisualBasicSyntaxNode)
        Dim items = Me.GetFoldingItemsInternal(root)

        For Each item In items
            Yield item
        Next

    End Function

    Private Iterator Function GetFoldingItemsInternal(node As VisualBasicSyntaxNode) As IEnumerable(Of Tuple(Of Integer, Integer, String))

        Dim kind = node.Kind()

        Select Case kind

            Case SyntaxKind.CompilationUnit
                ' #Region

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

                            Yield Tuple.Create(startLength, endLength, header)

                    End Select

                Next


            Case SyntaxKind.NamespaceBlock,
                 SyntaxKind.ClassBlock,
                 SyntaxKind.StructureBlock,
                 SyntaxKind.InterfaceBlock,
                 SyntaxKind.ModuleBlock,
                 SyntaxKind.EnumBlock,
                 SyntaxKind.ConstructorBlock,
                 SyntaxKind.OperatorBlock,
                 SyntaxKind.SubBlock,
                 SyntaxKind.FunctionBlock,
                 SyntaxKind.PropertyBlock,
                 SyntaxKind.GetAccessorBlock,
                 SyntaxKind.SetAccessorBlock,
                 SyntaxKind.EventBlock,
                 SyntaxKind.AddHandlerAccessorBlock,
                 SyntaxKind.RemoveHandlerAccessorBlock,
                 SyntaxKind.RaiseEventAccessorBlock

                ' ブロック系は、開始ノードが、開始と終了の両方の文字列位置を持っているので、取得する

                Dim startLength = node.Span.Start
                Dim endLength = node.Span.End

                Dim header = node.ToString()
                If header.Contains(Environment.NewLine) Then
                    header = header.Substring(0, header.IndexOf(Environment.NewLine))
                    header = $"{header} ..."
                End If

                Yield Tuple.Create(startLength, endLength, header)

        End Select

        ' Namespace -> Class -> Sub Method などのように、再帰して取得
        For Each child As VisualBasicSyntaxNode In node.ChildNodes()

            Dim childItems = Me.GetFoldingItemsInternal(child)
            For Each grandChild In childItems
                Yield grandChild
            Next

        Next

    End Function

    Public Function GetMethodIndentShape(methodRange As String) As Border

        Dim tree = VisualBasicSyntaxTree.ParseText(methodRange)
        Dim root = TryCast(tree.GetRoot(), VisualBasicSyntaxNode)
        Dim node = root.DescendantNodes().FirstOrDefault(Function(x) (x.Kind() = SyntaxKind.SubNewStatement) OrElse (x.Kind() = SyntaxKind.SubStatement) OrElse (x.Kind() = SyntaxKind.FunctionStatement) OrElse (x.Kind() = SyntaxKind.OperatorStatement))

        Dim border1 = New Border
        border1.BorderBrush = Brushes.Pink
        border1.BorderThickness = New Thickness(1)
        border1.Background = Brushes.FloralWhite
        border1.CornerRadius = New CornerRadius(8)
        border1.Margin = New Thickness(20)

        Dim dockpanel1 = New DockPanel
        border1.Child = dockpanel1

        Dim textblock1 = New TextBlock
        dockpanel1.Children.Add(textblock1)
        DockPanel.SetDock(textblock1, Dock.Top)
        textblock1.Text = "メソッド"
        textblock1.Margin = New Thickness(10, 5, 0, 0)

        Dim line1 = New Line
        dockpanel1.Children.Add(line1)
        DockPanel.SetDock(line1, Dock.Top)
        line1.Stroke = Brushes.Gainsboro
        line1.StrokeThickness = 1
        line1.Margin = New Thickness(10, 0, 10, 5)

        Dim bind1 = New Binding("ActualWidth")
        bind1.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
        line1.SetBinding(Line.X2Property, bind1)

        Dim textblock2 = New TextBlock
        dockpanel1.Children.Add(textblock2)
        DockPanel.SetDock(textblock2, Dock.Top)
        textblock2.Text = node.ToString()
        textblock2.Margin = New Thickness(10, 0, 50, 0)

        Dim itemscontrol1 = New ItemsControl
        dockpanel1.Children.Add(itemscontrol1)
        itemscontrol1.Margin = New Thickness(20, 0, 0, 20)

        Me.GetMethodIndentShapeInternal(root, itemscontrol1)

        Return border1

    End Function

    Private Sub GetMethodIndentShapeInternal(node As VisualBasicSyntaxNode, itemscontrol1 As ItemsControl)

        Select Case node.Kind()

            Case SyntaxKind.SingleLineIfStatement

                Dim textblock1 = New TextBlock
                itemscontrol1.Items.Add(textblock1)
                textblock1.Text = "↓"
                textblock1.Foreground = Brushes.Gainsboro
                textblock1.Margin = New Thickness(20, 10, 20, 0)

                Dim border1 = New Border
                itemscontrol1.Items.Add(border1)
                border1.BorderBrush = Brushes.ForestGreen
                border1.BorderThickness = New Thickness(1)
                border1.Background = Brushes.Honeydew
                border1.CornerRadius = New CornerRadius(8)
                border1.Margin = New Thickness(10)

                Dim dockpanel1 = New DockPanel
                border1.Child = dockpanel1

                Dim textblock2 = New TextBlock
                dockpanel1.Children.Add(textblock2)
                DockPanel.SetDock(textblock2, Dock.Top)   ' DockPanel は、DockPanel.SetDock(コントロール名、値)
                textblock2.Margin = New Thickness(10, 5, 10, 0)
                textblock2.Text = "If"

                Dim line2 = New Line
                dockpanel1.Children.Add(line2)
                DockPanel.SetDock(line2, Dock.Top)
                line2.Stroke = Brushes.Gainsboro
                line2.StrokeThickness = 1
                line2.Margin = New Thickness(10, 0, 10, 5)

                Dim bind2 = New Binding("ActualWidth")
                bind2.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                line2.SetBinding(Line.X2Property, bind2)

                Dim itemscontrol2 = New ItemsControl
                dockpanel1.Children.Add(itemscontrol2)
                itemscontrol2.Margin = New Thickness(20, 0, 0, 0)

                For Each grandChild As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.GetMethodIndentShapeInternal(grandChild, itemscontrol2)
                Next

                Dim textblock3 = New TextBlock
                itemscontrol1.Items.Add(textblock3)
                textblock3.Text = "↓"
                textblock3.Foreground = Brushes.Gainsboro
                textblock3.Margin = New Thickness(20, 0, 20, 10)


            Case SyntaxKind.MultiLineIfBlock

                Dim textblock1 = New TextBlock
                itemscontrol1.Items.Add(textblock1)
                textblock1.Text = "↓"
                textblock1.Foreground = Brushes.Gainsboro
                textblock1.Margin = New Thickness(20, 10, 20, 0)

                ' 条件分岐の数分、列を追加
                Dim childNodes = node.ChildNodes().Where(Function(x) (TypeOf x Is IfStatementSyntax) OrElse (TypeOf x Is ElseIfBlockSyntax) OrElse (TypeOf x Is ElseBlockSyntax))
                Dim grid1 = New Grid
                itemscontrol1.Items.Add(grid1)

                For i As Integer = 0 To childNodes.Count() - 1
                    grid1.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})
                Next

                ' 行は固定（合流の横線、各分岐開始の矢印線、各分岐、各分岐終了の矢印線、合流の横線）
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})

                Dim line1 = New Line
                grid1.Children.Add(line1)
                line1.SetValue(Grid.RowProperty, 0)
                line1.SetValue(Grid.ColumnProperty, 0)
                line1.SetValue(Grid.ColumnSpanProperty, childNodes.Count())
                line1.Stroke = Brushes.Gainsboro
                line1.StrokeThickness = 1
                line1.Margin = New Thickness(25, 0, 50, 0)

                Dim bind1 = New Binding("ActualWidth")
                bind1.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                line1.SetBinding(Line.X2Property, bind1)       ' Binding は、コントロール名.SetBinding(コントロールのプロパティ、値)

                ' 条件分岐の数だけ、開始矢印を追加
                For i As Integer = 0 To childNodes.Count() - 1

                    Dim textblock2 = New TextBlock
                    grid1.Children.Add(textblock2)
                    textblock2.Text = "↓"
                    textblock2.Foreground = Brushes.Gainsboro
                    textblock2.SetValue(Grid.RowProperty, 1)
                    textblock2.SetValue(Grid.ColumnProperty, i)

                    Dim islastNode = (i = childNodes.Count() - 1)
                    If islastNode Then
                        textblock2.HorizontalAlignment = HorizontalAlignment.Right
                        textblock2.Margin = New Thickness(0, 0, 45, 0)
                    Else
                        textblock2.HorizontalAlignment = HorizontalAlignment.Center
                    End If

                Next

                ' 条件分岐の命令
                For i As Integer = 0 To childNodes.Count() - 1

                    Dim child = childNodes(i)

                    Dim border1 = New Border
                    grid1.Children.Add(border1)
                    border1.SetValue(Grid.RowProperty, 2)     ' Grid は、コントロール名.SetValue(Gridプロパティ、値）
                    border1.SetValue(Grid.ColumnProperty, i)
                    border1.BorderBrush = Brushes.ForestGreen
                    border1.BorderThickness = New Thickness(1)
                    border1.Background = Brushes.Honeydew
                    border1.CornerRadius = New CornerRadius(8)
                    border1.Margin = New Thickness(10)

                    Dim dockpanel1 = New DockPanel
                    border1.Child = dockpanel1

                    Dim textblock2 = New TextBlock
                    dockpanel1.Children.Add(textblock2)
                    DockPanel.SetDock(textblock2, Dock.Top)   ' DockPanel は、DockPanel.SetDock(コントロール名、値)
                    textblock2.Margin = New Thickness(10, 5, 10, 0)

                    Select Case True
                        Case TypeOf child Is IfStatementSyntax : textblock2.Text = "If"
                        Case TypeOf child Is ElseIfBlockSyntax : textblock2.Text = "ElseIf"
                        Case TypeOf child Is ElseBlockSyntax : textblock2.Text = "Else"
                    End Select

                    Dim line2 = New Line
                    dockpanel1.Children.Add(line2)
                    DockPanel.SetDock(line2, Dock.Top)
                    line2.Stroke = Brushes.Gainsboro
                    line2.StrokeThickness = 1
                    line2.Margin = New Thickness(10, 0, 10, 5)

                    Dim bind2 = New Binding("ActualWidth")
                    bind2.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                    line2.SetBinding(Line.X2Property, bind2)

                    Dim itemscontrol2 = New ItemsControl
                    dockpanel1.Children.Add(itemscontrol2)
                    itemscontrol2.Margin = New Thickness(20, 0, 0, 0)

                    ' If の場合、Block ではなく Statement なので、この下階層（子リスト）には、ループや条件分岐が（ソース的にはあったとしても）含まれない
                    ' １階層上にさかのぼって、ElseIfBlock, または ElseBlock 以外で、ループや条件分岐がある場合、各自を再帰する
                    If TypeOf child Is IfStatementSyntax Then

                        Dim innerBlocks = node.ChildNodes().Where(Function(x) (TypeOf x IsNot IfStatementSyntax) AndAlso (TypeOf x IsNot ElseIfBlockSyntax) AndAlso (TypeOf x IsNot ElseBlockSyntax))
                        For Each innerBlock As VisualBasicSyntaxNode In innerBlocks
                            Me.GetMethodIndentShapeInternal(innerBlock, itemscontrol2)
                        Next

                    Else

                        For Each grandChild As VisualBasicSyntaxNode In child.ChildNodes()
                            Me.GetMethodIndentShapeInternal(grandChild, itemscontrol2)
                        Next

                    End If

                Next

                ' 条件分岐の数だけ、終了の矢印を追加
                For i As Integer = 0 To childNodes.Count() - 1

                    Dim textblock2 = New TextBlock
                    grid1.Children.Add(textblock2)
                    textblock2.Text = "↓"
                    textblock2.Foreground = Brushes.Gainsboro
                    textblock2.SetValue(Grid.RowProperty, 3)
                    textblock2.SetValue(Grid.ColumnProperty, i)

                    Dim islastNode = (i = childNodes.Count() - 1)
                    If islastNode Then
                        textblock2.HorizontalAlignment = HorizontalAlignment.Right
                        textblock2.Margin = New Thickness(0, 0, 45, 0)
                    Else
                        textblock2.HorizontalAlignment = HorizontalAlignment.Center
                    End If

                Next

                Dim line3 = New Line
                grid1.Children.Add(line3)
                line3.SetValue(Grid.RowProperty, 4)
                line3.SetValue(Grid.ColumnProperty, 0)
                line3.SetValue(Grid.ColumnSpanProperty, childNodes.Count())
                line3.Stroke = Brushes.Gainsboro
                line3.StrokeThickness = 1
                line3.Margin = New Thickness(25, 0, 50, 0)

                Dim bind3 = New Binding("ActualWidth")
                bind3.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                line3.SetBinding(Line.X2Property, bind3)

                Dim textblock3 = New TextBlock
                itemscontrol1.Items.Add(textblock3)
                textblock3.Text = "↓"
                textblock3.Foreground = Brushes.Gainsboro
                textblock3.Margin = New Thickness(20, 0, 20, 10)


            Case SyntaxKind.SelectBlock

                Dim textblock1 = New TextBlock
                itemscontrol1.Items.Add(textblock1)
                textblock1.Text = "↓"
                textblock1.Foreground = Brushes.Gainsboro
                textblock1.Margin = New Thickness(20, 10, 20, 0)

                ' 条件分岐の数分、列を追加
                Dim childNodes = node.ChildNodes().Where(Function(x) (TypeOf x Is CaseBlockSyntax))
                Dim grid1 = New Grid
                itemscontrol1.Items.Add(grid1)

                For i As Integer = 0 To childNodes.Count() - 1
                    grid1.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})
                Next

                ' 行は固定（合流の横線、各分岐開始の矢印線、各分岐、各分岐終了の矢印線、合流の横線）
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
                grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})

                Dim line1 = New Line
                grid1.Children.Add(line1)
                line1.SetValue(Grid.RowProperty, 0)
                line1.SetValue(Grid.ColumnProperty, 0)
                line1.SetValue(Grid.ColumnSpanProperty, childNodes.Count())
                line1.Stroke = Brushes.Gainsboro
                line1.StrokeThickness = 1
                line1.Margin = New Thickness(25, 0, 50, 0)

                Dim bind1 = New Binding("ActualWidth")
                bind1.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                line1.SetBinding(Line.X2Property, bind1)       ' Binding は、コントロール名.SetBinding(コントロールのプロパティ、値)

                ' 条件分岐の数だけ、開始矢印を追加
                For i As Integer = 0 To childNodes.Count() - 1

                    Dim textblock2 = New TextBlock
                    grid1.Children.Add(textblock2)
                    textblock2.Text = "↓"
                    textblock2.Foreground = Brushes.Gainsboro
                    textblock2.SetValue(Grid.RowProperty, 1)
                    textblock2.SetValue(Grid.ColumnProperty, i)

                    Dim islastNode = (i = childNodes.Count() - 1)
                    If islastNode Then
                        textblock2.HorizontalAlignment = HorizontalAlignment.Right
                        textblock2.Margin = New Thickness(0, 0, 45, 0)
                    Else
                        textblock2.HorizontalAlignment = HorizontalAlignment.Center
                    End If

                Next

                ' 条件分岐の命令
                For i As Integer = 0 To childNodes.Count() - 1

                    Dim child = childNodes(i)

                    Dim border1 = New Border
                    grid1.Children.Add(border1)
                    border1.SetValue(Grid.RowProperty, 2)     ' Grid は、コントロール名.SetValue(Gridプロパティ、値）
                    border1.SetValue(Grid.ColumnProperty, i)
                    border1.BorderBrush = Brushes.ForestGreen
                    border1.BorderThickness = New Thickness(1)
                    border1.Background = Brushes.Honeydew
                    border1.CornerRadius = New CornerRadius(8)
                    border1.Margin = New Thickness(10)

                    Dim dockpanel1 = New DockPanel
                    border1.Child = dockpanel1

                    Dim textblock2 = New TextBlock
                    dockpanel1.Children.Add(textblock2)
                    DockPanel.SetDock(textblock2, Dock.Top)   ' DockPanel は、DockPanel.SetDock(コントロール名、値)
                    textblock2.Margin = New Thickness(10, 5, 10, 0)

                    Select Case True
                        Case TypeOf child Is CaseBlockSyntax

                            If i = 0 Then
                                textblock2.Text = "Select Case"
                            ElseIf child.DescendantNodes().Any(Function(x) TypeOf x Is ElseCaseClauseSyntax) Then
                                textblock2.Text = "Case Else"
                            Else
                                textblock2.Text = "Case"
                            End If

                    End Select

                    Dim line2 = New Line
                    dockpanel1.Children.Add(line2)
                    DockPanel.SetDock(line2, Dock.Top)
                    line2.Stroke = Brushes.Gainsboro
                    line2.StrokeThickness = 1
                    line2.Margin = New Thickness(10, 0, 10, 5)

                    Dim bind2 = New Binding("ActualWidth")
                    bind2.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                    line2.SetBinding(Line.X2Property, bind2)

                    Dim itemscontrol2 = New ItemsControl
                    dockpanel1.Children.Add(itemscontrol2)
                    itemscontrol2.Margin = New Thickness(20, 0, 0, 0)

                    For Each grandChild As VisualBasicSyntaxNode In child.ChildNodes()
                        Me.GetMethodIndentShapeInternal(grandChild, itemscontrol2)
                    Next

                Next

                ' 条件分岐の数だけ、終了の矢印を追加
                For i As Integer = 0 To childNodes.Count() - 1

                    Dim textblock2 = New TextBlock
                    grid1.Children.Add(textblock2)
                    textblock2.Text = "↓"
                    textblock2.Foreground = Brushes.Gainsboro
                    textblock2.SetValue(Grid.RowProperty, 3)
                    textblock2.SetValue(Grid.ColumnProperty, i)

                    Dim islastNode = (i = childNodes.Count() - 1)
                    If islastNode Then
                        textblock2.HorizontalAlignment = HorizontalAlignment.Right
                        textblock2.Margin = New Thickness(0, 0, 45, 0)
                    Else
                        textblock2.HorizontalAlignment = HorizontalAlignment.Center
                    End If

                Next

                Dim line3 = New Line
                grid1.Children.Add(line3)
                line3.SetValue(Grid.RowProperty, 4)
                line3.SetValue(Grid.ColumnProperty, 0)
                line3.SetValue(Grid.ColumnSpanProperty, childNodes.Count())
                line3.Stroke = Brushes.Gainsboro
                line3.StrokeThickness = 1
                line3.Margin = New Thickness(25, 0, 50, 0)

                Dim bind3 = New Binding("ActualWidth")
                bind3.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                line3.SetBinding(Line.X2Property, bind3)


            Case SyntaxKind.ForBlock, SyntaxKind.ForEachBlock, SyntaxKind.WhileBlock, SyntaxKind.SimpleDoLoopBlock, SyntaxKind.DoLoopUntilBlock, SyntaxKind.DoLoopWhileBlock, SyntaxKind.UsingBlock

                Dim textblock1 = New TextBlock
                itemscontrol1.Items.Add(textblock1)
                textblock1.Text = "↓"
                textblock1.Foreground = Brushes.Gainsboro
                textblock1.Margin = New Thickness(20, 10, 20, 0)

                Dim border1 = New Border
                itemscontrol1.Items.Add(border1)
                border1.BorderBrush = Brushes.RoyalBlue
                border1.BorderThickness = New Thickness(1)
                border1.Background = Brushes.AliceBlue
                border1.CornerRadius = New CornerRadius(8)
                border1.Margin = New Thickness(10)

                Dim dockpanel1 = New DockPanel
                border1.Child = dockpanel1

                Dim textblock2 = New TextBlock
                dockpanel1.Children.Add(textblock2)
                DockPanel.SetDock(textblock2, Dock.Top)   ' DockPanel は、DockPanel.SetDock(コントロール名、値)
                textblock2.Margin = New Thickness(10, 5, 10, 0)

                Select Case True
                    Case node.Kind() = SyntaxKind.ForBlock : textblock2.Text = "For"
                    Case node.Kind() = SyntaxKind.ForEachBlock : textblock2.Text = "ForEach"
                    Case node.Kind() = SyntaxKind.WhileBlock : textblock2.Text = "While"
                    Case node.Kind() = SyntaxKind.SimpleDoLoopBlock : textblock2.Text = "Do-Loop"
                    Case node.Kind() = SyntaxKind.DoLoopUntilBlock : textblock2.Text = "Do-Until"
                    Case node.Kind() = SyntaxKind.DoLoopWhileBlock : textblock2.Text = "Do-While"
                    Case node.Kind() = SyntaxKind.UsingBlock : textblock2.Text = "Using"
                End Select

                Dim line2 = New Line
                dockpanel1.Children.Add(line2)
                DockPanel.SetDock(line2, Dock.Top)
                line2.Stroke = Brushes.Gainsboro
                line2.StrokeThickness = 1
                line2.Margin = New Thickness(10, 0, 10, 5)

                Dim bind2 = New Binding("ActualWidth")
                bind2.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                line2.SetBinding(Line.X2Property, bind2)

                Dim itemscontrol2 = New ItemsControl
                dockpanel1.Children.Add(itemscontrol2)
                itemscontrol2.Margin = New Thickness(20, 0, 0, 0)

                For Each grandChild As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.GetMethodIndentShapeInternal(grandChild, itemscontrol2)
                Next


            Case SyntaxKind.TryBlock

                Dim textblock1 = New TextBlock
                itemscontrol1.Items.Add(textblock1)
                textblock1.Text = "↓"
                textblock1.Foreground = Brushes.Gainsboro
                textblock1.Margin = New Thickness(20, 10, 20, 0)

                ' Try のブロック
                Dim border1 = New Border
                itemscontrol1.Items.Add(border1)
                border1.BorderBrush = Brushes.Tomato
                border1.BorderThickness = New Thickness(1)
                border1.Background = Brushes.Linen
                border1.CornerRadius = New CornerRadius(8)
                border1.Margin = New Thickness(10)

                Dim dockpanel1 = New DockPanel
                border1.Child = dockpanel1

                Dim textblock2 = New TextBlock
                dockpanel1.Children.Add(textblock2)
                DockPanel.SetDock(textblock2, Dock.Top)   ' DockPanel は、DockPanel.SetDock(コントロール名、値)
                textblock2.Text = "Try"
                textblock2.Margin = New Thickness(10, 5, 10, 0)

                Dim line2 = New Line
                dockpanel1.Children.Add(line2)
                DockPanel.SetDock(line2, Dock.Top)
                line2.Stroke = Brushes.Gainsboro
                line2.StrokeThickness = 1
                line2.Margin = New Thickness(10, 0, 10, 5)

                Dim bind2 = New Binding("ActualWidth")
                bind2.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                line2.SetBinding(Line.X2Property, bind2)

                Dim itemscontrol2 = New ItemsControl
                dockpanel1.Children.Add(itemscontrol2)
                itemscontrol2.Margin = New Thickness(20, 0, 0, 0)

                Dim innerBlocks = node.ChildNodes().Where(Function(x) (TypeOf x IsNot TryStatementSyntax) AndAlso (TypeOf x IsNot CatchBlockSyntax) AndAlso (TypeOf x IsNot FinallyBlockSyntax))
                For Each innerBlock As VisualBasicSyntaxNode In innerBlocks
                    Me.GetMethodIndentShapeInternal(innerBlock, itemscontrol2)
                Next

                Dim textblock3 = New TextBlock
                itemscontrol1.Items.Add(textblock3)
                textblock3.Text = "↓"
                textblock3.Foreground = Brushes.Gainsboro
                textblock3.Margin = New Thickness(20, 10, 20, 0)

                ' Catch ブロック
                If node.ChildNodes().Any(Function(x) TypeOf x Is CatchBlockSyntax) Then

                    Dim innerBlocks2 = node.ChildNodes().Where(Function(x) TypeOf x Is CatchBlockSyntax)
                    For Each innerBlock2 In innerBlocks2

                        Dim border2 = New Border
                        itemscontrol1.Items.Add(border2)
                        border2.BorderBrush = Brushes.Tomato
                        border2.BorderThickness = New Thickness(1)
                        border2.Background = Brushes.Linen
                        border2.CornerRadius = New CornerRadius(8)
                        border2.Margin = New Thickness(10)

                        Dim dockpanel2 = New DockPanel
                        border2.Child = dockpanel2

                        Dim textblock4 = New TextBlock
                        dockpanel2.Children.Add(textblock4)
                        DockPanel.SetDock(textblock4, Dock.Top)   ' DockPanel は、DockPanel.SetDock(コントロール名、値)
                        textblock4.Text = "Catch"
                        textblock4.Margin = New Thickness(10, 5, 10, 0)

                        Dim line3 = New Line
                        dockpanel2.Children.Add(line3)
                        DockPanel.SetDock(line3, Dock.Top)
                        line3.Stroke = Brushes.Gainsboro
                        line3.StrokeThickness = 1
                        line3.Margin = New Thickness(10, 0, 10, 5)

                        Dim bind3 = New Binding("ActualWidth")
                        bind3.RelativeSource = New RelativeSource(RelativeSourceMode.Self)
                        line3.SetBinding(Line.X2Property, bind3)

                        Dim itemscontrol3 = New ItemsControl
                        dockpanel2.Children.Add(itemscontrol3)
                        itemscontrol3.Margin = New Thickness(20, 0, 0, 0)

                        For Each inner As VisualBasicSyntaxNode In innerBlock2.ChildNodes()
                            Me.GetMethodIndentShapeInternal(inner, itemscontrol3)
                        Next

                        Dim textblock5 = New TextBlock
                        itemscontrol1.Items.Add(textblock5)
                        textblock5.Text = "↓"
                        textblock5.Foreground = Brushes.Gainsboro
                        textblock5.Margin = New Thickness(20, 10, 20, 0)

                    Next

                End If

                ' Finally ブロック
                If node.ChildNodes().Any(Function(x) TypeOf x Is FinallyBlockSyntax) Then

                    Dim innerBlocks2 = node.ChildNodes().Where(Function(x) TypeOf x Is FinallyBlockSyntax)
                    For Each innerBlock2 In innerBlocks2

                        Dim border2 = New Border
                        itemscontrol1.Items.Add(border2)
                        border2.BorderBrush = Brushes.Tomato
                        border2.BorderThickness = New Thickness(1)
                        border2.Background = Brushes.Linen
                        border2.CornerRadius = New CornerRadius(8)
                        border2.Margin = New Thickness(10)

                        Dim dockpanel2 = New DockPanel
                        border2.Child = dockpanel2

                        Dim textblock4 = New TextBlock
                        dockpanel2.Children.Add(textblock4)
                        DockPanel.SetDock(textblock4, Dock.Top)   ' DockPanel は、DockPanel.SetDock(コントロール名、値)
                        textblock4.Text = "Finally"
                        textblock4.Margin = New Thickness(10, 5, 10, 0)

                        Dim itemscontrol3 = New ItemsControl
                        dockpanel2.Children.Add(itemscontrol3)
                        itemscontrol3.Margin = New Thickness(20, 0, 0, 0)

                        For Each inner As VisualBasicSyntaxNode In innerBlock2.ChildNodes()
                            Me.GetMethodIndentShapeInternal(inner, itemscontrol3)
                        Next

                        Dim textblock5 = New TextBlock
                        itemscontrol1.Items.Add(textblock5)
                        textblock5.Text = "↓"
                        textblock5.Foreground = Brushes.Gainsboro
                        textblock5.Margin = New Thickness(20, 10, 20, 0)

                    Next

                End If


            Case Else

                For Each child As VisualBasicSyntaxNode In node.ChildNodes()
                    Me.GetMethodIndentShapeInternal(child, itemscontrol1)
                Next


        End Select

    End Sub

    Public Function DescendantTokens(methodRange As String) As IEnumerable(Of SyntaxToken)

        Dim tree = VisualBasicSyntaxTree.ParseText(methodRange)
        Dim root = TryCast(tree.GetRoot(), VisualBasicSyntaxNode)
        Dim items = root.DescendantTokens()
        Return items

    End Function

End Class
