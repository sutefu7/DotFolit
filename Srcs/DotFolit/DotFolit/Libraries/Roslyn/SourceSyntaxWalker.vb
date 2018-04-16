Imports System.Data
Imports System.IO
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax


' 「親情報の引継ぎ渡しの仕組み」がうまく思いつかなかったため、
' VisualBasicSyntaxWalker クラスを継承した実装は、いったん保留
' 以前の自前処理を使う


Public Class SourceSyntaxWalker

    Public Sub New()

    End Sub

    ' いるか？
    'Public Sub Parse(nsTable As DataTable, rootNamespace As String, sourceFile As String)

    '    Dim source = File.ReadAllText(sourceFile, EncodeResolver.GetEncoding(sourceFile))
    '    Dim tree = VisualBasicSyntaxTree.ParseText(source)
    '    Dim root = TryCast(tree.GetRoot(), VisualBasicSyntaxNode)

    '    Me.ParseInternal(nsTable, sourceFile, rootNamespace, "Namespace", root)

    'End Sub

    Public Sub Parse(nsTable As DataTable, rootNamespace As String, sourceFile As String, tree As SyntaxTree)

        Dim root = TryCast(tree.GetRoot(), VisualBasicSyntaxNode)
        Me.ParseInternal(nsTable, sourceFile, rootNamespace, "Namespace", root)

    End Sub

    ' 各 ParseXxx メソッドで 主な処理（DB 登録）をしている。ParseInternal メソッドは種類の振り分けをするだけの経由メソッド
    Private Sub ParseInternal(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim kind = node.Kind()

        Select Case kind

            Case SyntaxKind.CompilationUnit : Me.ParseCompilationUnit(nsTable, sourceFile, containerFullName, containerDefineKind, node)

                ' コンテナ系、または入れ子のコンテナ系
            Case SyntaxKind.NamespaceBlock : Me.ParseNamespaceBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.ClassBlock : Me.ParseClassBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.StructureBlock : Me.ParseStructureBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.InterfaceBlock : Me.ParseInterfaceBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.ModuleBlock : Me.ParseModuleBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.EnumBlock : Me.ParseEnumBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)

                ' メンバー（メソッド系。コンストラクタ、オペレータ、WindowsAPI、イベントハンドラ、メソッド）
            Case SyntaxKind.ConstructorBlock : Me.ParseConstructorBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.OperatorBlock : Me.ParseOperatorBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.SubBlock : Me.ParseSubBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.FunctionBlock : Me.ParseFunctionBlock(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.DeclareSubStatement : Me.ParseDeclareSubStatement(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.DeclareFunctionStatement : Me.ParseDeclareFunctionStatement(nsTable, sourceFile, containerFullName, containerDefineKind, node)

                ' メンバー（フィールド）
            Case SyntaxKind.FieldDeclaration : Me.ParseFieldDeclaration(nsTable, sourceFile, containerFullName, containerDefineKind, node)

                ' メンバー（プロパティ。通常プロパティ、自動実装プロパティ）
            Case SyntaxKind.PropertyBlock, SyntaxKind.PropertyStatement : Me.ParsePropertyBlockOrPropertyStatement(nsTable, sourceFile, containerFullName, containerDefineKind, node)

                ' 入れ子型（デリゲート）
            Case SyntaxKind.DelegateSubStatement : Me.ParseDelegateSubStatement(nsTable, sourceFile, containerFullName, containerDefineKind, node)
            Case SyntaxKind.DelegateFunctionStatement : Me.ParseDelegateFunctionStatement(nsTable, sourceFile, containerFullName, containerDefineKind, node)

                ' メンバー（イベント定義。カスタムイベント定義、イベント定義）
            Case SyntaxKind.EventBlock, SyntaxKind.EventStatement : Me.ParseEventBlockOrEventStatement(nsTable, sourceFile, containerFullName, containerDefineKind, node)

        End Select

    End Sub

    Private Sub ParseCompilationUnit(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        For Each child As VisualBasicSyntaxNode In node.ChildNodes
            Me.ParseInternal(nsTable, sourceFile, containerFullName, containerDefineKind, child)
        Next

    End Sub

    Private Sub ParseNamespaceBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

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
        row("DefineName") = defineName
        row("DisplayDefineName") = defineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = DBNull.Value
        row("MethodArguments") = DBNull.Value
        row("IsPartial") = DBNull.Value
        row("IsShared") = DBNull.Value
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

        For Each child As VisualBasicSyntaxNode In node.ChildNodes()
            Me.ParseInternal(nsTable, sourceFile, $"{containerFullName}.{defineName}", "Namespace", child)
        Next

    End Sub

    Private Sub ParseClassBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of ClassStatementSyntax)().FirstOrDefault()
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
        Dim displayDefineName = defineName

        ' ジェネリックを定義している場合、個数を追加
        If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

            Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
            Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
            defineName = $"{defineName}`{genericNodes.Count()}"

            Dim genericNames = genericNodes.Select(Function(x) x.ToString())
            displayDefineName = $"{displayDefineName}(Of {String.Join(", ", genericNames)})"

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
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = DBNull.Value
        row("MethodArguments") = DBNull.Value
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

        For Each child As VisualBasicSyntaxNode In node.ChildNodes()
            Me.ParseInternal(nsTable, sourceFile, defineFullName, "Class", child)
        Next

    End Sub

    Private Sub ParseStructureBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of StructureStatementSyntax)().FirstOrDefault()
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
        Dim displayDefineName = defineName

        ' ジェネリックを定義している場合、個数を追加
        If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

            Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
            Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
            defineName = $"{defineName}`{genericNodes.Count()}"

            Dim genericNames = genericNodes.Select(Function(x) x.ToString())
            displayDefineName = $"{displayDefineName}(Of {String.Join(", ", genericNames)})"

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
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = DBNull.Value
        row("MethodArguments") = DBNull.Value
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

        For Each child As VisualBasicSyntaxNode In node.ChildNodes()
            Me.ParseInternal(nsTable, sourceFile, defineFullName, "Structure", child)
        Next

    End Sub

    Private Sub ParseInterfaceBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of InterfaceStatementSyntax)().FirstOrDefault()
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
        Dim displayDefineName = defineName

        ' ジェネリックを定義している場合、個数を追加
        If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

            Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
            Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
            defineName = $"{defineName}`{genericNodes.Count()}"

            Dim genericNames = genericNodes.Select(Function(x) x.ToString())
            displayDefineName = $"{displayDefineName}(Of {String.Join(", ", genericNames)})"

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
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = DBNull.Value
        row("MethodArguments") = DBNull.Value
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

        For Each child As VisualBasicSyntaxNode In node.ChildNodes()
            Me.ParseInternal(nsTable, sourceFile, defineFullName, "Interface", child)
        Next

    End Sub

    Private Sub ParseModuleBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of ModuleStatementSyntax)().FirstOrDefault()
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
        Dim displayDefineName = defineName

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
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = DBNull.Value
        row("MethodArguments") = DBNull.Value
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

        For Each child As VisualBasicSyntaxNode In node.ChildNodes()
            Me.ParseInternal(nsTable, sourceFile, defineFullName, "Module", child)
        Next

    End Sub

    Private Sub ParseEnumBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of EnumStatementSyntax)().FirstOrDefault()
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
        Dim displayDefineName = defineName

        ' 内側のコンテナの場合、ドットつなぎではなくプラスつなぎにする
        Dim defineFullName = $"{containerFullName}.{defineName}"
        Dim containerKeywords = New List(Of String) From {"Class", "Structure", "Interface", "Module"}
        If containerKeywords.Contains(containerDefineKind) Then
            defineFullName = $"{containerFullName}+{defineName}"
        End If

        ' メンバー名
        Dim memberNames = node.ChildNodes().OfType(Of EnumMemberDeclarationSyntax)().Select(Function(x) x.Identifier.Text)

        Dim row = nsTable.NewRow()
        row("DefineKind") = "Enum"
        row("ContainerFullName") = containerFullName
        row("DefineFullName") = defineFullName
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = DBNull.Value
        row("MethodArguments") = String.Join(", ", memberNames)
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseConstructorBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of SubNewStatementSyntax)().FirstOrDefault()
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.NewKeyword).ToString()
        Dim displayDefineName = defineName

        ' ジェネリックを定義している場合、個数を追加
        If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

            Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
            Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
            defineName = $"{defineName}`{genericNodes.Count()}"

            Dim genericNames = genericNodes.Select(Function(x) x.ToString())
            displayDefineName = $"{displayDefineName}(Of {String.Join(", ", genericNames)})"

        End If

        Dim listNode = statementNode.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
        Dim methodArguments = Me.GetMethodArguments(listNode)

        Dim row = nsTable.NewRow()
        row("DefineKind") = "Constructor"
        row("ContainerFullName") = containerFullName
        row("DefineFullName") = $"{containerFullName}.{defineName}"
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = "Void"
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseOperatorBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of OperatorStatementSyntax)().FirstOrDefault()
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ToString()
        defineName = defineName.Substring(defineName.IndexOf("Operator ") + "Operator ".Length)
        defineName = defineName.Substring(0, defineName.IndexOf("("))
        Dim displayDefineName = defineName

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

            Dim genericNames = genericNodes.Select(Function(x) x.ToString())
            displayDefineName = $"{displayDefineName}(Of {String.Join(", ", genericNames)})"

        End If

        Dim listNode = statementNode.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
        Dim methodArguments = Me.GetMethodArguments(listNode)

        Dim typeNode = statementNode.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
        Dim typeName = Me.GetTypeName(typeNode)

        Dim row = nsTable.NewRow()
        row("DefineKind") = "Operator"
        row("ContainerFullName") = containerFullName
        row("DefineFullName") = $"{containerFullName}.{defineName}"
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = typeName
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseSubBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of MethodStatementSyntax)().FirstOrDefault()
        Dim isWinAPI = statementNode.AttributeLists.Any(Function(x) x.ToString().Contains("DllImport"))
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
        Dim displayDefineName = defineName

        ' ジェネリックを定義している場合、個数を追加
        If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

            Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
            Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
            defineName = $"{defineName}`{genericNodes.Count()}"

            Dim genericNames = genericNodes.Select(Function(x) x.ToString())
            displayDefineName = $"{displayDefineName}(Of {String.Join(", ", genericNames)})"

        End If

        Dim listNode = statementNode.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
        Dim methodArguments = Me.GetMethodArguments(listNode)

        Dim row = nsTable.NewRow()
        row("DefineKind") = If(isWinAPI, "WindowsAPI", "Method")
        row("ContainerFullName") = containerFullName
        row("DefineFullName") = $"{containerFullName}.{defineName}"
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = "Void"
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseFunctionBlock(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim statementNode = node.ChildNodes().OfType(Of MethodStatementSyntax)().FirstOrDefault()
        Dim isWinAPI = statementNode.AttributeLists.Any(Function(x) x.ToString().Contains("DllImport"))
        Dim isPartial = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = statementNode.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = statementNode.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
        Dim displayDefineName = defineName

        ' ジェネリックを定義している場合、個数を追加
        If statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

            Dim genericRangeNode = statementNode.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
            Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
            defineName = $"{defineName}`{genericNodes.Count()}"

            Dim genericNames = genericNodes.Select(Function(x) x.ToString())
            displayDefineName = $"{displayDefineName}(Of {String.Join(", ", genericNames)})"

        End If

        Dim listNode = statementNode.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
        Dim methodArguments = Me.GetMethodArguments(listNode)

        Dim typeNode = statementNode.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
        Dim typeName = Me.GetTypeName(typeNode)

        Dim row = nsTable.NewRow()
        row("DefineKind") = If(isWinAPI, "WindowsAPI", "Method")
        row("ContainerFullName") = containerFullName
        row("DefineFullName") = $"{containerFullName}.{defineName}"
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = typeName
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseDeclareSubStatement(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim isPartial = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = node.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).Text
        Dim displayDefineName = defineName

        Dim listNode = node.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
        Dim methodArguments = Me.GetMethodArguments(listNode)

        Dim row = nsTable.NewRow()
        row("DefineKind") = "WindowsAPI"
        row("ContainerFullName") = containerFullName
        row("DefineFullName") = $"{containerFullName}.{defineName}"
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = "Void"
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseDeclareFunctionStatement(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim isPartial = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = node.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).Text
        Dim displayDefineName = defineName

        Dim listNode = node.ChildNodes().OfType(Of ParameterListSyntax)().FirstOrDefault()
        Dim methodArguments = Me.GetMethodArguments(listNode)

        Dim typeNode = node.ChildNodes().OfType(Of SimpleAsClauseSyntax)().FirstOrDefault()
        Dim typeName = Me.GetTypeName(typeNode)

        Dim row = nsTable.NewRow()
        row("DefineKind") = "WindowsAPI"
        row("ContainerFullName") = containerFullName
        row("DefineFullName") = $"{containerFullName}.{defineName}"
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = typeName
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseFieldDeclaration(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

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
                row("DefineName") = defineName
                row("DisplayDefineName") = defineName
                row("DefineType") = $"{typeName}{bracket}"
                row("ReturnType") = DBNull.Value
                row("MethodArguments") = DBNull.Value
                row("IsPartial") = isPartial
                row("IsShared") = isShared
                row("SourceFile") = sourceFile
                row("StartLength") = startLength
                row("EndLength") = endLength
                nsTable.Rows.Add(row)

            Next

        Next

    End Sub

    Private Sub ParsePropertyBlockOrPropertyStatement(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)
        ' PropertyBlock / 通常プロパティ用（setter, getter いづれかの記述あり）
        ' PropertyStatement / 自動実装プロパティ用（setter, getter 記述無し）

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        If node.Kind() = SyntaxKind.PropertyBlock Then
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
        row("DefineName") = defineName
        row("DisplayDefineName") = defineName
        row("DefineType") = typeName
        row("ReturnType") = DBNull.Value
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseDelegateSubStatement(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

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
        row("DefineName") = defineName
        row("DisplayDefineName") = defineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = "Void"
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseDelegateFunctionStatement(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        Dim isPartial = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.PartialKeyword)
        Dim isShared = node.ChildTokens().Any(Function(x) x.Kind() = SyntaxKind.SharedKeyword)
        Dim defineName = node.ChildTokens().FirstOrDefault(Function(x) x.Kind() = SyntaxKind.IdentifierToken).ToString()
        Dim displayDefineName = defineName

        ' ジェネリックを定義している場合、個数を追加
        If node.ChildNodes().OfType(Of TypeParameterListSyntax)().Any() Then

            Dim genericRangeNode = node.ChildNodes().OfType(Of TypeParameterListSyntax)().FirstOrDefault()
            Dim genericNodes = genericRangeNode.ChildNodes().OfType(Of TypeParameterSyntax)()
            defineName = $"{defineName}`{genericNodes.Count()}"

            Dim genericNames = genericNodes.Select(Function(x) x.ToString())
            displayDefineName = $"{displayDefineName}(Of {String.Join(", ", genericNames)})"

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
        row("DefineName") = defineName
        row("DisplayDefineName") = displayDefineName
        row("DefineType") = DBNull.Value
        row("ReturnType") = typeName
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

    End Sub

    Private Sub ParseEventBlockOrEventStatement(nsTable As DataTable, sourceFile As String, containerFullName As String, containerDefineKind As String, node As VisualBasicSyntaxNode)

        Dim startLength = node.Span.Start
        Dim endLength = node.Span.End

        If node.Kind() = SyntaxKind.EventBlock Then
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
        row("DefineName") = defineName
        row("DisplayDefineName") = defineName
        row("DefineType") = typeName
        row("ReturnType") = DBNull.Value
        row("MethodArguments") = methodArguments
        row("IsPartial") = isPartial
        row("IsShared") = isShared
        row("SourceFile") = sourceFile
        row("StartLength") = startLength
        row("EndLength") = endLength
        nsTable.Rows.Add(row)

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

End Class
