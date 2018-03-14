Imports System.IO
Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Imports System.Reflection
Imports System.CodeDom.Compiler
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports Microsoft.VisualBasic.CompilerServices
Imports ICSharpCode.AvalonEdit.Folding
Imports DotFolit.Petzold.Media2D


Public Class EditorUserControl

#Region "フィールド、プロパティ"

    Private nsView As DataView = Nothing

    Private displayDefineFullName As String = String.Empty
    Private displayMethodFullName As String = String.Empty
    Private displayArgumentsFullName As String = String.Empty

    Private ClassCache As Dictionary(Of String, UIElement()) = Nothing
    Private MethodCache As Dictionary(Of Tuple(Of String, String), Border) = Nothing

#End Region

#Region "ツリービュー、ノード選択"

    Private Sub treeview1_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))

        Dim model = TryCast(e.NewValue, TreeViewItemModel)
        If (model Is Nothing) OrElse (model.StartLength = -1) Then
            Return
        End If

        ' ソースが違う場合は無視（画面クラス、Form1.vb と Form1.Designer.vb を想定｛名前空間が同名｝）
        If Me.texteditor1.Document.FileName <> model.FileName Then
            Return
        End If

        ' キャレット位置を、メンバー定義位置へ移動
        Me.texteditor1.TextArea.Caret.Offset = model.StartLength

        ' メンバー定義位置が見えるようにスクロール
        Dim jumpLine = Me.texteditor1.Document.GetLineByOffset(model.StartLength).LineNumber
        Me.texteditor1.ScrollToLine(jumpLine)

    End Sub

#End Region

#Region "エディター、キャレット移動"

    ' （注意）ShowNamespaceButton_Click イベントハンドラ内で、呼び出している。
    ' sender, または e　引数を使う場合は、再考慮すること

    Private Async Sub Caret_PositionChanged(sender As Object, e As EventArgs)

        Dim offset = Me.texteditor1.TextArea.Caret.Offset
        Dim sourceFile = Me.texteditor1.Document.FileName

        Dim defineFullName = String.Empty
        Dim defineKind = String.Empty

        ' クラスツリーへの表示処理、図形処理２つは、別スレッド上でおこなう。各処理の待機用リスト
        Dim tasks = New List(Of Task)

        ' メソッド系の場合、フローチャートを表示更新
        Me.nsView.RowFilter = $"SourceFile='{sourceFile}' AND DefineKind IN ('Constructor', 'Operator', 'WindowsAPI', 'EventHandler', 'Method') "
        Me.nsView.Sort = "StartLength, EndLength"

        ' DataView.RowFilter は、文字列の条件分岐がある場合、追加で数値の比較を書いても、うまくフィルタされない？ため別判定
        Dim foundMethod = False
        If Me.nsView.Count <> 0 Then

            Dim methodKinds = New List(Of String) From {"Constructor", "Operator", "WindowsAPI", "EventHandler", "Method"}
            For i As Integer = 0 To Me.nsView.Count - 1

                Dim row = Me.nsView(i)
                Dim startLength = CInt(row("StartLength"))
                Dim endLength = CInt(row("EndLength"))

                If startLength <= offset AndAlso offset <= endLength Then

                    Dim methodArguments = CStr(row("MethodArguments"))
                    defineFullName = CStr(row("DefineFullName"))
                    defineKind = CStr(row("DefineKind"))
                    foundMethod = True

                    If (Me.displayMethodFullName = defineFullName) AndAlso (Me.displayArgumentsFullName = methodArguments) Then
                        Exit For
                    End If
                    Me.displayMethodFullName = defineFullName
                    Me.displayArgumentsFullName = methodArguments

                    Dim task1 = Task.Run(Sub()

                                             Me.Dispatcher.BeginInvoke(Sub()

                                                                           Me.AddMethodFlowChartCanvas(defineFullName, methodArguments, startLength, endLength)

                                                                       End Sub)

                                         End Sub)
                    tasks.Add(task1)

                End If

            Next

        End If

        If Not foundMethod Then
            Me.displayMethodFullName = String.Empty
            Me.displayArgumentsFullName = String.Empty
            Me.ClearMethodFlowChartCanvas()
        End If

        ' クラスメンバーツリーを表示更新
        Me.nsView.RowFilter = $"SourceFile='{sourceFile}' AND DefineKind IN ('Class', 'Structure', 'Interface', 'Module', 'Enum', 'Delegate') "
        Me.nsView.Sort = "StartLength, EndLength"

        If Me.nsView.Count = 0 Then
            Me.treeview1.ItemsSource = Nothing
            Me.displayDefineFullName = String.Empty
            Me.displayMethodFullName = String.Empty
            Me.displayArgumentsFullName = String.Empty
            Me.ClearClassInheritsCanvas()
            Me.ClearMethodFlowChartCanvas()
            Return
        End If

        ' 以下のような内側のクラス等の場合、昇順で探すと外側のクラスを採用してしまう
        ' もっとも内側のコンテナを採用したいため、降順で探す（あらかじめ昇順でソートしておく）
        ' Class Class1(0-20)、※開始・終了の定義文字数位置
        '    Class InnerClass(11-17)、※開始・終了の定義文字数位置

        For i As Integer = Me.nsView.Count - 1 To 0 Step -1

            Dim row = Me.nsView(i)
            Dim startLength = CInt(row("StartLength"))
            Dim endLength = CInt(row("EndLength"))

            If startLength <= offset AndAlso offset <= endLength Then

                ' 見つけた候補が、Delegate、または Enum の場合、親コンテナに含まれていないかチェック、含まれている場合、その親を候補に取り替える
                defineFullName = CStr(row("DefineFullName"))
                defineKind = CStr(row("DefineKind"))

                If (defineKind = "Delegate") OrElse (defineKind = "Enum") Then

                    ' 入れ子の場合、ドット区切りではなくプラス区切りとなる
                    If defineFullName.Contains("+") Then

                        Dim containerFullName = defineFullName.Substring(0, defineFullName.LastIndexOf("+"))
                        Me.nsView.RowFilter = $"DefineFullName='{containerFullName}' AND DefineKind IN ('Class', 'Structure', 'Interface', 'Module') "
                        Me.nsView.Sort = "StartLength, EndLength"

                        If 0 < Me.nsView.Count Then

                            row = Me.nsView(Me.nsView.Count - 1)
                            defineFullName = CStr(row("DefineFullName"))
                            defineKind = CStr(row("DefineKind"))

                        End If

                    End If

                End If

                Exit For

            End If

        Next

        If defineFullName = String.Empty Then
            Me.treeview1.ItemsSource = Nothing
            Me.displayDefineFullName = String.Empty
            Me.displayMethodFullName = String.Empty
            Me.displayArgumentsFullName = String.Empty
            Me.ClearClassInheritsCanvas()
            Me.ClearMethodFlowChartCanvas()
            Return
        End If

        ' すでに同じデータを表示していないかチェック
        ' （矢印キーを押しっぱなしにした際の、応答性向上対応）
        If Me.displayDefineFullName = defineFullName Then
            Return
        Else
            Me.displayDefineFullName = defineFullName
        End If

        Dim containerType = Me.GetTypeFromAllAssemblies(defineFullName)
        If containerType Is Nothing Then
            Me.treeview1.ItemsSource = Nothing
            Me.displayDefineFullName = String.Empty
            Me.displayMethodFullName = String.Empty
            Me.displayArgumentsFullName = String.Empty
            Me.ClearClassInheritsCanvas()
            Me.ClearMethodFlowChartCanvas()
            Return
        End If

        If defineKind = "Class" Then

            Dim task2 = Task.Run(Sub()

                                     Me.Dispatcher.BeginInvoke(Sub()

                                                                   Me.AddClassInheritsCanvas(defineFullName)

                                                               End Sub)

                                 End Sub)
            tasks.Add(task2)

        End If

        Dim task3 = Task.Run(Sub()

                                 Me.Dispatcher.BeginInvoke(Sub()

                                                               Me.AddTreeView(defineFullName, defineKind, containerType)

                                                           End Sub)

                             End Sub)
        tasks.Add(task3)

        ' 各処理の完了を待機
        If tasks.Count <> 0 Then
            Await Task.WhenAll(tasks)
        End If

    End Sub

    Private Sub ClearClassInheritsCanvas()

        Dim parentWindow = TryCast(Window.GetWindow(Me), MainWindow)
        If parentWindow Is Nothing Then
            Return
        End If

        Dim inheritsCanvas = parentWindow.InheritsCanvas
        inheritsCanvas.Children.Clear()

    End Sub

    Private Sub ClearMethodFlowChartCanvas()

        Dim parentWindow = TryCast(Window.GetWindow(Me), MainWindow)
        If parentWindow Is Nothing Then
            Return
        End If

        Dim flowChartCanvas = parentWindow.FlowChartCanvas
        flowChartCanvas.Children.Clear()

    End Sub



    ' 「Type.GetType(type name : string) の結果、Nothing が返ってきてしまう現象」の対応メソッド
    ' 自分自身のアセンブリおよび mscorlib 内のクラス以外なら、アセンブリ名までを含んだ完全限定名（AssemblyQualifiedName）が必要とのこと
    ' xxx.xxx, Version=x.x.x.x, Culture=neutral, PublicKeyToken=xxx-xxx ... みたいな形式
    ' http://bbs.wankuma.com/index.cgi?mode=al2&namber=36296&KLOG=63
    ' https://stackoverflow.com/questions/1825147/type-gettypenamespace-a-b-classname-returns-null

    ' 上記より、アセンブリ情報まで付けなくても、読み込んでいるアセンブリ経由で名前空間を指定すれば、（あれば）取得できるとのこと

    Private Function GetTypeFromAllAssemblies(typeName As String) As Type

        Dim foundType = Type.GetType(typeName)
        If foundType IsNot Nothing Then
            Return foundType
        End If

        For Each otherAssembly In AppDomain.CurrentDomain.GetAssemblies()

            foundType = otherAssembly.GetType(typeName)
            If foundType IsNot Nothing Then
                Return foundType
            End If

        Next

        Return Nothing

    End Function

    Private Sub AddTreeView(defineFullName As String, defineKind As String, containerType As Type)

        Dim model As TreeViewItemModel = Nothing
        Dim isRemoveNamespace As Boolean = (Not Me.ShowNamespaceButton.IsChecked.GetValueOrDefault())

        Select Case defineKind
            Case "Class" : model = Me.GetClassModel(containerType, isRemoveNamespace)
            Case "Structure" : model = Me.GetStructureModel(containerType, isRemoveNamespace)
            Case "Interface" : model = Me.GetInterfaceModel(containerType, isRemoveNamespace)
            Case "Module" : model = Me.GetModuleModel(containerType, isRemoveNamespace)
            Case "Delegate" : model = Me.GetDelegateModel(containerType, isRemoveNamespace)
            Case "Enum" : model = Me.GetEnumModel(containerType, isRemoveNamespace)
        End Select

        Dim items = New ObservableCollection(Of TreeViewItemModel)
        items.Add(model)
        Me.treeview1.ItemsSource = items

        ' Delegate, Enum は返却。それ以外はメンバーも取得する
        If defineKind = "Delegate" OrElse defineKind = "Enum" Then
            Return
        End If


        ' コンテナ系（Class, Structure, Interface, Module）の場合

        ' ソースから取得した情報と、dll から取得した情報の２つから取得する（冗長ではある）
        ' 二重化により、①ソースから取得したメンバーを使うことで、リフレクション時に発生する、自動生成メンバーの判定を考慮しなくて済む（全てのメンバーがジャンプ先のソース行位置を持っている）し、
        ' ②そのメンバーを dll から取得することで、ソースに記載されていないような情報を取得することができる

        ' 表示メンバーを取得
        Me.nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind IN ('Field', 'Property', 'Event', 'Constructor', 'Operator', 'Method', 'Delegate', 'Enum') "
        If Me.nsView.Count = 0 Then
            Return
        End If

        ' そのクラス内で定義している全メンバーを取得（継承元メンバーは含まない）
        Dim flags = BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.Static Or BindingFlags.DeclaredOnly
        Dim members = containerType.GetMembers(flags)

        ' 表示メンバーに該当する、型データを取得
        Dim displayMembers = New List(Of MemberInfo)
        For Each row As DataRowView In Me.nsView

            Dim memberKind = CStr(row("DefineKind"))
            Dim memberName = CStr(row("DefineFullName"))
            memberName = memberName.Substring(memberName.LastIndexOf(".") + 1)

            Select Case memberKind

                Case "Field"

                    Dim foundMember = members.
                        Where(Function(x) x.MemberType = MemberTypes.Field).
                        Where(Function(x) x.Name = memberName).
                        FirstOrDefault()

                    If foundMember IsNot Nothing Then
                        displayMembers.Add(foundMember)
                    End If


                Case "Event"

                    Dim foundMember = members.
                        Where(Function(x) x.MemberType = MemberTypes.Event).
                        Where(Function(x) x.Name = memberName).
                        FirstOrDefault()

                    If foundMember IsNot Nothing Then
                        displayMembers.Add(foundMember)
                    End If


                Case "Property"

                    Dim candidateMembers = members.
                        Where(Function(x) x.MemberType = MemberTypes.Property).
                        Select(Function(x) TryCast(x, PropertyInfo)).
                        Where(Function(x) x.Name = memberName)

                    Dim methodArguments = CStr(row("MethodArguments"))
                    methodArguments = Me.RemoveNamespaceAll(methodArguments)
                    methodArguments = Me.ConvertToVBType(methodArguments)

                    Dim foundMember = candidateMembers.
                        Where(Function(x) Me.ConvertToVBType(Me.GetMethodArguments(x.GetIndexParameters(), True)) = methodArguments).
                        FirstOrDefault()

                    If foundMember IsNot Nothing Then
                        displayMembers.Add(foundMember)
                    End If


                Case "Constructor"

                    Dim candidateMembers = members.
                        Where(Function(x) x.MemberType = MemberTypes.Constructor).
                        Select(Function(x) TryCast(x, ConstructorInfo))

                    Dim methodArguments = CStr(row("MethodArguments"))
                    methodArguments = Me.RemoveNamespaceAll(methodArguments)
                    methodArguments = Me.ConvertToVBType(methodArguments)

                    Dim foundMember = candidateMembers.
                        Where(Function(x) Me.ConvertToVBType(Me.GetMethodArguments(x, True)) = methodArguments).
                        FirstOrDefault()

                    If foundMember IsNot Nothing Then
                        displayMembers.Add(foundMember)
                    End If


                Case "Operator", "Method"

                    ' ソース側ではメソッドの場合もジェネリック定義している場合、名称の後にジェネリック定義数を付与している（RoslynParser.vb）
                    ' しかしリフレクションの場合、クラス等にはジェネリック定義数が付与されるが、メソッドの場合は付与されない（＝名称の不一致）
                    ' そのため、ソース側のメソッド名を削る
                    Dim hasGenericDefinition = memberName.Contains("`")
                    Dim genericCount = 0

                    If hasGenericDefinition Then
                        genericCount = CInt(memberName.Substring(memberName.IndexOf("`") + 1))
                        memberName = memberName.Substring(0, memberName.IndexOf("`"))
                    End If

                    ' メソッド名の一致チェック
                    Dim candidateMembers = members.
                        Where(Function(x) x.MemberType = MemberTypes.Method).
                        Select(Function(x) TryCast(x, MethodInfo)).
                        Where(Function(x) x.Name = memberName)

                    Dim methodArguments = CStr(row("MethodArguments"))
                    methodArguments = Me.RemoveNamespaceAll(methodArguments)
                    methodArguments = Me.ConvertToVBType(methodArguments)

                    ' 引数の一致チェック
                    ' ジェネリック定義している場合、定義数も一致チェック
                    If hasGenericDefinition Then

                        Dim candidateMembers2 = candidateMembers.
                            Where(Function(x) x.IsGenericMethodDefinition).
                            Where(Function(x) x.GetGenericArguments().Count() = genericCount)

                        Dim foundMember = candidateMembers2.
                            Where(Function(x) Me.ConvertToVBType(Me.GetMethodArguments(x, True)) = methodArguments).
                            FirstOrDefault()

                        If foundMember IsNot Nothing Then
                            displayMembers.Add(foundMember)
                        End If

                    Else

                        Dim foundMember = candidateMembers.
                            Where(Function(x) Me.ConvertToVBType(Me.GetMethodArguments(x, True)) = methodArguments).
                            FirstOrDefault()

                        If foundMember IsNot Nothing Then
                            displayMembers.Add(foundMember)
                        End If

                    End If


                Case "Enum"

                    ' 入れ子の場合、ドット区切りではなくプラス区切りとなる
                    If memberName.Contains("+") Then
                        memberName = memberName.Substring(memberName.LastIndexOf("+") + 1)
                    End If

                    Dim foundMember = members.
                        Where(Function(x) x.MemberType = MemberTypes.NestedType).
                        Select(Function(x) TryCast(x, Type)).
                        Where(Function(x) x.IsEnum AndAlso x.Name = memberName).
                        FirstOrDefault()

                    If foundMember IsNot Nothing Then
                        displayMembers.Add(foundMember)
                    End If


                Case "Delegate"

                    ' 入れ子の場合、ドット区切りではなくプラス区切りとなる
                    If memberName.Contains("+") Then
                        memberName = memberName.Substring(memberName.LastIndexOf("+") + 1)
                    End If

                    ' 名前一致チェックでいったん区切る
                    Dim candidateMembers = members.
                        Where(Function(x) x.MemberType = MemberTypes.NestedType).
                        Select(Function(x) TryCast(x, Type)).
                        Where(Function(x) x.Name = memberName).
                        Where(Function(x) x.Equals(GetType([Delegate])) OrElse x.IsSubclassOf(GetType([Delegate])))

                    Dim methodArguments = CStr(row("MethodArguments"))
                    methodArguments = Me.RemoveNamespaceAll(methodArguments)
                    methodArguments = Me.ConvertToVBType(methodArguments)

                    ' 引数一致チェック
                    Dim foundMember = candidateMembers.
                        Where(Function(x) x.GetMethod("Invoke") IsNot Nothing).
                        Where(Function(x) Me.ConvertToVBType(Me.GetMethodArguments(x.GetMethod("Invoke"), True)) = methodArguments).
                        FirstOrDefault()

                    If foundMember IsNot Nothing Then
                        displayMembers.Add(foundMember)
                    End If


            End Select


        Next

        Me.AddMembers(model, displayMembers, isRemoveNamespace)

    End Sub

    Private Function GetClassModel(containerType As Type, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetTargetContainerModel(containerType, TreeNodeKinds.ClassNode, isRemoveNamespace)
    End Function

    Private Function GetStructureModel(containerType As Type, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetTargetContainerModel(containerType, TreeNodeKinds.StructureNode, isRemoveNamespace)
    End Function

    Private Function GetInterfaceModel(containerType As Type, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetTargetContainerModel(containerType, TreeNodeKinds.InterfaceNode, isRemoveNamespace)
    End Function

    Private Function GetModuleModel(containerType As Type, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetTargetContainerModel(containerType, TreeNodeKinds.ModuleNode, isRemoveNamespace)
    End Function

    Private Function GetTargetContainerModel(containerType As Type, kind As TreeNodeKinds, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel

        ' 名称
        Dim defineName = containerType.Name
        If containerType.IsGenericTypeDefinition Then

            Dim arguments = containerType.GetGenericArguments()
            Dim constraints = Me.GetGenericTypeNames(arguments, isRemoveNamespace)

            defineName = defineName.Substring(0, defineName.IndexOf("`"))
            defineName = $"{defineName}{constraints}"

        End If

        ' 定義開始位置
        Dim defineFullName = containerType.FullName
        If defineFullName.Contains("[") Then
            defineFullName = defineFullName.Substring(0, defineFullName.IndexOf("["))
        End If

        Me.nsView.RowFilter = $"DefineFullName='{defineFullName}' "
        Dim startLength = CInt(Me.nsView(0)("StartLength"))

        Dim containerModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = kind, .StartLength = startLength, .IsExpanded = True}
        Return containerModel

    End Function

    Private Function GetGenericTypeNames(arguments As Type(), Optional isRemoveNamespace As Boolean = True) As String

        Dim sb = New StringBuilder

        For Each argument In arguments

            If sb.Length <> 0 Then
                sb.Append(", ")
            End If

            ' 制約リスト
            Dim items = New List(Of String)
            Dim attributes = argument.GenericParameterAttributes
            Dim check = attributes And GenericParameterAttributes.SpecialConstraintMask

            ' 種類を指定している場合
            If (check And GenericParameterAttributes.ReferenceTypeConstraint) = GenericParameterAttributes.ReferenceTypeConstraint Then
                items.Add("Class")
            End If

            If (check And GenericParameterAttributes.NotNullableValueTypeConstraint) = GenericParameterAttributes.NotNullableValueTypeConstraint Then
                items.Add("NotNullableValueType")
            End If

            If (check And GenericParameterAttributes.DefaultConstructorConstraint) = GenericParameterAttributes.DefaultConstructorConstraint Then
                items.Add("New")
            End If

            ' 何らかのクラス、またはインターフェースを指定している場合
            If argument.GetGenericParameterConstraints().Any() Then

                For Each constraintsType In argument.GetGenericParameterConstraints()

                    Dim item = constraintsType.ToString()
                    item = Me.ConvertToVBType(item)
                    If isRemoveNamespace Then
                        item = Me.RemoveNamespaceAll(item)
                    End If

                    items.Add(constraintsType.ToString())

                Next

            End If

            sb.Append(argument.Name)
            If items.Any() Then
                Dim constraints = String.Join(", ", items)
                sb.Append($" As {"{"}{constraints}{"}"}")
            End If

        Next

        sb.Insert(0, "(Of ")
        sb.Append(")")
        Return sb.ToString()

    End Function

    Private Sub AddMembers(containerModel As TreeViewItemModel, members As List(Of MemberInfo), Optional isRemoveNamespace As Boolean = True)

        ' フィールド
        If members.Any(Function(x) x.MemberType = MemberTypes.Field) Then

            Dim items = members.Where(Function(x) x.MemberType = MemberTypes.Field)
            Dim fieldModel = New TreeViewItemModel With {.Text = "フィールド", .TreeNodeKind = TreeNodeKinds.FieldNode, .IsExpanded = True}
            containerModel.AddChild(fieldModel)

            For Each member In items

                Dim info = TryCast(member, FieldInfo)
                Dim item = Me.GetFieldModel(info, isRemoveNamespace)
                fieldModel.AddChild(item)

            Next

        End If

        ' プロパティ
        If members.Any(Function(x) x.MemberType = MemberTypes.Property) Then

            Dim items = members.Where(Function(x) x.MemberType = MemberTypes.Property)
            Dim propertyModel = New TreeViewItemModel With {.Text = "プロパティ", .TreeNodeKind = TreeNodeKinds.PropertyNode, .IsExpanded = True}
            containerModel.AddChild(propertyModel)

            For Each member In items

                Dim info = TryCast(member, PropertyInfo)
                Dim item = Me.GetPropertyModel(info, isRemoveNamespace)
                propertyModel.AddChild(item)

            Next

        End If

        ' コンストラクタ
        If members.Any(Function(x) x.MemberType = MemberTypes.Constructor) Then

            Dim items = members.Where(Function(x) x.MemberType = MemberTypes.Constructor)
            Dim constructorModel = New TreeViewItemModel With {.Text = "コンストラクタ", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
            containerModel.AddChild(constructorModel)

            For Each member In items

                Dim info = TryCast(member, ConstructorInfo)
                Dim item = Me.GetConstructorModel(info, isRemoveNamespace)
                constructorModel.AddChild(item)

            Next

        End If

        ' メソッド
        If members.Any(Function(x) x.MemberType = MemberTypes.Method) Then

            ' メソッドはさらに、WindowsAPI、Operator、EventHandler、Method と分ける

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.Method).
                Select(Function(x) TryCast(x, MethodInfo))

            Dim windowsAPIItems = New List(Of MethodInfo)
            Dim operatorItems = New List(Of MethodInfo)
            Dim eventHandlerItems = New List(Of MethodInfo)
            Dim methodItems = New List(Of MethodInfo)

            For Each item In items

                ' Windows API
                ' 属性が付与されている、その中の１つに、DllImport 属性がある
                ' Declare 版をリフレクションしたら、同じになるか？
                If item.CustomAttributes.Any(Function(x) x Is GetType(DllImportAttribute)) Then
                    windowsAPIItems.Add(item)
                    Continue For
                End If

                ' Operator
                ' 名前が op_Xxx で始まっている、SpecialName 属性が付与されている
                If item.Name.StartsWith("op_") AndAlso item.Attributes.ToString().Contains("SpecialName") Then
                    operatorItems.Add(item)
                    Continue For
                End If

                ' EventHandler
                ' 引数が２つある、かつ１つ目が Object 型、２つ目が EventArgs 型、またはその継承先クラスである
                Dim parameters = item.GetParameters()
                If parameters.Length = 2 Then
                    If (TypeOf parameters(0).ParameterType Is Object) Then
                        If (parameters(1).ParameterType Is GetType(System.EventArgs)) OrElse (parameters(1).ParameterType.IsSubclassOf(GetType(System.EventArgs))) Then
                            eventHandlerItems.Add(item)
                            Continue For
                        End If
                    End If
                End If

                ' Sub / Function
                ' それ以外
                methodItems.Add(item)

            Next

            ' Windows API
            If windowsAPIItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "Windows API", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                containerModel.AddChild(methodModel)

                For Each member In windowsAPIItems

                    Dim item = Me.GetWindowsAPIModel(member, isRemoveNamespace)
                    methodModel.AddChild(item)

                Next

            End If

            ' Operator
            If operatorItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "オペレータ", .TreeNodeKind = TreeNodeKinds.OperatorNode, .IsExpanded = True}
                containerModel.AddChild(methodModel)

                For Each member In operatorItems

                    Dim item = Me.GetOperatorModel(member, isRemoveNamespace)
                    methodModel.AddChild(item)

                Next

            End If

            ' EventHandler
            If eventHandlerItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "イベントハンドラ", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                containerModel.AddChild(methodModel)

                For Each member In eventHandlerItems

                    Dim item = Me.GetEventHandlerModel(member, isRemoveNamespace)
                    methodModel.AddChild(item)

                Next

            End If

            ' Method
            If methodItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "メソッド", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                containerModel.AddChild(methodModel)

                For Each member In methodItems

                    Dim item = Me.GetMethodModel(member, isRemoveNamespace)
                    methodModel.AddChild(item)

                Next

            End If


        End If

        ' イベント定義
        If members.Any(Function(x) x.MemberType = MemberTypes.Event) Then

            Dim items = members.Where(Function(x) x.MemberType = MemberTypes.Event)
            Dim eventModel = New TreeViewItemModel With {.Text = "イベント定義", .TreeNodeKind = TreeNodeKinds.EventNode, .IsExpanded = True}
            containerModel.AddChild(eventModel)

            For Each member In items

                Dim info = TryCast(member, EventInfo)
                Dim item = Me.GetEventModel(info, isRemoveNamespace)
                eventModel.AddChild(item)

            Next

        End If

        ' デリゲート、列挙体
        If members.Any(Function(x) x.MemberType = MemberTypes.NestedType) Then

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.NestedType).
                Select(Function(x) TryCast(x, Type))

            Dim enumItems = New List(Of Type)
            Dim delegateItems = New List(Of Type)

            For Each item In items

                Dim isDelegate = item.Equals(GetType([Delegate])) OrElse item.IsSubclassOf(GetType([Delegate]))
                If isDelegate Then
                    delegateItems.Add(item)
                    Continue For
                End If

                If item.IsEnum Then
                    enumItems.Add(item)
                    Continue For
                End If

            Next

            ' Delegate
            If delegateItems.Any() Then

                Dim delegateModel = New TreeViewItemModel With {.Text = "デリゲート", .TreeNodeKind = TreeNodeKinds.DelegateNode, .IsExpanded = True}
                containerModel.AddChild(delegateModel)

                For Each member In delegateItems

                    Dim item = Me.GetDelegateModel(member, isRemoveNamespace)
                    delegateModel.AddChild(item)

                Next

            End If

            ' Enum
            If enumItems.Any() Then

                Dim enumModel = New TreeViewItemModel With {.Text = "列挙体", .TreeNodeKind = TreeNodeKinds.EnumNode, .IsExpanded = True}
                containerModel.AddChild(enumModel)

                For Each member In enumItems

                    Dim item = Me.GetEnumModel(member, isRemoveNamespace)
                    enumModel.AddChild(item)

                Next

            End If

        End If

    End Sub

    Private Function GetEnumModel(containerType As Type, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel

        Dim defineName = containerType.Name

        ' 定義開始位置の取得
        Dim defineFullName = containerType.FullName
        If defineFullName.Contains("[") Then
            defineFullName = defineFullName.Substring(0, defineFullName.IndexOf("["))
        End If

        Me.nsView.RowFilter = $"DefineFullName='{defineFullName}' "
        Dim startLength = If(Me.nsView.Count = 0, -1, CInt(Me.nsView(0)("StartLength")))
        Dim fileName = If(Me.nsView.Count = 0, String.Empty, CStr(Me.nsView(0)("SourceFile")))
        Dim containerModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = TreeNodeKinds.EnumNode, .StartLength = startLength, .FileName = fileName, .IsExpanded = True}

        ' メンバー名の取得
        Dim enumNames = [Enum].GetNames(containerType)
        For Each enumName In enumNames

            Dim childModel = New TreeViewItemModel With {.Text = enumName, .TreeNodeKind = TreeNodeKinds.EnumItemNode}
            containerModel.AddChild(childModel)

        Next

        Return containerModel

    End Function

    Private Function GetDelegateModel(containerType As Type, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel

        Dim defineName = containerType.Name
        Dim invokeMethod = containerType.GetMethod("Invoke")
        Dim methodArguments = "()"
        Dim returnType = "Void"

        ' ジェネリック定義している場合
        If containerType.IsGenericTypeDefinition Then

            Dim arguments = containerType.GetGenericArguments()
            Dim constraints = Me.GetGenericTypeNames(arguments, isRemoveNamespace)

            defineName = defineName.Substring(0, defineName.IndexOf("`"))
            defineName = $"{defineName}{constraints}"

        End If

        If invokeMethod IsNot Nothing Then

            methodArguments = Me.GetMethodArguments(invokeMethod, isRemoveNamespace)
            methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))
            methodArguments = Me.ConvertToVBType(methodArguments)

            returnType = invokeMethod.ReturnType.ToString()
            returnType = Me.ConvertToVBType(returnType)
            If isRemoveNamespace Then
                returnType = Me.RemoveNamespaceAll(returnType)
            End If

        End If

        defineName = $"{defineName}{methodArguments}"
        If Not returnType.Contains("Void") Then
            defineName = $"{defineName} : {returnType}"
        End If

        ' 定義開始位置の取得
        Dim defineFullName = containerType.FullName
        If defineFullName.Contains("[") Then
            defineFullName = defineFullName.Substring(0, defineFullName.IndexOf("["))
        End If

        methodArguments = "0()"
        If invokeMethod IsNot Nothing Then

            methodArguments = Me.GetMethodArguments(invokeMethod, True)
            methodArguments = Me.ConvertToVBType(methodArguments)

        End If

        Me.nsView.RowFilter = String.Empty
        Dim foundRow = Me.nsView.Cast(Of DataRowView)().FirstOrDefault(
            Function(x)

                If x("MethodArguments") Is DBNull.Value Then
                    Return False
                End If

                Dim rowName = CStr(x("DefineFullName"))
                Dim rowArguments = CStr(x("MethodArguments"))
                rowArguments = Me.RemoveNamespaceAll(rowArguments)
                rowArguments = Me.ConvertToVBType(rowArguments)

                Dim b1 = (rowName = defineFullName)
                Dim b2 = (rowArguments = methodArguments)

                Return b1 AndAlso b2

            End Function)

        Dim startLength = If(foundRow Is Nothing, -1, CInt(foundRow("StartLength")))
        Dim fileName = If(foundRow Is Nothing, String.Empty, CStr(foundRow("SourceFile")))
        Dim containerModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = TreeNodeKinds.DelegateNode, .StartLength = startLength, .FileName = fileName}
        Return containerModel

    End Function

    Private Function GetMethodArguments(info As MethodBase, Optional isRemoveNamespace As Boolean = True) As String
        Return Me.GetMethodArguments(info.GetParameters(), isRemoveNamespace)
    End Function

    Private Function GetMethodArguments(arguments As ParameterInfo(), Optional isRemoveNamespace As Boolean = True) As String

        ' MethodArguments は、引数の数(引数の型、引数の型、・・・)という形式
        Dim sb = New StringBuilder
        Dim methodArguments = String.Empty

        For Each argument In arguments

            If sb.Length <> 0 Then
                sb.Append(", ")
            End If

            Dim isByRef = argument.ParameterType.ToString().EndsWith("&")
            Dim isOptional = argument.IsOptional
            Dim isParamArray = argument.CustomAttributes.Any(Function(x) x.AttributeType Is GetType(ParamArrayAttribute))
            Dim parameterType = argument.ParameterType.ToString()

            If isRemoveNamespace Then
                parameterType = Me.RemoveNamespaceAll(parameterType)
            End If

            ' ByRef one, ByRef Optional, ByVal all, ByVal ParamArray
            If isByRef Then

                parameterType = parameterType.Substring(0, parameterType.LastIndexOf("&"))

                If isOptional Then
                    sb.Append($"[ByRef {parameterType}]")
                Else
                    sb.Append($"ByRef {parameterType}")
                End If

                Continue For

            End If

            If isOptional Then
                sb.Append($"[{parameterType}]")
            ElseIf isParamArray Then
                sb.Append($"ParamArray {parameterType}")
            Else
                sb.Append($"{parameterType}")
            End If

        Next

        methodArguments = $"{arguments.Length}({sb.ToString()})"
        Return methodArguments

    End Function

    Private Function ConvertToVBType(target As String) As String

        ' ビルド後は、記号名から対応する文字列名称に変換される
        Dim langTable = MemoryDB.Instance.DB.Tables("LanguageConversion")
        For Each langRow As DataRow In langTable.Rows

            Dim fxType = CStr(langRow("NETFramework"))
            Dim vbType = CStr(langRow("VBNet"))

            If target.Contains(fxType) Then
                target = target.Replace(fxType, vbType)
            End If

        Next

        If target.Contains("`") Then
            target = Regex.Replace(target, "(`\d+\[)", "(Of ")
        End If

        If target.Contains("[") Then
            target = target.Replace("[", "(")
        End If

        If target.Contains("]") Then
            target = target.Replace("]", ")")
        End If

        Return target

    End Function

    'Private Function RemoveNamespace(parameterType As String) As String

    '    If parameterType.Contains(".") Then
    '        parameterType = parameterType.Substring(parameterType.LastIndexOf(".") + 1)
    '    End If

    '    Return parameterType

    'End Function

    Private Function RemoveNamespaceAll(target As String) As String

        If Regex.IsMatch(target, "([\w`]+[.])+") Then
            target = Regex.Replace(target, "([\w`]+[.])+", String.Empty)
        End If

        Return target

    End Function

    Private Function GetFieldModel(info As FieldInfo, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel

        Dim defineName = info.Name
        Dim typeName = info.FieldType.ToString()
        typeName = Me.ConvertToVBType(typeName)
        If isRemoveNamespace Then
            typeName = Me.RemoveNamespaceAll(typeName)
        End If

        defineName = $"{defineName} : {typeName}"

        ' 定義開始位置の取得
        ' 親コンテナの名前空間
        Dim defineFullName = info.ReflectedType.FullName
        If defineFullName.Contains("[") Then
            defineFullName = defineFullName.Substring(0, defineFullName.IndexOf("["))
        End If

        ' メンバーの名前
        defineFullName = $"{defineFullName}.{info.Name}"
        Me.nsView.RowFilter = $"DefineFullName='{defineFullName}' "
        Dim startLength = If(Me.nsView.Count = 0, -1, CInt(Me.nsView(0)("StartLength")))
        Dim fileName = If(Me.nsView.Count = 0, String.Empty, CStr(Me.nsView(0)("SourceFile")))

        Dim containerModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = TreeNodeKinds.FieldNode, .StartLength = startLength, .FileName = fileName}
        Return containerModel

    End Function

    Private Function GetPropertyModel(info As PropertyInfo, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel

        Dim defineName = info.Name
        Dim typeName = info.PropertyType.ToString()
        typeName = Me.ConvertToVBType(typeName)
        If isRemoveNamespace Then
            typeName = Me.RemoveNamespaceAll(typeName)
        End If

        ' インデクサ（indexer）の場合
        Dim methodArguments = String.Empty
        Dim isIndexer = info.GetIndexParameters().Any()
        If isIndexer Then

            methodArguments = Me.GetMethodArguments(info.GetIndexParameters(), isRemoveNamespace)
            methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))
            methodArguments = Me.ConvertToVBType(methodArguments)
            defineName = $"{defineName}{methodArguments}"

        End If

        defineName = $"{defineName} : {typeName}"

        ' 定義開始位置の取得
        ' メンバー名
        Dim defineFullName = info.ReflectedType.FullName
        If defineFullName.Contains("[") Then
            defineFullName = defineFullName.Substring(0, defineFullName.IndexOf("["))
        End If

        defineFullName = $"{defineFullName}.{info.Name}"

        ' 引数
        methodArguments = "0()"
        If isIndexer Then

            methodArguments = Me.GetMethodArguments(info.GetIndexParameters(), True)
            methodArguments = Me.ConvertToVBType(methodArguments)

        End If

        Me.nsView.RowFilter = String.Empty
        Dim foundRow = Me.nsView.Cast(Of DataRowView)().FirstOrDefault(
            Function(x)

                If x("MethodArguments") Is DBNull.Value Then
                    Return False
                End If

                Dim rowName = CStr(x("DefineFullName"))
                Dim rowArguments = CStr(x("MethodArguments"))
                rowArguments = Me.RemoveNamespaceAll(rowArguments)
                rowArguments = Me.ConvertToVBType(rowArguments)

                Dim b1 = (rowName = defineFullName)
                Dim b2 = (rowArguments = methodArguments)

                Return b1 AndAlso b2

            End Function)

        Dim startLength = If(foundRow Is Nothing, -1, CInt(foundRow("StartLength")))
        Dim fileName = If(foundRow Is Nothing, String.Empty, CStr(foundRow("SourceFile")))
        Dim containerModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = TreeNodeKinds.PropertyNode, .StartLength = startLength, .FileName = fileName}
        Return containerModel

    End Function

    Private Function GetEventModel(info As EventInfo, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel

        Dim defineName = info.Name
        Dim typeName = info.EventHandlerType.ToString()
        typeName = Me.ConvertToVBType(typeName)
        If isRemoveNamespace Then
            typeName = Me.RemoveNamespaceAll(typeName)
        End If

        defineName = $"{defineName} : {typeName}"

        ' 定義開始位置の取得
        ' 親コンテナの名前空間
        Dim defineFullName = info.ReflectedType.FullName
        If defineFullName.Contains("[") Then
            defineFullName = defineFullName.Substring(0, defineFullName.IndexOf("["))
        End If

        ' メンバーの名前
        defineFullName = $"{defineFullName}.{info.Name}"
        Me.nsView.RowFilter = $"DefineFullName='{defineFullName}' "
        Dim startLength = If(Me.nsView.Count = 0, -1, CInt(Me.nsView(0)("StartLength")))
        Dim fileName = If(Me.nsView.Count = 0, String.Empty, CStr(Me.nsView(0)("SourceFile")))

        Dim containerModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = TreeNodeKinds.EventNode, .StartLength = startLength, .FileName = fileName}
        Return containerModel

    End Function

    Private Function GetConstructorModel(info As ConstructorInfo, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetMethodModelInternal(info, TreeNodeKinds.MethodNode, isRemoveNamespace)
    End Function

    Private Function GetMethodModel(info As MethodInfo, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetMethodModelInternal(info, TreeNodeKinds.MethodNode, isRemoveNamespace)
    End Function

    Private Function GetWindowsAPIModel(info As MethodInfo, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetMethodModelInternal(info, TreeNodeKinds.MethodNode, isRemoveNamespace)
    End Function

    Private Function GetOperatorModel(info As MethodInfo, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetMethodModelInternal(info, TreeNodeKinds.OperatorNode, isRemoveNamespace)
    End Function

    Private Function GetEventHandlerModel(info As MethodInfo, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel
        Return Me.GetMethodModelInternal(info, TreeNodeKinds.MethodNode, isRemoveNamespace)
    End Function

    Private Function GetMethodModelInternal(info As MethodBase, kind As TreeNodeKinds, Optional isRemoveNamespace As Boolean = True) As TreeViewItemModel

        Dim isConstructor = (TryCast(info, ConstructorInfo) IsNot Nothing)
        Dim isMethod = (TryCast(info, MethodInfo) IsNot Nothing)
        Dim defineName = info.Name

        ' コンストラクタの場合、メソッド名を VB 用に変換（.ctor → New）
        If isConstructor Then
            defineName = "New"
        End If

        Dim methodArguments = String.Empty
        Dim returnType = "Void"

        ' Operator の場合、特殊文字列名から定義名に、メソッド名を戻す
        If kind = TreeNodeKinds.OperatorNode Then
            defineName = Me.ConvertToVBType(defineName)
        End If

        ' ジェネリック定義している場合、メソッド名に追加
        If info.IsGenericMethodDefinition Then

            Dim arguments = info.GetGenericArguments()
            Dim constraints = Me.GetGenericTypeNames(arguments, isRemoveNamespace)
            defineName = $"{defineName}{constraints}"

        End If

        methodArguments = Me.GetMethodArguments(info, isRemoveNamespace)
        methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))
        methodArguments = Me.ConvertToVBType(methodArguments)

        If isMethod Then

            Dim mi = TryCast(info, MethodInfo)
            returnType = mi.ReturnType.ToString()
            returnType = Me.ConvertToVBType(returnType)
            If isRemoveNamespace Then
                returnType = Me.RemoveNamespaceAll(returnType)
            End If

        End If

        defineName = $"{defineName}{methodArguments}"
        If Not returnType.Contains("Void") Then
            defineName = $"{defineName} : {returnType}"
        End If

        ' 定義開始位置の取得
        ' 親コンテナの名前空間
        Dim defineFullName = info.ReflectedType.FullName
        If defineFullName.Contains("[") Then
            defineFullName = defineFullName.Substring(0, defineFullName.IndexOf("["))
        End If

        ' メンバーの名前
        If isConstructor Then
            defineFullName = $"{defineFullName}.New"
        Else
            defineFullName = $"{defineFullName}.{info.Name}"
        End If

        If info.IsGenericMethodDefinition Then

            Dim arguments = info.GetGenericArguments()
            defineFullName = $"{defineFullName}`{arguments.Count()}"

        End If

        methodArguments = Me.GetMethodArguments(info, True)
        methodArguments = Me.ConvertToVBType(methodArguments)

        Me.nsView.RowFilter = String.Empty
        Dim foundRow = Me.nsView.Cast(Of DataRowView)().FirstOrDefault(
            Function(x)

                If x("MethodArguments") Is DBNull.Value Then
                    Return False
                End If

                Dim rowName = CStr(x("DefineFullName"))
                Dim rowArguments = CStr(x("MethodArguments"))
                rowArguments = Me.RemoveNamespaceAll(rowArguments)
                rowArguments = Me.ConvertToVBType(rowArguments)

                Dim b1 = (rowName = defineFullName)
                Dim b2 = (rowArguments = methodArguments)

                Return b1 AndAlso b2

            End Function)

        Dim startLength = If(foundRow Is Nothing, -1, CInt(foundRow("StartLength")))
        Dim fileName = If(foundRow Is Nothing, String.Empty, CStr(foundRow("SourceFile")))
        Dim containerModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = kind, .StartLength = startLength, .FileName = fileName}
        Return containerModel

    End Function

    Private Sub AddClassInheritsCanvas(defineFullName As String)

        Dim parentWindow = TryCast(Window.GetWindow(Me), MainWindow)
        If parentWindow Is Nothing Then
            Return
        End If

        ' 以前作成済みのキャッシュデータがあれば、こちらを再利用する
        If Me.ClassCache Is Nothing Then
            Me.ClassCache = New Dictionary(Of String, UIElement())
        End If

        Dim cacheDatas As UIElement() = Nothing
        If Me.ClassCache.TryGetValue(defineFullName, cacheDatas) Then
            ' キャッシュデータをそのまま利用する

            Dim inheritsCanvas = parentWindow.InheritsCanvas
            inheritsCanvas.Children.Clear()

            For Each cacheData As UIElement In cacheDatas
                inheritsCanvas.Children.Add(cacheData)
            Next

        Else
            ' 新規データを作成して、キャッシュしておく

            ' ターゲットクラスと、その継承元クラス全ての Type を取得
            Dim items = New List(Of Type)
            Dim classType = Me.GetTypeFromAllAssemblies(defineFullName)
            If classType Is Nothing Then
                Return
            End If

            ' 継承元クラスは、都度前に追加していく（キャンバスの表示順となる）
            items.Add(classType)
            Dim baseType As Type = classType.BaseType
            While True

                If baseType Is Nothing Then
                    Exit While
                End If

                items.Insert(0, baseType)
                baseType = baseType.BaseType

            End While

            Dim inheritsCanvas = parentWindow.InheritsCanvas
            inheritsCanvas.Children.Clear()

            For Each item In items
                Me.AddClassInheritsCanvasInternal(inheritsCanvas, item, (item IsNot classType))
            Next

            ' キャッシュしておく
            ' 参照コピーではダメみたいなので、インスタンスの複製の方のコピーでキャッシュ
            ReDim cacheDatas(inheritsCanvas.Children.Count - 1)
            inheritsCanvas.Children.CopyTo(cacheDatas, 0)
            Me.ClassCache.Add(defineFullName, cacheDatas)

        End If

    End Sub

    Private Sub AddClassInheritsCanvasInternal(inheritsCanvas As Canvas, classType As Type, isBaseClass As Boolean)

        Dim newThumb = New ResizableThumb
        newThumb.Template = TryCast(Me.Resources("ClassTreeRectangleTemplate"), ControlTemplate)
        newThumb.ApplyTemplate()

        ' インスタンス生成後に、どこかで UpdateLayout メソッドを呼び出さないと、
        ' Width/Height, ActualXxx, Desired.Xxx など、サイズに関する全ての値が 0 のままになってしまう（後で表示位置の計算をしたいのに、できない）
        ' これは、Canvas.Children に追加して、画面表示した後の状態でも！継続されるため、このタイミングで呼び出してしまう
        newThumb.UpdateLayout()

        ' クラス名
        Dim isRemoveNamespace As Boolean = (Not Me.ShowNamespaceButton.IsChecked.GetValueOrDefault())
        Dim defineName = classType.Name

        ' ジェネリック定義があれば追加
        If classType.IsGenericTypeDefinition Then

            Dim arguments = classType.GetGenericArguments()
            Dim constraints = Me.GetGenericTypeNames(arguments, isRemoveNamespace)

            defineName = defineName.Substring(0, defineName.IndexOf("`"))
            defineName = $"{defineName}{constraints}"

        End If

        ' 継承元クラスがあれば追加
        If classType.BaseType IsNot Nothing Then

            Dim baseName = classType.BaseType.ToString()
            If isRemoveNamespace Then
                baseName = Me.RemoveNamespaceAll(baseName)
            End If

            defineName = $"{defineName} : {baseName}"

        End If

        Dim border1 = TryCast(newThumb.Template.FindName("Border1", newThumb), Border)
        If isBaseClass Then
            ' 継承元クラスは赤系
            border1.BorderBrush = Brushes.Tomato
            border1.Background = Brushes.LavenderBlush
        Else
            ' ターゲットクラスは青系
            border1.BorderBrush = Brushes.Blue
            border1.Background = Brushes.AliceBlue
        End If

        Dim expander1 = TryCast(newThumb.Template.FindName("Expander1", newThumb), Expander)
        expander1.Header = defineName
        expander1.IsExpanded = True

        Dim inheritsTree = TryCast(newThumb.Template.FindName("InheritsTree", newThumb), TreeView)
        Dim treeItems = New ObservableCollection(Of TreeViewItemModel)
        inheritsTree.ItemsSource = treeItems

        ' メンバー（Delegate, Enum 含む）
        Dim flags = BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static Or BindingFlags.Instance Or BindingFlags.DeclaredOnly
        Dim members = classType.GetMembers(flags)

        ' フィールド
        If members.Any(Function(x) x.MemberType = MemberTypes.Field) Then

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.Field).
                Select(Function(x) TryCast(x, FieldInfo))

            If isBaseClass Then
                items = items.Where(Function(x) x.IsPrivate = False)
            End If

            If items.Any() Then

                Dim fieldModel = New TreeViewItemModel With {.Text = "フィールド", .TreeNodeKind = TreeNodeKinds.FieldNode, .IsExpanded = True}
                treeItems.Add(fieldModel)

                For Each item In items

                    Dim childModel = Me.GetFieldModel(item, isRemoveNamespace)
                    fieldModel.AddChild(childModel)

                Next

            End If

        End If

        ' プロパティ
        If members.Any(Function(x) x.MemberType = MemberTypes.Property) Then

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.Property).
                Select(Function(x) TryCast(x, PropertyInfo))

            If isBaseClass Then
                items = items.
                    Where(Function(x) x.CanRead AndAlso x.GetGetMethod() IsNot Nothing AndAlso x.GetGetMethod().IsPrivate = False).
                    Where(Function(x) x.CanWrite AndAlso x.GetSetMethod() IsNot Nothing AndAlso x.GetSetMethod().IsPrivate = False)
            End If

            If items.Any() Then

                Dim propertyModel = New TreeViewItemModel With {.Text = "プロパティ", .TreeNodeKind = TreeNodeKinds.PropertyNode, .IsExpanded = True}
                treeItems.Add(propertyModel)

                For Each item In items

                    Dim childModel = Me.GetPropertyModel(item, isRemoveNamespace)
                    propertyModel.AddChild(childModel)

                Next

            End If

        End If

        ' コンストラクタ
        If members.Any(Function(x) x.MemberType = MemberTypes.Constructor) Then

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.Constructor).
                Select(Function(x) TryCast(x, ConstructorInfo))

            If isBaseClass Then
                items = items.Where(Function(x) x.IsPrivate = False)
            End If

            If items.Any() Then

                Dim constructorModel = New TreeViewItemModel With {.Text = "コンストラクタ", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                treeItems.Add(constructorModel)

                For Each item In items

                    Dim childModel = Me.GetConstructorModel(item, isRemoveNamespace)
                    constructorModel.AddChild(childModel)

                Next

            End If

        End If

        ' メソッド
        If members.Any(Function(x) x.MemberType = MemberTypes.Method) Then

            ' メソッドはさらに、WindowsAPI、Operator、EventHandler、Method と分ける

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.Method).
                Select(Function(x) TryCast(x, MethodInfo))

            If isBaseClass Then
                items = items.Where(Function(x) x.IsPrivate = False)
            End If

            Dim windowsAPIItems = New List(Of MethodInfo)
            Dim operatorItems = New List(Of MethodInfo)
            Dim eventHandlerItems = New List(Of MethodInfo)
            Dim methodItems = New List(Of MethodInfo)

            For Each item In items

                ' Windows API
                ' 属性が付与されている、その中の１つに、DllImport 属性がある
                ' Declare 版をリフレクションしたら、同じになるか？
                If item.CustomAttributes.Any(Function(x) x Is GetType(DllImportAttribute)) Then
                    windowsAPIItems.Add(item)
                    Continue For
                End If

                ' Operator
                ' 名前が op_Xxx で始まっている、SpecialName 属性が付与されている
                If item.Name.StartsWith("op_") AndAlso item.Attributes.ToString().Contains("SpecialName") Then
                    operatorItems.Add(item)
                    Continue For
                End If

                ' EventHandler
                ' 引数が２つある、かつ１つ目が Object 型、２つ目が EventArgs 型、またはその継承先クラスである
                Dim parameters = item.GetParameters()
                If parameters.Length = 2 Then
                    If (TypeOf parameters(0).ParameterType Is Object) Then
                        If (parameters(1).ParameterType Is GetType(System.EventArgs)) OrElse (parameters(1).ParameterType.IsSubclassOf(GetType(System.EventArgs))) Then
                            eventHandlerItems.Add(item)
                            Continue For
                        End If
                    End If
                End If

                ' Sub / Function
                ' それ以外
                methodItems.Add(item)

            Next

            ' Windows API
            If windowsAPIItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "Windows API", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                treeItems.Add(methodModel)

                For Each member In windowsAPIItems

                    Dim item = Me.GetWindowsAPIModel(member, isRemoveNamespace)
                    methodModel.AddChild(item)

                Next

            End If

            ' Operator
            If operatorItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "オペレータ", .TreeNodeKind = TreeNodeKinds.OperatorNode, .IsExpanded = True}
                treeItems.Add(methodModel)

                For Each member In operatorItems

                    Dim item = Me.GetOperatorModel(member, isRemoveNamespace)
                    methodModel.AddChild(item)

                Next

            End If

            ' EventHandler
            If eventHandlerItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "イベントハンドラ", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                treeItems.Add(methodModel)

                For Each member In eventHandlerItems

                    Dim item = Me.GetEventHandlerModel(member, isRemoveNamespace)
                    methodModel.AddChild(item)

                Next

            End If

            ' Method
            If methodItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "メソッド", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                treeItems.Add(methodModel)

                For Each member In methodItems

                    Dim item = Me.GetMethodModel(member, isRemoveNamespace)
                    methodModel.AddChild(item)

                Next

            End If

        End If

        ' イベント定義
        If members.Any(Function(x) x.MemberType = MemberTypes.Event) Then

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.Event).
                Select(Function(x) TryCast(x, EventInfo))

            Dim eventModel = New TreeViewItemModel With {.Text = "イベント定義", .TreeNodeKind = TreeNodeKinds.EventNode, .IsExpanded = True}
            treeItems.Add(eventModel)

            For Each member In items

                Dim item = Me.GetEventModel(member, isRemoveNamespace)
                eventModel.AddChild(item)

            Next

        End If

        ' デリゲート、列挙体
        If members.Any(Function(x) x.MemberType = MemberTypes.NestedType) Then

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.NestedType).
                Select(Function(x) TryCast(x, Type))

            If isBaseClass Then
                items = items.Where(Function(x) x.IsNestedPrivate = False)
            End If

            Dim enumItems = New List(Of Type)
            Dim delegateItems = New List(Of Type)

            For Each item In items

                Dim isDelegate = item.Equals(GetType([Delegate])) OrElse item.IsSubclassOf(GetType([Delegate]))
                If isDelegate Then
                    delegateItems.Add(item)
                    Continue For
                End If

                If item.IsEnum Then
                    enumItems.Add(item)
                    Continue For
                End If

            Next

            ' Delegate
            If delegateItems.Any() Then

                Dim delegateModel = New TreeViewItemModel With {.Text = "デリゲート", .TreeNodeKind = TreeNodeKinds.DelegateNode, .IsExpanded = True}
                treeItems.Add(delegateModel)

                For Each member In delegateItems

                    Dim item = Me.GetDelegateModel(member, isRemoveNamespace)
                    delegateModel.AddChild(item)

                Next

            End If

            ' Enum
            If enumItems.Any() Then

                Dim enumModel = New TreeViewItemModel With {.Text = "列挙体", .TreeNodeKind = TreeNodeKinds.EnumNode, .IsExpanded = True}
                treeItems.Add(enumModel)

                For Each member In enumItems

                    Dim item = Me.GetEnumModel(member, isRemoveNamespace)
                    enumModel.AddChild(item)

                Next

            End If

        End If

        Dim pos = Me.GetNewLocation(inheritsCanvas)
        Canvas.SetLeft(newThumb, pos.X)
        Canvas.SetTop(newThumb, pos.Y)

        ' キャンバスに登録する前に、１つ前に登録した図形を取得して置く
        Dim previousThumb As ResizableThumb = Nothing
        If inheritsCanvas.Children.OfType(Of ResizableThumb)().Count <> 0 Then

            Dim items = inheritsCanvas.Children.OfType(Of ResizableThumb)()
            previousThumb = items.LastOrDefault()

        End If

        ' 今回分の図形を新規登録
        inheritsCanvas.Children.Add(newThumb)

        If previousThumb Is Nothing Then
            Return
        End If

        ' 前回と今回の図形同士を、矢印線でつなげる
        Dim arrow = New ArrowLine
        arrow.Stroke = Brushes.Green
        arrow.StrokeThickness = 1
        inheritsCanvas.Children.Add(arrow)

        previousThumb.StartLines.Add(arrow)
        newThumb.EndLines.Add(arrow)

        Me.UpdateLineLocation(previousThumb)
        Me.UpdateLineLocation(newThumb)

        ' なぜか、最後の図形だけ、矢印線が左上の角を指してしまう不具合
        ' → ActualWidth, ActualHeight が 0 だから。いったん画面表示させないとダメか？
        ' → Measure メソッドを呼び出して、希望サイズを更新する。こちらで矢印線の位置を調整する
        Dim newSize = New Size(inheritsCanvas.ActualWidth, inheritsCanvas.ActualHeight)
        inheritsCanvas.Measure(newSize)
        Me.UpdateLineLocation(newThumb)

    End Sub

    Private Function GetNewLocation(inheritsCanvas As Canvas) As Point

        ' 最も右下に位置する ResizableThumb を探す
        Dim items = inheritsCanvas.Children.OfType(Of ResizableThumb)()
        If items.Count() = 0 Then
            Return New Point(10, 10)
        End If

        ' 既に表示されている図形のうち、１つ目の図形位置を基準として、今回の図形位置を計算する
        Dim item = items(0)
        Dim newWidth As Double = item.ActualWidth
        Dim newHeight As Double = item.ActualHeight

        Dim newX As Double = Canvas.GetLeft(item) + newWidth + 40
        Dim newY As Double = Canvas.GetTop(item) + 40 ' Math.Min(40, newHeight / 2)

        ' 予測計算した位置・コントロールの大きさの内に、すでに他のコントロールが配置されていないかチェック
        ' 一部が重なり合う場合、重ならないように予測位置を修正して、再度全チェック
        Dim found = False

        While True

            ' 初期化してチャレンジ、または再チャレンジ
            found = False
            For i As Integer = 1 To items.Count - 1

                item = items(i)
                Dim currentX = Canvas.GetLeft(item)
                Dim currentY = Canvas.GetTop(item)

                Dim currentWidth = item.ActualWidth
                Dim currentHeight = item.ActualHeight

                Dim currentRect = New Rect(currentX, currentY, currentWidth, currentHeight)
                Dim newRect = New Rect(newX, newY, newWidth, newHeight)

                Select Case True
                    Case currentRect.Contains(newRect.TopLeft),
                     currentRect.Contains(newRect.TopRight),
                     currentRect.Contains(newRect.BottomLeft),
                     currentRect.Contains(newRect.BottomRight)

                        ' 重なっている図形の【右下ちょい上くらい】まで移動
                        found = True
                        newX = currentX + currentWidth + 40
                        newY = currentY + 40 ' Math.Min(40, currentHeight / 2)

                End Select

            Next

            ' 既存配置しているコントロールリスト全てに重ならなかったので、この位置で決定
            If Not found Then
                Exit While
            End If

            ' １つ以上重なっていたことにより、予測位置を修正したので、もう一度コントロールリスト全部と再チェック
        End While

        Dim result = New Point(newX, newY)
        Return result

    End Function

    ' ResizableThumb.vb, MethodWindow.xaml.vb 側にも同じメソッドがあるので同期すること
    Private Sub UpdateLineLocation(target As ResizableThumb)

        Dim newX = Canvas.GetLeft(target)
        Dim newY = Canvas.GetTop(target)

        Dim newWidth = target.ActualWidth
        Dim newHeight = target.ActualHeight

        If newWidth = 0 Then newWidth = target.DesiredSize.Width
        If newHeight = 0 Then newHeight = target.DesiredSize.Height

        For i As Integer = 0 To target.StartLines.Count - 1
            target.StartLines(i).X1 = newX + newWidth
            target.StartLines(i).Y1 = newY + (newHeight / 2)
        Next

        For i As Integer = 0 To target.EndLines.Count - 1
            target.EndLines(i).X2 = newX
            target.EndLines(i).Y2 = newY + (newHeight / 2)
        Next

    End Sub

    Private Sub AddMethodFlowChartCanvas(defineFullName As String, methodArguments As String, startLength As Integer, endLength As Integer)

        ' 以前作成済みのキャッシュデータがあれば、こちらを再利用する
        If Me.MethodCache Is Nothing Then
            Me.MethodCache = New Dictionary(Of Tuple(Of String, String), Border)
        End If

        Dim cacheData As Border = Nothing
        If Me.MethodCache.TryGetValue(Tuple.Create(defineFullName, methodArguments), cacheData) Then
            ' キャッシュデータをそのまま利用する
        Else
            ' 新規データを作成して、キャッシュしておく

            Dim methodRange = Me.texteditor1.Text.Substring(startLength, endLength - startLength)
            Dim parser = New RoslynParser
            Dim border1 = parser.GetMethodIndentShape(methodRange)

            ' キャッシュしておく
            cacheData = border1
            Me.MethodCache.Add(Tuple.Create(defineFullName, methodArguments), cacheData)

        End If

        Dim parentWindow = TryCast(Window.GetWindow(Me), MainWindow)
        If parentWindow Is Nothing Then
            Return
        End If

        Dim flowChartCanvas = parentWindow.FlowChartCanvas
        flowChartCanvas.Children.Clear()
        flowChartCanvas.Children.Add(cacheData)

        ' 初期表示範囲が、左上の隅から少し右下にずれてしまう現象の対応
        Dim scrollviewer1 = TryCast(flowChartCanvas.Parent, ScrollViewer)
        scrollviewer1.ScrollToTop()
        scrollviewer1.ScrollToLeftEnd()

    End Sub

#End Region

#Region "メンバーツリーを表示トグルボタン、クリック"

    Private Sub ShowMemberTreeButton_Click(sender As Object, e As RoutedEventArgs)

        ' 横列の制御は、TreeView を配置している１つ目だけの変更だと、うまく動作しなかった（開閉）
        ' 横全列を状況に合わせて変更することで、うまく動作するようになった

        If Me.ShowMemberTreeButton.IsChecked Then

            Me.grid1.ColumnDefinitions(0).Width = New GridLength(1, GridUnitType.Star)
            Me.grid1.ColumnDefinitions(1).Width = New GridLength(1, GridUnitType.Auto)
            Me.grid1.ColumnDefinitions(2).Width = New GridLength(2, GridUnitType.Star)

        Else

            Me.grid1.ColumnDefinitions(0).Width = New GridLength(0, GridUnitType.Pixel)
            Me.grid1.ColumnDefinitions(1).Width = New GridLength(0, GridUnitType.Pixel)
            Me.grid1.ColumnDefinitions(2).Width = New GridLength(1, GridUnitType.Star)

        End If

    End Sub

#End Region

#Region "名前空間を表示トグルボタン、クリック"

    Private Sub ShowNamespaceButton_Click(sender As Object, e As RoutedEventArgs)

        ' （注意）Caret_PositionChanged イベントハンドラ内で、sender, または e 引数を使っていないか確認すること
        Me.displayDefineFullName = String.Empty
        Me.Caret_PositionChanged(Me, EventArgs.Empty)

    End Sub

#End Region

#Region "エディタのマウスホイール変更（ズームイン、アウト）"

    ' [WPF]コントロールキーあるいはシフトキーが押されているかどうかを取得する
    ' https://gist.github.com/pinzolo/3080481


    ' うまく動作せず
    'Private Sub texteditor1_MouseWheel(sender As Object, e As MouseWheelEventArgs)

    '    Dim isDownLeftControlKey = ((Keyboard.GetKeyStates(Key.LeftCtrl) And KeyStates.Down) = KeyStates.Down)
    '    Dim isDownRightControlKey = ((Keyboard.GetKeyStates(Key.RightCtrl) And KeyStates.Down) = KeyStates.Down)
    '    Dim isDownControlKey = isDownLeftControlKey OrElse isDownRightControlKey

    '    If isDownControlKey Then

    '        If 0 < e.Delta Then

    '            texteditorScaleTransform.ScaleX *= 1.1
    '            texteditorScaleTransform.ScaleY *= 1.1

    '        Else

    '            texteditorScaleTransform.ScaleX /= 1.1
    '            texteditorScaleTransform.ScaleY /= 1.1

    '        End If

    '    End If

    'End Sub

    Private Sub texteditor1_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)

        Dim isDownLeftControlKey = ((Keyboard.GetKeyStates(Key.LeftCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownRightControlKey = ((Keyboard.GetKeyStates(Key.RightCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownControlKey = isDownLeftControlKey OrElse isDownRightControlKey

        If isDownControlKey Then

            If 0 < e.Delta Then
                Me.texteditor1.FontSize *= 1.1
            Else
                Me.texteditor1.FontSize /= 1.1
            End If

        End If

    End Sub

#End Region

#Region "メソッド"

    Public Sub InitializeData(sourceFile As String)

        ' XAML 上で設定していない部分の設定
        ' タブはスペース変換して表示する
        Me.texteditor1.Options.ConvertTabsToSpaces = True

        ' 現在行の背景色を表示する
        Me.texteditor1.Options.HighlightCurrentLine = True

        ' ソースを表示
        Dim source = File.ReadAllText(sourceFile, EncodeResolver.GetEncoding(sourceFile))
        Me.texteditor1.Document.Text = source
        Me.texteditor1.Document.FileName = sourceFile

        ' ソースを表示してから、折りたたみ機能を設定
        Dim strategy = New VBNetFoldingStrategy
        Dim manager = FoldingManager.Install(Me.texteditor1.TextArea)
        strategy.UpdateFoldings(manager, Me.texteditor1.Document)

        ' 作業用としてキャッシュしておく
        Dim nsTable = MemoryDB.Instance.DB.Tables("NamespaceResolution")
        Me.nsView = nsTable.AsDataView()

        ' イベントハンドラ内で、nsView 変数を使っているため、インスタンス確保した後で、
        ' キャレット移動イベントの購読
        AddHandler Me.texteditor1.TextArea.Caret.PositionChanged, AddressOf Me.Caret_PositionChanged

    End Sub

#End Region

End Class
