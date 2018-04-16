Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports System.Runtime.InteropServices


' ほとんど、クラスの継承関係図専用になっているので、改名したいところ

Public Class ReflectionHelper

    Private nsView As DataView = Nothing

    Public Function GetInheritsItems(nsView As DataView, defineFullName As String) As ObservableCollection(Of InheritsItemModel)

        Dim classType = Me.GetTypeFromAllAssemblies(defineFullName)
        If classType Is Nothing Then
            Return New ObservableCollection(Of InheritsItemModel)
        End If

        ' 継承元クラスは、都度前に追加していく（キャンバスの表示順となる）
        Dim items = New List(Of Type) From {classType}
        Dim baseType = classType.BaseType
        While True

            If baseType Is Nothing Then
                Exit While
            End If

            items.Insert(0, baseType)
            baseType = baseType.BaseType

        End While

        Dim result = New ObservableCollection(Of InheritsItemModel)
        Me.nsView = nsView

        For Each item In items
            result.Add(Me.GetInheritsItemModel(item, item IsNot classType))
        Next

        Return result

    End Function



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

    Private Function GetInheritsItemModel(classType As Type, isBaseClass As Boolean) As InheritsItemModel

        ' 以前作成したキャッシュデータがある場合は、こちらを返却する（高速化対応）
        If MemoryDB.Instance.InheritsModelCache Is Nothing Then
            MemoryDB.Instance.InheritsModelCache = New Dictionary(Of String, InheritsItemModel)
        End If

        Dim modelCache = MemoryDB.Instance.InheritsModelCache
        If modelCache.ContainsKey($"{classType.FullName}{isBaseClass}") Then
            Dim cache = modelCache($"{classType.FullName}{isBaseClass}").NewInstance()
            Return cache
        End If

        ' クラス名
        Dim defineName = classType.Name

        ' ジェネリック定義があれば追加
        If classType.IsGenericTypeDefinition Then

            Dim arguments = classType.GetGenericArguments()
            Dim constraints = Me.GetGenericTypeNames(arguments)

            defineName = defineName.Substring(0, defineName.IndexOf("`"))
            defineName = $"{defineName}{constraints}"

        End If

        ' 継承元クラスがあれば追加
        If classType.BaseType IsNot Nothing Then

            Dim baseName = classType.BaseType.ToString()
            baseName = Me.RemoveNamespaceAll(baseName)
            defineName = $"{defineName} : {baseName}"

        End If

        Dim model = New InheritsItemModel
        model.ClassName = defineName
        model.IsTargetClass = Not isBaseClass
        model.TreeItems = Me.GetMemberTreeItems(classType, isBaseClass)

        ' キャッシュしておく
        modelCache.Add($"{classType.FullName}{isBaseClass}", model.NewInstance())

        Return model

    End Function

    Private Function GetMemberTreeItems(classType As Type, isBaseClass As Boolean) As ObservableCollection(Of TreeViewItemModel)

        Dim result = New ObservableCollection(Of TreeViewItemModel)

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
                result.Add(fieldModel)

                For Each item In items

                    Dim childModel = Me.GetFieldModel(item)
                    fieldModel.Children.Add(childModel)

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
                result.Add(propertyModel)

                For Each item In items

                    Dim childModel = Me.GetPropertyModel(item)
                    propertyModel.Children.Add(childModel)

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
                result.Add(constructorModel)

                For Each item In items

                    Dim childModel = Me.GetConstructorModel(item)
                    constructorModel.Children.Add(childModel)

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
                result.Add(methodModel)

                For Each member In windowsAPIItems

                    Dim item = Me.GetWindowsAPIModel(member)
                    methodModel.Children.Add(item)

                Next

            End If

            ' Operator
            If operatorItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "オペレータ", .TreeNodeKind = TreeNodeKinds.OperatorNode, .IsExpanded = True}
                result.Add(methodModel)

                For Each member In operatorItems

                    Dim item = Me.GetOperatorModel(member)
                    methodModel.Children.Add(item)

                Next

            End If

            ' EventHandler
            If eventHandlerItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "イベントハンドラ", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                result.Add(methodModel)

                For Each member In eventHandlerItems

                    Dim item = Me.GetEventHandlerModel(member)
                    methodModel.Children.Add(item)

                Next

            End If

            ' Method
            If methodItems.Any() Then

                Dim methodModel = New TreeViewItemModel With {.Text = "メソッド", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                result.Add(methodModel)

                For Each member In methodItems

                    Dim item = Me.GetMethodModel(member)
                    methodModel.Children.Add(item)

                Next

            End If

        End If

        ' イベント定義
        If members.Any(Function(x) x.MemberType = MemberTypes.Event) Then

            Dim items = members.
                Where(Function(x) x.MemberType = MemberTypes.Event).
                Select(Function(x) TryCast(x, EventInfo))

            Dim eventModel = New TreeViewItemModel With {.Text = "イベント定義", .TreeNodeKind = TreeNodeKinds.EventNode, .IsExpanded = True}
            result.Add(eventModel)

            For Each member In items

                Dim item = Me.GetEventModel(member)
                eventModel.Children.Add(item)

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
                result.Add(delegateModel)

                For Each member In delegateItems

                    Dim item = Me.GetDelegateModel(member)
                    delegateModel.Children.Add(item)

                Next

            End If

            ' Enum
            If enumItems.Any() Then

                Dim enumModel = New TreeViewItemModel With {.Text = "列挙体", .TreeNodeKind = TreeNodeKinds.EnumNode, .IsExpanded = True}
                result.Add(enumModel)

                For Each member In enumItems

                    Dim item = Me.GetEnumModel(member)
                    enumModel.Children.Add(item)

                Next

            End If

        End If

        Return result

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
            containerModel.Children.Add(childModel)

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

End Class
