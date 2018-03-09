Imports System.Globalization
Imports System.Windows.Controls.Primitives


Public Class TreeNodeKindsToImageConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert

        Dim img = New BitmapImage(New Uri("/Images/Miscellaneousfile.png", UriKind.Relative))

        If TypeOf value IsNot TreeNodeKinds Then
            Return img
        End If

        Dim kind = CType(value, TreeNodeKinds)
        Select Case kind

            Case TreeNodeKinds.SolutionNode : img = New BitmapImage(New Uri("/Images/Solution.png", UriKind.Relative))
            Case TreeNodeKinds.ProjectNode : img = New BitmapImage(New Uri("/Images/VBProject.png", UriKind.Relative))
            Case TreeNodeKinds.SourceNode : img = New BitmapImage(New Uri("/Images/VBFile.png", UriKind.Relative))
            Case TreeNodeKinds.DependencyNode : img = New BitmapImage(New Uri("/Images/Dependencies.png", UriKind.Relative))
            Case TreeNodeKinds.NamespaceNode : img = New BitmapImage(New Uri("/Images/Namespace.png", UriKind.Relative))
            Case TreeNodeKinds.ClassNode : img = New BitmapImage(New Uri("/Images/Class.png", UriKind.Relative))
            Case TreeNodeKinds.StructureNode : img = New BitmapImage(New Uri("/Images/Structure.png", UriKind.Relative))
            Case TreeNodeKinds.InterfaceNode : img = New BitmapImage(New Uri("/Images/Interface.png", UriKind.Relative))
            Case TreeNodeKinds.ModuleNode : img = New BitmapImage(New Uri("/Images/Module.png", UriKind.Relative))
            Case TreeNodeKinds.EnumNode : img = New BitmapImage(New Uri("/Images/Enum.png", UriKind.Relative))
            Case TreeNodeKinds.EnumItemNode : img = New BitmapImage(New Uri("/Images/EnumItem.png", UriKind.Relative))
            Case TreeNodeKinds.DelegateNode : img = New BitmapImage(New Uri("/Images/Delegate.png", UriKind.Relative))
            Case TreeNodeKinds.FolderNode : img = New BitmapImage(New Uri("/Images/Folder_Collapse.png", UriKind.Relative))
            Case TreeNodeKinds.GeneratedFileNode : img = New BitmapImage(New Uri("/Images/Generatedfile.png", UriKind.Relative))

            Case TreeNodeKinds.EventNode : img = New BitmapImage(New Uri("/Images/Event.png", UriKind.Relative))
            Case TreeNodeKinds.FieldNode : img = New BitmapImage(New Uri("/Images/Field.png", UriKind.Relative))
            Case TreeNodeKinds.PropertyNode : img = New BitmapImage(New Uri("/Images/Property.png", UriKind.Relative))
            Case TreeNodeKinds.OperatorNode : img = New BitmapImage(New Uri("/Images/Operator.png", UriKind.Relative))
            Case TreeNodeKinds.MethodNode : img = New BitmapImage(New Uri("/Images/Method.png", UriKind.Relative))
            Case TreeNodeKinds.ConditionStatementNode : img = New BitmapImage(New Uri("/Images/Condition.png", UriKind.Relative))
            Case TreeNodeKinds.LoopStatementNode : img = New BitmapImage(New Uri("/Images/Loop_Start.png", UriKind.Relative))
            Case TreeNodeKinds.ProcedureStatementNode : img = New BitmapImage(New Uri("/Images/Procedure.png", UriKind.Relative))

        End Select

        ' Type インスタンスを取得しないと判断できない
        'Case TreeNodeKinds.WindowsFormNode : img = New BitmapImage(New Uri("/Images/xxx.png", UriKind.Relative))
        'Case TreeNodeKinds.UserControlNode : img = New BitmapImage(New Uri("/Images/xxx.png", UriKind.Relative))
        'Case TreeNodeKinds.ComponentNode : img = New BitmapImage(New Uri("/Images/xxx.png", UriKind.Relative))


        Return img

    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function

End Class
