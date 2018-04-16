
Imports System.Windows.Markup

' WPF で IDE ライクなビューを作る (ドッキングウィンドウ・AvalonDock)
' https://qiita.com/lriki/items/475a7fc0e01ef62ef62a
' https://github.com/lriki/WPFSkeletonIDE/tree/GenericTheme1


<ContentProperty("Items")>
Public Class LayoutItemContainerStyleSelector
    Inherits StyleSelector

    Public Property Items As List(Of LayoutItemTypedStyle)

    Public Sub New()

        Items = New List(Of LayoutItemTypedStyle)

    End Sub

    Public Overrides Function SelectStyle(item As Object, container As DependencyObject) As Style

        Dim styleData = Items.Find(Function(x) item.GetType().IsSubclassOf(x.DataType))
        If styleData IsNot Nothing Then
            Return styleData.TargetStyle
        End If

        Return MyBase.SelectStyle(item, container)

    End Function

End Class

<ContentProperty("TargetStyle")>
Public Class LayoutItemTypedStyle

    Public Property DataType As Type
    Public Property TargetStyle As Style

End Class