
Imports System.Windows.Markup

' WPF で IDE ライクなビューを作る (ドッキングウィンドウ・AvalonDock)
' https://qiita.com/lriki/items/475a7fc0e01ef62ef62a
' https://github.com/lriki/WPFSkeletonIDE/tree/GenericTheme1

<ContentProperty("Items")>
Public Class LayoutItemTemplateSelector
    Inherits DataTemplateSelector

    Public Property Items As List(Of DataTemplate)

    Public Sub New()

        Items = New List(Of DataTemplate)

    End Sub

    Public Overrides Function SelectTemplate(item As Object, container As DependencyObject) As DataTemplate

        Dim template = Items.Find(Function(x) item.GetType().Equals(x.DataType))
        If template IsNot Nothing Then
            Return template
        End If

        Return MyBase.SelectTemplate(item, container)

    End Function

End Class
