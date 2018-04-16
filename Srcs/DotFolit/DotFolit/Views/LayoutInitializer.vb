
Imports System.Windows.Markup
Imports System.Reflection
Imports Xceed.Wpf.AvalonDock.Layout

' WPF で IDE ライクなビューを作る (ドッキングウィンドウ・AvalonDock)
' https://qiita.com/lriki/items/475a7fc0e01ef62ef62a
' https://github.com/lriki/WPFSkeletonIDE/tree/GenericTheme1


<ContentProperty("Items")>
Public Class LayoutInitializer
    Implements ILayoutUpdateStrategy

    Public Property Items As List(Of LayoutInsertTarget)

    Public Sub New()

        Items = New List(Of LayoutInsertTarget)

    End Sub

    Public Function BeforeInsertAnchorable(layout As LayoutRoot, anchorableToShow As LayoutAnchorable, destinationContainer As ILayoutContainer) As Boolean Implements ILayoutUpdateStrategy.BeforeInsertAnchorable

        Dim dummy = TryCast(destinationContainer, LayoutAnchorablePane)
        If (dummy IsNot Nothing) AndAlso (destinationContainer.FindParent(Of LayoutFloatingWindow)() IsNot Nothing) Then
            Return False
        End If

        Dim vm = TryCast(anchorableToShow.Content, LayoutInsertTarget)
        If vm Is Nothing Then
            Return False
        End If

        Dim target = Items.Find(Function(x) x.ContentId = vm.ContentId)
        If target Is Nothing Then
            Return False
        End If

        Dim pane = layout.Descendents().OfType(Of LayoutAnchorablePane)().FirstOrDefault(Function(x) x.Name = target.TargetLayoutName)
        If pane Is Nothing Then
            Return False
        End If

        pane.Children.Add(anchorableToShow)
        Return True

    End Function

    Public Sub AfterInsertAnchorable(layout As LayoutRoot, anchorableShown As LayoutAnchorable) Implements ILayoutUpdateStrategy.AfterInsertAnchorable

    End Sub

    Public Function BeforeInsertDocument(layout As LayoutRoot, anchorableToShow As LayoutDocument, destinationContainer As ILayoutContainer) As Boolean Implements ILayoutUpdateStrategy.BeforeInsertDocument

        Return False

    End Function

    Public Sub AfterInsertDocument(layout As LayoutRoot, anchorableShown As LayoutDocument) Implements ILayoutUpdateStrategy.AfterInsertDocument

    End Sub

End Class

Public Class LayoutInsertTarget

    Public Property ContentId As String
    Public Property TargetLayoutName As String

End Class
