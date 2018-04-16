Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Interactivity
Imports Xceed.Wpf.AvalonDock

Public Class DocumentClosingBehavior
    Inherits Behavior(Of DockingManager)

#Region "DocumentClosingCommand 依存関係プロパティ"

    Public Shared ReadOnly DocumentClosingCommandProperty As DependencyProperty =
        DependencyProperty.Register(
        "DocumentClosingCommand",
        GetType(ICommand),
        GetType(DocumentClosingBehavior),
        New PropertyMetadata())

    Public Property DocumentClosingCommand As ICommand
        Get
            Return TryCast(GetValue(DocumentClosingCommandProperty), ICommand)
        End Get
        Set(value As ICommand)
            SetValue(DocumentClosingCommandProperty, value)
        End Set
    End Property

#End Region

    Protected Overrides Sub OnAttached()

        MyBase.OnAttached()
        AddHandler Me.AssociatedObject.DocumentClosing, AddressOf DockingManager_DocumentClosing

    End Sub

    Protected Overrides Sub OnDetaching()

        MyBase.OnDetaching()
        RemoveHandler Me.AssociatedObject.DocumentClosing, AddressOf DockingManager_DocumentClosing

    End Sub

    Private Sub DockingManager_DocumentClosing(sender As Object, e As DocumentClosingEventArgs)

        If Me.DocumentClosingCommand.CanExecute(e) Then
            Me.DocumentClosingCommand.Execute(e)
        End If

    End Sub

End Class
