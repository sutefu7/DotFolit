Imports System.Collections.ObjectModel
Imports Livet

Public Class InheritsItemModel
    Inherits ViewModel

#Region "IsTargetClass変更通知プロパティ"
    Private _IsTargetClass As Boolean

    Public Property IsTargetClass() As Boolean
        Get
            Return _IsTargetClass
        End Get
        Set(ByVal value As Boolean)
            If (_IsTargetClass = value) Then Return
            _IsTargetClass = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "ClassName変更通知プロパティ"
    Private _ClassName As String

    Public Property ClassName() As String
        Get
            Return _ClassName
        End Get
        Set(ByVal value As String)
            If (_ClassName = value) Then Return
            _ClassName = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "TreeItems変更通知プロパティ"
    Private _TreeItems As ObservableCollection(Of TreeViewItemModel)

    Public Property TreeItems() As ObservableCollection(Of TreeViewItemModel)
        Get
            Return _TreeItems
        End Get
        Set(ByVal value As ObservableCollection(Of TreeViewItemModel))
            If (value Is Nothing) Then Return
            _TreeItems = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "メソッド"

    Public Sub New()

        Me._IsTargetClass = False
        Me._ClassName = String.Empty
        Me._TreeItems = Nothing

    End Sub

    Public Function NewInstance() As InheritsItemModel

        Dim model = New InheritsItemModel
        model.IsTargetClass = Me.IsTargetClass
        model.ClassName = Me.ClassName

        model.TreeItems = New ObservableCollection(Of TreeViewItemModel)
        For Each child In Me.TreeItems
            model.TreeItems.Add(child.NewInstance())
        Next

        Return model

    End Function

#End Region

End Class
