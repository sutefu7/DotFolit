Imports System.Collections.ObjectModel
Imports Livet


Public Class TreeViewItemModel
    Inherits ViewModel

#Region "Text変更通知プロパティ"
    Private _Text As String

    Public Property Text() As String
        Get
            Return _Text
        End Get
        Set(ByVal value As String)
            If (_Text = value) Then Return
            _Text = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "IsSelected変更通知プロパティ"
    Private _IsSelected As Boolean

    Public Property IsSelected() As Boolean
        Get
            Return _IsSelected
        End Get
        Set(ByVal value As Boolean)
            If (_IsSelected = value) Then Return
            _IsSelected = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "IsExpanded変更通知プロパティ"
    Private _IsExpanded As Boolean

    Public Property IsExpanded() As Boolean
        Get
            Return _IsExpanded
        End Get
        Set(ByVal value As Boolean)
            If (_IsExpanded = value) Then Return
            _IsExpanded = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "Children変更通知プロパティ"
    Private _Children As ObservableCollection(Of TreeViewItemModel)

    Public Property Children() As ObservableCollection(Of TreeViewItemModel)
        Get
            Return _Children
        End Get
        Set(ByVal value As ObservableCollection(Of TreeViewItemModel))
            If (value Is Nothing) Then Return
            _Children = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "TreeNodeKind変更通知プロパティ"
    Private _TreeNodeKind As TreeNodeKinds

    Public Property TreeNodeKind() As TreeNodeKinds
        Get
            Return _TreeNodeKind
        End Get
        Set(ByVal value As TreeNodeKinds)
            If (_TreeNodeKind = value) Then Return
            _TreeNodeKind = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "FileName変更通知プロパティ"
    Private _FileName As String

    Public Property FileName() As String
        Get
            Return _FileName
        End Get
        Set(ByVal value As String)
            If (_FileName = value) Then Return
            _FileName = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "ContainerName変更通知プロパティ"
    Private _ContainerName As String

    Public Property ContainerName() As String
        Get
            Return _ContainerName
        End Get
        Set(ByVal value As String)
            If (_ContainerName = value) Then Return
            _ContainerName = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "StartLength変更通知プロパティ"
    Private _StartLength As Integer

    Public Property StartLength() As Integer
        Get
            Return _StartLength
        End Get
        Set(ByVal value As Integer)
            If (_StartLength = value) Then Return
            _StartLength = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "メソッド"

    Public Sub New()

        Me._Text = String.Empty
        Me._IsSelected = False
        Me._IsExpanded = False
        Me._Children = New ObservableCollection(Of TreeViewItemModel)
        Me._TreeNodeKind = TreeNodeKinds.None
        Me._FileName = String.Empty
        Me._ContainerName = String.Empty
        Me._StartLength = -1

    End Sub

    Public Function NewInstance() As TreeViewItemModel

        Dim model = New TreeViewItemModel
        model.Text = Me.Text
        model.IsSelected = Me.IsSelected
        model.IsExpanded = Me.IsExpanded

        model.Children = New ObservableCollection(Of TreeViewItemModel)
        For Each child In Me.Children
            model.Children.Add(child.NewInstance())
        Next

        model.TreeNodeKind = Me.TreeNodeKind
        model.FileName = Me.FileName
        model.ContainerName = Me.ContainerName
        model.StartLength = Me.StartLength

        Return model

    End Function

#End Region

End Class
