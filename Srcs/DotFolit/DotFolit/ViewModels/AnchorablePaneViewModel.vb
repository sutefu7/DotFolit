Imports Livet

Public MustInherit Class AnchorablePaneViewModel
    Inherits ViewModel

    Public MustOverride ReadOnly Property Title As String
    Public MustOverride ReadOnly Property ContentId As String

#Region "CanHide変更通知プロパティ"
    Private _CanHide As Boolean

    Public Property CanHide() As Boolean
        Get
            Return _CanHide
        End Get
        Set(ByVal value As Boolean)
            If (_CanHide = value) Then Return
            _CanHide = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "CanClose変更通知プロパティ"
    Private _CanClose As Boolean

    Public Property CanClose() As Boolean
        Get
            Return _CanClose
        End Get
        Set(ByVal value As Boolean)
            If (_CanClose = value) Then Return
            _CanClose = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "CanFloat変更通知プロパティ"
    Private _CanFloat As Boolean

    Public Property CanFloat() As Boolean
        Get
            Return _CanFloat
        End Get
        Set(ByVal value As Boolean)
            If (_CanFloat = value) Then Return
            _CanFloat = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "IsActive変更通知プロパティ"
    Private _IsActive As Boolean

    Public Property IsActive() As Boolean
        Get
            Return _IsActive
        End Get
        Set(ByVal value As Boolean)
            If (_IsActive = value) Then Return
            _IsActive = value
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


End Class
