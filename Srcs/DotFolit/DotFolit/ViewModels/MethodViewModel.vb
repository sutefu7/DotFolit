Imports Livet

Public Class MethodViewModel
    Inherits ViewModel

#Region "SelectedNode変更通知プロパティ"
    Private _SelectedNode As TreeViewItemModel

    Public Property SelectedNode() As TreeViewItemModel
        Get
            Return _SelectedNode
        End Get
        Set(ByVal value As TreeViewItemModel)
            If (value Is Nothing) Then Return
            _SelectedNode = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

    ' Dummy ViewModel

    ' MVVM 形式での実装が難しいため、実際にはコードビハインドで実装している
    ' Messenger 経由でダイアログ表示するため、対応する ViewModel が必要なため定義している



End Class
